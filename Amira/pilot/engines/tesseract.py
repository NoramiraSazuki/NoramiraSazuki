"""Local OCR of one rendered page with the Tesseract command-line tool (no network).

The page is rendered with its /Rotate applied, at a capped resolution, into the batch temp
folder. If the first pass finds little text, Tesseract orientation detection is tried and the
better of the two passes is kept. Word boxes and Tesseract's confidences are recorded; the
confidence is a hint for review, not proof of correctness.
"""

from __future__ import annotations

import csv
import io
import shutil
import subprocess
import time
from functools import lru_cache

from ..pdfio import render_page
from .base import Block, Engine, EngineCancelled, EngineTimeout, PageContext, PageResult


@lru_cache(maxsize=1)
def tesseract_version() -> str:
    exe = shutil.which("tesseract")
    if not exe:
        return ""
    out = subprocess.run([exe, "--version"], capture_output=True, text=True, timeout=30)
    first = (out.stdout or out.stderr).strip().splitlines()
    return first[0].strip() if first else "tesseract-unknown"


def _run(cmd: list[str], timeout: float, should_stop) -> subprocess.CompletedProcess:
    """Run a subprocess, polling for cancel; kill it on timeout or cancel."""
    proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    deadline = time.monotonic() + timeout
    while True:
        try:
            out, err = proc.communicate(timeout=max(0.005, min(0.5, deadline - time.monotonic())))
            if time.monotonic() > deadline:  # finished, but only after the allowed time
                raise EngineTimeout(f"OCR exceeded {timeout:.0f}s")
            return subprocess.CompletedProcess(cmd, proc.returncode, out, err)
        except subprocess.TimeoutExpired:
            if should_stop():
                proc.kill()
                proc.communicate()
                raise EngineCancelled("cancelled")
            if time.monotonic() > deadline:
                proc.kill()
                proc.communicate()
                raise EngineTimeout(f"OCR exceeded {timeout:.0f}s")


def parse_tsv(tsv: str, scale: float) -> tuple[list[Block], list[float]]:
    """Group Tesseract word rows into lines; convert pixel boxes to points (top-left origin)."""
    lines: dict[tuple, dict] = {}
    confs = []
    for row in csv.DictReader(io.StringIO(tsv), delimiter="\t", quoting=csv.QUOTE_NONE):
        if row.get("level") != "5":
            continue
        word = (row.get("text") or "").strip()
        conf = float(row.get("conf") or -1)
        if not word or conf < 0:
            continue
        confs.append(conf)
        key = (int(row["page_num"]), int(row["block_num"]), int(row["par_num"]), int(row["line_num"]))
        x, y, w, h = (int(row[k]) for k in ("left", "top", "width", "height"))
        ln = lines.setdefault(key, {"words": [], "confs": [], "box": [x, y, x + w, y + h]})
        ln["words"].append(word)
        ln["confs"].append(conf)
        b = ln["box"]
        ln["box"] = [min(b[0], x), min(b[1], y), max(b[2], x + w), max(b[3], y + h)]
    blocks = []
    for key in sorted(lines):
        ln = lines[key]
        blocks.append(Block(" ".join(ln["words"]), [round(v / scale, 2) for v in ln["box"]], "line",
                            round(sum(ln["confs"]) / len(ln["confs"]), 1)))
    return blocks, confs


class TesseractEngine(Engine):
    name = "tesseract"
    label = "Local OCR (Tesseract)"

    def __init__(self, lang: str = "eng+msa", dpi: int = 300, max_megapixels: float = 40,
                 timeout: float = 180, psm: int = 3):
        self.lang, self.dpi, self.max_megapixels, self.timeout, self.psm = lang, dpi, max_megapixels, timeout, psm

    @property
    def version(self) -> str:
        return tesseract_version() or "unavailable"

    def settings(self) -> dict:
        return {"lang": self.lang, "dpi": self.dpi, "max_megapixels": self.max_megapixels, "psm": self.psm,
                "osd_fallback": True}

    def available(self) -> tuple[bool, str]:
        return (True, "") if tesseract_version() else (False, "tesseract not found on PATH")

    def _ocr(self, img_path, dpi, scale, ctx) -> tuple[list[Block], list[float]]:
        cmd = ["tesseract", str(img_path), "stdout", "-l", self.lang, "--psm", str(self.psm),
               "--dpi", str(dpi), "tsv"]
        proc = _run(cmd, self.timeout, ctx.should_stop)
        if proc.returncode != 0:
            # Only the exit code and last stderr line (an error name, not document text).
            last = (proc.stderr.strip().splitlines() or ["(no message)"])[-1][:120]
            raise RuntimeError(f"tesseract exit {proc.returncode}: {last}")
        return parse_tsv(proc.stdout, scale)

    def _osd_rotation(self, img_path, dpi, ctx) -> int:
        try:
            proc = _run(["tesseract", str(img_path), "stdout", "--psm", "0", "--dpi", str(dpi)],
                        min(self.timeout, 60), ctx.should_stop)
        except EngineTimeout:
            return 0
        rot, conf = 0, 0.0
        for line in proc.stdout.splitlines():
            if line.startswith("Rotate:"):
                rot = int(line.split(":")[1])
            elif line.startswith("Orientation confidence:"):
                conf = float(line.split(":")[1])
        return rot if conf >= 2.0 else 0

    def extract_page(self, ctx: PageContext) -> PageResult:
        img, scale = render_page(ctx.page, self.dpi, self.max_megapixels)
        eff_dpi = round(72 * scale)
        ctx.tmp_dir.mkdir(parents=True, exist_ok=True)
        img_path = ctx.tmp_dir / f"p{ctx.page_no:06d}.png"
        warnings, rotated = [], 0
        try:
            img.save(img_path)
            size = img.size
            blocks, confs = self._ocr(img_path, eff_dpi, scale, ctx)
            if sum(len(b.text) for b in blocks) < 40:
                rot = self._osd_rotation(img_path, eff_dpi, ctx)
                if rot:
                    alt_path = ctx.tmp_dir / f"p{ctx.page_no:06d}_rot{rot}.png"
                    img.rotate(-rot, expand=True).save(alt_path)
                    alt_blocks, alt_confs = self._ocr(alt_path, eff_dpi, scale, ctx)
                    if not ctx.keep_temp:
                        alt_path.unlink(missing_ok=True)
                    if sum(len(b.text) for b in alt_blocks) > sum(len(b.text) for b in blocks):
                        blocks, confs, rotated = alt_blocks, alt_confs, rot
                        warnings.append(f"page image rotated {rot}° for OCR; boxes are in the rotated image")
        finally:
            img.close()
            if not ctx.keep_temp:
                img_path.unlink(missing_ok=True)
        if eff_dpi < self.dpi:
            warnings.append(f"rendered at {eff_dpi} dpi (size cap) instead of {self.dpi}")
        text = "\n".join(b.text for b in blocks)
        return PageResult(
            text=text,
            blocks=blocks,
            bbox_space="rendered_view_points_origin_top_left" + (f"_rotated_{rotated}" if rotated else ""),
            reading_order="Tesseract block/paragraph/line order (not verified)",
            stats={"chars": sum(1 for c in text if not c.isspace()), "words": len(confs),
                   "mean_conf": round(sum(confs) / len(confs), 1) if confs else None,
                   "render_dpi": eff_dpi, "image_px": list(size), "osd_rotation": rotated},
            warnings=warnings,
        )
