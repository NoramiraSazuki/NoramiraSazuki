"""Run OCRmyPDF locally on selected pages, writing only to a workspace temp folder."""

from __future__ import annotations

import logging
import os
import shutil
import subprocess
from pathlib import Path

log = logging.getLogger(__name__)


class OcrError(RuntimeError):
    pass


def ocr_available() -> bool:
    return shutil.which("ocrmypdf") is not None and shutil.which("tesseract") is not None


def _page_spec(pages: list[int]) -> str:
    """[1,2,3,7,9,10] -> '1-3,7,9-10'"""
    ranges, start, prev = [], pages[0], pages[0]
    for p in pages[1:] + [None]:
        if p is not None and p == prev + 1:
            prev = p
            continue
        ranges.append(f"{start}-{prev}" if start != prev else str(start))
        if p is not None:
            start = prev = p
    return ",".join(ranges)


def run_ocr(
    source: Path,
    pages: list[int],
    work_dir: Path,
    language: str = "eng+msa",
    jobs: int = 2,
    timeout: int = 3600,
) -> Path:
    """OCR only ``pages`` of ``source`` into a new PDF inside ``work_dir``.

    The source is only read. OCRmyPDF and Tesseract scratch files are kept inside
    ``work_dir`` by pointing TMPDIR there.
    """
    if not ocr_available():
        raise OcrError("ocrmypdf/tesseract not found on PATH")
    scratch = work_dir / "scratch"
    scratch.mkdir(parents=True, exist_ok=True)
    out_pdf = work_dir / "ocr_output.pdf"

    cmd = [
        "ocrmypdf",
        "--mode", "force",          # rasterise + OCR the selected pages (temp copy only)
        "--pages", _page_spec(sorted(pages)),
        "--language", language,
        "--output-type", "pdf",     # skip PDF/A conversion; this file is temporary
        "--optimize", "0",
        "--jobs", str(jobs),
        "--quiet",
        str(source),
        str(out_pdf),
    ]
    env = {**os.environ, "TMPDIR": str(scratch)}
    try:
        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout, env=env)
    except subprocess.TimeoutExpired as exc:
        raise OcrError(f"OCR timed out after {timeout}s") from exc
    if proc.returncode != 0:
        # Only the last stderr line is kept: it names the error, not document text.
        last = (proc.stderr.strip().splitlines() or ["(no message)"])[-1][:200]
        raise OcrError(f"ocrmypdf exit code {proc.returncode}: {last}")
    return out_pdf
