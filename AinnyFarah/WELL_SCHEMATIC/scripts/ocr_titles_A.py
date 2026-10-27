#!/usr/bin/env python3
"""900-dpi OCR of each page's title / header lines (independent source 't' for the well-name check)."""
import csv, io, os, re, subprocess
from pathlib import Path
import pymupdf as fitz
from PIL import Image, ImageFilter, ImageOps
PDF = "/Volumes/homes/public/DulangOSS/20261016/DULANG_WELLS/Well_Schematic/1. Well Schematic/Dulang A/Dulang A Well Schematic updated rev1.pdf"
P = Path(__file__).resolve().parent.parent / "output_A" / "ocr_passes"
fd = os.open(PDF, os.O_RDONLY); b = b""
while (x := os.read(fd, 1 << 20)): b += x
try: os.close(fd)
except OSError: pass
doc = fitz.open(stream=b, filetype="pdf")
for i, pg in enumerate(doc, 1):
    n = f"p{i:02d}"
    lines = {}
    for k in ("a", "b", "c"):
        for r in csv.DictReader(io.StringIO((P / f"{n}_{k}.tsv").read_text(errors="ignore")), delimiter="\t", quoting=csv.QUOTE_NONE):
            if (r.get("text") or "").strip():
                lines.setdefault((k, r["block_num"], r["par_num"], r["line_num"]), []).append(r)
    bands = []
    for ws in lines.values():
        t = " ".join(w["text"] for w in ws)
        if re.search(r"PRODUCER|INJECTOR|DIAGRAM|WELL|Well|DLA|ELL", t):
            top = min(int(w["top"]) for w in ws); bot = max(int(w["top"]) + int(w["height"]) for w in ws)
            if not any(abs(top - t0) < 40 for t0, _ in bands):
                bands.append((top, bot))
    out = []
    for top, bot in sorted(bands)[:8]:
        sc = 72 / 600; h = bot - top
        clip = fitz.Rect(0, max(0, (top - 0.5 * h) * sc), pg.rect.width, (bot + 0.5 * h) * sc)
        g = P / f"{n}_title_tmp.png"
        pg.get_pixmap(dpi=900, clip=clip).save(g)
        im = ImageOps.autocontrast(Image.open(g).convert("L")).filter(ImageFilter.SHARPEN).point(lambda v: 255 if v > 140 else 0)
        im.save(g)
        out.append(subprocess.run(["tesseract", str(g), "-", "--psm", "7"], capture_output=True, text=True).stdout.strip())
        g.unlink()
    (P / f"{n}_t.txt").write_text("\n".join(out) + "\n")
print("title passes written")
