#!/usr/bin/env python3
"""Extra title-line OCR variants (1200 dpi, 4 thresholds) for pages whose well name lacks a 2nd agreeing reading."""
import csv, io, os, re, subprocess, sys
from pathlib import Path
import pymupdf as fitz
from PIL import Image, ImageFilter, ImageOps
PDF = "/Volumes/homes/public/DulangOSS/20261016/DULANG_WELLS/Well_Schematic/1. Well Schematic/Dulang A/Dulang A Well Schematic updated rev1.pdf"
P = Path(__file__).resolve().parent.parent / "output_A" / "ocr_passes"
pages = [int(x) for x in sys.argv[1:]]
fd = os.open(PDF, os.O_RDONLY); b = b""
while (x := os.read(fd, 1 << 20)): b += x
try: os.close(fd)
except OSError: pass
doc = fitz.open(stream=b, filetype="pdf")
for i in pages:
    pg, n = doc[i - 1], f"p{i:02d}"
    bands = []
    for k in ("a", "b", "c"):
        rows = {}
        for r in csv.DictReader(io.StringIO((P / f"{n}_{k}.tsv").read_text(errors="ignore")), delimiter="\t", quoting=csv.QUOTE_NONE):
            if (r.get("text") or "").strip():
                rows.setdefault((r["block_num"], r["par_num"], r["line_num"]), []).append(r)
        for ws in rows.values():
            t = " ".join(w["text"] for w in ws)
            if re.search(r"PRODUCER|INJECTOR|ELL\s*NAME|Well\s*No|WELL\s*:", t, re.I):
                top = min(int(w["top"]) for w in ws); bot = max(int(w["top"]) + int(w["height"]) for w in ws)
                if not any(abs(top - t0) < 40 for t0, _ in bands):
                    bands.append((top, bot))
    for v, thr in enumerate((110, 130, 170, 190), 1):
        out = []
        for top, bot in sorted(bands)[:6]:
            sc, h = 72 / 600, bot - top
            clip = fitz.Rect(0, max(0, (top - 0.6 * h) * sc), pg.rect.width, (bot + 0.6 * h) * sc)
            g = P / f"{n}_v_tmp.png"
            pg.get_pixmap(dpi=1200, clip=clip).save(g)
            ImageOps.autocontrast(Image.open(g).convert("L")).filter(ImageFilter.SHARPEN).point(lambda x, t=thr: 255 if x > t else 0).save(g)
            out.append(subprocess.run(["tesseract", str(g), "-", "--psm", "7"], capture_output=True, text=True).stdout.strip())
            g.unlink()
        (P / f"{n}_v{v}.txt").write_text("\n".join(out) + "\n")
    print(n, "variants:", [l for v in range(1, 5) for l in (P / f"{n}_v{v}.txt").read_text().splitlines() if re.search(r"ELL|Well", l)][:6])
