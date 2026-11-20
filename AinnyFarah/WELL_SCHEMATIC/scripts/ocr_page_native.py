#!/usr/bin/env python3
"""Local OCR of one PDF page's embedded schematic image (same 6 variants as ocr_A1). Usage: ocr_page_native.py <page_no> <out_dir>
Extracts the largest embedded image (native pixels), upscales x4 (Lanczos), autocontrast + sharpen, thresholds 140/175,
and runs Tesseract psm 6/4/11 (txt + tsv). Also dumps the page's exact PDF text layer (typed boxes) to text_layer.json.
Reads the source PDF only; everything is written inside <out_dir>.
"""
import io
import json
import os
import subprocess
import sys
from pathlib import Path

import pymupdf as fitz
from PIL import Image, ImageFilter, ImageOps

PDF = "/Volumes/homes/public/DulangOSS/20261016/DULANG_WELLS/Well_Schematic/1. Well Schematic/Dulang A/Dulang A Well Schematic updated rev1.pdf"
TESS = "/usr/local/bin/tesseract"


def read_pdf():
    fd = os.open(PDF, os.O_RDONLY); b = []
    while (x := os.read(fd, 1 << 20)):
        b.append(x)
    try:
        os.close(fd)
    except OSError:
        pass
    return b"".join(b)


def main():
    pno, out = int(sys.argv[1]), Path(sys.argv[2])
    out.mkdir(parents=True, exist_ok=True)
    doc = fitz.open(stream=read_pdf(), filetype="pdf")
    page = doc[pno - 1]
    # the schematic may be split into several stacked images: paste every large image (> 100k px) at its page position
    big = [i for i in page.get_images(full=True) if i[2] * i[3] > 100_000]
    parts = [(Image.open(io.BytesIO(doc.extract_image(i[0])["image"])).convert("L"), page.get_image_rects(i[0])[0]) for i in big]
    k = max(im.width / r.width for im, r in parts)            # native pixels per point
    x0, y0 = min(r.x0 for _, r in parts), min(r.y0 for _, r in parts)
    native = Image.new("L", (round((max(r.x1 for _, r in parts) - x0) * k), round((max(r.y1 for _, r in parts) - y0) * k)), 255)
    for im, r in parts:
        native.paste(im.resize((round(r.width * k), round(r.height * k)), Image.LANCZOS), (round((r.x0 - x0) * k), round((r.y0 - y0) * k)))
    imgs = big
    native.save(out / "img_native.png")
    x4 = ImageOps.autocontrast(native.resize((native.width * 4, native.height * 4), Image.LANCZOS)).filter(ImageFilter.SHARPEN)
    x4.save(out / "img_x4.png")
    for t in (140, 175):
        x4.point(lambda v, t=t: 255 if v > t else 0).save(out / f"img_x4_t{t}.png")
    for img, psm in (("img_x4", 6), ("img_x4", 4), ("img_x4", 11), ("img_x4_t140", 6), ("img_x4_t140", 4), ("img_x4_t175", 6)):
        base = out / f"{img}_psm{psm}"
        for ext in ("txt", "tsv"):
            subprocess.run([TESS, str(out / f"{img}.png"), str(base), "--psm", str(psm), "-l", "eng"] + ([] if ext == "txt" else ["tsv"]),
                           check=True, capture_output=True)
    layer = [{"text": " ".join(s["text"] for s in l["spans"]).strip(), "bbox": l["bbox"]}
             for b in page.get_text("dict")["blocks"] if b["type"] == 0 for l in b["lines"]]
    (out / "text_layer.json").write_text(json.dumps([l for l in layer if l["text"]], indent=1, ensure_ascii=False))
    print(f"page {pno}: native image {native.size}, {len(imgs)} images, {len(layer)} text-layer lines")


if __name__ == "__main__":
    main()
