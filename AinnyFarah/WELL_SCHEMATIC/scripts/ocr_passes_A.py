#!/usr/bin/env python3
"""Local multi-pass OCR of every page of the Dulang A well-schematic PDF (read-only source; nothing leaves the Mac)."""
import hashlib, json, os, subprocess, sys
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
import pymupdf as fitz
from PIL import Image, ImageFilter, ImageOps

PDF = "/Volumes/homes/public/DulangOSS/20261016/DULANG_WELLS/Well_Schematic/1. Well Schematic/Dulang A/Dulang A Well Schematic updated rev1.pdf"
OUT = Path(__file__).resolve().parent.parent / "output_A" / "ocr_passes"


def read(p):
    fd = os.open(p, os.O_RDONLY); b = []
    try:
        while (x := os.read(fd, 1 << 20)): b.append(x)
    finally:
        try: os.close(fd)
        except OSError: pass
    return b"".join(b)


def ocr(img, psm, base):
    subprocess.run(["tesseract", str(img), str(base), "--psm", str(psm), "-c", "preserve_interword_spaces=1", "txt", "tsv"],
                   capture_output=True, check=True)


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    b = read(PDF)
    doc = fitz.open(stream=b, filetype="pdf")
    jobs, layer = [], {}
    for i, pg in enumerate(doc, 1):
        n = f"p{i:02d}"
        g = OUT / f"{n}_600.png"
        pg.get_pixmap(dpi=600).save(g)
        im = ImageOps.autocontrast(Image.open(g).convert("L")).filter(ImageFilter.SHARPEN)
        im.save(g)
        bw = OUT / f"{n}_600bw.png"
        im.point(lambda v: 255 if v > 150 else 0).save(bw)
        jobs += [(g, 6, OUT / f"{n}_a"), (g, 4, OUT / f"{n}_b"), (bw, 6, OUT / f"{n}_c")]
        layer[n] = [dict(text="".join(s["text"] for s in l["spans"]).strip(), bbox=[round(v, 1) for v in l["bbox"]])
                    for bl in pg.get_text("dict")["blocks"] for l in bl.get("lines", []) if "".join(s["text"] for s in l["spans"]).strip()]
    with ThreadPoolExecutor(4) as ex:
        list(ex.map(lambda j: ocr(*j), jobs))
    (OUT / "text_layer.json").write_text(json.dumps(layer, indent=1, ensure_ascii=False))
    (OUT / "source_sha256.txt").write_text(hashlib.sha256(b).hexdigest() + "\n")
    print(doc.page_count, "pages,", len(jobs), "OCR passes ->", OUT)


if __name__ == "__main__":
    main()
