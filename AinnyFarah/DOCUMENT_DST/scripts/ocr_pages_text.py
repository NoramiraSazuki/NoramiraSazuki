#!/usr/bin/env python3
"""OCR selected pages of the DULANG-7 searchable PDF at high DPI with local Tesseract.

Usage: ocr_pages_text.py OUT_DIR PAGE [PAGE ...]   (ranges like 46-58 allowed)
Writes OUT_DIR/pNNN.txt per page.
"""
import os
import subprocess
import sys
from concurrent.futures import ProcessPoolExecutor
from pathlib import Path

OUT_ROOT = Path("/Users/farahyusof/Documents/GitHub/NoramiraSazuki/AinnyFarah/DOCUMENT_DST")
PDF = OUT_ROOT / "searchable_pdf" / "DULANG-7_EPMI Well Test Report DST 1,1A,2,2A_ 01 Dec. 1990_OCR.pdf"
TMP = OUT_ROOT / "tmp"


def ocr(args):
    page, out_dir = args
    env = dict(os.environ, TMPDIR=str(TMP), OMP_THREAD_LIMIT="1")
    png = TMP / f"ocr_p{page:03d}.png"
    subprocess.run(["gs", "-q", "-dNOPAUSE", "-dBATCH", "-dSAFER", "-sDEVICE=pnggray", "-r400",
                    f"-dFirstPage={page}", f"-dLastPage={page}", f"-sOutputFile={png}", str(PDF)],
                   check=True, env=env)
    txt = subprocess.run(["tesseract", str(png), "stdout", "-l", "eng", "--psm", "6"],
                         capture_output=True, text=True, env=env).stdout
    png.unlink(missing_ok=True)
    (out_dir / f"p{page:03d}.txt").write_text(txt)
    return page


def main():
    out_dir = Path(sys.argv[1])
    out_dir.mkdir(parents=True, exist_ok=True)
    pages = []
    for a in sys.argv[2:]:
        if "-" in a:
            s, e = map(int, a.split("-"))
            pages += range(s, e + 1)
        else:
            pages.append(int(a))
    with ProcessPoolExecutor(max_workers=os.cpu_count() or 2) as ex:
        for p in ex.map(ocr, [(p, out_dir) for p in pages]):
            print("done", p)


if __name__ == "__main__":
    main()
