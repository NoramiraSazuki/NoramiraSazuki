#!/usr/bin/env python3
"""Render a region of one page of the DULANG-7 searchable PDF at high DPI for visual reading.

Usage: crop_page.py PAGE DPI X0 Y0 X1 Y1 [OUT] [--rot DEG]
  X0..Y1 are fractions (0-1) of page width/height. --rot rotates the crop (e.g. 90).
Writes to DOCUMENT_DST/tmp/crops/ by default.
"""
import os
import subprocess
import sys
from pathlib import Path

from PIL import Image

OUT_ROOT = Path("/Users/farahyusof/Documents/GitHub/NoramiraSazuki/AinnyFarah/DOCUMENT_DST")
PDF = OUT_ROOT / "searchable_pdf" / "DULANG-7_EPMI Well Test Report DST 1,1A,2,2A_ 01 Dec. 1990_OCR.pdf"
TMP = OUT_ROOT / "tmp"
CROPS = TMP / "crops"


def main():
    args = sys.argv[1:]
    rot = 0
    if "--rot" in args:
        i = args.index("--rot")
        rot = int(args[i + 1])
        del args[i:i + 2]
    page, dpi = int(args[0]), int(args[1])
    x0, y0, x1, y1 = map(float, args[2:6])
    out = Path(args[6]) if len(args) > 6 else CROPS / f"p{page:03d}_{dpi}_{x0:.2f}_{y0:.2f}_{x1:.2f}_{y1:.2f}_r{rot}.png"
    CROPS.mkdir(parents=True, exist_ok=True)
    full = TMP / f"full_p{page:03d}_{dpi}.png"
    if not full.exists():
        subprocess.run(["gs", "-q", "-dNOPAUSE", "-dBATCH", "-dSAFER", "-sDEVICE=pnggray", f"-r{dpi}",
                        f"-dFirstPage={page}", f"-dLastPage={page}", f"-sOutputFile={full}", str(PDF)],
                       check=True, env=dict(os.environ, TMPDIR=str(TMP)))
    im = Image.open(full)
    w, h = im.size
    im = im.crop((int(x0 * w), int(y0 * h), int(x1 * w), int(y1 * h)))
    if rot:
        im = im.rotate(rot, expand=True)
    im.save(out)
    print(out, im.size)


if __name__ == "__main__":
    main()
