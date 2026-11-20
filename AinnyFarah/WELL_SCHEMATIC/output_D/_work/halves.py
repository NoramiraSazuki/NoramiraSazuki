#!/usr/bin/env python3
"""Split an upscaled crop into overlapping top/bottom views for visual transcription: halves.py media_png/<file> [n_parts]"""
import sys
from PIL import Image
p = sys.argv[1]; n = int(sys.argv[2]) if len(sys.argv) > 2 else 2
im = Image.open(p); w, h = im.size
step = h // n
for i in range(n):
    y0 = max(0, i * step - 60); y1 = min(h, (i + 1) * step + 60)
    im.crop((0, y0, w, y1)).save(p.replace(".png", f"_part{i + 1}.png"))
    print(p.replace(".png", f"_part{i + 1}.png"), y0, y1)
