#!/usr/bin/env python3
"""Crop every slide picture to its PowerPoint srcRect (the part shown on the slide) -> media_png/sNN_<image>.png (x3 LANCZOS)."""
import json, glob, os, io, struct
from PIL import Image
os.makedirs("media_png", exist_ok=True)
def load(p):
    if p.endswith(".emf"):
        b = open(p, "rb").read(); off = 0
        while off + 8 <= len(b):
            t, s = struct.unpack_from("<II", b, off)
            if t == 81:
                r = b[off:off + s]
                ob, cb, obits, cbits = struct.unpack_from("<IIII", r, 48)
                bmi, bits = r[ob:ob + cb], r[obits:obits + cbits]
                hdr = b"BM" + struct.pack("<IHHI", 14 + len(bmi) + len(bits), 0, 0, 14 + len(bmi))
                return Image.open(io.BytesIO(hdr + bmi + bits))
            if s < 8: break
            off += s
        return None
    return Image.open(p)
for f in sorted(glob.glob("dump/slide*.json")):
    d = json.load(open(f))
    for o in d["objects"]:
        if not o.get("image"): continue
        p = "unzip/ppt/" + o["image"].replace("../", "")
        im = load(p)
        if im is None or min(im.size) < 150: continue
        w, h = im.size; c = o.get("crop") or {}
        l, t, r, b = (int(c.get(k, 0)) / 100000 for k in "ltrb")
        box = (max(0, int(w * l)), max(0, int(h * t)), min(w, int(w * (1 - r))), min(h, int(h * (1 - b))))
        im = im.convert("RGB").crop(box)
        im = im.resize((im.size[0] * 3, im.size[1] * 3), Image.LANCZOS)
        out = f"media_png/s{d['slide']:02d}_{os.path.basename(p).split('.')[0]}.png"
        im.save(out); print(out, box, im.size)
