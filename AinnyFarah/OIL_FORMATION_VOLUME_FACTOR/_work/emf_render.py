"""Minimal EMF renderer (polygons/polylines/polypolygons/beziers as polylines) to PNG for OCR."""
import struct, sys
from pathlib import Path
from PIL import Image, ImageDraw
SCALE_W = 3000
for f in sys.argv[1:]:
    b = Path(f).read_bytes()
    recs, i = [], 0
    while i + 8 <= len(b):
        t, sz = struct.unpack_from("<II", b, i)
        if sz < 8: break
        recs.append((t, b[i+8:i+sz])); i += sz
    objs, brush, pen = {}, (255, 255, 255), (0, 0, 0)
    shapes, cur, path = [], (0, 0), []
    def pts16(d, off, n): return [struct.unpack_from("<hh", d, off + 4*k) for k in range(n)]
    for t, d in recs:
        if t == 39:  # CREATEBRUSHINDIRECT
            ih, style, col = struct.unpack_from("<III", d, 0)
            objs[ih] = ("brush", None if style == 1 else (col & 255, col >> 8 & 255, col >> 16 & 255))
        elif t in (38, 95):
            ih = struct.unpack_from("<I", d, 0)[0]
            col = struct.unpack_from("<I", d, 16 if t == 38 else 24)[0]
            objs[ih] = ("pen", (col & 255, col >> 8 & 255, col >> 16 & 255))
        elif t == 37:
            ih = struct.unpack_from("<I", d, 0)[0]
            if ih & 0x80000000:
                s = ih & 0xFF
                if s in (0, 1, 2, 3, 4): brush = [(255,)*3, (192,)*3, (128,)*3, (64,)*3, (0,)*3][s]
                if s == 5: brush = None
                if s == 7: pen = (0, 0, 0)
                if s == 8: pen = None
            elif ih in objs:
                k, v = objs[ih]
                if k == "brush": brush = v
                else: pen = v
        elif t in (86, 87, 85, 88):  # POLYGON16, POLYLINE16, POLYBEZIER16, POLYBEZIERTO16
            n = struct.unpack_from("<I", d, 16)[0]
            p = pts16(d, 20, n)
            if t == 88: p = [cur] + p
            shapes.append(("poly" if t == 86 else "line", p, brush, pen))
            if p: cur = p[-1]
        elif t == 91:  # POLYPOLYGON16
            npoly, tot = struct.unpack_from("<II", d, 16)
            counts = struct.unpack_from("<%dI" % npoly, d, 24)
            off = 24 + 4*npoly
            polys = []
            for c in counts:
                polys.append(pts16(d, off, c)); off += 4*c
            shapes.append(("pp", polys, brush, pen))
        elif t == 27: cur = struct.unpack_from("<ii", d, 0)
        elif t == 54:
            p = struct.unpack_from("<ii", d, 0); shapes.append(("line", [cur, p], brush, pen)); cur = p
    allp = [q for s in shapes for q in (s[1] if s[0] != "pp" else [x for pl in s[1] for x in pl])]
    if not allp: print(f, "no shapes"); continue
    x0 = min(p[0] for p in allp); x1 = max(p[0] for p in allp); y0 = min(p[1] for p in allp); y1 = max(p[1] for p in allp)
    sc = SCALE_W / max(1, x1 - x0)
    img = Image.new("RGB", (SCALE_W + 40, int((y1 - y0) * sc) + 40), "white"); dr = ImageDraw.Draw(img)
    T = lambda p: ((p[0] - x0) * sc + 20, (p[1] - y0) * sc + 20)
    for kind, p, br, pn in shapes:
        if kind == "pp":
            for pl in p:
                if len(pl) > 2: dr.polygon([T(q) for q in pl], fill=br, outline=pn)
        elif kind == "poly" and len(p) > 2:
            dr.polygon([T(q) for q in p], fill=br, outline=pn)
        elif len(p) > 1 and pn:
            dr.line([T(q) for q in p], fill=pn, width=2)
    out = Path("render") / (Path(f).stem + ".png"); img.save(out); print(out, img.size, len(shapes))
