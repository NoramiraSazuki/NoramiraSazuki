#!/usr/bin/env python3
"""Dump every slide of the local PPTX copy (unzipped in _work/unzip) to _work/dump/slideN.json + slideN.txt.

Walks sp / grpSp / graphicFrame (tables) / pic / cxnSp recursively, applying group transforms so every object gets an
absolute position in EMU and in slide percent; records hidden flags, alt text (descr), text runs, table cells, images.
Also dumps notes slides. Read-only on the source.
"""
import json
import re
import sys
from pathlib import Path
from xml.etree import ElementTree as ET

W = Path(__file__).resolve().parent
UZ, OUT = W / "unzip", W / "dump"
NS = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main",
      "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
RID = "{%s}" % NS["r"]


def q(tag):
    p, t = tag.split(":")
    return "{%s}%s" % (NS[p], t)


def rels(path):
    rp = path.parent / "_rels" / (path.name + ".rels")
    if not rp.exists():
        return {}
    return {r.get("Id"): r.get("Target") for r in ET.parse(rp).getroot()}


def para_text(txBody):
    out = []
    for p in txBody.findall("a:p", NS):
        s = "".join((t.text or "") if t.tag == q("a:t") else "\n" if t.tag == q("a:br") else ""
                    for t in p.iter() if t.tag in (q("a:t"), q("a:br")))
        out.append(s)
    while out and not out[-1].strip():
        out.pop()
    return "\n".join(out)


class Xf:
    """affine map child EMU -> slide EMU"""
    def __init__(self, sx=1, sy=1, dx=0, dy=0):
        self.sx, self.sy, self.dx, self.dy = sx, sy, dx, dy

    def apply(self, x, y, cx, cy):
        return self.dx + x * self.sx, self.dy + y * self.sy, cx * self.sx, cy * self.sy

    def child(self, xfrm):
        off, ext = xfrm.find("a:off", NS), xfrm.find("a:ext", NS)
        coff, cext = xfrm.find("a:chOff", NS), xfrm.find("a:chExt", NS)
        x, y, cx, cy = (int(off.get("x")), int(off.get("y")), int(ext.get("cx")), int(ext.get("cy")))
        X, Y, CX, CY = self.apply(x, y, cx, cy)
        chx, chy = int(coff.get("x")), int(coff.get("y"))
        chcx, chcy = int(cext.get("cx")) or 1, int(cext.get("cy")) or 1
        sx, sy = CX / chcx, CY / chcy
        return Xf(sx, sy, X - chx * sx, Y - chy * sy)


def geom(el, xf):
    xfrm = None
    for path in ("p:spPr/a:xfrm", "p:xfrm", "p:grpSpPr/a:xfrm"):
        xfrm = el.find(path, NS)
        if xfrm is not None:
            break
    if xfrm is None or xfrm.find("a:off", NS) is None:
        return None
    off, ext = xfrm.find("a:off", NS), xfrm.find("a:ext", NS)
    X, Y, CX, CY = xf.apply(int(off.get("x")), int(off.get("y")), int(ext.get("cx")), int(ext.get("cy")))
    return dict(x=round(X), y=round(Y), cx=round(CX), cy=round(CY), rot=int(xfrm.get("rot", 0)) / 60000,
                flipV=xfrm.get("flipV") == "1", flipH=xfrm.get("flipH") == "1")


def walk(parent, xf, rel, sw, sh, path, out):
    for el in parent:
        tag = el.tag.split("}")[1]
        if tag not in ("sp", "grpSp", "graphicFrame", "pic", "cxnSp"):
            continue
        nv = el.find(".//p:cNvPr", NS)
        d = dict(kind=tag, id=nv.get("id"), name=nv.get("name"), path=path, hidden=nv.get("hidden") == "1",
                 descr=nv.get("descr") or nv.get("title"))
        g = geom(el, xf)
        if g:
            d.update(g)
            d["pct"] = [round(100 * g["x"] / sw, 2), round(100 * g["y"] / sh, 2),
                        round(100 * (g["x"] + g["cx"]) / sw, 2), round(100 * (g["y"] + g["cy"]) / sh, 2)]
        if tag == "sp":
            tb = el.find("p:txBody", NS)
            if tb is not None:
                d["text"] = para_text(tb)
            pg = el.find("p:spPr/a:prstGeom", NS)
            d["geom"] = pg.get("prst") if pg is not None else ("custom" if el.find("p:spPr/a:custGeom", NS) is not None else None)
            sf = el.find("p:spPr/a:solidFill/a:srgbClr", NS)
            if sf is not None:
                d["fill"] = sf.get("val")
            bf = el.find("p:spPr/a:blipFill/a:blip", NS)
            if bf is not None:
                d["image"] = rel.get(bf.get(RID + "embed"))
        elif tag == "pic":
            b = el.find(".//a:blip", NS)
            if b is not None:
                d["image"] = rel.get(b.get(RID + "embed"))
                ext = b.find(".//*[@" + RID + "embed]")
                if ext is not None and ext is not b:
                    d["image_alt"] = rel.get(ext.get(RID + "embed"))
            sc = el.find("p:blipFill/a:srcRect", NS)
            if sc is not None:
                d["crop"] = dict(sc.attrib)
        elif tag == "graphicFrame":
            tbl = el.find(".//a:tbl", NS)
            if tbl is not None:
                rows = []
                for tr in tbl.findall("a:tr", NS):
                    rows.append([para_text(tc.find("a:txBody", NS)) if tc.find("a:txBody", NS) is not None else ""
                                 for tc in tr.findall("a:tc", NS)])
                d["table"] = rows
            else:
                gd = el.find(".//a:graphicData", NS)
                d["graphic"] = gd.get("uri") if gd is not None else None
                for sub in el.iter():
                    for k, v in sub.attrib.items():
                        if k.startswith(RID):
                            d.setdefault("refs", []).append(rel.get(v))
        out.append(d)
        if tag == "grpSp":
            gx = el.find("p:grpSpPr/a:xfrm", NS)
            walk(el, xf.child(gx) if gx is not None and gx.find("a:chOff", NS) is not None else xf, rel, sw, sh,
                 path + [d["name"]], out)


def main():
    OUT.mkdir(exist_ok=True)
    pres = ET.parse(UZ / "ppt/presentation.xml").getroot()
    sz = pres.find("p:sldSz", NS)
    sw, sh = int(sz.get("cx")), int(sz.get("cy"))
    prel = rels(UZ / "ppt/presentation.xml")
    order = [prel[s.get(RID + "id")] for s in pres.find("p:sldIdLst", NS)]
    summary = []
    for n, target in enumerate(order, 1):
        sp = UZ / "ppt" / target
        root = ET.parse(sp).getroot()
        rel = {k: v for k, v in rels(sp).items()}
        objs = []
        walk(root.find("p:cSld/p:spTree", NS), Xf(), rel, sw, sh, [], objs)
        notes = ""
        for t in rel.values():
            if "notesSlide" in t:
                nr = ET.parse((sp.parent / t).resolve()).getroot()
                notes = "\n".join(para_text(tb) for tb in nr.iter(q("p:txBody"))).strip()
        info = dict(slide=n, file=target, hidden=root.get("show") == "0", size=[sw, sh], notes=notes, objects=objs)
        (OUT / f"slide{n:02d}.json").write_text(json.dumps(info, indent=1, ensure_ascii=False))
        with open(OUT / f"slide{n:02d}.txt", "w") as f:
            f.write(f"# slide {n} ({target}) hidden={info['hidden']} objects={len(objs)}\n# notes: {notes!r}\n")
            for o in objs:
                if o.get("text", "").strip() or o.get("table") or o.get("image") or o["kind"] == "graphicFrame":
                    f.write(f"[{o['kind']} id={o['id']} '{o['name']}' pct={o.get('pct')} hidden={o['hidden']}"
                            f"{' img=' + o['image'] if o.get('image') else ''}{' descr=' + repr(o['descr']) if o.get('descr') else ''}]\n")
                    if o.get("text", "").strip():
                        f.write("    " + o["text"].replace("\n", "\n    ") + "\n")
                    if o.get("table"):
                        for r in o["table"]:
                            f.write("    | " + " | ".join(c.replace("\n", " / ") for c in r) + "\n")
        summary.append((n, target, info["hidden"], len(objs), sum(1 for o in objs if o.get("text", "").strip()),
                        sum(1 for o in objs if o.get("image")), sum(1 for o in objs if o.get("table"))))
    for s in summary:
        print(*s)


if __name__ == "__main__":
    sys.exit(main())
