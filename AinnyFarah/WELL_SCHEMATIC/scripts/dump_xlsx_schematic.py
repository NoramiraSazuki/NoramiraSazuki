#!/usr/bin/env python3
"""Dump everything readable from a well-schematic .xlsx working copy (read-only): sheets + state, merged ranges, every
non-empty cell (value, formula, cached value), every drawing-layer object (shape / text box / picture / connector) with its
anchor cell range and text, comments, and embedded media. Usage: dump_xlsx_schematic.py <xlsx> <out.json>
"""
import html
import json
import re
import sys
import zipfile

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"


def drawing_objects(z, path):
    xml = z.read(path).decode("utf8")
    out = []
    for m in re.finditer(r"<xdr:(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b.*?</xdr:\1>", xml, re.S):
        a = m.group(0)
        fr = re.search(r"<xdr:from><xdr:col>(\d+)</xdr:col><xdr:colOff>(-?\d+)</xdr:colOff><xdr:row>(\d+)</xdr:row><xdr:rowOff>(-?\d+)</xdr:rowOff></xdr:from>", a)
        to = re.search(r"<xdr:to><xdr:col>(\d+)</xdr:col><xdr:colOff>(-?\d+)</xdr:colOff><xdr:row>(\d+)</xdr:row><xdr:rowOff>(-?\d+)</xdr:rowOff></xdr:to>", a)
        # one anchor may hold a group with many shapes; record each shape with its own name + text
        for s in re.finditer(r"<xdr:(sp|pic|cxnSp|graphicFrame)\b.*?</xdr:\1>", a, re.S):
            body = s.group(0)
            name = re.search(r'<xdr:cNvPr [^>]*name="([^"]*)"', body)
            hidden = 'hidden="1"' in (re.search(r"<xdr:cNvPr [^>]*>", body).group(0) if re.search(r"<xdr:cNvPr [^>]*>", body) else "")
            paras = []
            for p in re.findall(r"<a:p>.*?</a:p>|<a:p .*?</a:p>", body, re.S):
                t = "".join(html.unescape(x) for x in re.findall(r"<a:t>(.*?)</a:t>", p, re.S))
                if t.strip():
                    paras.append(t)
            emb = re.search(r'r:embed="([^"]+)"', body)
            off = re.search(r'<a:off x="(-?\d+)" y="(-?\d+)"/>', body)
            out.append({"kind": s.group(1), "name": name.group(1) if name else "", "hidden": hidden,
                        "from": f"{get_column_letter(int(fr.group(1)) + 1)}{int(fr.group(3)) + 1}" if fr else None,
                        "to": f"{get_column_letter(int(to.group(1)) + 1)}{int(to.group(3)) + 1}" if to else None,
                        "off_emu": [int(off.group(1)), int(off.group(2))] if off else None,
                        "text": paras, "embed": emb.group(1) if emb else None})
    return out


def main():
    src, dst = sys.argv[1], sys.argv[2]
    wf = load_workbook(src, data_only=False)
    wv = load_workbook(src, data_only=True)
    z = zipfile.ZipFile(src)
    res = {"file": src, "sheets": [], "media": [n for n in z.namelist() if n.startswith("xl/media/")]}
    for ws in wf.worksheets:
        wsv = wv[ws.title]
        cells = []
        for row in ws.iter_rows():
            for c in row:
                if c.value is None:
                    continue
                v = wsv[c.coordinate].value
                d = {"ref": c.coordinate, "value": v if not hasattr(v, "isoformat") else v.isoformat()}
                if isinstance(c.value, str) and c.value.startswith("="):
                    d["formula"] = c.value
                cells.append(d)
        comments = [{"ref": c.coordinate, "text": c.comment.text} for row in ws.iter_rows() for c in row if c.comment]
        res["sheets"].append({"title": ws.title, "state": ws.sheet_state, "dims": ws.dimensions,
                              "merged": [str(r) for r in ws.merged_cells.ranges], "cells": cells, "comments": comments})
    for n in z.namelist():
        if re.match(r"xl/drawings/drawing\d+\.xml$", n):
            res.setdefault("drawings", {})[n] = drawing_objects(z, n)
        if re.match(r"xl/drawings/_rels/drawing\d+\.xml\.rels$", n):
            res.setdefault("drawing_rels", {})[n] = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', z.read(n).decode()))
    json.dump(res, open(dst, "w"), indent=1, ensure_ascii=False, default=str)
    for s in res["sheets"]:
        print(f"sheet {s['title']!r} state={s['state']} dims={s['dims']} cells={len(s['cells'])} merged={len(s['merged'])} "
              f"formulas={sum(1 for c in s['cells'] if 'formula' in c)} comments={len(s['comments'])}")
    for n, objs in res.get("drawings", {}).items():
        kinds = {}
        for o in objs:
            kinds[o["kind"]] = kinds.get(o["kind"], 0) + 1
        print(n, kinds, "with text:", sum(1 for o in objs if o["text"]), "hidden:", sum(1 for o in objs if o["hidden"]))
    print("media:", res["media"])


if __name__ == "__main__":
    main()
