#!/usr/bin/env python3
"""Draft row pairing for the table-style slides (rows/slideNN.json) -> drafts/slideNN.json + .txt (review aid).

Per side (S = short string columns, L = long string columns): every equipment-name cell is paired with the nearest unused
DEPTH / MIN ID / STATUS cell of the same side within a vertical tolerance. Sand boxes ('SAND ...') are parsed as zones;
depth cells left unpaired are listed so the reviewer can see bracketing depths (e.g. blast-joint top/bottom).
"""
import json
import re
from pathlib import Path

W = Path(__file__).resolve().parent
NUM = re.compile(r"^\s*\d{1,4}(\.\d+)?\s*$")
TOL_D, TOL_ID, TOL_ST = 0.75, 0.75, 1.0


def main():
    (W / "drafts").mkdir(exist_ok=True)
    for f in sorted((W / "rows").glob("slide*.json")):
        d = json.loads(f.read_text())
        cells = d["cells"]
        used = set()
        rows, zones, notes = [], [], []
        for c in cells:
            t = c["text"].strip()
            if re.match(r"(?i)^\s*SAND\s*:", t) or re.search(r"(?i)PERF?T\s*:", t):
                zones.append(c)
                used.add(c["id"])
        for side in "SL":
            names = [c for c in cells if c["col"] == f"NAME_{side}" and c["id"] not in used]
            for n in names:
                row = {"side": side, "name": n}
                used.add(n["id"])
                for col, tol in ((f"DEPTH_{side}", TOL_D), (f"MINID_{side}", TOL_ID), (f"STATUS_{side}", TOL_ST)):
                    cand = [c for c in cells if c["col"] == col and c["id"] not in used and abs(c["y"] - n["y"]) <= tol]
                    if cand:
                        best = min(cand, key=lambda c: abs(c["y"] - n["y"]))
                        row[col.split("_")[0].lower()] = best
                        used.add(best["id"])
                rows.append(row)
        left = [c for c in cells if c["id"] not in used]
        out = dict(slide=d["slide"], rows=rows, zones=zones, unpaired=left)
        (W / "drafts" / f.name).write_text(json.dumps(out, indent=1, ensure_ascii=False))
        with open(W / "drafts" / f.name.replace(".json", ".txt"), "w") as fh:
            for side in "SL":
                fh.write(f"--- {side}\n")
                for r in sorted((r for r in rows if r["side"] == side), key=lambda r: r["name"]["y"]):
                    g = lambda k: (r[k]["text"].replace("\n", " / ") + f"#{r[k]['id']}") if k in r else "-"
                    fh.write(f"{r['name']['y']:6.2f} {r['name']['text'].replace(chr(10), ' / ')!r}#{r['name']['id']} | D {g('depth')} | ID {g('minid')} | ST {g('status')}\n")
            fh.write("--- zones\n")
            for z in zones:
                fh.write(f"{z['y']:6.2f} #{z['id']} {z['text']!r}\n")
            fh.write("--- unpaired\n")
            for c in left:
                fh.write(f"{c['y']:6.2f} {c['col']:<12} #{c['id']} {c['text']!r}\n")


if __name__ == "__main__":
    main()
