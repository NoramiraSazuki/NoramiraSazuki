#!/usr/bin/env python3
"""Row-align the native text of each table-style slide (dump/slideNN.json) into _work/rows/slideNN.txt + .json.

Columns are located from the header boxes (STATUS | MIN ID | SHORT STRING | DEPTH || DEPTH | LONG STRING | MIN ID | STATUS);
every text box below the header row is assigned to the column whose band contains its x-centre, then grouped into rows by
y-centre. The output is a review aid: the build script takes values only from object ids listed in its per-well specs.
"""
import json
import re
from pathlib import Path

W = Path(__file__).resolve().parent
OUT = W / "rows"
# single-string layout (slide 13: DEPTH | COMPLETION | MIN ID | STATUS, no short/long headers) -> one 'L' side
MANUAL = {13: [(49.0, "DEPTH_L"), (63.0, "NAME_L"), (77.0, "MINID_L"), (85.5, "STATUS_L")]}
HDR = [("STATUS", r"STATUS"), ("MINID", r"MIN\s*ID.*"), ("NAME_S", r"SHORT\s*STRING"), ("DEPTH", r"DEPTH.*")]


def cx(o):
    return (o["pct"][0] + o["pct"][2]) / 2


def cy(o):
    return (o["pct"][1] + o["pct"][3]) / 2


def main():
    OUT.mkdir(exist_ok=True)
    for f in sorted((W / "dump").glob("slide*.json")):
        d = json.loads(f.read_text())
        objs = [o for o in d["objects"] if o.get("text", "").strip() and o.get("pct")]
        heads = []
        for o in objs:
            t = re.sub(r"\s+", " ", o["text"]).strip()
            if 22 <= o["pct"][1] <= 30 and re.fullmatch(r"STATUS|MIN ID( \(in\))?|SHORT STRING|(COMPLETION )?LONG STRING|DEPTH( m-? ?MDTHF)?", t):
                heads.append((cx(o), t, o["pct"][3]))
        if d["slide"] in MANUAL:
            names = MANUAL[d["slide"]]
            heads = [(x, n, 28.5) for x, n in names]
        elif len(heads) < 6:
            continue
        heads.sort()
        names = [] if d["slide"] not in MANUAL else names
        for i, (x, t, _) in enumerate(heads if d["slide"] not in MANUAL else []):
            side = "S" if x < 50 else "L"
            key = {"STATUS": "STATUS", "SHORT STRING": "NAME", "LONG STRING": "NAME", "COMPLETION LONG STRING": "NAME"}.get(t)
            key = key or ("MINID" if t.startswith("MIN") else "DEPTH")
            names.append((x, f"{key}_{side}"))
        if d["slide"] not in MANUAL and not any(n == "NAME_L" for _, n in names):  # older layout: long-string name column has no header box
            dl = [x for x, n in names if n == "DEPTH_L"][0]
            ml = [x for x, n in names if n == "MINID_L"][0]
            names.append(((dl + ml) / 2, "NAME_L"))
            names.sort()
        bounds = [(names[i][0] + names[i + 1][0]) / 2 for i in range(len(names) - 1)]
        top = max(h[2] for h in heads)
        cells = []
        for o in objs:
            if cy(o) <= top or o["text"].lstrip().upper().startswith("LAST ENTRY"):
                continue
            x = cx(o)
            col = names[sum(1 for b in bounds if x > b)][1]
            if x < names[0][0] - (12 if d["slide"] in MANUAL else 5):
                col = "LEFT_MARGIN"
            elif x > names[-1][0] + 5:
                col = "RIGHT_MARGIN"
            cells.append(dict(id=o["id"], col=col, y=round(cy(o), 2), x=round(x, 2), h=round(o["pct"][3] - o["pct"][1], 2),
                              text=o["text"], name=o["name"], path=o["path"]))
        cells.sort(key=lambda c: c["y"])
        (OUT / f.name).write_text(json.dumps(dict(slide=d["slide"], columns=names, cells=cells), indent=1, ensure_ascii=False))
        with open(OUT / f.name.replace(".json", ".txt"), "w") as fh:
            fh.write(f"# slide {d['slide']} columns {[(round(x, 1), n) for x, n in names]}\n")
            for c in cells:
                fh.write(f"{c['y']:6.2f} {c['col']:<12} id={c['id']:<6} {c['text']!r}\n")
        print(d["slide"], len(cells))


if __name__ == "__main__":
    main()
