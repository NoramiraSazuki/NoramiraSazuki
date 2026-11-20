#!/usr/bin/env python3
"""Write output_A/Dulang_A17ST2_v1.json = output_A/Dulang_A17ST2.json + the user's header-size ID rule only (28 Oct 2026).

User-defined mapping rule (all components): a size in a component header with no ID/OD designation is treated as ID.
Header table: sheet 'OP5 single A17 DIAL eSTMZ FID', column V8 'SIZE' (no ID/OD designation).
Liner (V13) and tubing (V14/V15) are already ID in the initial JSON -> unchanged here.
Equipment-table sizes come from explicit 'MAX OD' (B19/AB19) and 'MIN ID' (D19/Z19) columns -> retained.
The initial JSON and the Excel source are only read; the output is never overwritten.
"""
import copy
import json
from pathlib import Path

import openpyxl

OUT_A = Path(__file__).resolve().parent.parent / "output_A"
SRC, DST = OUT_A / "Dulang_A17ST2.json", OUT_A / "Dulang_A17ST2_v1.json"
XLSX = Path("/Volumes/homes/public/DulangOSS/20261016/DULANG_WELLS/Well_Schematic/1. Well Schematic/Dulang A/"
            "Dulang A17ST2 Well Schematic updated .xlsx")
SHEET = "OP5 single A17 DIAL eSTMZ FID"
NOTE = "REVISED 2026-10-28 (user correction)"

# component id -> (header cell, expected text, size in inches)
CHANGES = {
    "conductor": ("V9", '26"', 26.0),
    "int-casing": ("V10", '13-3/8"', 13.375),
    "casing": ("V11", '9-5/8"', 9.625),
}


def main():
    if DST.exists():
        raise SystemExit(f"{DST} exists — not overwriting without confirmation")
    ws = openpyxl.load_workbook(XLSX, read_only=True, data_only=True)[SHEET]
    assert ws["V8"].value == "SIZE"
    assert (ws["B19"].value, ws["D19"].value, ws["Z19"].value, ws["AB19"].value) == ("MAX OD", "MIN ID", "MIN ID", "MAX OD")

    raw = SRC.read_text(encoding="utf-8")
    orig = json.loads(raw)
    doc = copy.deepcopy(orig)
    c = {x["id"]: x for x in doc["components"]}

    for cid, (cell, text, size) in CHANGES.items():
        assert ws[cell].value == text, (cell, ws[cell].value)
        x = c[cid]
        assert x["outerDiameter"] == size and x["innerDiameter"] is None, cid
        x["innerDiameter"], x["outerDiameter"] = x["outerDiameter"], None
        x["source"] += (f"; {NOTE}: header size {cell} '{text}' (column V8 'SIZE') has no ID/OD designation -> "
                        "recorded as innerDiameter (user-defined mapping rule); outerDiameter missing")

    # liner / tubing already moved to ID in the initial JSON
    for cid in ("liner-7", "tubing-upper", "tubing-lower"):
        assert c[cid]["outerDiameter"] is None and c[cid]["innerDiameter"] is not None, cid

    DST.write_text(json.dumps(doc, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")

    # ---------------- validation: only the intended fields differ
    back = json.loads(DST.read_text(encoding="utf-8"))
    assert SRC.read_text(encoding="utf-8") == raw
    assert list(back) == list(orig)
    for k in orig:
        if k != "components":
            assert back[k] == orig[k], k
    assert [x["id"] for x in back["components"]] == [x["id"] for x in orig["components"]]
    diffs = []
    for a, b in zip(orig["components"], back["components"]):
        assert list(a) == list(b), a["id"]
        diffs += [(a["id"], k, a[k], b[k]) for k in a if a[k] != b[k]]
    expected = {(i, k) for i in CHANGES for k in ("outerDiameter", "innerDiameter", "source")}
    assert {(i, k) for i, k, _, _ in diffs} == expected, {(i, k) for i, k, _, _ in diffs} ^ expected
    assert all(isinstance(v, (int, float)) or v is None for x in back["components"]
               for k, v in x.items() if k in ("topMd", "bottomMd", "outerDiameter", "innerDiameter"))
    print("valid JSON; components", len(back["components"]), "(unchanged count); differences limited to the requested fields:")
    for i, k, a, b in diffs:
        if k != "source":
            print(f"  {i}.{k}: {a!r} -> {b!r}")


if __name__ == "__main__":
    main()
