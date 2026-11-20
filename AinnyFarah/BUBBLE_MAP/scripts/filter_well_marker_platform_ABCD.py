#!/usr/bin/env python3
"""Write output_bubble-map/Well_Marker_15Jan2026_Platform_ABCD.xlsx = input_Well_Marker_15Jan2026.xlsx keeping only
rows whose well is on platform A, B, C or D (29 Oct 2026).

Platform rule: Sheet2 column E 'WELL NAME' matches DULG-<A|B|C|D><digit> (e.g. DULG-A01, DULG-B25S4, DULG-D19P).
Exploration/appraisal wells (DULG-1 ... DULG-87, DULG-6G-*, DULG-8S1) are dropped.
Sheet1 and Sheet2 are row-aligned (same row = same marker), so the same rows are kept in both sheets.
Cell values and styles, column widths, freeze panes and autofilter are kept; Sheet2 formulas in C/D are renumbered.
The source workbook is only read; the output is never overwritten.
"""
import re
from pathlib import Path

import openpyxl

BASE = Path(__file__).resolve().parent.parent
SRC = BASE / "input_Well_Marker_15Jan2026.xlsx"
DST = BASE / "output_bubble-map" / "Well_Marker_15Jan2026_Platform_ABCD.xlsx"
PLATFORM = re.compile(r"^DULG-[A-D]\d")


def compact(ws, keep):
    """Move kept data rows up under the header, then drop the tail."""
    ncol = ws.max_column
    for new, old in enumerate(keep, start=2):
        if new != old:
            for c in range(1, ncol + 1):
                s, d = ws.cell(old, c), ws.cell(new, c)
                d.value, d._style = s.value, s._style
        if ws.title == "Sheet2":
            ws.cell(new, 3).value = f"=CONCATENATE(A{new},B{new})"
            ws.cell(new, 4).value = f"=UPPER(C{new})"
    last = len(keep) + 1
    ws.delete_rows(last + 1, ws.max_row - last)
    if ws.auto_filter.ref:
        ws.auto_filter.ref = f"A1:{openpyxl.utils.get_column_letter(ncol)}{last}"


def main():
    if DST.exists():
        raise SystemExit(f"{DST} exists — not overwriting without confirmation")
    src_bytes = SRC.read_bytes()
    wb = openpyxl.load_workbook(SRC)
    s1, s2 = wb["Sheet1"], wb["Sheet2"]
    assert s1.max_row == s2.max_row and s1["L1"].value == "Well" and s2["B1"].value == "WELL" and s2["E1"].value == "WELL NAME"
    for r in range(2, s1.max_row + 1):
        assert s1.cell(r, 12).value == s2.cell(r, 2).value, r

    names = {r: s2.cell(r, 5).value for r in range(2, s2.max_row + 1)}
    keep = [r for r, n in names.items() if PLATFORM.match(str(n or ""))]
    orig1 = {r: [c.value for c in s1[r]] for r in keep}
    orig2 = {r: [c.value for c in s2[r]][:2] + [s2.cell(r, 5).value] for r in keep}

    compact(s1, keep)
    compact(s2, keep)
    DST.parent.mkdir(exist_ok=True)
    wb.save(DST)

    # ---------------- validation
    assert SRC.read_bytes() == src_bytes
    out = openpyxl.load_workbook(DST)
    o1, o2 = out["Sheet1"], out["Sheet2"]
    assert o1.max_row == o2.max_row == len(keep) + 1
    for new, old in enumerate(keep, start=2):
        assert [c.value for c in o1[new]] == orig1[old], old
        assert [o2.cell(new, 1).value, o2.cell(new, 2).value, o2.cell(new, 5).value] == orig2[old], old
    kept_names = {o2.cell(r, 5).value for r in range(2, o2.max_row + 1)}
    assert all(PLATFORM.match(n) for n in kept_names)
    dropped = sorted({n for r, n in names.items() if r not in set(keep)}, key=str)
    assert not any(PLATFORM.match(str(n)) for n in dropped)

    per = {p: sum(1 for r in keep if names[r][5] == p) for p in "ABCD"}
    print(f"source rows {len(names)} -> kept {len(keep)} ({per}); wells kept {len(kept_names)}, dropped {len(dropped)}")
    print("dropped wells:", ", ".join(map(str, dropped)))
    print("written", DST)


if __name__ == "__main__":
    main()
