#!/usr/bin/env python3
"""Fill Sheet1 column K 'UPDATED WELL NAME' in output_bubble-map/Well_Marker_15Jan2026_Platform_ABCD.xlsx, in place (29 Oct 2026).

Source: Sheet2 column E 'WELL NAME' (row-aligned with Sheet1).
Rule: DULG-<platform><3-digit well no.>[<suffix letter><2-digit suffix no.>]; ST -> S; DULG- never doubled.
  A1 -> DULG-A001, A1W1 -> DULG-A001W01, A33ST3 -> DULG-A033S03
Names that do not fit the rule exactly are left blank and reported for review.
Only column K of Sheet1 changes; everything else is checked to be identical after the save.
The workbook is regenerable from input_Well_Marker_15Jan2026.xlsx with filter_well_marker_platform_ABCD.py.
"""
import re
from collections import Counter
from pathlib import Path

import openpyxl

DST = Path(__file__).resolve().parent.parent / "output_bubble-map" / "Well_Marker_15Jan2026_Platform_ABCD.xlsx"
NAME = re.compile(r"^(?:DULG-)?([A-D])(\d{1,3})(?:(ST|S|W|H)(\d{1,2}))?$", re.I)


def updated(name):
    m = NAME.match(str(name or "").strip())
    if not m:
        return None
    plat, num, suf, sufnum = m.groups()
    out = f"DULG-{plat.upper()}{int(num):03d}"
    if suf:
        out += ("S" if suf.upper() == "ST" else suf.upper()) + f"{int(sufnum):02d}"
    return out


def snapshot(wb):
    """Values (except Sheet1!K) + resolved styles; style indices are renumbered on save, so compare by content."""
    def style(c):
        return tuple(repr(x) for x in (c.font, c.fill, c.border, c.alignment, c.protection)) + (c.number_format,)
    sheets = {ws.title: [[(None if ws.title == "Sheet1" and c.column == 11 else c.value, style(c)) for c in row]
                         for row in ws.iter_rows()] for ws in wb}
    sheets["_layout"] = [(ws.title, ws.freeze_panes, ws.auto_filter.ref, ws.dimensions,
                          {k: (d.width, d.hidden) for k, d in ws.column_dimensions.items()}) for ws in wb]
    return sheets


def main():
    for raw, want in {"A1": "DULG-A001", "A1W1": "DULG-A001W01", "A33ST3": "DULG-A033S03", "C10ST2": "DULG-C010S02",
                      "DULG-B25S4": "DULG-B025S04", "DULG-D19P": None}.items():
        assert updated(raw) == want, (raw, updated(raw))

    wb = openpyxl.load_workbook(DST)
    s1, s2 = wb["Sheet1"], wb["Sheet2"]
    assert s1["K1"].value == "UPDATED WELL NAME" and s2["E1"].value == "WELL NAME" and s1.max_row == s2.max_row
    before = snapshot(wb)
    review = Counter()
    for r in range(2, s1.max_row + 1):
        assert s1.cell(r, 12).value == s2.cell(r, 2).value, r
        src = s2.cell(r, 5).value
        new = updated(src)
        s1.cell(r, 11).value = new
        if new is None:
            review[(src, s2.cell(r, 2).value)] += 1
    wb.save(DST)

    # ---------------- validation: only Sheet1!K values changed
    back = openpyxl.load_workbook(DST)
    assert snapshot(back) == before
    b1, b2 = back["Sheet1"], back["Sheet2"]
    filled = Counter(b1.cell(r, 11).value for r in range(2, b1.max_row + 1))
    for r in range(2, b1.max_row + 1):
        assert b1.cell(r, 11).value == updated(b2.cell(r, 5).value)
    print(f"rows {b1.max_row - 1}; filled {b1.max_row - 1 - filled[None]} ({len(filled) - (None in filled)} distinct names); blank {filled[None]}")
    for (src, well), n in review.items():
        print(f"  REVIEW: WELL NAME {src!r} (Well {well!r}) x{n} -> left blank")


if __name__ == "__main__":
    main()
