#!/usr/bin/env python3
"""Split Dulang_Unit_Water_Injection.xlsx into one worksheet per FAULT BLOCK (same file, backup kept)."""
import datetime as dt
import shutil
from copy import copy
from pathlib import Path

from openpyxl import load_workbook

OUT = Path(__file__).resolve().parent.parent / "output_water-injection"
import os
STEM = os.environ.get("DULANG_AREA", "Dulang Unit").replace(" ", "_") + "_Water_Injection"
F = OUT / f"{STEM}.xlsx"
REV = OUT / f"{STEM}_Review.xlsx"
ORDER = ["VIII", "IX", "N1", "N2", "N3AB", "N3B", "N3BDEFSE1", "N3C", "N3D", "S1", "S1A", "S1B", "S1C", "S1CD", "S1D", "S2", "S3", "SE1", "SE2"]


def main():
    bdir = OUT / "backups"
    bdir.mkdir(exist_ok=True)
    backup = bdir / f"{F.stem}_before_split_{dt.datetime.now():%Y%m%d_%H%M%S}.xlsx"
    shutil.copy2(F, backup)
    wb = load_workbook(F)
    src = wb["Sheet1"]
    hdr = [c for c in src[1]]
    widths = {k: v.width for k, v in src.column_dimensions.items()}
    rows = [list(r) for r in src.iter_rows(min_row=2, values_only=True)]
    groups = {}
    for r in rows:
        groups.setdefault(r[4] or "No Fault Block", []).append(r)
    names = [n for n in ORDER if n in groups] + sorted(set(groups) - set(ORDER) - {"No Fault Block"}) + (["No Fault Block"] if "No Fault Block" in groups else [])
    for n in names:
        ws = wb.create_sheet(n)
        for c in hdr:
            nc = ws.cell(1, c.column, c.value)
            nc.font, nc.fill, nc.alignment, nc.border, nc.number_format = copy(c.font), copy(c.fill), copy(c.alignment), copy(c.border), c.number_format
        ws.row_dimensions[1].height = src.row_dimensions[1].height
        for r in groups[n]:
            ws.append(r)
        for row in ws.iter_rows(min_row=2, min_col=7, max_col=7):
            row[0].number_format = "yyyy-mm-dd"
        for k, w in widths.items():
            ws.column_dimensions[k].width = w
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{ws.cell(1, len(hdr)).column_letter}{ws.max_row}"
    wb.remove(src)
    wb.save(F)

    # validation against the backup
    old = [tuple(r) for r in load_workbook(backup, read_only=True)["Sheet1"].iter_rows(min_row=2, values_only=True)]
    new_wb = load_workbook(F, read_only=True)
    new, ok_hdr = [], True
    for n in new_wb.sheetnames:
        it = list(new_wb[n].iter_rows(values_only=True))
        ok_hdr &= list(it[0]) == [c.value for c in hdr]
        ok_fb = all((r[4] or "No Fault Block") == n for r in it[1:])
        new += [tuple(r) for r in it[1:]]
        print(f"{n:15} rows {len(it) - 1:5}  {'ok' if ok_fb else 'WRONG FB'}")
    print("PASS" if ok_hdr else "FAIL", "- every sheet has the template header + extra columns")
    print("PASS" if sorted(map(str, old)) == sorted(map(str, new)) else "FAIL", f"- all {len(old)} rows present exactly once, values unchanged")
    print("sheets:", len(new_wb.sheetnames), "| backup:", backup.name)

    rv = load_workbook(REV)
    s = rv["Summary"]
    s.cell(s.max_row + 1, 1, "Layout").font = copy(s["A1"].font)
    s.cell(s.max_row, 2, f"Output split into one worksheet per FAULT BLOCK ({', '.join(names)}); same columns on every sheet. Backup of the single-sheet version: backups/{backup.name}")
    rv.save(REV)


if __name__ == "__main__":
    main()
