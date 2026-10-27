#!/usr/bin/env python3
"""Add well-identification columns to the Dulang 'Well Injection Report' .xls exports (inputs read-only).

For each input <name>.xls writes <name>_Processed.xlsx in the same folder, same worksheet name:
  title block (as in the source) / blank row / one header row / every source data row (incl. the ' Total' row).
Columns: Original Well Name | Field | Platform | Well | String Type | Reservoir | <all original columns that hold any value,
in source order> | Review Flag.  Two-row source headers (e.g. 'Injection Rate' over 'Gas (scf/d)') are joined into one
header 'Injection Rate - Gas (scf/d)'. Merged-cell filler columns that are empty in every row are not carried over.
Cells the source highlights red ('Missing / incorrect mandatory data' legend) keep the red fill.

Name pattern found in both files: 'Dulang-<A-D><3-digit no.>[S|W<2 digits>]-<SS|LS|TS>'. Parsing is case-insensitive and
tolerant of spaces / '_' / '-' separators; anything that does not match is left empty and flagged.
"""
import datetime as dt
import re
import sys
from pathlib import Path

import xlrd
from openpyxl import Workbook, load_workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Alignment, Font, PatternFill

HERE = Path(__file__).resolve().parent.parent
FILES = ["Well Injection Report - Dulang_2013-2015.xls", "Well Injection Report - Dulang_2023-2024.xls"]
NAME_RE = re.compile(r"^\s*DULANG\s*[-_ ]?\s*([ABCD])\s*[-_ ]?\s*(\d{3})\s*((?:S|W)\s*\d{2})?\s*[-_ ]\s*(SS|LS|TS)\s*$", re.I)


def _exact_string(value, _orig=None):
    """openpyxl writes floats with '%.16g', which can drop the 17th significant digit; write the exact round-trip repr."""
    if isinstance(value, float) and value == value and value not in (float("inf"), float("-inf")):
        return repr(value)
    return _ORIG_SAFE(value)


import openpyxl.cell._writer as _cw          # noqa: E402  (patch only inside this script's process)
import openpyxl.compat.strings as _cs        # noqa: E402
_ORIG_SAFE = _cs.safe_string
_cs.safe_string = _cw.safe_string = _exact_string

RED = PatternFill("solid", fgColor="FFFF0000")
BOLD = Font(bold=True)
ID_COLS = ["Original Well Name", "Field", "Platform", "Well", "String Type", "Reservoir"]


def parse(name):
    m = NAME_RE.match(str(name))
    if not m:
        return None
    letter, num, suf, st = m.group(1).upper(), m.group(2), (m.group(3) or "").upper().replace(" ", ""), m.group(4).upper()
    return f"Dulang-{letter}", f"Dulang-{letter}{num}{suf}", st


def merged_anchor(sheet):
    anc = {}
    for r0, r1, c0, c1 in sheet.merged_cells:
        for r in range(r0, r1):
            for c in range(c0, c1):
                anc[(r, c)] = (r0, c0)
    return anc


def process(src):
    b = xlrd.open_workbook(src, formatting_info=True)
    assert b.nsheets == 1, "expected one worksheet"
    s = b.sheet_by_index(0)
    hdr = [r for r in range(s.nrows) if any(str(s.cell_value(r, c)).strip() == "Well Name" for c in range(s.ncols))]
    assert len(hdr) == 1
    h = hdr[0]
    wcol = [c for c in range(s.ncols) if str(s.cell_value(h, c)).strip() == "Well Name"][0]
    anc = merged_anchor(s)
    first = h + 2                                                    # two header rows
    keep = [c for c in range(s.ncols) if c != wcol and any(str(s.cell_value(r, c)).strip() for r in range(h, s.nrows))]

    def top(c):                                                       # header text, following merges across columns
        r0, c0 = anc.get((h, c), (h, c))
        return str(s.cell_value(r0, c0)).replace("\n", " ").strip()

    names = []
    for c in keep:
        sub = str(s.cell_value(h + 1, c)).replace("\n", " ").strip()
        names.append(f"{top(c)} - {sub}" if sub else top(c))
    date_cols = {c for c in keep if top(c) == "Date"}

    def is_red(r, c):
        r0, c0 = anc.get((r, c), (r, c))
        for rr, cc in {(r, c), (r0, c0)}:
            bg = b.xf_list[s.cell_xf_index(rr, cc)].background
            if bg.fill_pattern and b.colour_map.get(bg.pattern_colour_index) == (255, 0, 0):
                return True
        return False

    out = Path(src).with_name(Path(src).stem + "_Processed.xlsx")
    if out.exists():
        raise SystemExit(f"{out.name} already exists — not overwriting")
    wb = Workbook(write_only=True)
    ws = wb.create_sheet(s.name)
    ws.freeze_panes = f"A{h + 3}"
    # title block, values as in the source (empty columns closed up), legend swatch kept red
    for r in range(h):
        row = []
        for c in range(s.ncols):
            v = s.cell_value(r, c)
            if str(v).strip():
                cell = WriteOnlyCell(ws, value=v)
                if r == 0:
                    cell.font = BOLD
                row.append(cell)
                if str(v).startswith("Missing / incorrect"):
                    sw = WriteOnlyCell(ws, value=None); sw.fill = RED; row.insert(len(row) - 1, sw)
        ws.append(row)
    ws.append([])
    header = []
    for n in ID_COLS + names + ["Review Flag"]:
        cell = WriteOnlyCell(ws, value=n); cell.font = BOLD; cell.alignment = Alignment(wrap_text=True, vertical="top")
        header.append(cell)
    ws.append(header)

    flags, n_rows, n_red, preview = [], 0, 0, []
    for r in range(first, s.nrows):
        vals = [s.cell_value(r, c) for c in range(s.ncols)]
        if not any(str(v).strip() for v in vals):
            continue
        n_rows += 1
        orig = s.cell_value(r, wcol)
        p = parse(orig)
        if p:
            ident, flag = [orig, "Dulang", *p, None], None
        elif str(orig).strip().lower() == "total":
            ident, flag = [orig, None, None, None, None, None], "Summary 'Total' row - not a well; identifiers left empty"
        else:
            ident, flag = [orig, "Dulang", None, None, None, None], "Well name does not match the expected pattern - identifiers left empty"
        if flag:
            flags.append((r + 1, orig, flag))
        row = [WriteOnlyCell(ws, value=v) for v in ident]
        if is_red(r, wcol):
            row[0].fill = RED; n_red += 1
        for c in keep:
            v = s.cell_value(r, c)
            ct = s.cell_type(r, c)
            if ct == xlrd.XL_CELL_EMPTY or (ct == xlrd.XL_CELL_TEXT and v == ""):
                v = None
            elif ct == xlrd.XL_CELL_DATE:
                v = xlrd.xldate_as_datetime(v, b.datemode)
            cell = WriteOnlyCell(ws, value=v)
            if c in date_cols and isinstance(v, dt.datetime):
                cell.number_format = "dd-mm-yyyy"
            if is_red(r, c):
                cell.fill = RED; n_red += 1
            row.append(cell)
        row.append(WriteOnlyCell(ws, value=flag))
        ws.append(row)
        if len(preview) < 10:
            preview.append(ident[:5] + [None])
    wb.save(out)
    return out, s, h, wcol, keep, names, first, n_rows, n_red, flags, preview


def verify(src, out, s, h, wcol, keep, first):
    wb = load_workbook(out, read_only=True)
    ws = wb[s.name]
    allr = list(ws.iter_rows(values_only=True))
    hi = [i for i, r in enumerate(allr) if r and r[0] == "Original Well Name"]
    assert len(hi) == 1
    width = len(allr[hi[0]])
    rows = [tuple(r) + (None,) * (width - len(r)) for r in allr[hi[0] + 1:]]   # read-only mode drops trailing empty cells
    src_rows = [r for r in range(first, s.nrows) if any(str(x).strip() for x in s.row_values(r))]
    ok_count = len(rows) == len(src_rows)
    ok_names = all(o[0] == s.cell_value(r, wcol) for o, r in zip(rows, src_rows))
    b = xlrd.open_workbook(src)
    mism = 0
    for o, r in zip(rows, src_rows):
        for k, c in enumerate(keep):
            v, ct = s.cell_value(r, c), s.cell_type(r, c)
            if ct == xlrd.XL_CELL_DATE:
                v = xlrd.xldate_as_datetime(v, b.datemode)
            if ct == xlrd.XL_CELL_EMPTY or v == "":
                v = None
            if o[6 + k] != v:
                mism += 1
    ok_res = all(o[5] is None for o in rows)
    return len(src_rows), len(rows), ok_count, ok_names, mism, ok_res, wb.sheetnames


def main():
    for f in FILES:
        src = str(HERE / f)
        out, s, h, wcol, keep, names, first, n_rows, n_red, flags, preview = process(src)
        n_src, n_out, ok_count, ok_names, mism, ok_res, sheets = verify(src, out, s, h, wcol, keep, first)
        print(f"\n=== {f}\n -> {out.name}  (sheet {sheets})")
        print(f" columns: {ID_COLS + names + ['Review Flag']}")
        print(f" data rows: source {n_src} | output {n_out} | {'PASS' if ok_count else 'FAIL'}")
        print(f" Original Well Name unchanged: {'PASS' if ok_names else 'FAIL'};  other values identical: {'PASS' if mism == 0 else f'FAIL ({mism})'};"
              f"  Reservoir empty: {'PASS' if ok_res else 'FAIL'};  red 'missing/incorrect' cells carried: {n_red}")
        print(f" rows flagged for review: {len(flags)}")
        for fl in flags:
            print("   source row", fl)
        print(" preview (Original Well Name | Field | Platform | Well | String Type | Reservoir):")
        for p in preview:
            print("   ", " | ".join("" if x is None else str(x) for x in p))


if __name__ == "__main__":
    main()
