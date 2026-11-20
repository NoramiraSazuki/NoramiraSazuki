#!/usr/bin/env python3
"""Print non-empty cells of sheets (local copies in _work_pvt_scan/src). Usage: dump_sheet.py <file-substring> [sheet-substring ...]
Without sheet arguments lists the sheets. Works for OOXML (any extension) and BIFF .xls."""
import io, sys
from pathlib import Path
import openpyxl, xlrd
SRC = Path(__file__).resolve().parent.parent / "_work_pvt_scan" / "src"
f = [p for p in SRC.iterdir() if sys.argv[1] in p.name]
assert len(f) == 1, f
p = f[0]
want = sys.argv[2:]
if p.read_bytes()[:2] == b"PK":
    wb = openpyxl.load_workbook(io.BytesIO(p.read_bytes()), data_only=True)
    sheets = [(ws.title, ws.sheet_state, ws) for ws in wb.worksheets]
    if not want:
        print([(t, s) for t, s, _ in sheets]); sys.exit()
    for t, s, ws in sheets:
        if not any((t == w[1:]) if w.startswith("=") else (w.lower() in t.lower()) for w in want): continue
        print(f"=== {t} ({s})")
        for row in ws.iter_rows():
            vals = [(c.coordinate, c.value) for c in row if c.value not in (None, "")]
            if vals: print("  ", vals)
else:
    wb = xlrd.open_workbook(str(p))
    if not want:
        print([(s.name, s.visibility) for s in wb.sheets()]); sys.exit()
    for s in wb.sheets():
        if not any((s.name == w[1:]) if w.startswith("=") else (w.lower() in s.name.lower()) for w in want): continue
        print(f"=== {s.name} ({s.visibility})")
        for r in range(s.nrows):
            vals = [(f"{xlrd.colname(c)}{r+1}", s.cell_value(r, c)) for c in range(s.ncols) if s.cell_value(r, c) not in ("", None)]
            if vals: print("  ", vals)
