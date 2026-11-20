#!/usr/bin/env python3
"""One workbook per fault block from the Dulang Unit + Dulang West water-injection outputs.

Changes vs the combined files: FIELD 'DULG' -> 'DULANG'; header 'CUM_WTR.INJ BBL' -> 'CUM_WATER_INJ BBL'.
The combined source workbooks are only read. Output: output_water-injection/water_injection_by_fault_block/.
"""
import hashlib
from copy import copy
from pathlib import Path

from openpyxl import Workbook, load_workbook

OUT = Path(__file__).resolve().parent.parent / "output_water-injection"
SRC = [OUT / "Dulang_Unit_Water_Injection.xlsx", OUT / "Dulang_West_Water_Injection.xlsx"]
DEST = OUT / "water_injection_by_fault_block"
sha = lambda p: hashlib.sha256(p.read_bytes()).hexdigest()


def main():
    before = {p: sha(p) for p in SRC}
    DEST.mkdir(exist_ok=True)
    seen, report = {}, []
    for p in SRC:
        wb = load_workbook(p)
        for ws in wb.worksheets:
            assert ws.title not in seen, f"fault block {ws.title} in two files"
            seen[ws.title] = p.name
            hdr = list(ws[1])
            names = [c.value for c in hdr]
            wi = next(i for i, h in enumerate(names) if h and h.strip().upper().startswith("CUM_WTR.INJ"))
            fi = names.index("FIELD")
            nb = Workbook()
            ns = nb.active
            ns.title = ws.title[:31]
            for c in hdr:
                v = c.value.replace("CUM_WTR.INJ", "CUM_WATER_INJ") if c.column - 1 == wi else c.value
                n = ns.cell(1, c.column, v)
                n.font, n.fill, n.alignment, n.border = copy(c.font), copy(c.fill), copy(c.alignment), copy(c.border)
            ns.row_dimensions[1].height = ws.row_dimensions[1].height
            rows = 0
            for r in ws.iter_rows(min_row=2, values_only=True):
                r = list(r)
                if r[fi] == "DULG":
                    r[fi] = "DULANG"
                ns.append(r)
                rows += 1
            for row in ns.iter_rows(min_row=2, min_col=7, max_col=7):
                row[0].number_format = "yyyy-mm-dd"
            for k, d in ws.column_dimensions.items():
                ns.column_dimensions[k].width = d.width
            ns.freeze_panes = "A2"
            ns.auto_filter.ref = f"A1:{ns.cell(1, len(hdr)).column_letter}{ns.max_row}"
            name = "water_injection_" + ws.title.replace(" ", "_") + ".xlsx"
            out = DEST / name
            assert not out.exists(), out
            nb.save(out)
            report.append((name, rows, ws, wi, fi))

    # validation
    ok_all = True
    for name, rows, ws, wi, fi in report:
        got = list(load_workbook(DEST / name, read_only=True).active.iter_rows(values_only=True))
        src = [list(r) for r in ws.iter_rows(values_only=True)]
        hdr_ok = got[0][wi] == "CUM_WATER_INJ BBL " and all(a == b for k, (a, b) in enumerate(zip(got[0], src[0])) if k != wi)
        fld_ok = all(g[fi] != "DULG" for g in got[1:]) and all((g[fi] == "DULANG") == (s[fi] in ("DULG", "DULANG")) for g, s in zip(got[1:], src[1:]))
        val_ok = len(got) == len(src) and all(all(a == b for k, (a, b) in enumerate(zip(g, s)) if k != fi) for g, s in zip(got[1:], src[1:]))
        ok = hdr_ok and fld_ok and val_ok
        ok_all &= ok
        print(f"{'PASS' if ok else 'FAIL'}  {name:42} rows {rows:5}")
    print("PASS" if ok_all else "FAIL", f"- {len(report)} workbooks, {sum(r[1] for r in report)} rows; header renamed, FIELD = DULANG, all other values unchanged")
    print("PASS" if all(sha(p) == h for p, h in before.items()) else "FAIL", "- combined source workbooks unchanged")


if __name__ == "__main__":
    main()
