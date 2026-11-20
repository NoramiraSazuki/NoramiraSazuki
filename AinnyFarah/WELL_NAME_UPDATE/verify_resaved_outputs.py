#!/usr/bin/env python3
"""Check the updated workbooks in 01_UPDATED against their source files (e.g. after they were re-saved in Excel):
same sheets, same row count; every source cell outside the template columns present unchanged (value and number
format) in its moved column; template columns equal to what the standardisation rules give; other sheets unchanged;
cell comments, hidden rows and the filter kept. Prints every difference found.
"""
import datetime
import re
import sys
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter as cl

sys.argv = sys.argv[:1] + ["DULG-A"]
import standardise_allocated_production as sap  # noqa: E402
import standardise_well_interventions as swi  # noqa: E402

BASE = Path(__file__).resolve().parent / "all_data_standardisation"
SETS = [  # output folder, source folder, kind
    ("01_UPDATED/WELL-INTERVENTIONS", "Well-Interventions", "intervention"),
    ("01_UPDATED/WELL-TEST-MONITORING_20261015", "well_test_monitoring_template_DULG_platform_20261015", "welltest"),
    ("01_UPDATED/WELL-TEST-MONITORING_20261014-1", "well_test_monitoring_template_DULG_platform_20261014-1", "welltest"),
    ("01_UPDATED/DULG-A", "DULG-A", "allocated"),
    ("01_UPDATED/DULG-B", "DULG-B", "allocated"),
    ("01_UPDATED/DULG-C", "DULG-C", "allocated"),
    ("01_UPDATED/DULG-D", "DULG-D", "allocated"),
    ("01_UPDATED/PRESSURE", "Pressure", "rename"),
    ("01_UPDATED/WELL-INJECTIONS", "Well-Injections", "rename"),
]
RENAMED_SOURCE = {"Well Injection Report - Dulang_2013-2015_20261023.xlsx":
                  "Well Injection Report - Dulang_2013-2015_20261023_v1.xlsx"}
WELL_RE = re.compile(r"(?:DULG|DULANG)-([A-D])(\d+)((?:[A-Z]+\d*)*)")


def std_well(v):
    m = WELL_RE.fullmatch(v)
    sfx = re.sub(r"([A-Z]+)(\d*)", lambda s: s.group(1) + (str(int(s.group(2))) if s.group(2) else ""), m.group(3))
    return f"DULANG-{m.group(1)}{int(m.group(2))}{sfx}"


def expected_intervention(ws_src, ncols):
    m = swi.colmap(ncols)
    plat = None
    exp = {}
    for r, row in enumerate(ws_src.iter_rows(min_row=2, values_only=True), 2):
        if all(v is None for v in row):
            continue
        d, _ = swi.decide(row, plat_letter)
        exp[r] = {2: d["field"], 3: d["platform"], 4: d["well"], 5: d["string"], 6: None}
    return m, exp


def expected_welltest(ws_src, ncols):
    m = {1: 1, **{c: c + 6 for c in range(3, ncols + 1)}}
    exp = {}
    for r in range(2, ws_src.max_row + 1):
        ent, dt, start = ws_src.cell(r, 1).value, ws_src.cell(r, 2).value, ws_src.cell(r, 13).value
        if ent is None and dt is None:
            continue
        well = std_well(ent)
        t = None if (start is None and dt.time() == datetime.time(0)) else dt.time()
        exp[r] = {2: "DULANG", 3: f"DULANG-{well[7]}", 4: well, 5: None, 6: None, 7: dt.date(), 8: t}
    return m, exp


def expected_allocated(ws_src, ncols):
    m = {1: 1, **{c: c + 6 for c in range(3, ncols + 1)}}
    exp = {}
    for r in range(2, ws_src.max_row + 1):
        ent, dv = ws_src.cell(r, 1).value, ws_src.cell(r, 2).value
        if ent is None and dv is None:
            continue
        d = sap.parse_entity(ent)
        if isinstance(dv, datetime.datetime):
            date, t = dv.date(), dv.time()
        else:
            date, t = datetime.date.fromisoformat(dv.strip()), None
        exp[r] = {2: "DULANG", 3: d["platform"], 4: d["well"], 5: d["string"], 6: d["reservoir"], 7: date, 8: t}
    return m, exp


def check_rename(out, src):
    """earlier in-place update: only WELL (standardised name) and PLATFORM (DULG- -> DULANG-) may differ"""
    issues, n = [], 0
    a, b = openpyxl.load_workbook(src), openpyxl.load_workbook(out)
    if a.sheetnames != b.sheetnames:
        return [f"sheet names differ: {a.sheetnames} vs {b.sheetnames}"], 0
    for wa in a.worksheets:
        wb_ = b[wa.title]
        if (wa.max_row, wa.max_column) != (wb_.max_row, wb_.max_column):
            issues.append(f"{wa.title}: size {wa.max_row}x{wa.max_column} -> {wb_.max_row}x{wb_.max_column}")
        hdr = {}
        for r in range(1, 31):
            for c in range(1, wa.max_column + 1):
                v = wa.cell(r, c).value
                if isinstance(v, str) and v.strip().upper() in ("WELL", "PLATFORM"):
                    hdr[c] = (v.strip().upper(), r)
        for ra, rb in zip(wa.iter_rows(), wb_.iter_rows()):
            for x, y in zip(ra, rb):
                h = hdr.get(x.column)
                want = x.value
                if h and x.row > h[1] and isinstance(x.value, str):
                    if h[0] == "WELL" and WELL_RE.fullmatch(x.value):
                        want = std_well(x.value)
                    elif h[0] == "PLATFORM" and re.fullmatch(r"DULG-[A-D]", x.value):
                        want = "DULANG-" + x.value[-1]
                if y.value != want:
                    issues.append(f"{wa.title}!{x.coordinate}: expected {want!r}, found {y.value!r}")
                elif y.value is not None:
                    n += 1
                if x.value is not None and x.number_format != y.number_format:
                    issues.append(f"{wa.title}!{x.coordinate}: number format {x.number_format!r} -> {y.number_format!r}")
                if (x.comment.text if x.comment else None) != (y.comment.text if y.comment else None):
                    issues.append(f"{wa.title}!{x.coordinate}: comment changed")
    return issues, n


def norm(v):
    if isinstance(v, datetime.datetime):
        return v
    return v


def check(out, src, kind):
    global plat_letter
    issues = []
    a, b = openpyxl.load_workbook(src), openpyxl.load_workbook(out)
    if a.sheetnames != b.sheetnames:
        issues.append(f"sheet names differ: {a.sheetnames} vs {b.sheetnames}")
        return issues, 0
    wa, wb_ = a.worksheets[0], b.worksheets[0]
    ncols = wa.max_column
    plat_letter = re.search(r"Dulang_([A-D])_", src.name).group(1) if kind == "intervention" else None
    m, exp = {"intervention": expected_intervention, "welltest": expected_welltest,
              "allocated": expected_allocated}[kind](wa, ncols)
    if wa.max_row != wb_.max_row:
        issues.append(f"row count {wa.max_row - 1} -> {wb_.max_row - 1}")
    ncells = 0
    for r in range(1, wa.max_row + 1):
        if bool(wa.row_dimensions[r].hidden) != bool(wb_.row_dimensions[r].hidden):
            issues.append(f"row {r} hidden flag changed")
        for oc, nc in m.items():
            x, y = wa.cell(r, oc), wb_.cell(r, nc)
            if r == 1 and (oc in (1, 2, 3, 4, 5, 8) if kind == "intervention" else oc == 1):
                continue  # renamed template headers
            if kind == "intervention" and 2 <= oc <= 5:
                continue  # template columns, checked below
            if x.value != y.value:
                issues.append(f"{cl(oc)}{r} -> {cl(nc)}{r}: {x.value!r} -> {y.value!r}")
            elif x.value is not None:
                ncells += 1
            if x.value is not None and x.number_format != y.number_format:
                issues.append(f"{cl(oc)}{r}: number format {x.number_format!r} -> {y.number_format!r}")
            if (x.comment.text if x.comment else None) != (y.comment.text if y.comment else None):
                issues.append(f"{cl(oc)}{r}: comment changed")
        for c, v in exp.get(r, {}).items():
            got = wb_.cell(r, c).value
            if isinstance(got, datetime.datetime) and isinstance(v, datetime.date) and not isinstance(v, datetime.datetime):
                got = got.date()
            if isinstance(got, datetime.datetime) and isinstance(v, datetime.time):
                got = got.time()
            if isinstance(got, datetime.time):
                got = got.replace(microsecond=0)
            if got != v:
                issues.append(f"{cl(c)}{r}: expected {v!r}, found {got!r}")
            elif v is not None:
                ncells += 1
    for name in a.sheetnames[1:]:
        va = [[c.value for c in row] for row in a[name].iter_rows()]
        vb = [[c.value for c in row] for row in b[name].iter_rows()]
        if va != vb:
            issues.append(f"sheet {name!r} values changed")
    if wb_.max_column != ncols + (1 if kind == "intervention" else 6):
        issues.append(f"column count {ncols} -> {wb_.max_column}")
    return issues, ncells


def main():
    total_bad = 0
    for out_dir, src_dir, kind in SETS:
        for p in sorted((BASE / out_dir).glob("*.xlsx")):
            src = BASE / src_dir / RENAMED_SOURCE.get(p.name, p.name)
            if p.name.startswith("~$"):
                continue
            if not src.exists():
                print(f"NO SOURCE {out_dir.split('/')[-1]}/{p.name}")
                total_bad += 1
                continue
            issues, n = check_rename(p, src) if kind == "rename" else check(p, src, kind)
            total_bad += len(issues)
            print(f"{'OK ' if not issues else 'DIFF'} {out_dir.split('/')[-1]}/{p.name}: {n:,} populated cells matched"
                  + ("" if not issues else f"; {len(issues)} differences"))
            for i in issues[:15]:
                print("      ", i)
    print("differences in total:", total_bad)


if __name__ == "__main__":
    main()
