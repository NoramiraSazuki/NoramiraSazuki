#!/usr/bin/env python3
"""Consolidate the Dulang deviation survey workbooks into one WELL_TRAJECTORY sheet per platform.

Inputs (read only): one workbook per platform, one worksheet per well/wellbore, header in row 1, stations below.
Outputs: WELL_TRAJECTORY/DULANG-<X>_Well_Trajectory.xlsx (never overwritten) and a review workbook in _reports/.

Rules agreed with the owner:
- Seven template columns first (ORIGINAL WELL NAME = worksheet name, FIELD, PLATFORM, WELL, STRING TYPE, RESERVOIR,
  DATE); the deviation sheets state no string type, reservoir or date, so those three stay blank.
- Survey columns are aligned only where the header means the same thing (AZ/AZM, E/EAST/EASTING, N/NORT/NORTHING);
  'AZM (TN)' (true north stated) is kept in its own column. No sheet states units; values are copied as supplied.
- A station written on several consecutive rows (same MD, values never conflicting) becomes one row.
- D_A10 rows 27-50 (an unlabelled MD/depth table under the survey) are kept as D_A10 rows in UNLABELLED columns.
- Side tables outside the survey columns, rows without MD, blank rows and empty worksheets are left out and reported.
Run with a Python that has openpyxl and xlrd, e.g. ../DOCUMENT_PVT_Fluid_Analysis/.venv/bin/python.
"""
import collections
import datetime
import os
import re
import sys
import warnings
from pathlib import Path

import openpyxl
import openpyxl.cell._writer
import xlrd
from openpyxl.compat import safe_string
from openpyxl.styles import Font

warnings.filterwarnings("ignore", message="wmf image format")
# openpyxl writes floats with %.16g, which drops the 17th digit Excel stores for some formula results
# (-104.24000000000001 -> -104.24); write the shortest exact form instead so values stay as supplied
openpyxl.cell._writer.safe_string = lambda v: repr(v) if isinstance(v, float) else safe_string(v)

SRC = Path("/Volumes/homes/public/DulangOSS/20261016/DULANG_WELLS/Well_Trajectory/3. Well Trajectory")
OUT = Path(__file__).resolve().parent
REPORTS = OUT / "_reports"
INPUTS = {"A": "Dulang A Deviation Survey Rev 1.xlsx", "B": "Dulang B Deviation Survey.xls",
          "C": "Dulang C Deviation Survey.xls", "D": "Dulang D Deviation Survey.xls"}
TEMPLATE = ["ORIGINAL WELL NAME", "FIELD", "PLATFORM", "WELL", "STRING TYPE", "RESERVOIR", "DATE"]

# header text -> meaning; output order follows this list
MEANINGS = [("MD", {"MD"}), ("AZ", {"AZ", "AZM"}), ("AZ_TN", {"AZM (TN)"}), ("INC", {"INC"}), ("TVD", {"TVD"}),
            ("TVDSS", {"TVDSS"}), ("EAST", {"E", "EAST", "EASTING"}), ("NORTH", {"N", "NORT", "NORTHING"}),
            ("DLS", {"DLS"}), ("VS", {"VS"})]
MEANING_OF = {h: k for k, hs in MEANINGS for h in hs}
# standard output headers (owner's request); other meanings use the header most of the platform's sheets use
OUTPUT_HEADER = {"EAST": "EASTING", "NORTH": "NORTHING"}
# D_A12_S1 column E has no header but holds =(TVD-32)*-1, the same TVDSS formula the other A sheets use
HEADER_FIX = {("A", "D_A12_S1", 5): "TVDSS"}
# D_A10: rows 27-50 are a separate unlabelled table in columns A-E (D = A - $C$27, E = B - $C$27)
UNLABELLED_BLOCK = {("A", "D_A10"): (27, 50, 5)}
NAME_RE = re.compile(r"(?:D_)?([A-D])(\d+)((?:_?[A-Z]+\d+)*)")


def blank(v):
    return v is None or (isinstance(v, str) and not v.strip())


def read_workbook(letter):
    """[(sheet name, visibility, grid)], grid[r][c] = (value, number format), 1-based r/c via padding"""
    path = SRC / INPUTS[letter]
    sheets = []
    if path.suffix == ".xlsx":
        wb = openpyxl.load_workbook(path, data_only=True)
        for ws in wb.worksheets:
            grid = [[(c.value, c.number_format) for c in row] for row in ws.iter_rows()]
            sheets.append((ws.title, ws.sheet_state, grid))
    else:
        bk = xlrd.open_workbook(path, formatting_info=True)
        for s in bk.sheets():
            grid = []
            for r in range(s.nrows):
                row = []
                for c in range(s.ncols):
                    t = s.cell_type(r, c)
                    v = None if t in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK) else s.cell_value(r, c)
                    if t == xlrd.XL_CELL_ERROR:
                        raise ValueError(f"{letter}/{s.name} R{r + 1}C{c + 1}: error cell")
                    fmt = bk.format_map[bk.xf_list[s.cell_xf_index(r, c)].format_key].format_str
                    row.append((v, fmt))
                grid.append(row)
            sheets.append((s.name, ["visible", "hidden", "veryHidden"][s.visibility], grid))
    return sheets


def parse_name(name):
    m = NAME_RE.fullmatch(name)
    if not m:
        return None, None
    sfx = re.sub(r"_?([A-Z]+)(\d+)", lambda s: s.group(1) + str(int(s.group(2))), m.group(3))
    return m.group(1), f"DULANG-{m.group(1)}{int(m.group(2))}{sfx}"


def cell(grid, r, c):
    return grid[r - 1][c - 1] if r <= len(grid) and c <= len(grid[r - 1]) else (None, "General")


def extract(letter, name, grid, log):
    """stations of one sheet: [(source rows, {meaning: (value, fmt)})]"""
    ncols = max((len(r) for r in grid), default=0)
    cols = {}  # source column -> meaning
    for c in range(1, ncols + 1):
        h = HEADER_FIX.get((letter, name, c), cell(grid, 1, c)[0])
        if blank(h):
            continue
        h = h.strip()
        if h not in MEANING_OF:
            raise ValueError(f"{letter}/{name}: unknown header {h!r} in column {c}")
        cols[c] = MEANING_OF[h]
        log["headers"].append((letter, name, openpyxl.utils.get_column_letter(c), cell(grid, 1, c)[0], cols[c]))
    survey_cols = range(1, max(cols) + 1)
    assert all(c in cols for c in survey_cols if any(not blank(cell(grid, r, c)[0]) for r in range(2, len(grid) + 1))), \
        f"{letter}/{name}: data under a blank header"
    block = UNLABELLED_BLOCK.get((letter, name))

    # cells outside the survey columns (side tables, stray spaces)
    side = collections.defaultdict(list)
    for r in range(1, len(grid) + 1):
        for c in range(max(cols) + 1, ncols + 1):
            v = cell(grid, r, c)[0]
            if v is not None:
                side[c].append((r, v))
    for c, items in side.items():
        cl = openpyxl.utils.get_column_letter(c)
        what = "whitespace only" if all(blank(v) for _, v in items) else "side table / values outside the survey columns"
        log["excluded"].append((letter, name, f"{cl}{items[0][0]}:{cl}{items[-1][0]}", len(items), what,
                                "; ".join(f"{cl}{r}={v!r}" for r, v in items[:12]) + (" ..." if len(items) > 12 else "")))

    rows = []
    for r in range(2, len(grid) + 1):
        vals = {c: cell(grid, r, c) for c in survey_cols}
        if all(blank(v) for v, _ in vals.values()):
            log["blank_rows"][(letter, name)].append(r)
            continue
        if block and block[0] <= r <= block[1]:
            assert all(blank(cell(grid, r, c)[0]) for c in survey_cols if c > block[2]), (name, r)
            rows.append(([r], {f"UNLABELLED COL {openpyxl.utils.get_column_letter(c)}": cell(grid, r, c)
                               for c in range(1, block[2] + 1) if not blank(cell(grid, r, c)[0])}))
            continue
        md_col = next(c for c, k in cols.items() if k == "MD")
        if blank(vals[md_col][0]):
            log["excluded"].append((letter, name, f"row {r}", sum(not blank(v) for v, _ in vals.values()),
                                    "row without MD (not a survey station)",
                                    "; ".join(f"{openpyxl.utils.get_column_letter(c)}{r}={v!r}"
                                              for c, (v, _) in vals.items() if not blank(v))))
            continue
        for c, (v, _) in vals.items():
            if isinstance(v, str) or (v is not None and not isinstance(v, (int, float))):
                raise ValueError(f"{letter}/{name} {openpyxl.utils.get_column_letter(c)}{r}: non-numeric {v!r}")
        rows.append(([r], {cols[c]: vf for c, vf in vals.items() if not blank(vf[0])}))

    # one row per station: merge consecutive rows with the same MD whose values never conflict
    stations = []
    for src_rows, vals in rows:
        prev = stations[-1] if stations else None
        if (prev and "MD" in vals and "MD" in prev[1] and prev[1]["MD"][0] == vals["MD"][0]
                and all(prev[1][k][0] == v[0] for k, v in vals.items() if k in prev[1])):
            prev[0].extend(src_rows)
            for k, v in vals.items():
                prev[1].setdefault(k, v)
        else:
            stations.append((list(src_rows), dict(vals)))
    same_md = [(a[0][0], b[0][0]) for a, b in zip(stations, stations[1:])
               if "MD" in a[1] and "MD" in b[1] and a[1]["MD"][0] == b[1]["MD"][0]]
    for a, b in same_md:
        log["flags"].append((letter, name, f"rows {a} and {b} share an MD but differ in other values - kept both"))
    return stations, cols


def build(letter, log):
    sheets = read_workbook(letter)
    data, labels = [], collections.defaultdict(collections.Counter)
    for name, state, grid in sheets:
        if not any(not blank(v) for row in grid for v, _ in row):
            log["skipped"].append((letter, name, state, "worksheet is empty (no cells with data)"))
            continue
        stations, cols = extract(letter, name, grid, log)
        for c, k in cols.items():
            labels[k][str(cell(grid, 1, c)[0] or HEADER_FIX.get((letter, name, c))).strip()] += 1
        plat, well = parse_name(name)
        if well is None:
            log["flags"].append((letter, name, "worksheet name could not be parsed - PLATFORM and WELL left blank"))
        elif plat != letter:
            log["flags"].append((letter, name, f"platform letter {plat} differs from the {letter} workbook"))
        n_src = sum(1 for r in range(2, len(grid) + 1)
                    if any(not blank(v) for v, _ in grid[r - 1][:max(cols)]))
        log["recon"].append({"platform": letter, "sheet": name, "well": well, "source_rows": n_src,
                             "blank_rows": len(log["blank_rows"][(letter, name)]),
                             "rows_without_md": sum(1 for e in log["excluded"] if e[:2] == (letter, name)
                                                    and e[4].startswith("row without MD")),
                             "rows_merged": sum(len(s[0]) - 1 for s in stations), "stations": len(stations)})
        data.append((name, plat, well, stations))
    keys = [k for k, _ in MEANINGS if k in labels]
    unl = sorted({k for _, _, _, st in data for _, v in st for k in v if k.startswith("UNLABELLED")})
    heads = [OUTPUT_HEADER.get(k, labels[k].most_common(1)[0][0]) for k in keys] + unl
    for k in keys:
        if set(labels[k]) != {heads[keys.index(k)]}:
            log["aligned"].append((letter, heads[keys.index(k)], dict(labels[k])))
    return data, keys + unl, heads


def write(letter, data, keys, heads, path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "WELL_TRAJECTORY"
    ws.append(TEMPLATE + heads)
    for c in ws[1]:
        c.font = Font(bold=True)
    for name, plat, well, stations in data:
        for _, vals in stations:
            r = ws.max_row + 1
            for i, v in enumerate([name, "DULANG", f"DULANG-{plat}" if plat else None, well, None, None, None], 1):
                ws.cell(r, i, v)
            for i, k in enumerate(keys, len(TEMPLATE) + 1):
                if k in vals:
                    c = ws.cell(r, i, vals[k][0])
                    c.number_format = vals[k][1]
    ws.freeze_panes = "A2"
    for i, w in enumerate([20, 9, 11, 16, 12, 11, 11], 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    tmp = path.with_name(path.stem + ".partial.xlsx")
    with open(tmp, "xb") as fh:
        wb.save(fh)
    return tmp


def validate(letter, tmp, data, keys, heads):
    """re-read the written file and check it against the source workbook read afresh"""
    wb = openpyxl.load_workbook(tmp)
    assert wb.sheetnames == ["WELL_TRAJECTORY"], wb.sheetnames
    ws = wb["WELL_TRAJECTORY"]
    hdr = [c.value for c in ws[1]]
    assert hdr[:7] == TEMPLATE and hdr[7:] == heads, hdr
    out = collections.defaultdict(list)
    order = []
    for r in range(2, ws.max_row + 1):
        row = [c.value for c in ws[r]]
        if not order or order[-1] != row[0]:
            assert row[0] not in order, f"{row[0]} rows are not contiguous"
            order.append(row[0])
        out[row[0]].append(row)
    assert order == [d[0] for d in data], "worksheet order changed"
    src = {name: grid for name, _, grid in read_workbook(letter)}
    n = 0
    for name, plat, well, stations in data:
        rows = out[name]
        assert len(rows) == len(stations), (name, len(rows), len(stations))
        for row, (src_rows, _) in zip(rows, stations):
            assert row[1:7] == ["DULANG", f"DULANG-{plat}", well, None, None, None], (name, row[:7])
            got = dict(zip(keys, row[7:]))
            # every source cell of the station's rows equals the output value under the matching meaning
            grid = src[name]
            for r in src_rows:
                for c in range(1, len(grid[r - 1]) + 1):
                    v = grid[r - 1][c - 1][0]
                    if blank(v):
                        continue
                    h = HEADER_FIX.get((letter, name, c), grid[0][c - 1][0] if c <= len(grid[0]) else None)
                    k = (f"UNLABELLED COL {openpyxl.utils.get_column_letter(c)}"
                         if (letter, name) in UNLABELLED_BLOCK and UNLABELLED_BLOCK[(letter, name)][0] <= r
                         <= UNLABELLED_BLOCK[(letter, name)][1] else MEANING_OF.get(str(h).strip()))
                    if k is None:
                        continue  # side-table column, reported as excluded
                    assert got[k] == v, (name, r, c, got[k], v)
                    n += 1
            # and nothing in the output row came from anywhere else
            for k, v in got.items():
                if v is not None:
                    assert any(not blank(grid[r - 1][c - 1][0]) and grid[r - 1][c - 1][0] == v
                               for r in src_rows for c in range(1, len(grid[r - 1]) + 1)), (name, k, v)
    return ws.max_row - 1, n


def main():
    targets = {x: OUT / f"DULANG-{x}_Well_Trajectory.xlsx" for x in INPUTS}
    existing = [p for p in targets.values() if p.exists() or p.with_name(p.stem + ".partial.xlsx").exists()]
    if existing:
        sys.exit(f"STOP: output already exists - not overwritten: {[p.name for p in existing]}")
    log = {"headers": [], "excluded": [], "skipped": [], "flags": [], "recon": [], "aligned": [],
           "blank_rows": collections.defaultdict(list)}
    log["flags"] += [
        ("A", "D_A12_S1", "column E has no header; mapped to TVDSS (its formula =(TVD-32)*-1 matches TVDSS elsewhere)"),
        ("A", "D_A10", "rows 27-50 are an unlabelled table (A, B, C27=13.88, D=A-$C$27, E=B-$C$27); kept as D_A10 "
                       "rows in UNLABELLED COL A-E, meaning unconfirmed; the labelled survey ends at MD 370.7"),
        ("A", "D_A4ST1", "azimuth header 'AZM (TN)' states true north; kept separate from AZM (reference not stated)"),
        ("*", "*", "no worksheet states units; values copied as supplied. STRING TYPE, RESERVOIR and DATE are not "
                   "stated in any deviation worksheet, so they are blank"),
        ("A", "D_*", "worksheet names carry a 'D_' prefix and '_' before the suffix (D_A11_S1); WELL drops both "
                     "(DULANG-A11S1), ORIGINAL WELL NAME keeps the name as written"),
    ]
    summary = []
    for x in INPUTS:
        data, keys, heads = build(x, log)
        tmp = write(x, data, keys, heads, targets[x])
        nrows, ncells = validate(x, tmp, data, keys, heads)
        os.replace(tmp, targets[x])
        summary.append((x, targets[x], len(data), nrows, ncells, heads))
        print(f"{targets[x].name}: {len(data)} wells, {nrows} station rows, {ncells} source values checked; "
              f"columns {heads}")
    write_report(log, summary)


def write_report(log, summary):
    REPORTS.mkdir(exist_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Summary"
    ws.append(["PLATFORM", "OUTPUT", "WELLS", "STATION ROWS", "SOURCE VALUES CHECKED", "SURVEY COLUMNS"])
    for x, p, w, n, c, heads in summary:
        ws.append([f"DULANG-{x}", str(p), w, n, c, ", ".join(heads)])
    sheets = [
        ("Reconciliation", ["PLATFORM", "WORKSHEET", "WELL", "SOURCE DATA ROWS", "BLANK ROWS SKIPPED",
                            "ROWS WITHOUT MD EXCLUDED", "REPEAT ROWS MERGED", "OUTPUT STATIONS", "CHECK"],
         [[f"DULANG-{d['platform']}", d["sheet"], d["well"], d["source_rows"], d["blank_rows"], d["rows_without_md"],
           d["rows_merged"], d["stations"],
           "OK" if d["source_rows"] - d["rows_without_md"] - d["rows_merged"] == d["stations"] else "MISMATCH"]
          for d in log["recon"]]),
        ("Column_Mapping", ["PLATFORM", "WORKSHEET", "SOURCE COLUMN", "SOURCE HEADER", "MEANING"],
         [[f"DULANG-{a}", b, c, d, e] for a, b, c, d, e in log["headers"]]),
        ("Aligned_Headers", ["PLATFORM", "OUTPUT HEADER", "SOURCE HEADERS (worksheets)"],
         [[f"DULANG-{a}", b, str(c)] for a, b, c in log["aligned"]]),
        ("Excluded_Cells", ["PLATFORM", "WORKSHEET", "RANGE", "CELLS", "REASON", "VALUES"],
         [[f"DULANG-{a}", *rest] for a, *rest in log["excluded"]]),
        ("Skipped_Worksheets", ["PLATFORM", "WORKSHEET", "VISIBILITY", "REASON"],
         [[f"DULANG-{a}", *rest] for a, *rest in log["skipped"]]),
        ("Flags", ["PLATFORM", "WORKSHEET", "NOTE"], [[f"DULANG-{a}", *rest] for a, *rest in log["flags"]]),
    ]
    for title, head, rows in sheets:
        s = wb.create_sheet(title)
        s.append(head)
        for r in rows:
            s.append(r)
    for s in wb.worksheets:
        for c in s[1]:
            c.font = Font(bold=True)
        s.freeze_panes = "A2"
    p = REPORTS / f"well_trajectory_consolidation_review_{datetime.datetime.now():%Y%m%d_%H%M%S}.xlsx"
    with open(p, "xb") as fh:
        wb.save(fh)
    print("report:", p)
    bad = [d for d in sheets[0][2] if d[-1] != "OK"]
    print("reconciliation mismatches:", len(bad))


if __name__ == "__main__":
    main()
