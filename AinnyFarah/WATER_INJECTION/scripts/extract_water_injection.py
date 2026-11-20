#!/usr/bin/env python3
"""Extract water-injection data from the Dulang Unit MBAL workbooks into water_injection_template.xlsx.

Decisions (user): raw MBAL blocks only; well-level monthly rows -> running cumulative per well (flagged derived);
include every raw block (zero injection too); keep all copies, flag duplicates (never merged).

Kind A  raw MBAL block: header 'Date | PRESSURESURV | Cum_Oil bbl | Cum_Gas cf | Cum_Water bbl | Cum_Gas_Inj | Cum_Wtr.Inj bbl'
        whose values are in field units (oil >= 1000 bbl and gas/oil ratio > 10, i.e. not /1e6 or /1e3 copies),
        on sheets that are not modified/adjusted versions.
Kind C  well-level monthly rows (Well/String/FB/Sand + GASINJ/WTRINJ) -> running cumulative per group.
NAS is read-only (bytes read into memory); the template is copied, never edited.
"""
import datetime as dt
from copy import copy
import hashlib
import json
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

import xlrd

sys.path.insert(0, str(Path(__file__).resolve().parent))
from scan_injection import AREA, NAS, TAG, read, sheets  # noqa: E402
UNIT = AREA == "Dulang Unit"
STEM = AREA.replace(" ", "_") + "_Water_Injection"

ROOT = Path(__file__).resolve().parent.parent
TPL = ROOT / "water_injection_template.xlsx"
OUT = ROOT / "output_water-injection"
HDR = ["date", "pressuresurv", "cum_oil", "cum_gas", "cum_water", "cum_gas_in", "cum_wtr.in"]
MODIFIED = re.compile(r"mod|adjust|result|revis", re.I)
norm = lambda v: re.sub(r"\s+", "", str(v)).lower() if v is not None else ""


def as_date(v, datemode=0):
    if isinstance(v, dt.datetime):
        return v
    if isinstance(v, float) and 10000 < v < 80000:
        return xlrd.xldate.xldate_as_datetime(v, datemode)
    return None


def num(v):
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        try:
            return float(v.replace(",", "").strip())
        except ValueError:
            return None
    return None


def raw_blocks(rows):
    """yield (header_row_idx, col_idx, header_texts, data_rows[(src_row, date, p, co, cg, cw, cgi, cwi)])."""
    for i, r in enumerate(rows):
        for j in range(len(r) - 6):
            if norm(r[j]) != "date" or not all(norm(r[j + k]).startswith(HDR[k]) for k in range(1, 7)):
                continue
            hdr = [str(r[j + k]).strip() for k in range(7)]
            data = []
            for ii in range(i + 1, len(rows)):
                rr = rows[ii] + [None] * 8
                d = as_date(rr[j])
                if d is None:
                    if isinstance(rr[j], str) and rr[j].strip():
                        break                      # next header / text block
                    continue
                vals = [num(rr[j + k]) for k in range(1, 7)]
                if all(x is None for x in vals):
                    continue                       # date only, no data
                data.append((ii + 1, d, *vals))
            yield i, j, hdr, data


def title_context(rows, i, j):
    out = []
    for ii in range(max(0, i - 5), i):
        for jj in range(max(0, j - 1), min(len(rows[ii]), j + 8)):
            v = rows[ii][jj]
            if isinstance(v, str) and v.strip() and norm(v) not in ("mbalformat", "chart", "mbal"):
                out.append(v.strip())
    return out


def is_raw(data):
    oil = max((r[3] or 0) for r in data) if data else 0
    gas = max((r[4] or 0) for r in data) if data else 0
    return oil >= 1000 and (gas == 0 or gas / oil > 10)


def col(n):
    s = ""
    while n:
        n, k = divmod(n - 1, 26)
        s = chr(65 + k) + s
    return s


def inventory():
    scan = json.loads((OUT / f"scan_results{TAG}.json").read_text())
    hit = defaultdict(set)
    for r in scan:
        if r.get("n_hits"):
            hit[r["file"]].add(r["sheet"])
    blocks, skipped = [], []
    for f in sorted(hit):
        for name, state, rows in sheets(read(NAS / f)):
            if name not in hit[f]:
                continue
            found = False
            for i, j, hdr, data in raw_blocks(rows):
                found = True
                ref = f"{col(j + 1)}{i + 1}"
                b = dict(file=f, sheet=name, header_cell=ref, header=hdr, rows=len(data), title=title_context(rows, i, j), data=data)
                if MODIFIED.search(name):
                    skipped.append((f, name, ref, "modified/adjusted sheet — not raw"))
                elif not data:
                    skipped.append((f, name, ref, "header without data rows"))
                elif not is_raw(data):
                    skipped.append((f, name, ref, "scaled copy (MM or M units under raw header)"))
                else:
                    blocks.append(b)
            if not found:
                skipped.append((f, name, "", "no raw MBAL block (derived MM columns / well-level / model sheet)"))
    return blocks, skipped


# ------------------------------------------------------------------ labels (fault block, reservoir, basis) per raw block
LABELS_UNIT = {
    ("N1/N1 - E2022, E23, E32 (Done)/Production Data N1 - E2022_ E23_ E32.xlsx", "Cum ali", "A2"): ("N1", "E2022/E23/E32", "folder name"),
    ("N1/N1 - E81213 (Done)/N1 E81213 basic data.xlsx", "Sheet1", "A6"): ("N1", "E8/12/13", "sheet title 'N1 E8/12/13'"),
    ("N1/N1 - E81213 (Done)/Working File/N1 - E1213/N1 - E1213 (water injection).xlsx", "Prod Data", "A5"): ("N1", "E12/13", "sheet title 'E12/13'"),
    ("N1/N1 - E81213 (Done)/Working File/N1 - E8/N1 - E8.xlsx", "Data", "A3"): ("N1", "E8", "sheet title 'N1 E8'"),
    ("N1/N1 - LD1,LD3/N1 LD1-LD3 basic data.xlsx", "Sheet1", "D6"): ("N1", "LD1-LD3", "file name"),
    ("N2/N2 - E1011, E1213, E14 (done)/working file/Data N2 - E1011_ E1213_ E14 (Single tank).xlsx", "Cum Ali", "A5"): ("N2", "E10/14", "sheet title 'E10/14'"),
    ("N2/N2 - E1011, E1213, E14 (done)/working file/Data N2 E10E14.xlsx", "1014 (ali)", "A5"): ("N2", "E10/14", "sheet title 'E10/14'"),
    ("N3AB/N3AB - E10 to E14 (Done)/N3AB E10-E14 Basic Data.xlsx", "Sheet1", "C6"): ("N3AB", "E10-14", "sheet title 'N3AB' / 'E10-14'"),
    ("N3AB/N3AB - E40,42 (Done - no pressure point to be matched)/N3AB - E40_42.xlsx", "Prod", "B3"): ("N3AB", "E40/42", "folder/file name (sheet title reads 'S1A - LD1')"),
    ("N3AB/N3AB - E6,7,8 (Done)/N3AB - E6_7_8.xlsx", "Basic Data", "A4"): ("N3AB", "E6/7/8", "sheet title 'N3AB' / 'E6/7/8'"),
    ("N3AB/N3AB - E6,7,8 (Done)/N3AB - E6_7_8.xlsx", "Sheet2", "A1"): ("N3AB", "E6/7/8", "file name"),
    ("N3B/N3B LD1-LD3 basic data.xlsx", "Sheet1", "D6"): ("N3B", "LD1-LD3", "file name"),
    ("N3BDEFSE1/E3436.xlsx", "Sheet1", "D6"): ("N3BDEFSE1", "E34/36", "folder name / file name 'E3436'"),
    ("N3C/N3C LD1-LD3 basic data.xlsx", "Sheet1", "D6"): ("N3C", "LD1-LD3", "file name"),
    ("N3C/N3cE10-14 basic data.xlsx", "Sheet1", "D5"): ("N3C", "E10-14", "file name"),
    ("N3C/N3cE34-46 basic data.xlsx", "Sheet1", "D6"): ("N3C", "E34-46", "file name"),
    ("N3C/N3cE7 basic data.xlsx", "Sheet1", "D6"): ("N3C", "E7", "file name"),
    ("N3D/N3D - E10E14 (done)/N3D - E10E14.xlsx", "Prod", "B3"): ("N3D", "E10/14", "sheet title 'N3D E10/14'"),
    ("S1A/S1A - LD1 (done)/S1A - LD1.xlsx", "Prod", "B3"): ("S1A", "LD1", "sheet title 'S1A - LD1'"),
    ("S1B/S1B - E10,11,12,13 (Done)/S1B - E10_11_12_13.xlsx", "Prod", "B3"): ("S1B", "E10/13", "sheet title 'S1B' / 'E10/13'"),
    ("S1B/S1B - E67 (Done)/S1B - E67.xlsx", "Prod Data", "A3"): ("S1B", "E67", "sheet title 'S1b' / 'E67'"),
    ("S1B/S1B-LD1 (Done)/S1B -LD1.xlsx", "Sheet1", "A3"): ("S1B", "LD1", "sheet title 'LD1' / 'S1B'"),
    ("S1C/S1C - E10,13 (Done)/S1C - E10_13.xlsx", "Prod", "B3"): ("S1C", "E10/13", "sheet title 'S1C - E10/13'"),
    ("S1C/S1C E6-E7-E8.xlsx", "Sheet1", "D6"): ("S1C", "E6-E7-E8", "file name"),
    ("S1CD/S1CD LD1-E3 basic data.xlsx", "Sheet1", "D6"): ("S1CD", "LD1-E3", "file name (model file says 'ld1-e2')"),
    ("S1D/S1D E678 basic data.xlsx", "Sheet1", "D6"): ("S1D", "E678 (E6-E8 + E10-E11)", "file name; values equal split E6-E8 + E10-E11"),
    ("S1D/S1D E678 basic data.xlsx", "split", "C5"): ("S1D", "E6-E8", "block title 'E6-E8'"),
    ("S1D/S1D E678 basic data.xlsx", "split", "U5"): ("S1D", "E10-E11", "block title 'E10-E11'"),
    ("S2/S2 - E1 (Done)/S2 - E1.xlsx", "Prod", "B3"): ("S2", "E1", "sheet title 'S2 - E1'"),
    ("S2/S2 - E7 (Done)/S2 - E7.xlsx", "Data", "A3"): ("S2", "E7", "sheet title 'E7'"),
    ("S2/S2-E34-36-42-46/S2 E34-36-40-42.xlsx", "Sheet1", "D6"): ("S2", "E34-36-40-42", "file name (folder says E34-36-42-46)"),
    ("S2/S2-E8 to E14/S2 E10-E14 Basic Data.xlsx", "Sheet1", "C4"): ("S2", "E10-E14", "file name (folder says E8 to E14)"),
    ("S3/S3 - E12,14 (Done)/S3 - E12_14.xlsx", "Prod", "B3"): ("S3", "E12/14", "sheet title 'S3 - E12/14'"),
    ("S3/S3 - E7 (Done)/S3 - E7.xlsx", "Basic Data", "A3"): ("S3", "E7", "sheet title 'E7'"),
    ("SE1/SE1 - E14/SE1-E14.xlsx", "Sheet1", "D6"): ("SE1", "E14", "file name"),
    ("SE1/SE1 - E14/SE1-E14.xlsx", "split", "C5"): ("SE1", "E6-E8", "block title 'E6-E8' (folder SE1; identical to S1D split)"),
    ("SE1/SE1 - E14/SE1-E14.xlsx", "split", "U5"): ("SE1", "E10-E11", "block title 'E10-E11' (folder SE1; identical to S1D split)"),
    ("SE1/SE1 - E7 (Done)/SE1 - E7.xlsx", "Prod", "B3"): ("SE1", "E7", "sheet title 'Se1 - e7'"),
    ("SE1/SE1 - LD3, E1, E2 (Done)/SE1 - LD3_E1_E2.xlsx", "Prod", "B3"): ("SE1", "LD3/E1/E2", "folder name (sheet title reads 'LD3/E2')"),
    ("SE2/E12-E13/SE2-E12-E13.xlsx", "Sheet1", "D6"): ("SE2", "E12/E13", "folder / file name"),
    ("SE2/E12-E13/SE2-E12-E13.xlsx", "Sheet2", "D5"): ("SE2", "E12/E13", "block title 'E12/E13' ('Combine with SE1')"),
    ("SE2/E12-E13/SE2-E12-E13.xlsx", "Sheet2", "V5"): ("S1", "E12/E13", "block title 'S1 E12/E13' as written (possibly SE1)"),
    ("SE2/E20-E23/SE2 E20-23 basic data.xlsx", "Sheet1", "D6"): ("SE2", "E20-E23", "folder / file name"),
}
LABELS_WEST = {
    ("IX/IX - E1,2 (Done)/IX - E1_2.xlsx", "Prod", "B3"): ("IX", "E1/2", "folder / file name (sheet title reads 'S1A - LD1')"),
    ("IX/IX - E6,7 (Done)/IX - E6_7.xlsx", "Prod", "B3"): ("IX", "E6/7", "sheet title 'IX - E6/7'"),
    ("VIII/VIII - E1,2 (Done)/VIII - E1_2.xlsx", "Prod", "B3"): ("VIII", "E1/E2", "sheet title 'VIII - E1/E2'"),
    ("VIII/VIII - E10,11 (Done)/VIII - E10_11.xlsx", "Prod", "B3"): ("VIII", "E10/11", "folder / file name (sheet title reads 'S1A - LD1')"),
}
LABELS = LABELS_UNIT if UNIT else LABELS_WEST
COLS = ["FIELD", "PLATFORM", "WELL", "STRING TYPE", "FAULT BLOCK", "RESERVOIR", "DATE", "PRESSURE SURVEY  ", "CUM_OIL BBL ",
        "CUM_GAS CF ", "CUM_WATER BBL ", "CUM_GAS_INJ MCF ", "CUM_WTR.INJ BBL "]
EXTRA = ["RECORD TYPE", "SOURCE FILE", "SOURCE SHEET", "SOURCE ROW", "LABEL BASIS", "DUPLICATE OF"]
WELL_FILE = "N2/N2 - E1011, E1213, E14 (done)/Data by pressure (B1 and B9).xlsx"
INJDAT_FILE = "N2/N2 - E1011, E1213, E14 (done)/working file/Data N2 - E1011_ E1213_ E14 (Single tank).xlsx"


def well_level():
    """Kind C: monthly rows -> running cumulative per (field, well, string, FB, sand)."""
    groups, notes = [], []
    src = {WELL_FILE: {s: rows for s, st, rows in sheets(read(NAS / WELL_FILE)) if s in ("inj", "B9")},
           INJDAT_FILE: {s: rows for s, st, rows in sheets(read(NAS / INJDAT_FILE)) if s == "InjDat(ori)"}}
    specs = [  # file, sheet, key columns (field, well, string, fb, sand), date col, oil, gas, water, gasinj, wtrinj
        (WELL_FILE, "inj", (0, 1, 2, 3, 4), 5, None, None, None, 6, 7),
        (WELL_FILE, "B9", (0, 1, 2, 3, 4), 5, 6, 7, 8, 9, 10),
        (INJDAT_FILE, "InjDat(ori)", (None, None, None, 0, 1), 2, None, None, None, 5, 6),
    ]
    for f, s, keyc, dc, oc, gc, wc, gic, wic in specs:
        rows = src[f][s]
        hdr = [str(x).strip() if x is not None else "" for x in rows[0]]
        acc = defaultdict(lambda: [0.0, 0.0, 0.0, 0.0, 0.0])
        order = []
        for ii, r in enumerate(rows[1:], 2):
            r = list(r) + [None] * 12
            d = as_date(r[dc])
            if d is None:
                if any(x not in (None, "") for x in r):
                    notes.append((f, s, ii, "row without date (column total / note) — not used", [x for x in r if x not in (None, "")][:4]))
                continue
            key = tuple((str(r[k]).strip() if r[k] not in (None, "") else "") if k is not None else "" for k in keyc)
            vals = [num(r[c]) if c is not None else None for c in (oc, gc, wc, gic, wic)]
            a = acc[key]
            cum = []
            for k, x in enumerate(vals):
                if (oc, gc, wc, gic, wic)[k] is None:
                    cum.append(None)
                else:
                    a[k] += x or 0.0
                    cum.append(a[k])
            order.append(dict(key=key, date=d, row=ii, cum=cum, sheet=s, file=f))
        groups.append((f, s, hdr, order))
    return groups, notes


def sha(b):
    return hashlib.sha256(b).hexdigest()


def build():
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    t_hash = sha(TPL.read_bytes())
    blocks, skipped = inventory()
    src_hash = {b["file"]: sha(read(NAS / b["file"])) for b in blocks}
    if UNIT:
        src_hash[WELL_FILE] = sha(read(NAS / WELL_FILE))
        src_hash[INJDAT_FILE] = sha(read(NAS / INJDAT_FILE))
    missing = [(b["file"], b["sheet"], b["header_cell"]) for b in blocks if (b["file"], b["sheet"], b["header_cell"]) not in LABELS]
    assert not missing, missing
    # duplicates: identical data series
    sig = {}
    for b in blocks:
        b["sig"] = tuple((r[1].date(),) + tuple(r[2:]) for r in b["data"])
        b["id"] = f"{b['file'].split('/')[-1]} [{b['sheet']}] {b['header_cell']}"
        b["dup"] = sig.get(b["sig"], "")
        sig.setdefault(b["sig"], b["id"])
    for b in blocks:   # mark the first copy too
        if not b["dup"] and sum(1 for x in blocks if x["sig"] == b["sig"]) > 1:
            b["dup"] = "; ".join(x["id"] for x in blocks if x["sig"] == b["sig"] and x is not b)

    out_rows = []
    for b in blocks:
        fb, res, basis = LABELS[(b["file"], b["sheet"], b["header_cell"])]
        for (srow, d, p, co, cg, cw, cgi, cwi) in b["data"]:
            out_rows.append(["DULG", None, None, None, fb, res, d, p, co, cg, cw, cgi, cwi,
                             "Tank cumulative (MBAL format, as reported)", b["file"], b["sheet"], srow, basis, b["dup"]])
    groups, wnotes = well_level() if UNIT else ([], [])
    wl_summary = []
    for f, s, hdr, order in groups:
        for o in order:
            field, well, string, fb, sand = o["key"]
            co, cg, cw, cgi, cwi = o["cum"]
            out_rows.append([field or "DULG", None, well or None, string or None, fb or None, sand or None, o["date"], None, co, cg, cw, cgi, cwi,
                             "Well cumulative (DERIVED: running sum of monthly values)", f, s, o["row"],
                             "row columns Field/Well/String/FB/Sand" if s != "InjDat(ori)" else "row columns FB/Sand (no well column)",
                             "same injection volumes as 'inj' sheet (total 802,680 bbl)" if s == "InjDat(ori)" else ""])
        per = Counter(o["key"] for o in order)
        for k, n in per.items():
            last = [o for o in order if o["key"] == k][-1]
            wl_summary.append((f, s, " / ".join(x or "—" for x in k), n, last["date"], last["cum"][4]))

    # ---------------- populated template (copy; header untouched; extra columns after M)
    OUT.mkdir(exist_ok=True)
    out_path = OUT / f"{STEM}.xlsx"
    k = 2
    while out_path.exists():
        out_path = OUT / f"{STEM}_v{k}.xlsx"
        k += 1
    wb = load_workbook(TPL)
    ws = wb["Sheet1"]
    assert [c.value for c in ws[1]][:13] == COLS
    hstyle = ws["M1"]
    for i, h in enumerate(EXTRA, 14):
        c = ws.cell(1, i, h)
        c.font, c.fill, c.alignment, c.border = copy(hstyle.font), copy(hstyle.fill), copy(hstyle.alignment), copy(hstyle.border)
    for r in out_rows:
        ws.append(r)
    for row in ws.iter_rows(min_row=2, min_col=7, max_col=7):
        row[0].number_format = "yyyy-mm-dd"
    for i, w in enumerate([8, 9, 10, 8, 12, 22, 11, 12, 13, 15, 13, 13, 14, 44, 70, 20, 10, 46, 50], 1):
        ws.column_dimensions[get_column_letter(i)].width = max(ws.column_dimensions[get_column_letter(i)].width or 0, w)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(13 + len(EXTRA))}{ws.max_row}"
    wb.save(out_path)

    # ---------------- validation
    val = []
    chk = load_workbook(out_path, read_only=True)["Sheet1"]
    got = list(chk.iter_rows(values_only=True))
    val.append(("Template header A1:M1 unchanged; extra columns after M", got[0][:13] == tuple(COLS) and list(got[0][13:]) == EXTRA))
    val.append((f"Rows written = rows extracted ({len(out_rows)})", len(got) - 1 == len(out_rows)))
    ok = all((g[6].date() if isinstance(g[6], dt.datetime) else g[6]) == (e[6].date()) and
             all((a is None and b is None) or (a is not None and b is not None and abs(float(a) - float(b)) <= 1e-9 * max(1.0, abs(float(b)))) for a, b in zip(g[7:13], e[7:13]))
             for g, e in zip(got[1:], out_rows))
    val.append(("Every written date and value equals the extracted source value", ok))
    # re-read sources and compare raw blocks independently
    reb, _ = inventory()
    val.append(("Re-read of sources reproduces every raw block (no omissions/changes)",
                [tuple(x[1:] for x in b["data"]) for b in reb] == [tuple(x[1:] for x in b["data"]) for b in blocks]))
    for f, s, hdr, order in groups:
        rows = {n: r for n, st, r in sheets(read(NAS / f))}[s]
        wic = {"inj": 7, "B9": 10, "InjDat(ori)": 6}[s]
        dc = {"inj": 5, "B9": 5, "InjDat(ori)": 2}[s]
        tot = sum(num(r[wic]) or 0 for r in rows[1:] if as_date(r[dc]) is not None)
        cum_end = sum(max((o["cum"][4] or 0) for o in order if o["key"] == k) for k in {o["key"] for o in order})
        val.append((f"Well-level [{s}]: sum of final cumulatives = sum of monthly WTRINJ ({tot:,.0f} bbl)", abs(tot - cum_end) < 1e-6))
    val.append((f"All {len(LABELS)} raw blocks labelled (fault block + reservoir)", not missing and len(blocks) == len(LABELS)))
    val.append(("Template unchanged", sha(TPL.read_bytes()) == t_hash))
    val.append(("All source workbooks unchanged (sha256 before/after)", all(sha(read(NAS / f)) == h for f, h in src_hash.items())))

    # ---------------- review workbook
    rev = OUT / out_path.name.replace(STEM, STEM + "_Review")
    scan = json.loads((OUT / f"scan_results{TAG}.json").read_text())
    rw = Workbook()
    s0 = rw.active
    s0.title = "Summary"
    nz = [b for b in blocks if max((r[7] or 0) for r in b["data"]) > 0]
    info = [("Output", str(out_path)), ("Template (read-only)", f"{TPL}  sha256 {t_hash[:16]}…"),
            ("Source folder (read-only)", str(NAS)),
            ("Scanned", f"{len({r['file'] for r in scan})} workbooks, {sum(1 for r in scan if r.get('sheet'))} worksheets; {sum(1 for r in scan if r.get('n_hits'))} worksheets matched injection terms"),
            ("Extracted", f"{len(blocks)} raw MBAL tank blocks ({len(nz)} with non-zero water injection) = {sum(len(b['data']) for b in blocks)} rows; "
                          f"well-level running cumulatives = {sum(len(o) for *_, o in groups)} rows; total {len(out_rows)} rows"),
            ("Decisions", "Raw MBAL blocks only; well-level monthly -> running cumulative (derived, flagged); zero-injection blocks included; copies kept and flagged"),
            ("FIELD", "'DULG' — as written in the well-level rows; tank blocks take it from the 'Dulang Unit' folder"),
            ("PLATFORM", "Not stated in any source — blank"),
            ("Units", "As reported: oil/water/water-inj bbl, gas cf, gas-inj Mcf — identical to template; no conversion"),
            ("Processed", dt.datetime.now().strftime("%Y-%m-%d %H:%M") + " — locally")]
    for i, (a, b2) in enumerate(info, 1):
        s0.cell(i, 1, a).font = Font(bold=True)
        s0.cell(i, 2, b2)
    s0.column_dimensions["A"].width, s0.column_dimensions["B"].width = 24, 150
    blk = [(b["file"], b["sheet"], b["header_cell"], *LABELS[(b["file"], b["sheet"], b["header_cell"])], len(b["data"]),
            b["data"][0][1].date(), b["data"][-1][1].date(), max((r[7] or 0) for r in b["data"]), " | ".join(b["header"]), b["dup"]) for b in blocks]
    tables = [
        ("Extracted Blocks", ["Source File", "Sheet", "Header Cell", "FAULT BLOCK", "RESERVOIR", "Label Basis", "Rows", "First Date", "Last Date",
                              "Max CUM_WTR.INJ BBL", "Source Header", "Duplicate Of"], blk, [62, 16, 8, 10, 22, 46, 6, 11, 11, 16, 90, 60]),
        ("Well-level Groups", ["Source File", "Sheet", "Field / Well / String / FB / Sand", "Rows", "Last Date", "Final CUM_WTR.INJ BBL"], wl_summary,
         [62, 12, 40, 6, 12, 16]),
        ("Not Extracted", ["Source File", "Sheet", "Header Cell", "Reason"], skipped, [70, 26, 8, 70]),
        ("Matched Sheets", ["Source File", "Sheet", "Visibility", "Rows", "Cols", "Matching cells (first)"],
         [(r["file"], r["sheet"], r["state"], r["rows"], r["cols"], "; ".join(f"{a}={b}" for a, b in r["hits"][:6])) for r in scan if r.get("n_hits")],
         [70, 26, 9, 8, 6, 100]),
        ("Flags", ["Item", "Detail"], ([
            ("N3AB - E40_42.xlsx [Prod]", "Sheet title reads 'S1A - LD1' (copied layout); labelled N3AB E40/42 from folder/file name."),
            ("SE1-E14.xlsx [split]", "Both blocks are identical to S1D E678 basic data.xlsx [split] (E6-E8, E10-E11) — kept, fault block from folder SE1, flagged duplicate."),
            ("S1D E678 basic data.xlsx [Sheet1]", "Values equal the sum of the split E6-E8 + E10-E11 blocks (final Wi 22,326,485 = 16,813,092 + 5,513,393)."),
            ("SE2-E12-E13.xlsx [Sheet2] V5", "Block titled 'S1 E12/E13' — fault block written 'S1' as titled (may mean SE1)."),
            ("SE1 - LD3_E1_E2.xlsx [Prod]", "Sheet title 'LD3/E2'; folder says LD3, E1, E2 — folder used."),
            ("Well-level rows", "CUM_* values are running sums of monthly OIL/GAS/WATER/GASINJ/WTRINJ per group, starting at the first row in that sheet (not total since first production if earlier months are absent)."),
            ("B9 rows without well", "43 rows give FIELD 'DULG' + WTRINJ only (146,749 bbl total) — kept as their own group with WELL/STRING/FB/RESERVOIR blank."),
            ("InjDat(ori)", "No well column; several rows per month per FB/Sand — cumulative is per FB/Sand; same total as 'inj' sheet (802,680 bbl) — flagged copy."),
            ("Kind C sheets not used", "Prod&Inj, E1011, E1213, E14 (tank-level monthly + derived MM cumulatives) — the raw cumulative for N2 E10/14 is taken from the Cum Ali / 1014 (ali) blocks."),
            ("Rows with date only", "Daily/monthly date rows with no values (e.g. N3AB E10-14, N2 1014 (ali)) are not written."),
        ] if UNIT else [
            ("IX - E1_2.xlsx [Prod] / VIII - E10_11.xlsx [Prod]", "Sheet title reads 'S1A - LD1' (copied layout); labelled from folder / file name."),
            ("Mod / Mod by Adam (Not applicable)", "Revised or MM-scaled copies of the Prod data — not extracted (raw MBAL blocks only)."),
            ("WESTERN IX E10-14.xlsx", "No injection terms in any worksheet — not extracted."),
            ("Rows with date only", "Date rows with no values are not written."),
        ]) + [(f"{a} [{b}] row {c}", f"{d}: {e}") for a, b, c, d, e in wnotes], [50, 130]),
        ("Validation", ["Check", "Result"], [(a, "PASS" if b else "FAIL") for a, b in val], [100, 10]),
    ]
    for title, hdr, data, widths in tables:
        sh = rw.create_sheet(title)
        sh.append(hdr)
        for r in data:
            sh.append(list(r))
        for c in sh[1]:
            c.font, c.fill = Font(bold=True), PatternFill("solid", fgColor="DDEBF7")
        for i, w in enumerate(widths, 1):
            sh.column_dimensions[get_column_letter(i)].width = w
        sh.freeze_panes = "A2"
    rw.save(rev)
    for a, b in val:
        print("PASS" if b else "FAIL", "-", a)
    print(out_path.name, rev.name, "rows", len(out_rows))
    by = Counter((r[4] or "—") for r in out_rows)
    print(dict(sorted(by.items())))


if __name__ == "__main__":
    build()
