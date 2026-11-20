#!/usr/bin/env python3
"""Populate a copy of the Dulang bubble-map template from the FPC production allocation (29 Oct 2026).

INPUT 1  input_FPC_Prod Allocation_11Jun126.xlsx  (Sheet1, headers rows 1-2)                     read-only
INPUT 2  input_Well Header Info & Data Availability Listing.xlsx  ('Availability Listing_USER')  read-only
TEMPLATE bubble_map_dulang_sample/bubble_map_template.xlsx  ('Bubble Map Sample', 15 columns)      read-only
OUTPUT   bubble_map_dulang_sample/bubble_map_dulang_populated.xlsx (numbered version if it exists)

User decisions (29 Oct 2026):
  * layout: keep the template's 15 columns unchanged, append the extra attributes to the right
  * well matching: naming rule DULG-<P><NNN>[S<NN>][W01|H01]-<LS|SS|TS> -> DULANG-<P><N>[S<n>|ST<n>];
    -LS/-SS/-TS strings and W01/H01-tagged strings resolve to their parent wellbore
"""
import calendar
import collections
import copy
import datetime as dt
import hashlib
import math
import re
import shutil
from pathlib import Path

import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

BASE = Path(__file__).resolve().parent.parent
PROD = BASE / "input_FPC_Prod Allocation_11Jun126.xlsx"
HDR = BASE / "input_Well Header Info & Data Availability Listing.xlsx"
TEMPLATE = BASE / "bubble_map_dulang_sample" / "bubble_map_template.xlsx"
OUT = BASE / "bubble_map_dulang_sample" / "bubble_map_dulang_populated.xlsx"
TEMPLATE_SHEET = "Bubble Map Sample"
PROD_SHEET, HDR_SHEET = "Sheet1", "Availability Listing_USER"

TEMPLATE_HEADERS = ["FIELD NAME", "WELL ID", "WELL NAME", "LAT", "LONG", "X UTM", "Y UTM", "isDummyCoord", "DATE",
                    "WATER CUT, WC (%)", "CUMULATIVE OIL PRODUCTION, NP", "GOR", "NPRADIUS",
                    "CUMULATIVE WATER PRODUCTION, WP", "PRODUCTIVITY"]
APPENDED = ["ORGINAL WELL NAME", "RESERVOIR UNIT", "Alloc. Month", "DAYS", "Producing Days", "Cumulative PD",
            "OIL CUT (%)", "Oil PD Rate, Np (bbls/d)", "Water PD Rate, Wp (bbls/d)", "Pay Thickness, h (ft)",
            "Porosity, F (fraction)", "Water Saturation, Sw (fraction)", "Hydrocarbon Saturation, Shc (fraction)",
            "Bo (rb/stb)", "SOURCE ROW (INPUT 1 Sheet1)", "QC FLAGS"]
HEADERS = TEMPLATE_HEADERS + APPENDED
COL = {h: get_column_letter(i) for i, h in enumerate(HEADERS, 1)}

# INPUT 1 columns (0-based) — verified against the row-2 headers below
P = dict(region=0, field=1, string=2, perf=3, ru=4, block=5, res=6, block_res=7, month=8, pd=9,
         oil_vol=12, gas_vol=13, wat_vol=15, oil_pd=19, wat_pd=22, wc=33, gor=35)
P_HEAD = {1: "Field", 2: "String", 3: "Perforation", 4: "Reservoir Unit", 12: "Alloc. Oil Volume(bbl)",
          13: "Alloc. Gas Volume(scf)", 15: "Alloc. Water Volume(bbl)", 19: "Oil PD Rate (bbls/d)",
          22: "Water PD Rate(bbls/d)", 33: "WC (%)", 35: "GOR (scf/bbl)"}
P_HEAD1 = {8: "Alloc. Month", 9: "Producing Days"}
NAME_RE = re.compile(r"DULG-([A-D])(\d{3})(?:S(\d{2}))?((?:W|H)\d{2})?-(LS|SS|TS)")


def sha(p):
    return hashlib.sha256(p.read_bytes()).hexdigest()


def num(x):
    return isinstance(x, (int, float)) and not isinstance(x, bool)


def out_path():
    if not OUT.exists():
        return OUT
    k = 2
    while (p := OUT.with_name(f"{OUT.stem}_v{k}{OUT.suffix}")).exists():
        k += 1
    return p


def read_header():
    ws = openpyxl.load_workbook(HDR, read_only=True, data_only=True)[HDR_SHEET]
    head = next(ws.iter_rows(min_row=2, max_row=2, values_only=True))
    assert (head[2], head[16], head[17]) == ("Wellbore Name", "Latitude WGS84", "Longitude WGS84"), head
    wells = collections.defaultdict(list)
    for i, r in enumerate(ws.iter_rows(min_row=3, values_only=True), 3):
        if r[2]:
            wells[r[2].strip().upper()].append(dict(row=i, name=r[2], lat=r[16], lon=r[17], remark=r[22]))
    return wells


def match(string, wells):
    """-> (header record or None, header name, method, status)"""
    exact = wells.get(string.strip().upper())
    if exact:
        return (exact[0] if len(exact) == 1 else None, exact[0]["name"], "exact (trim/case)",
                "MATCHED" if len(exact) == 1 else "MULTIPLE HEADER MATCHES")
    m = NAME_RE.fullmatch(string)
    if not m:
        return None, "", "no naming-rule pattern", "UNMATCHED"
    plat, n, side, tag, _ = m.groups()
    base = f"DULANG-{plat}{int(n)}"
    cands = [base] if not side else [f"{base}S{int(side)}", f"{base}ST{int(side)}"]
    hits = [c for c in cands if c in wells]
    method = (f"user-approved naming rule: DULG->DULANG, well no. {n}->{int(n)}"
              + (f", S{side}->S{int(side)}/ST{int(side)}" if side else "")
              + (f", {tag} tag -> parent wellbore" if tag else "") + ", string suffix -> same well")
    if len(hits) != 1:
        return None, " | ".join(hits) or " | ".join(cands), method, "MULTIPLE HEADER MATCHES" if hits else "UNMATCHED"
    recs = wells[hits[0]]
    if len(recs) != 1:
        return None, hits[0], method, "MULTIPLE HEADER MATCHES"
    return recs[0], hits[0], method, "MATCHED"


def main():
    before = {p: sha(p) for p in (PROD, HDR, TEMPLATE)}
    dst = out_path()

    # ---------------- INPUT 1
    ws = openpyxl.load_workbook(PROD, read_only=True, data_only=True)[PROD_SHEET]
    h1, h2 = (next(ws.iter_rows(min_row=k, max_row=k, values_only=True)) for k in (1, 2))
    for i, t in P_HEAD.items():
        assert h2[i] == t, (i, h2[i], t)
    for i, t in P_HEAD1.items():
        assert h1[i] == t, (i, h1[i], t)
    recs = [(i, r) for i, r in enumerate(ws.iter_rows(min_row=3, values_only=True), 3) if any(v is not None for v in r)]
    key = lambda r: (r[P["field"]], r[P["string"]], r[P["perf"]], r[P["ru"]])
    dups = [k for k, v in collections.Counter(key(r) + (r[P["month"]],) for _, r in recs).items() if v > 1]
    recs.sort(key=lambda x: key(x[1]) + (x[1][P["month"]], x[0]))

    # ---------------- INPUT 2 + matching
    wells = read_header()
    per_string = {}
    for s in sorted({r[P["string"]] for _, r in recs}):
        rec, hname, method, status = match(s, wells)
        coord = None
        if rec is not None:
            lat, lon = rec["lat"], rec["lon"]
            if num(lat) and num(lon):
                assert -90 <= lat <= 90 and -180 <= lon <= 180 and 0 < lat < 10 and 100 < lon < 110, (s, lat, lon)
                coord = (lat, lon)
            else:
                status = "MATCHED - HEADER HAS NO COORDINATES"
        per_string[s] = dict(rec=rec, hname=hname, method=method, status=status, coord=coord, n=0)

    # ---------------- output workbook from a copy of the template
    shutil.copy2(TEMPLATE, dst)
    wb = openpyxl.load_workbook(dst)
    out = wb[TEMPLATE_SHEET]
    assert [c.value for c in out[1]][:15] == TEMPLATE_HEADERS
    for row in out.iter_rows(min_row=2):          # remove any sample data (template currently has none)
        for c in row:
            c.value = None
    hstyle = out["A1"]
    for j, h in enumerate(APPENDED, len(TEMPLATE_HEADERS) + 1):
        c = out.cell(1, j, h)
        c.font, c.fill, c.border, c.alignment = (copy.copy(hstyle.font), copy.copy(hstyle.fill),
                                                 copy.copy(hstyle.border), copy.copy(hstyle.alignment))
        c.number_format = hstyle.number_format
        out.column_dimensions[get_column_letter(j)].width = max(12, min(40, len(h) + 2))
    for h in ("FIELD NAME", "WELL NAME", "ORGINAL WELL NAME", "QC FLAGS"):
        out.column_dimensions[COL[h]].width = {"QC FLAGS": 60, "ORGINAL WELL NAME": 30}.get(h, 18)

    fmt = {"DATE": "yyyy-mm", "Alloc. Month": "yyyy-mm", "LAT": "0.0000000", "LONG": "0.0000000",
           "WATER CUT, WC (%)": "0.00", "OIL CUT (%)": "0.00", "GOR": "0.00", "Producing Days": "0.##",
           "Cumulative PD": "0.##", "Oil PD Rate, Np (bbls/d)": "0.00", "Water PD Rate, Wp (bbls/d)": "0.00"}
    c_ = COL
    flag_counts = collections.Counter()
    issue_rows = collections.defaultdict(list)
    ent = collections.OrderedDict()
    prev_key, blank_seen = None, 0
    for n, (src, r) in enumerate(recs, 2):
        k = key(r)
        if k != prev_key:
            prev_key, blank_seen = k, 0
        month, pdays = r[P["month"]], r[P["pd"]]
        days = calendar.monthrange(month.year, month.month)[1]
        info = per_string[r[P["string"]]]
        info["n"] += 1
        flags = []
        if pdays is None:
            blank_seen += 1
            flags.append("PD_BLANK")
        elif not 0 <= pdays <= days:
            flags.append(f"PD_OUTSIDE_0_{days}")
            issue_rows["PD outside 0..DAYS"].append((src, r[P["perf"]], month, pdays, days))
        if blank_seen:
            flags.append(f"CUM_PD_EXCLUDES_{blank_seen}_BLANK_PD_MONTH(S)")
        for nm, vol, rate in (("OIL", r[P["oil_vol"]], r[P["oil_pd"]]), ("WATER", r[P["wat_vol"]], r[P["wat_pd"]])):
            if rate is None and num(vol) and vol != 0 and not (num(pdays) and pdays > 0):
                flags.append(f"{nm}_VOLUME_WITHOUT_PD_RATE_BLANK")
                issue_rows[f"{nm.title()} volume with zero/blank Producing Days"].append((src, r[P["perf"]], month, pdays, vol))
        wc, ov, wv = r[P["wc"]], r[P["oil_vol"]], r[P["wat_vol"]]
        if num(wc) and num(ov) and num(wv) and ov + wv > 0 and abs(wv / (ov + wv) * 100 - wc) >= 1e-6:
            flags.append("WC_DIFFERS_FROM_VOLUMES")
            issue_rows["WC differs from volumes"].append((src, r[P["perf"]], month, wc, wv / (ov + wv) * 100))
        if info["coord"] is None:
            flags.append("NO_COORD: " + info["status"])
        for f in flags:
            flag_counts[re.sub(r"\d+", "n", f.split(":")[0]) if not f.startswith("NO_COORD") else f] += 1

        e = ent.setdefault(k, dict(first=month, last=month, n=0, blank=0, months=[]))
        e["last"], e["n"], e["blank"] = month, e["n"] + 1, e["blank"] + (pdays is None)
        e["months"].append(month)

        lat_lon = info["coord"] or (None, None)
        row = {
            "FIELD NAME": r[P["field"]], "WELL NAME": r[P["string"]], "LAT": lat_lon[0], "LONG": lat_lon[1],
            "DATE": f"={c_['Alloc. Month']}{n}", "WATER CUT, WC (%)": wc, "GOR": r[P["gor"]],
            "ORGINAL WELL NAME": r[P["perf"]], "RESERVOIR UNIT": r[P["ru"]], "Alloc. Month": month,
            "DAYS": f"=DAY(EOMONTH({c_['Alloc. Month']}{n},0))", "Producing Days": pdays,
            "Cumulative PD": (f"=IF(AND({c_['FIELD NAME']}{n}={c_['FIELD NAME']}{n-1},{c_['WELL NAME']}{n}={c_['WELL NAME']}{n-1},"
                              f"{c_['ORGINAL WELL NAME']}{n}={c_['ORGINAL WELL NAME']}{n-1},"
                              f"{c_['RESERVOIR UNIT']}{n}={c_['RESERVOIR UNIT']}{n-1}),{c_['Cumulative PD']}{n-1},0)"
                              f"+N({c_['Producing Days']}{n})"),
            "OIL CUT (%)": f'=IF({c_["WATER CUT, WC (%)"]}{n}="","",100-{c_["WATER CUT, WC (%)"]}{n})',
            "Oil PD Rate, Np (bbls/d)": r[P["oil_pd"]], "Water PD Rate, Wp (bbls/d)": r[P["wat_pd"]],
            "Hydrocarbon Saturation, Shc (fraction)": (f'=IF({c_["Water Saturation, Sw (fraction)"]}{n}="","",'
                                                       f'1-{c_["Water Saturation, Sw (fraction)"]}{n})'),
            "SOURCE ROW (INPUT 1 Sheet1)": src, "QC FLAGS": "; ".join(flags) or None,
        }
        for h, v in row.items():
            if v is not None:
                cell = out[f"{c_[h]}{n}"]
                cell.value = v
                if h in fmt:
                    cell.number_format = fmt[h]
    last_row = len(recs) + 1
    out.auto_filter.ref = f"A1:{get_column_letter(len(HEADERS))}{last_row}"
    out.freeze_panes = "A2"

    write_review(wb, dst, recs, per_string, wells, ent, dups, flag_counts, issue_rows, before, last_row)
    wb.save(dst)
    verify(dst, recs, per_string, ent, before)


def write_review(wb, dst, recs, per_string, wells, ent, dups, flag_counts, issue_rows, before, last_row):
    ws = wb.create_sheet("Mapping_Review")
    bold = Font(bold=True)
    rowp = [1]

    def section(title, head, rows):
        ws.cell(rowp[0], 1, title).font = Font(bold=True, size=12)
        rowp[0] += 1
        for j, h in enumerate(head, 1):
            ws.cell(rowp[0], j, h).font = bold
        rowp[0] += 1
        for r in rows:
            for j, v in enumerate(r, 1):
                c = ws.cell(rowp[0], j, v)
                if isinstance(v, dt.datetime):
                    c.number_format = "yyyy-mm"
            rowp[0] += 1
        rowp[0] += 1

    matched = [s for s, i in per_string.items() if i["coord"]]
    nrec_coord = sum(per_string[r[P["string"]]]["coord"] is not None for _, r in recs)
    months = sorted({r[P["month"]] for _, r in recs})
    section("1. SUMMARY", ["Item", "Value"], [
        ("Run date", "2026-10-29"),
        ("Output workbook", str(dst)),
        ("INPUT 1 (read-only)", f"{PROD} > sheet '{PROD_SHEET}' (header rows 1-2, data rows 3-{max(i for i, _ in recs)})"),
        ("INPUT 2 (read-only)", f"{HDR} > sheet '{HDR_SHEET}' (header row 2)"),
        ("Template (read-only)", f"{TEMPLATE} > sheet '{TEMPLATE_SHEET}' (15 headers; no sample data, no formulas)"),
        *[(f"SHA-256 before/after: {p.name}", before[p]) for p in before],
        ("Production records in INPUT 1", len(recs)),
        ("Records written to output", last_row - 1),
        ("Records with LAT/LONG", nrec_coord),
        ("Records without LAT/LONG", len(recs) - nrec_coord),
        ("Production strings (WELL NAME)", len(per_string)),
        ("Strings with unique header match + coordinates", len(matched)),
        ("Strings without coordinates", len(per_string) - len(matched)),
        ("Distinct header wells used", len({per_string[s]['hname'] for s in matched})),
        ("Production entities (Cumulative PD groups)", len(ent)),
        ("Allocation months", f"{months[0]:%Y-%m} to {months[-1]:%Y-%m} ({len(months)} distinct)"),
        ("Duplicate entity/month records", len(dups)),
        ("Output row order", "sorted by FIELD NAME, WELL NAME, ORGINAL WELL NAME, RESERVOIR UNIT, then Alloc. Month "
                             "(required for the running Cumulative PD); SOURCE ROW gives the INPUT 1 row"),
    ])

    S1 = f"INPUT 1 '{PROD_SHEET}'"
    section("2. ATTRIBUTE MAPPING", ["Output column", "Template attribute", "Source", "Rule / formula", "Notes"], [
        ("A", "FIELD NAME", f"{S1} col B 'Field'", "copied", "all records 'DULG'; Region col A = 'PMA' (not mapped)"),
        ("B", "WELL ID", "—", "left blank", "no defined source in the mapping table"),
        ("C", "WELL NAME", f"{S1} col C 'String'", "copied exactly", "e.g. DULG-A002-LS (string suffix kept)"),
        ("D", "LAT", f"INPUT 2 col Q 'Latitude WGS84'", "copied only after a unique well match",
         "decimal degrees as stored; datum WGS84 as labelled in the INPUT 2 header (Q2); no conversion"),
        ("E", "LONG", f"INPUT 2 col R 'Longitude WGS84'", "copied only after a unique well match",
         "decimal degrees as stored; datum WGS84 as labelled in the INPUT 2 header (R2); no conversion"),
        ("F", "X UTM", "—", "left blank", "INPUT 2 has no UTM easting; no CRS transformation performed"),
        ("G", "Y UTM", "—", "left blank", "INPUT 2 has no UTM northing; no CRS transformation performed"),
        ("H", "isDummyCoord", "—", "left blank", "no source"),
        ("I", "DATE", "= Alloc. Month (col R)", "=R<n>, format yyyy-mm", "template DATE column, same month as Alloc. Month"),
        ("J", "WATER CUT, WC (%)", f"{S1} col AH 'WC (%)'", "copied; 0-100 scale (source unit %)",
         "blank where the source is blank; 100 values differ from water/(oil+water) volumes (flagged, source kept)"),
        ("K", "CUMULATIVE OIL PRODUCTION, NP", "—", "left blank",
         "the mapping table gives Np as a producing-day RATE (col W); no cumulative-oil definition supplied"),
        ("L", "GOR", f"{S1} col AJ 'GOR (scf/bbl)'", "copied; unit scf/bbl, no conversion",
         "verified = Alloc. Gas Volume(scf) / Alloc. Oil Volume(bbl) for all populated values"),
        ("M", "NPRADIUS", "—", "left blank",
         "MISSING DEFINITION: template has no formula and no source formula exists; rate-vs-cumulative basis cannot be confirmed"),
        ("N", "CUMULATIVE WATER PRODUCTION, WP", "—", "left blank", "no definition supplied (Wp mapped as a rate, col X)"),
        ("O", "PRODUCTIVITY", "—", "left blank", "no source"),
        ("P", "ORGINAL WELL NAME", f"{S1} col D 'Perforation'", "copied exactly (template spelling kept)",
         "always equals String + '-' + Reservoir Unit"),
        ("Q", "RESERVOIR UNIT", f"{S1} col E 'Reservoir Unit'", "copied", ""),
        ("R", "Alloc. Month", f"{S1} col I 'Alloc. Month'", "Excel date (1st of month), format yyyy-mm", ""),
        ("S", "DAYS", "calculated", "=DAY(EOMONTH(R<n>,0))", "calendar days incl. leap years"),
        ("T", "Producing Days", f"{S1} col J 'Producing Days'", "copied; blanks kept blank (not 0)",
         "non-integer values kept as in source"),
        ("U", "Cumulative PD", "calculated",
         "=IF(AND(A<n>=A<n-1>,C<n>=C<n-1>,P<n>=P<n-1>,Q<n>=Q<n-1>),U<n-1>,0)+N(T<n>)",
         "running total per FIELD+WELL NAME+ORGINAL WELL NAME+RESERVOIR UNIT in month order, from the entity's first "
         "month in INPUT 1 (available-history total, not lifetime); blank PD months add nothing and are flagged "
         "CUM_PD_EXCLUDES_n_BLANK_PD_MONTH(S) on that and every later row of the entity"),
        ("V", "OIL CUT (%)", "calculated", '=IF(J<n>="","",100-J<n>)', "0-100 scale, matching WC (%)"),
        ("W", "Oil PD Rate, Np (bbls/d)", f"{S1} col T 'Oil PD Rate (bbls/d)'", "copied (producing-day rate)",
         "verified = Alloc. Oil Volume / Producing Days wherever populated; no rate had to be calculated; "
         "calendar-day rate (col AA) not used"),
        ("X", "Water PD Rate, Wp (bbls/d)", f"{S1} col W 'Water PD Rate(bbls/d)'", "copied (producing-day rate)",
         "verified = Alloc. Water Volume / Producing Days wherever populated; calendar-day rate (col AD) not used"),
        ("Y", "Pay Thickness, h (ft)", "—", "left blank", "not in either input"),
        ("Z", "Porosity, F (fraction)", "—", "left blank", "not in either input"),
        ("AA", "Water Saturation, Sw (fraction)", "—", "left blank", "not in either input; water cut NOT used as Sw"),
        ("AB", "Hydrocarbon Saturation, Shc (fraction)", "calculated", '=IF(AA<n>="","",1-AA<n>)',
         "returns blank because Sw is unavailable"),
        ("AC", "Bo (rb/stb)", "—", "left blank", "not in either input"),
        ("AD", "SOURCE ROW (INPUT 1 Sheet1)", "INPUT 1 row number", "added for traceability", ""),
        ("AE", "QC FLAGS", "derived", "record-level issues, '; '-separated", "see section 5"),
        ("—", "(not mapped)", f"{S1} cols F-H 'Block', 'Reservoir', 'Block Reservoir'", "not needed to keep records apart",
         "entity/month key is already unique; 59 perforations change Block/Reservoir over time"),
    ])

    rows = []
    for s, i in sorted(per_string.items()):
        rec = i["rec"]
        rows.append((s, i["hname"], i["method"], i["status"],
                     f"INPUT 2 '{HDR_SHEET}' row {rec['row']}" if rec else "",
                     f"Q{rec['row']}" if rec else "", f"R{rec['row']}" if rec else "",
                     i["coord"][0] if i["coord"] else None, i["coord"][1] if i["coord"] else None, i["n"],
                     (rec or {}).get("remark")))
    section("3. WELL MATCHING (one row per production string)",
            ["Production WELL NAME (String)", "Header Wellbore Name", "Matching method", "Status",
             "Header source row", "LAT cell", "LONG cell", "LAT", "LONG", "Records", "Header REMARKS"], rows)

    used = {i["hname"] for i in per_string.values() if i["rec"]}
    section("3b. HEADER WELLS WITH NO COORDINATES OR NOT USED", ["Header Wellbore Name", "Header row", "LAT", "LONG", "Used by production"],
            [(v[0]["name"], v[0]["row"], v[0]["lat"], v[0]["lon"], k in used) for k, v in sorted(wells.items())
             if not (num(v[0]["lat"]) and num(v[0]["lon"]))])

    erows = []
    for k, e in ent.items():
        exp = (e["last"].year - e["first"].year) * 12 + e["last"].month - e["first"].month + 1
        erows.append((*k, e["first"], e["last"], e["n"], exp - e["n"], e["blank"]))
    section("4. CUMULATIVE PD GROUPS (FIELD, WELL NAME, ORGINAL WELL NAME, RESERVOIR UNIT)",
            ["FIELD NAME", "WELL NAME", "ORGINAL WELL NAME", "RESERVOIR UNIT", "Start month", "End month", "Records",
             "Missing months inside range", "Blank Producing Days"], erows)

    section("5. QC FLAG COUNTS", ["Flag (in QC FLAGS column)", "Records"], sorted(flag_counts.items()) + [
        ("Duplicate entity/month records", len(dups))])
    heads = {"PD outside 0..DAYS": ["SOURCE ROW", "ORGINAL WELL NAME", "Alloc. Month", "Producing Days", "DAYS"],
             "WC differs from volumes": ["SOURCE ROW", "ORGINAL WELL NAME", "Alloc. Month", "Source WC (%)", "water/(oil+water) x 100"],
             "Oil volume with zero/blank Producing Days": ["SOURCE ROW", "ORGINAL WELL NAME", "Alloc. Month", "Producing Days", "Alloc. Oil Volume(bbl)"],
             "Water volume with zero/blank Producing Days": ["SOURCE ROW", "ORGINAL WELL NAME", "Alloc. Month", "Producing Days", "Alloc. Water Volume(bbl)"]}
    for t, h in heads.items():
        section(f"5. ISSUE LIST: {t} ({len(issue_rows[t])} records; source values kept, rate left blank where applicable)",
                h, issue_rows[t])
    ws.column_dimensions["A"].width = 34
    for col in "BCDE":
        ws.column_dimensions[col].width = 30
    ws.column_dimensions["C"].width = 60


def verify(dst, recs, per_string, ent, before):
    wb = openpyxl.load_workbook(dst, read_only=True)
    ws = wb[TEMPLATE_SHEET]
    it = ws.iter_rows(values_only=True)
    assert list(next(it)) == HEADERS
    src_by_row = dict(recs)
    seen, n, ncoord = set(), 0, 0
    prev, cum = None, 0
    ix = {h: i for i, h in enumerate(HEADERS)}
    for n, row in enumerate(it, 2):
        src = row[ix["SOURCE ROW (INPUT 1 Sheet1)"]]
        assert src not in seen
        seen.add(src)
        r = src_by_row[src]
        assert row[ix["WELL NAME"]] == r[P["string"]] and row[ix["ORGINAL WELL NAME"]] == r[P["perf"]]
        assert row[ix["RESERVOIR UNIT"]] == r[P["ru"]] and row[ix["FIELD NAME"]] == r[P["field"]]
        assert row[ix["Alloc. Month"]] == r[P["month"]]
        for h, k in (("Producing Days", "pd"), ("WATER CUT, WC (%)", "wc"), ("GOR", "gor"), ("Oil PD Rate, Np (bbls/d)", "oil_pd"),
                     ("Water PD Rate, Wp (bbls/d)", "wat_pd")):
            a, b = row[ix[h]], r[P[k]]
            assert a == b or (num(a) and num(b) and math.isclose(a, b, rel_tol=1e-14, abs_tol=1e-12)), (src, h, a, b)
        assert row[ix["DAYS"]] == f"=DAY(EOMONTH(R{n},0))" and row[ix["DATE"]] == f"=R{n}"
        assert row[ix["Cumulative PD"]].startswith(f"=IF(AND(A{n}=A{n-1},C{n}=C{n-1},P{n}=P{n-1},Q{n}=Q{n-1}),U{n-1},0)+N(T{n})")
        info = per_string[r[P["string"]]]
        lat, lon = row[ix["LAT"]], row[ix["LONG"]]
        if lat is not None or lon is not None:
            assert info["status"] == "MATCHED" and (lat, lon) == info["coord"] and 0 < lat < 10 and 100 < lon < 110
            ncoord += 1
        else:
            assert info["coord"] is None
        k = (r[P["field"]], r[P["string"]], r[P["perf"]], r[P["ru"]])
        if k == prev:
            assert r[P["month"]] > pm
        cum = (cum if k == prev else 0) + (r[P["pd"]] or 0)
        prev, pm = k, r[P["month"]]
        for h in ("X UTM", "Y UTM", "Pay Thickness, h (ft)", "Porosity, F (fraction)", "Water Saturation, Sw (fraction)", "Bo (rb/stb)", "NPRADIUS"):
            assert row[ix[h]] is None
        assert not any(isinstance(v, str) and "#REF" in v for v in row)
    assert len(seen) == len(recs) == n - 1
    assert "Mapping_Review" in wb.sheetnames
    after = {p: sha(p) for p in before}
    assert after == before, "an input/template file changed"
    print(f"OK: {dst}\n  records {len(recs)} (all accounted for, no duplicates/removals); with LAT/LONG {ncoord}")
    print(f"  strings {len(per_string)}: matched with coordinates {sum(1 for i in per_string.values() if i['coord'])}; "
          + str(collections.Counter(i['status'] for i in per_string.values())))
    print("  inputs and template unchanged (SHA-256)")


if __name__ == "__main__":
    main()
