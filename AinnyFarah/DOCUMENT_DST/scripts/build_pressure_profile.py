#!/usr/bin/env python3
"""Re-extract the Pressure Profile template only, from the Dulang 1992-2021 pressure-survey export.

Columns: template headers (Well, Date, Reservoir Pressure (psig)) + added 'Type', 'TVD ft (SS)'.
Rules
- Every survey record (row with WELL) gives one output row, in source order. No rows merged/deleted.
- Well  <- WELL (as-is); Date <- SURVEY DATE (date part; full timestamp kept in the log).
- Reservoir Pressure (psig) <- DATUM PRESS (psig) ONLY when TYPE = 'SGS'; blank for FGS/PBU.
- Type <- TYPE (as-is).
- TVD ft (SS) <- DATUM TVD (ft SS): the depth reference of the datum pressure.
- Blank source values stay blank; zeros are kept. No unit conversion (psig -> psig, ft SS -> ft SS).
Inputs are read-only and sha256-verified unchanged; existing output files are never overwritten.
"""
import hashlib
from collections import Counter, defaultdict
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BASE = Path("/Users/farahyusof/Library/CloudStorage/OneDrive-MRCHURROSF&BSDN.BHD/Ainny Farah/07_STE/13_DULANG")
SRC = BASE / "1992-2021_Pressure Survey Export_Dulang.xlsx"
TEMPLATE = BASE / "01_EDAFY_Asset_Templates" / "pressure_profile_template.xlsx"
OUT_DIR = BASE / "02_Mapped_Output"
OUT_DATA = OUT_DIR / "Dulang_Pressure_Profile_1992-2021.xlsx"
OUT_LOG = OUT_DIR / "Dulang_Pressure_Profile_Log.xlsx"
FIRST_DATA_ROW = 9
I = {"WELL": 0, "RESERVOIR": 1, "SUBUNIT": 2, "SURVEY DATE": 3, "TYPE": 4, "QUALITY": 5, "DATUM TVD": 15, "DATUM PRESS": 16}
EXTRA = ["Type", "TVD ft (SS)"]


def sha(p):
    return hashlib.sha256(p.read_bytes()).hexdigest()


def blank(x):
    return x is None or (isinstance(x, str) and not x.strip())


def head(ws, widths):
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="DDEBF7")
        c.alignment = Alignment(wrap_text=True, vertical="top")
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def main():
    for p in (OUT_DATA, OUT_LOG):
        assert not p.exists(), f"refusing to overwrite {p}"
    before = {p: sha(p) for p in (SRC, TEMPLATE)}

    tmpl_header = [c.value for c in load_workbook(TEMPLATE)["Pressure Profile"][1] if c.value not in (None, "")]
    assert tmpl_header == ["Well", "Date", "Reservoir Pressure (psig)"], tmpl_header
    header = tmpl_header + EXTRA

    src_ws = load_workbook(SRC, data_only=True)["SurveyData"]
    records = [(i, r) for i, r in enumerate(src_ws.iter_rows(min_row=FIRST_DATA_ROW, values_only=True), FIRST_DATA_ROW)
               if not blank(r[I["WELL"]])]

    data = Workbook()
    ws = data.active
    ws.title = "Pressure Profile"
    ws.append(header)
    trace, issues = [], []
    for src_row, r in records:
        g = lambda k: None if blank(r[I[k]]) else r[I[k]]
        typ, dt = g("TYPE"), g("SURVEY DATE")
        press = g("DATUM PRESS") if typ == "SGS" else None
        ws.append([g("WELL"), dt.date() if dt else None, press, typ, g("DATUM TVD")])
        out_row = ws.max_row
        ws.cell(out_row, 2).number_format = "yyyy-mm-dd"
        if typ == "SGS":
            rule = "SGS: Reservoir Pressure <- DATUM PRESS" if press is not None else "SGS but DATUM PRESS blank in source -> left blank"
        else:
            rule = f"{typ}: Reservoir Pressure left blank by rule (DATUM PRESS {g('DATUM PRESS')} not loaded)"
        trace.append([out_row, "SurveyData", src_row, g("WELL"), g("RESERVOIR"), g("SUBUNIT"), dt, typ, g("QUALITY"),
                      g("DATUM TVD"), g("DATUM PRESS"), press, rule])
        if typ == "SGS" and press is None:
            issues.append(["Missing value", out_row, src_row, g("WELL"), dt, "Reservoir Pressure (psig)",
                           "SGS record with no DATUM PRESS in source; row kept with pressure blank."])
        if typ == "SGS" and press is not None and g("DATUM TVD") is None:
            issues.append(["Missing depth", out_row, src_row, g("WELL"), dt, "TVD ft (SS)",
                           "Datum pressure present but DATUM TVD blank — depth reference unknown."])
        if g("QUALITY") == "BAD" and press is not None:
            issues.append(["Data quality", out_row, src_row, g("WELL"), dt, "Reservoir Pressure (psig)",
                           "Source QUALITY = BAD (no quality column in template); value loaded as-is."])
        if g("SUBUNIT") is None:
            issues.append(["Missing field", out_row, src_row, g("WELL"), dt, "SUBUNIT", "SUBUNIT blank in source."])

    # suspected duplicates: same Well + Date + Type (date part)
    grp = defaultdict(list)
    for t in trace:
        grp[(t[3], t[6].date(), t[7])].append(t)
    for (well, d, typ), ts in grp.items():
        if len(ts) > 1:
            pressures = {t[11] for t in ts if t[11] is not None}
            for t in ts:
                issues.append(["Suspected duplicate", t[0], t[2], well, t[6], "Well + Date + Type",
                               f"{len(ts)} {typ} rows for {well} on {d} (source rows {', '.join(str(x[2]) for x in ts)}; "
                               f"times {', '.join(x[6].strftime('%H:%M') for x in ts)}; subunits "
                               f"{', '.join(str(x[5]) for x in ts)}"
                               + (f"; differing reservoir pressures {sorted(pressures)}" if len(pressures) > 1 else "") + ")"])
    head(ws, [22, 12, 24, 8, 14])
    OUT_DIR.mkdir(exist_ok=True)
    data.save(OUT_DATA)

    log = Workbook()
    a = log.active
    a.title = "Attribute Mapping"
    a.append(["Target Column", "Source Sheet", "Source Column", "Source Unit", "Target Unit", "Rule / Transformation"])
    for row in [
        ["Well", "SurveyData", "WELL", "", "", "Direct copy, all records."],
        ["Date", "SurveyData", "SURVEY DATE", "date/time", "date", "Direct; date part written (YYYY-MM-DD). Full timestamp in Record Traceability."],
        ["Reservoir Pressure (psig)", "SurveyData", "DATUM PRESS", "psig", "psig",
         "Direct copy ONLY when TYPE = SGS. FGS and PBU rows left blank. No conversion."],
        ["Type (added column)", "SurveyData", "TYPE", "", "", "Direct copy (SGS / FGS / PBU)."],
        ["TVD ft (SS) (added column)", "SurveyData", "DATUM TVD", "ft (SS)", "ft (SS)",
         "Direct copy of the datum depth (the depth reference of DATUM PRESS), all records. "
         "Gauge TVD and mid-perf TVD not used."],
    ]:
        a.append(row)
    head(a, [28, 12, 16, 12, 12, 100])

    t = log.create_sheet("Record Traceability")
    t.append(["Output Row", "Source Sheet", "Source Row", "WELL", "RESERVOIR", "SUBUNIT", "SURVEY DATE (full)", "TYPE",
              "QUALITY", "DATUM TVD (ft SS)", "DATUM PRESS (psig, source)", "Reservoir Pressure loaded", "Rule Applied"])
    for x in trace:
        t.append(x)
    for row in t.iter_rows(min_row=2, min_col=7, max_col=7):
        row[0].number_format = "yyyy-mm-dd hh:mm"
    head(t, [10, 12, 10, 20, 14, 30, 18, 7, 9, 14, 16, 16, 60])

    s = log.create_sheet("Issues")
    s.append(["Issue Type", "Output Row", "Source Row", "WELL", "SURVEY DATE", "Column", "Detail"])
    for x in issues:
        s.append(x)
    for row in s.iter_rows(min_row=2, min_col=5, max_col=5):
        row[0].number_format = "yyyy-mm-dd hh:mm"
    head(s, [20, 10, 10, 20, 17, 26, 110])

    types = Counter(x[7] for x in trace)
    loaded = sum(1 for x in trace if x[11] is not None)
    m = log.create_sheet("Summary")
    for row in [["Item", "Count"], ["Source records (rows with WELL)", len(records)], ["Output rows", len(trace)],
                *[[f"   TYPE = {k}", v] for k, v in sorted(types.items())],
                ["Rows with Reservoir Pressure loaded (SGS with DATUM PRESS)", loaded],
                ["SGS rows with blank Reservoir Pressure (no DATUM PRESS in source)",
                 sum(1 for x in trace if x[7] == "SGS" and x[11] is None)],
                ["FGS/PBU rows (pressure blank by rule)", sum(1 for x in trace if x[7] != "SGS")],
                ["Rows with TVD ft (SS) populated", sum(1 for x in trace if x[9] is not None)],
                *[[f"Issues: {k}", v] for k, v in Counter(i[0] for i in issues).most_common()]]:
        m.append(row)
    head(m, [60, 10])
    log.save(OUT_LOG)

    # validate by reopening
    d = load_workbook(OUT_DATA)["Pressure Profile"]
    rows = list(d.iter_rows(values_only=True))
    assert list(rows[0]) == header
    assert len(rows) - 1 == len(records)
    assert all((r[2] is None) or r[3] == "SGS" for r in rows[1:])
    lt = list(load_workbook(OUT_LOG)["Record Traceability"].iter_rows(min_row=2, values_only=True))
    assert len(lt) == len(rows) - 1
    assert all(sha(p) == h for p, h in before.items()), "input changed"
    print(f"rows={len(rows) - 1} types={dict(types)} pressure_loaded={loaded} "
          f"tvd_filled={sum(1 for r in rows[1:] if r[4] is not None)} issues={dict(Counter(i[0] for i in issues))}")
    print("validated; inputs unchanged")
    print(OUT_DATA)
    print(OUT_LOG)


if __name__ == "__main__":
    main()
