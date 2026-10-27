#!/usr/bin/env python3
"""Re-extract the Field Production template only, from the Dulang 1992-2021 pressure-survey export.

Rules (as specified by the user)
- Applicable records: survey rows with at least one WELL TEST rate (LIQUID / OIL / WATER).
- Date           <- SURVEY DATE (direct; date part written, full timestamp in log)
- Gross (bfpd)   <- WELL TEST LIQUID RATE (bbl/d, direct)
- Oil Obs (bopd) <- WELL TEST OIL RATE (bbl/d, direct)
- Water (bwpd)   <- WELL TEST WATER RATE (bbl/d, direct)
One output row per source record; nothing summed or merged. Source values are well-level
test rates; the well/subunit of each row is kept in the log (template has no well column).
Zeros kept; blanks stay blank. Inputs read-only (sha256 verified); outputs never overwritten.
"""
import hashlib
from collections import Counter, defaultdict
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BASE = Path("/Users/farahyusof/Library/CloudStorage/OneDrive-MRCHURROSF&BSDN.BHD/Ainny Farah/07_STE/13_DULANG")
SRC = BASE / "1992-2021_Pressure Survey Export_Dulang.xlsx"
TEMPLATE = BASE / "01_EDAFY_Asset_Templates" / "field_production_template.xlsx"
OUT_DIR = BASE / "02_Mapped_Output"
OUT_DATA = OUT_DIR / "Dulang_Field_Production_1992-2021.xlsx"
OUT_LOG = OUT_DIR / "Dulang_Field_Production_Log.xlsx"
FIRST_DATA_ROW = 9
I = {"WELL": 0, "SUBUNIT": 2, "SURVEY DATE": 3, "TYPE": 4, "QUALITY": 5, "LIQUID": 17, "OIL": 18, "WATER": 20}


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

    header = [c.value for c in load_workbook(TEMPLATE)["Field Production"][1] if c.value not in (None, "")]
    assert header == ["Date", "Gross (bfpd)", "Oil Obs (bopd)", "Water (bwpd)"], header

    src_ws = load_workbook(SRC, data_only=True)["SurveyData"]
    records = [(i, r) for i, r in enumerate(src_ws.iter_rows(min_row=FIRST_DATA_ROW, values_only=True), FIRST_DATA_ROW)
               if not blank(r[I["WELL"]])]

    data = Workbook()
    ws = data.active
    ws.title = "Field Production"
    ws.append(header)
    trace, issues, skipped = [], [], 0
    for src_row, r in records:
        g = lambda k: None if blank(r[I[k]]) else r[I[k]]
        liq, oil, wat = g("LIQUID"), g("OIL"), g("WATER")
        if liq is None and oil is None and wat is None:
            skipped += 1
            continue
        dt = g("SURVEY DATE")
        ws.append([dt.date() if dt else None, liq, oil, wat])
        out_row = ws.max_row
        ws.cell(out_row, 1).number_format = "yyyy-mm-dd"
        trace.append([out_row, "SurveyData", src_row, g("WELL"), g("SUBUNIT"), dt, g("TYPE"), g("QUALITY"), liq, oil, wat])
        if all(isinstance(x, (int, float)) for x in (liq, oil, wat)) and abs(liq - (oil + wat)) > max(1, 0.02 * liq):
            issues.append(["Conflicting values", out_row, src_row, g("SUBUNIT"), dt, "Gross vs Oil Obs + Water",
                           f"LIQUID {liq:.3f} != OIL {oil:.3f} + WATER {wat:.3f} = {oil + wat:.3f}; loaded as-is."])
        if g("QUALITY") == "BAD":
            issues.append(["Data quality", out_row, src_row, g("SUBUNIT"), dt, "", "Source survey QUALITY = BAD; row kept."])

    grp = defaultdict(list)
    for t in trace:
        grp[t[5].date()].append(t)
    for d, ts in grp.items():
        if len(ts) > 1:
            for t in ts:
                issues.append(["Multiple rows same date", t[0], t[2], t[4], t[5], "Date",
                               f"{len(ts)} rows on {d} from different well tests ({', '.join(str(x[4]) for x in ts)}). "
                               "Not summed into a field total; rows are indistinguishable in the template without the log."])

    head(ws, [12, 14, 15, 14])
    data.save(OUT_DATA)

    log = Workbook()
    a = log.active
    a.title = "Attribute Mapping"
    a.append(["Target Column", "Source Sheet", "Source Column", "Source Unit", "Target Unit", "Rule / Transformation"])
    for row in [
        ["Date", "SurveyData", "SURVEY DATE", "date/time", "date", "Direct; date part written. Full timestamp in Record Traceability."],
        ["Gross (bfpd)", "SurveyData", "WELL TEST LIQUID RATE", "bbl/d", "bfpd", "Direct copy; no conversion."],
        ["Oil Obs (bopd)", "SurveyData", "WELL TEST OIL RATE", "bbl/d", "bopd", "Direct copy; no conversion."],
        ["Water (bwpd)", "SurveyData", "WELL TEST WATER RATE", "bbl/d", "bwpd", "Direct copy; no conversion."],
        ["(record selection)", "SurveyData", "LIQUID / OIL / WATER RATE", "", "",
         f"Rows with at least one rate: {len(trace)} of {len(records)} records; {skipped} records have no well-test rates."],
        ["(level note)", "", "", "", "",
         "Source rates are WELL-level test rates measured during surveys; the template is field-level. One row per source "
         "record, nothing aggregated. Well and subunit of each row are in Record Traceability."],
    ]:
        a.append(row)
    head(a, [20, 12, 26, 10, 10, 110])

    t = log.create_sheet("Record Traceability")
    t.append(["Output Row", "Source Sheet", "Source Row", "WELL", "SUBUNIT", "SURVEY DATE (full)", "TYPE", "QUALITY",
              "LIQUID RATE", "OIL RATE", "WATER RATE"])
    for x in trace:
        t.append(x)
    for row in t.iter_rows(min_row=2, min_col=6, max_col=6):
        row[0].number_format = "yyyy-mm-dd hh:mm"
    head(t, [10, 12, 10, 20, 30, 18, 7, 9, 12, 12, 12])

    s = log.create_sheet("Issues")
    s.append(["Issue Type", "Output Row", "Source Row", "SUBUNIT", "SURVEY DATE", "Column", "Detail"])
    for x in issues:
        s.append(x)
    for row in s.iter_rows(min_row=2, min_col=5, max_col=5):
        row[0].number_format = "yyyy-mm-dd hh:mm"
    head(s, [24, 10, 10, 30, 17, 22, 110])

    m = log.create_sheet("Summary")
    for row in [["Item", "Count"], ["Source records (rows with WELL)", len(records)],
                ["Records with no well-test rate (not applicable)", skipped], ["Output rows", len(trace)],
                *[[f"   TYPE = {k}", v] for k, v in sorted(Counter(x[6] for x in trace).items())],
                ["Distinct dates", len(grp)], ["Dates with more than one row", sum(1 for v in grp.values() if len(v) > 1)],
                *[[f"Issues: {k}", v] for k, v in Counter(i[0] for i in issues).most_common()]]:
        m.append(row)
    head(m, [50, 10])
    log.save(OUT_LOG)

    rows = list(load_workbook(OUT_DATA)["Field Production"].iter_rows(values_only=True))
    assert list(rows[0]) == header and len(rows) - 1 == len(trace) == len(records) - skipped
    lt = list(load_workbook(OUT_LOG)["Record Traceability"].iter_rows(min_row=2, values_only=True))
    for r, x in zip(rows[1:], lt):
        assert (r[1], r[2], r[3]) == (x[8], x[9], x[10])
    assert all(sha(p) == h for p, h in before.items()), "input changed"
    print(f"rows={len(trace)} (of {len(records)}; {skipped} without rates) distinct_dates={len(grp)} "
          f"multi_row_dates={sum(1 for v in grp.values() if len(v) > 1)} issues={dict(Counter(i[0] for i in issues))}")
    print("validated; inputs unchanged")
    print(OUT_DATA)
    print(OUT_LOG)


if __name__ == "__main__":
    main()
