#!/usr/bin/env python3
"""Re-extract the Well Test Monitoring template only, from the Dulang 1992-2021 pressure-survey export.

Rules (as specified by the user)
- Applicable records: survey rows with at least one WELL TEST rate (LIQUID / OIL / WATER).
- Entity       <- SUBUNIT (direct)
- Date         <- SURVEY DATE (direct; date part written, full timestamp in log)
- Gross (BFPD) <- WELL TEST LIQUID RATE (bbl/d, direct)
- Oil (BOPD)   <- WELL TEST OIL RATE (bbl/d, direct)
- WC (%)       = WATER RATE / (OIL RATE + WATER RATE) x 100
                 blank if oil or water is blank, or oil + water = 0 (undefined)
- Water (BWPD) <- WELL TEST WATER RATE (bbl/d, direct)
- Freq, Ampere, Voltage, Pi, Tm, RH: empty (no source data)
Zeros kept; blanks stay blank. Inputs read-only (sha256 verified); outputs never overwritten.
"""
import hashlib
import shutil
import sys
from collections import Counter, defaultdict
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BASE = Path("/Users/farahyusof/Library/CloudStorage/OneDrive-MRCHURROSF&BSDN.BHD/Ainny Farah/07_STE/13_DULANG")
SRC = BASE / "1992-2021_Pressure Survey Export_Dulang.xlsx"
TEMPLATE = BASE / "01_EDAFY_Asset_Templates" / "well_test_monitoring_template.xlsx"
OUT_DIR = BASE / "02_Mapped_Output"
OUT_DATA = OUT_DIR / "Dulang_Well_Test_Monitoring_1992-2021.xlsx"
OUT_LOG = OUT_DIR / "Dulang_Well_Test_Monitoring_Log.xlsx"
FIRST_DATA_ROW = 9
BACKUP_DIR = Path("/Users/farahyusof/Documents/GitHub/NoramiraSazuki/AinnyFarah/DOCUMENT_DST/backups")

# Full source column list (index, name, unit) shared with the asset-mapping script.
sys.path.insert(0, str(Path(__file__).parent))
from build_dulang_asset_mapping import SRC_COLS  # noqa: E402

# Source columns already used by the template columns; all others are appended after 'RH (hr)'.
USED = {"SUBUNIT", "SURVEY DATE", "WELL TEST LIQUID RATE", "WELL TEST OIL RATE", "WELL TEST WATER RATE"}
EXTRA = [(i, n, u) for i, n, u in SRC_COLS if n not in USED]
I = {"WELL": 0, "RESERVOIR": 1, "SUBUNIT": 2, "SURVEY DATE": 3, "TYPE": 4, "QUALITY": 5,
     "LIQUID": 17, "OIL": 18, "GAS": 19, "WATER": 20, "BSW": 23}


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
    # Previous versions are backed up (outside OneDrive) before being replaced.
    stamp = __import__("datetime").datetime.now().strftime("%Y%m%d_%H%M%S")
    for p in (OUT_DATA, OUT_LOG):
        if p.exists():
            BACKUP_DIR.mkdir(exist_ok=True)
            shutil.copy2(p, BACKUP_DIR / f"{p.stem}_backup_{stamp}{p.suffix}")
    before = {p: sha(p) for p in (SRC, TEMPLATE)}

    tmpl_header = [c.value for c in load_workbook(TEMPLATE)["Well Test Monitoring"][1] if c.value not in (None, "")]
    header = tmpl_header + [f"{n} [{u}]" if u else n for _, n, u in EXTRA]
    assert tmpl_header == ["Entity", "Date", "Gross (BFPD)", "Oil (BOPD)", "WC (%)", "Water (BWPD)", "Freq (Hz)", "Ampere",
                      "Voltage (Volt)", "Pi (psig)", "Tm (deg-F)", "RH (hr)"], tmpl_header

    src_ws = load_workbook(SRC, data_only=True)["SurveyData"]
    records = [(i, r) for i, r in enumerate(src_ws.iter_rows(min_row=FIRST_DATA_ROW, values_only=True), FIRST_DATA_ROW)
               if not blank(r[I["WELL"]])]

    data = Workbook()
    ws = data.active
    ws.title = "Well Test Monitoring"
    ws.append(header)
    trace, issues, skipped = [], [], 0
    for src_row, r in records:
        g = lambda k: None if blank(r[I[k]]) else r[I[k]]
        liq, oil, wat = g("LIQUID"), g("OIL"), g("WATER")
        if liq is None and oil is None and wat is None:
            skipped += 1
            continue
        dt = g("SURVEY DATE")
        if oil is None or wat is None:
            wc, wc_note = None, "WC blank: oil or water rate missing"
        elif oil + wat == 0:
            wc, wc_note = None, "WC blank: oil + water = 0 (undefined)"
        else:
            wc = wat / (oil + wat) * 100
            wc_note = f"WC = {wat} / ({oil} + {wat}) x 100"
        extra = [None if blank(r[i]) else r[i] for i, _, _ in EXTRA]
        ws.append([g("SUBUNIT"), dt.date() if dt else None, liq, oil, wc, wat, None, None, None, None, None, None] + extra)
        out_row = ws.max_row
        ws.cell(out_row, 2).number_format = "yyyy-mm-dd"
        trace.append([out_row, "SurveyData", src_row, g("SUBUNIT"), g("WELL"), dt, g("TYPE"), g("QUALITY"),
                      liq, oil, wat, wc, g("BSW"), wc_note])

        if g("SUBUNIT") is None:
            issues.append(["Missing Entity", out_row, src_row, g("WELL"), dt, "Entity", "SUBUNIT blank in source; Entity left blank."])
        if wc is None:
            issues.append(["WC not calculated", out_row, src_row, g("SUBUNIT"), dt, "WC (%)", wc_note])
        if all(isinstance(x, (int, float)) for x in (liq, oil, wat)) and abs(liq - (oil + wat)) > max(1, 0.02 * liq):
            issues.append(["Conflicting values", out_row, src_row, g("SUBUNIT"), dt, "Gross vs Oil + Water",
                           f"LIQUID {liq:.3f} != OIL {oil:.3f} + WATER {wat:.3f} = {oil + wat:.3f}; loaded as-is."])
        bsw = g("BSW")
        if wc is not None and isinstance(bsw, (int, float)) and abs(wc - bsw) > 1:
            issues.append(["WC vs source BSW", out_row, src_row, g("SUBUNIT"), dt, "WC (%)",
                           f"Calculated WC {wc:.2f} % differs from source BSW {bsw} % (informational)."])
        if g("QUALITY") == "BAD":
            issues.append(["Data quality", out_row, src_row, g("SUBUNIT"), dt, "", "Source survey QUALITY = BAD; row kept."])

    grp = defaultdict(list)
    for t in trace:
        grp[(t[3], t[5].date())].append(t)
    for (ent, d), ts in grp.items():
        if len(ts) > 1:
            for t in ts:
                issues.append(["Suspected duplicate", t[0], t[2], ent, t[5], "Entity + Date",
                               f"{len(ts)} rows for {ent} on {d} (source rows {', '.join(str(x[2]) for x in ts)}; "
                               f"types {', '.join(str(x[6]) for x in ts)}; times {', '.join(x[5].strftime('%H:%M') for x in ts)}). Kept all."])

    head(ws, [30, 12, 13, 13, 10, 13, 10, 9, 13, 10, 11, 9] + [max(12, min(len(h) + 2, 40)) for h in header[12:]])
    for row in ws.iter_rows(min_row=2, min_col=5, max_col=5):
        row[0].number_format = "0.00"
    data.save(OUT_DATA)

    log = Workbook()
    a = log.active
    a.title = "Attribute Mapping"
    a.append(["Target Column", "Source Sheet", "Source Column", "Source Unit", "Target Unit", "Rule / Transformation"])
    for row in [
        ["Entity", "SurveyData", "SUBUNIT", "", "", "Direct copy (well string + reservoir sub-unit, e.g. DULG-A025-TS-N3-E2)."],
        ["Date", "SurveyData", "SURVEY DATE", "date/time", "date", "Direct; date part written. Full timestamp in Record Traceability."],
        ["Gross (BFPD)", "SurveyData", "WELL TEST LIQUID RATE", "bbl/d", "BFPD", "Direct copy; no conversion."],
        ["Oil (BOPD)", "SurveyData", "WELL TEST OIL RATE", "bbl/d", "BOPD", "Direct copy; no conversion."],
        ["WC (%)", "SurveyData", "WELL TEST WATER RATE, WELL TEST OIL RATE", "bbl/d", "%",
         "Calculated: WATER RATE / (OIL RATE + WATER RATE) x 100. Blank if a rate is missing or oil + water = 0. "
         "Written as percent (0-100); note the template's example rows show fractions (0.6895)."],
        ["Water (BWPD)", "SurveyData", "WELL TEST WATER RATE", "bbl/d", "BWPD", "Direct copy; no conversion."],
        ["Freq (Hz)", "", "", "", "Hz", "Empty — no source data."],
        ["Ampere", "", "", "", "", "Empty — no source data."],
        ["Voltage (Volt)", "", "", "", "Volt", "Empty — no source data."],
        ["Pi (psig)", "", "", "", "psig", "Empty — ESP pump-intake pressure; survey gauge pressure is not equivalent."],
        ["Tm (deg-F)", "", "", "", "deg-F", "Empty — no source data."],
        ["RH (hr)", "", "", "", "hr", "Empty — no source data."],
        *[[h, "SurveyData", n, u, u, "Additional column after RH (hr): direct copy of the source column (not a template column)."]
          for h, (_, n, u) in zip(header[12:], EXTRA)],
        ["(record selection)", "SurveyData", "LIQUID / OIL / WATER RATE", "", "",
         f"Rows with at least one rate: {len(trace)} of {len(records)} records; {skipped} records have no well-test rates and are not applicable."],
    ]:
        a.append(row)
    head(a, [20, 12, 36, 10, 10, 110])

    t = log.create_sheet("Record Traceability")
    t.append(["Output Row", "Source Sheet", "Source Row", "SUBUNIT (Entity)", "WELL", "SURVEY DATE (full)", "TYPE", "QUALITY",
              "LIQUID RATE", "OIL RATE", "WATER RATE", "WC (%) calculated", "Source BSW (%)", "WC Calculation"])
    for x in trace:
        t.append(x)
    for row in t.iter_rows(min_row=2, min_col=6, max_col=6):
        row[0].number_format = "yyyy-mm-dd hh:mm"
    head(t, [10, 12, 10, 30, 20, 18, 7, 9, 12, 12, 12, 14, 12, 50])

    s = log.create_sheet("Issues")
    s.append(["Issue Type", "Output Row", "Source Row", "Entity / WELL", "SURVEY DATE", "Column", "Detail"])
    for x in issues:
        s.append(x)
    for row in s.iter_rows(min_row=2, min_col=5, max_col=5):
        row[0].number_format = "yyyy-mm-dd hh:mm"
    head(s, [22, 10, 10, 30, 17, 20, 100])

    m = log.create_sheet("Summary")
    for row in [["Item", "Count"], ["Source records (rows with WELL)", len(records)],
                ["Records with no well-test rate (not applicable)", skipped], ["Output rows", len(trace)],
                *[[f"   TYPE = {k}", v] for k, v in sorted(Counter(x[6] for x in trace).items())],
                ["Rows with WC (%) calculated", sum(1 for x in trace if x[11] is not None)],
                ["Rows with WC (%) blank", sum(1 for x in trace if x[11] is None)],
                *[[f"Issues: {k}", v] for k, v in Counter(i[0] for i in issues).most_common()]]:
        m.append(row)
    head(m, [50, 10])
    log.save(OUT_LOG)

    # validate by reopening
    rows = list(load_workbook(OUT_DATA)["Well Test Monitoring"].iter_rows(values_only=True))
    assert list(rows[0]) == header and len(rows) - 1 == len(trace) == len(records) - skipped
    src_by_row = dict(records)
    for r, x in zip(rows[1:], trace):
        assert all(v is None for v in r[6:12])
        src = src_by_row[x[2]]
        for k, (i, _, _) in enumerate(EXTRA):
            exp = None if blank(src[i]) else src[i]
            got = r[12 + k]
            # Excel keeps ~15-16 significant digits; allow float round-trip differences only.
            same = (got == exp) or (isinstance(got, (int, float)) and isinstance(exp, (int, float))
                                    and abs(got - exp) <= 1e-12 * max(1.0, abs(exp)))
            assert same, (x[2], header[12 + k], got, exp)
        if r[4] is not None:
            assert abs(r[4] - r[5] / (r[3] + r[5]) * 100) < 1e-9
    assert all(sha(p) == h for p, h in before.items()), "input changed"
    print(f"columns={len(header)} (12 template + {len(EXTRA)} source) rows={len(trace)} (of {len(records)} records; {skipped} without rates) "
          f"wc_calculated={sum(1 for x in trace if x[11] is not None)} issues={dict(Counter(i[0] for i in issues))}")
    print("validated; inputs unchanged")
    print(OUT_DATA)
    print(OUT_LOG)


if __name__ == "__main__":
    main()
