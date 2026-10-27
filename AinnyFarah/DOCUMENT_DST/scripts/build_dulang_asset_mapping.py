#!/usr/bin/env python3
"""Map the Dulang 1992-2021 pressure-survey export onto the 15 EDAFY asset templates.

Inputs (read-only, sha256-verified unchanged):
  13_DULANG/1992-2021_Pressure Survey Export_Dulang.xlsx   (sheet 'SurveyData')
  13_DULANG/01_EDAFY_Asset_Templates/*.xlsx                (15 template files)
Outputs (new folder 13_DULANG/02_Mapped_Output):
  Dulang_Asset_Data_Mapped_1992-2021.xlsx  - one sheet per template, template header layout
  Dulang_Mapping_Log.xlsx                  - Attribute Mapping, Record Traceability, Issues, Summary

Mapping principles
- Survey conditions are not mixed: SGS (static gradient) datum pressure -> Pressure Profile;
  FGS (flowing gradient) datum pressure and explicit Pwf/BHFP -> BHP; PBU datum pressures are
  ambiguous (some are flowing, e.g. row 43) and are not loaded.
- Depth references are not interchanged: only DATUM pressures go to datum/reservoir columns;
  gauge and mid-perf pressures are not substituted.
- No unit conversion is needed for any populated column (units identical); none applied.
- Blank source cells stay blank; zeros are preserved.
- No record is deleted or merged; suspected duplicates and conflicts go to the Issues log.
"""
import hashlib
from collections import Counter, OrderedDict, defaultdict
from copy import copy
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BASE = Path("/Users/farahyusof/Library/CloudStorage/OneDrive-MRCHURROSF&BSDN.BHD/Ainny Farah/07_STE/13_DULANG")
SRC = BASE / "1992-2021_Pressure Survey Export_Dulang.xlsx"
TDIR = BASE / "01_EDAFY_Asset_Templates"
OUT_DIR = BASE / "02_Mapped_Output"
OUT_DATA = OUT_DIR / "Dulang_Asset_Data_Mapped_1992-2021.xlsx"
OUT_LOG = OUT_DIR / "Dulang_Mapping_Log.xlsx"
SRC_SHEET = "SurveyData"
FIRST_DATA_ROW = 9

# --------------------------------------------------------------------------- source columns
SRC_COLS = [  # (index, name as printed incl. group, unit)
    (0, "WELL", ""), (1, "RESERVOIR", ""), (2, "SUBUNIT", ""), (3, "SURVEY DATE", "date/time"),
    (4, "TYPE", "SGS/FGS/PBU"), (5, "QUALITY", "GOOD/FAIR/BAD"),
    (6, "SURVEY/GAUGE MD", "ft (BDF)"), (7, "SURVEY/GAUGE TVD", "ft (SS)"), (8, "SURVEY/GAUGE PRESS", "psig"),
    (9, "SURVEY/GAUGE TEMP", "F"), (10, "SURVEY/GAUGE PRESSURE GRADIENT", "psi/ft"),
    (11, "SURVEY/GAUGE RESV FLUID GRADIENT", "psi/ft"), (12, "MID-PERF MD", "ft (BDF)"), (13, "MID-PERF TVD", "ft (SS)"),
    (14, "MID-PERF PRESS", "psig"), (15, "DATUM TVD", "ft (SS)"), (16, "DATUM PRESS", "psig"),
    (17, "WELL TEST LIQUID RATE", "bbl/d"), (18, "WELL TEST OIL RATE", "bbl/d"), (19, "WELL TEST GAS RATE", "Mscf/d"),
    (20, "WELL TEST WATER RATE", "bbl/d"), (21, "WELL TEST WATER INJ RATE", "bbl/d"), (22, "WELL TEST GAS INJ RATE", "Mscf/d"),
    (23, "WELL TEST BSW", "%"), (24, "WELL TEST THP", "psig"), (25, "WELL TEST CHP", "psig"), (26, "WELL TEST FLP", "psig"),
    (27, "PERM", "mD"), (28, "THICKNESS", "ft"), (29, "FLOW CAPACITY", "mDft"), (30, "SKIN FACTOR", ""),
    (31, "RADIUS OF INVESTIGATION", "ft"), (32, "FLOW EFFICIENCY", "%"), (33, "SKIN DRAWDOWN", "psi"),
    (34, "NON-DARCY EFFECT", ""), (35, "Pwf / BHFP", "psig"), (36, "P*", "psig"), (37, "PBAR", "psig"),
    (38, "PMAX", "psig"), (39, "PI or II", "bbl/d/psi"), (40, "MODEL USED", ""), (41, "REMARK", "text"),
]
C = {n: i for i, n, _ in SRC_COLS}
UNIT = {n: u for _, n, u in SRC_COLS}

# --------------------------------------------------------------------------- templates
# file -> (data sheet, last header row, output sheet name)
TEMPLATES = OrderedDict([
    ("allocated_production_import_template.xlsx", ("Sheet1", 1, "Allocated Production")),
    ("bhp_template.xlsx", ("BHP", 1, "BHP")),
    ("chemical_usage_template.xlsx", ("Chemical Usage", 8, "Chemical Usage")),
    ("esp_brand_template.xlsx", ("ESP Brand", 1, "ESP Brand")),
    ("esp_monitoring_template.xlsx", ("Well Information", 9, "ESP Monitoring")),
    ("field_production_template.xlsx", ("Field Production", 1, "Field Production")),
    ("perforation_split_template.xlsx", ("Split Ratios", 1, "Perforation Split")),
    ("pressure_profile_template.xlsx", ("Pressure Profile", 1, "Pressure Profile")),
    ("production_deferment_template.xlsx", ("Production Deferment", 1, "Production Deferment")),
    ("spatial_data_template.xlsx", ("Spatial Data", 1, "Spatial Data")),
    ("stock_oil_template.xlsx", ("Stock Oil", 1, "Stock Oil")),
    ("tank_template.xlsx", ("Tank", 1, "Tank")),
    ("water_injection.xlsx", ("Water Injection Monitoring", 31, "Water Injection")),
    ("well_reservoir_props_template.xlsx", ("Well Reservoir Props", 1, "Well Reservoir Props")),
    ("well_test_monitoring_template.xlsx", ("Well Test Monitoring", 1, "Well Test Monitoring")),
])
OTHER_SHEETS_NOTE = {
    "perforation_split_template.xlsx": "'Instructions' sheet is guidance only (not a data sheet).",
    "stock_oil_template.xlsx": "'Instructions' sheet is guidance only.",
    "tank_template.xlsx": "'Fields' sheet is a reference list (Dulang not listed).",
    "water_injection.xlsx": ("Other sheets ('Real Production PB#36', 'Real Test PB#34', 'Daily Chemical Operation New', "
                             "'Well Test Actual Due to PB-30', 'Actual PB-16') are example/working data for other wells; "
                             "'Water Injection Monitoring' used as the template sheet."),
    "esp_monitoring_template.xlsx": "Per-well block layout; block labels 'WELL PB-35'/'WELL PB-02' are template placeholders.",
}

# Status vocabulary
DIRECT, TRANSF, REVIEW, NOSRC, AMBIG = ("Mapped - direct", "Mapped - transformation", "Mapped - review required",
                                        "Not mapped - no source data", "Not mapped - ambiguous")

# Populated columns: (output sheet, target column) -> (source column, status, transformation, rationale)
MAP = {
    ("BHP", "Well"): ("WELL", DIRECT, "Copied as-is (well/string code, e.g. DULG-B005-SS).",
                      "Same identifier convention as other Dulang templates (e.g. Perforation Split 'DULG-A001W01-TS')."),
    ("BHP", "Survey Date"): ("SURVEY DATE", TRANSF, "Datetime -> date (YYYY-MM-DD). Time of day kept in Record Traceability.",
                             "Template example is date-only."),
    ("BHP", "Pwf - BHFP (psig)"): ("Pwf / BHFP", DIRECT, "None (psig -> psig).",
                                   "Source column is explicitly 'Pwf / BHFP' in psig (3 PBU records)."),
    ("BHP", "Datum Pressure (psig)"): ("DATUM PRESS", REVIEW, "None (psig -> psig). FGS records only.",
                                       "FGS = flowing gradient survey, so this is FLOWING pressure at datum. Template does not "
                                       "state the survey condition or datum depth; datum TVD (ft SS) varies by reservoir "
                                       "(e.g. 3970-4282) and has no template column. Confirm the BHP template expects flowing datum pressure."),
    ("Pressure Profile", "Well"): ("WELL", DIRECT, "Copied as-is.", "Same identifier convention."),
    ("Pressure Profile", "Date"): ("SURVEY DATE", TRANSF, "Datetime -> date; time kept in traceability.", "Template example is date-only."),
    ("Pressure Profile", "Reservoir Pressure (psig)"): (
        "DATUM PRESS", REVIEW, "None (psig -> psig). SGS records only.",
        "SGS = static gradient survey; datum pressure is the static reservoir pressure referenced to the reservoir datum. "
        "Datum TVD has no template column (depth reference lost) and QUALITY (644 BAD records overall) has no template column."),
    ("Well Reservoir Props", "Well"): ("WELL", DIRECT, "Copied as-is.", "Same identifier convention."),
    ("Well Reservoir Props", "Reservoir"): (
        "RESERVOIR", REVIEW, "Copied as-is (e.g. DULG-S3-E).",
        "Source RESERVOIR is a block-formation code; template example 'E70' looks like a sand/reservoir unit. "
        "Sand-level unit exists only inside SUBUNIT (e.g. ...-S3-E7) and was not parsed."),
    ("Well Reservoir Props", "Permeability (md)"): ("PERM", DIRECT, "None (mD = md).", "PBU-derived permeability; same unit."),
    ("Well Reservoir Props", "Net Pay (ft)"): (
        "THICKNESS", REVIEW, "None (ft -> ft).",
        "Well-test THICKNESS (h) used in PBU analysis; normally net pay but not stated explicitly."),
    ("Well Reservoir Props", "KH (md·ft)"): ("FLOW CAPACITY", DIRECT, "None (mDft = md·ft).", "Flow capacity = kh; same unit."),
    ("Well Test Monitoring", "Entity"): ("WELL", DIRECT, "Copied as-is.", "Template Entity is the well/string identifier."),
    ("Well Test Monitoring", "Date"): ("SURVEY DATE", TRANSF, "Datetime -> date; time kept in traceability.", "Template example is date-only."),
    ("Well Test Monitoring", "Gross (BFPD)"): ("WELL TEST LIQUID RATE", DIRECT, "None (bbl/d = BFPD).", "Measured well-test liquid rate."),
    ("Well Test Monitoring", "Oil (BOPD)"): ("WELL TEST OIL RATE", DIRECT, "None (bbl/d = BOPD).", "Measured well-test oil rate."),
    ("Well Test Monitoring", "Water (BWPD)"): ("WELL TEST WATER RATE", DIRECT, "None (bbl/d = BWPD).", "Measured well-test water rate."),
}

# Explicitly considered but NOT populated: (sheet, column) -> (candidate source, status, reason)
NOT_MAPPED = {
    ("Well Test Monitoring", "WC (%)"): ("WELL TEST BSW (%)", AMBIG,
        "BS&W includes sediment (not identical to water cut) and the template header says % while its example values are "
        "fractions (0.6895). Left blank; can be enabled once the convention is confirmed."),
    ("Well Test Monitoring", "Pi (psig)"): ("SURVEY/GAUGE PRESS / MID-PERF PRESS", AMBIG,
        "Template Pi is ESP pump-intake pressure; survey gauge pressures are not the same measurement point."),
    ("Well Test Monitoring", "Tm (deg-F)"): ("SURVEY/GAUGE TEMP", AMBIG, "Template Tm is ESP motor temperature; gauge temperature is not equivalent."),
    ("Well Test Monitoring", "RH (hr)"): ("", NOSRC, "No running-hours data."),
    ("BHP", "Pwf - BHFP (psig)#FGS"): ("MID-PERF PRESS (FGS)", AMBIG,
        "FGS mid-perf pressure is a flowing BHP at mid-perf depth but is not labelled Pwf; not substituted (different depth/label). Review."),
    ("Pressure Profile", "Reservoir Pressure (psig)#PBU"): ("PBAR / P* / DATUM PRESS (PBU)", AMBIG,
        "PBU datum pressures are sometimes flowing values (row 43: datum 513 vs Pwf 493 vs PBAR 1082); PBAR/P* have no stated "
        "depth reference. Not loaded; listed in Issues for review."),
    ("Well Reservoir Props", "Porosity (fraction)"): ("", NOSRC, "No porosity in the source."),
    ("Well Reservoir Props", "Water Saturation (fraction)"): ("", NOSRC, "No water saturation in the source."),
    ("Well Reservoir Props", "Source"): ("", NOSRC, "Provenance label is not a source field; left blank (traceability log records origin)."),
    ("Well Reservoir Props", "Sand"): ("SUBUNIT (embedded)", AMBIG, "Sand unit is embedded in SUBUNIT text; parsing it would be a derived value. Left blank."),
    ("Perforation Split", "*"): ("WELL / RESERVOIR / SUBUNIT", AMBIG,
        "Source identifiers match the perforation-code convention, but the source has no split ratios (Oil/Gas/Water %), "
        "well type, status or ratio start date. Creating identifier-only rows would misrepresent split data; sheet left empty."),
    ("ESP Monitoring", "*"): ("WELL TEST THP / FLP", AMBIG,
        "THP is a wellhead pressure, but ESP Monitoring is a per-well daily ESP log; no ESP data or daily records in the source."),
    ("Allocated Production", "*"): ("WELL TEST rates", AMBIG, "Measured test rates are not allocated production volumes/rates."),
    ("Spatial Data", "*"): ("", NOSRC, "No coordinates in the source."),
    ("Field Production", "*"): ("", NOSRC, "Source is well-level survey data, not field daily production."),
    ("Water Injection", "*"): ("WELL TEST WATER INJ RATE", NOSRC, "Water injection rate column is empty in the source (0 values)."),
}
for s in ("Chemical Usage", "ESP Brand", "Production Deferment", "Stock Oil", "Tank"):
    NOT_MAPPED[(s, "*")] = ("", NOSRC, "No related data in a pressure-survey export.")


def sha(p):
    return hashlib.sha256(p.read_bytes()).hexdigest()


def load_template(path, sheet):
    try:
        wb = load_workbook(path)
        return wb[sheet], False
    except ValueError:      # tank_template.xlsx: invalid data-validation XML -> read-only fallback
        wb = load_workbook(path, read_only=True)
        return wb[sheet], True


def header_cells(ws, last_row, read_only):
    cells = {}
    for row in ws.iter_rows(min_row=1, max_row=last_row):
        for c in row:
            if c.value not in (None, ""):
                cells[(c.row, c.column)] = c
    merges = [] if read_only else [m for m in ws.merged_cells.ranges if m.max_row <= last_row]
    widths = {} if read_only else {k: v.width for k, v in ws.column_dimensions.items() if v.width}
    return cells, merges, widths


def flat_headers(cells, last_row, sheet_out):
    """Column labels for the attribute log: join non-empty header texts in each column."""
    cols = defaultdict(list)
    for (r, c), cell in sorted(cells.items()):
        cols[c].append((r, str(cell.value).replace("\n", " ").strip()))
    out = []
    for c in sorted(cols):
        parts = [t for r, t in cols[c] if not (sheet_out in ("Chemical Usage", "Water Injection") and r < 6
                                              or sheet_out == "Water Injection" and r == 27)]
        if parts:
            out.append((c, " | ".join(parts)))
    return out


# --------------------------------------------------------------------------- main
def main():
    inputs = [SRC, *sorted(TDIR.glob("*.xlsx"))]
    before = {p: sha(p) for p in inputs}
    assert len(list(TDIR.glob("*.xlsx"))) == 15 and set(TEMPLATES) == {p.name for p in TDIR.glob("*.xlsx")}

    src_ws = load_workbook(SRC, data_only=True)[SRC_SHEET]
    records = [(i, r) for i, r in enumerate(src_ws.iter_rows(min_row=FIRST_DATA_ROW, values_only=True), FIRST_DATA_ROW)
               if r[C["WELL"]] not in (None, "")]
    blank_rows = src_ws.max_row - FIRST_DATA_ROW + 1 - len(records)

    def v(r, name):
        x = r[C[name]]
        return None if x in (None, "") else x

    out = Workbook()
    out.remove(out.active)
    trace, issues, attr_rows, summary = [], [], [], []
    tmpl_info = {}

    # ---- build every template sheet with its exact header block
    for fname, (tsheet, hrow, osheet) in TEMPLATES.items():
        tws, ro = load_template(TDIR / fname, tsheet)
        cells, merges, widths = header_cells(tws, hrow, ro)
        ws = out.create_sheet(osheet)
        for (r, c), cell in cells.items():
            n = ws.cell(row=r, column=c, value=cell.value)
            if not ro and cell.has_style:
                n.font, n.alignment, n.border, n.fill = copy(cell.font), copy(cell.alignment), copy(cell.border), copy(cell.fill)
                n.number_format = cell.number_format
        for m in merges:
            ws.merge_cells(str(m))
        for k, w in widths.items():
            ws.column_dimensions[k].width = w
        flat = flat_headers(cells, hrow, osheet)
        tmpl_info[osheet] = {"file": fname, "tsheet": tsheet, "hrow": hrow, "cells": {k: c.value for k, c in cells.items()},
                             "cols": flat, "next_row": hrow + 1}
        if osheet in ("BHP", "Pressure Profile", "Well Reservoir Props", "Well Test Monitoring"):
            for c, label in flat:
                ws.column_dimensions[get_column_letter(c)].width = max(ws.column_dimensions[get_column_letter(c)].width or 0, len(label) + 4)

    def col_index(osheet, label):
        for c, lab in tmpl_info[osheet]["cols"]:
            if lab == label:
                return c
        raise KeyError((osheet, label))

    def write_row(osheet, values, src_row, rec, populated, note=""):
        ws = out[osheet]
        r_out = tmpl_info[osheet]["next_row"]
        tmpl_info[osheet]["next_row"] += 1
        for label, val in values.items():
            if val is None:
                continue
            cell = ws.cell(row=r_out, column=col_index(osheet, label), value=val)
            if isinstance(val, datetime) or hasattr(val, "isoformat") and not isinstance(val, str):
                cell.number_format = "yyyy-mm-dd"
        trace.append([osheet, r_out, SRC_SHEET, src_row, v(rec, "WELL"), v(rec, "RESERVOIR"), v(rec, "SUBUNIT"),
                      v(rec, "SURVEY DATE"), v(rec, "TYPE"), v(rec, "QUALITY"), "; ".join(populated), note])
        return r_out

    applicable = Counter()
    excluded = defaultdict(list)
    used_rows = set()

    for src_row, rec in records:
        typ, dt = v(rec, "TYPE"), v(rec, "SURVEY DATE")
        d = dt.date() if dt else None

        # Pressure Profile: SGS static datum pressure
        if typ == "SGS":
            applicable["Pressure Profile"] += 1
            p = v(rec, "DATUM PRESS")
            if p is None:
                excluded["Pressure Profile"].append((src_row, "SGS record without DATUM PRESS (gauge/mid-perf pressures not substituted)"))
            else:
                write_row("Pressure Profile", {"Well": v(rec, "WELL"), "Date": d, "Reservoir Pressure (psig)": p},
                          src_row, rec, ["Well<-WELL", "Date<-SURVEY DATE", "Reservoir Pressure (psig)<-DATUM PRESS"],
                          f"Datum TVD {v(rec, 'DATUM TVD')} ft SS (not in template)")
                used_rows.add(src_row)

        # BHP: FGS flowing datum pressure; explicit Pwf/BHFP (PBU)
        if typ == "FGS" or v(rec, "Pwf / BHFP") is not None:
            applicable["BHP"] += 1
            vals = {"Well": v(rec, "WELL"), "Survey Date": d, "Pwf - BHFP (psig)": v(rec, "Pwf / BHFP"),
                    "Datum Pressure (psig)": v(rec, "DATUM PRESS") if typ == "FGS" else None}
            if vals["Pwf - BHFP (psig)"] is None and vals["Datum Pressure (psig)"] is None:
                excluded["BHP"].append((src_row, "FGS record without DATUM PRESS or Pwf/BHFP"))
            else:
                pop = ["Well<-WELL", "Survey Date<-SURVEY DATE"]
                pop += ["Pwf - BHFP (psig)<-Pwf / BHFP"] if vals["Pwf - BHFP (psig)"] is not None else []
                pop += ["Datum Pressure (psig)<-DATUM PRESS (FGS flowing)"] if vals["Datum Pressure (psig)"] is not None else []
                note = f"Datum TVD {v(rec, 'DATUM TVD')} ft SS" if vals["Datum Pressure (psig)"] is not None else \
                    f"PBU: datum pressure {v(rec, 'DATUM PRESS')} not loaded (condition ambiguous)"
                write_row("BHP", vals, src_row, rec, pop, note)
                used_rows.add(src_row)

        # Well Reservoir Props: PBU interpretation results
        if any(v(rec, n) is not None for n in ("PERM", "THICKNESS", "FLOW CAPACITY")):
            applicable["Well Reservoir Props"] += 1
            write_row("Well Reservoir Props",
                      {"Well": v(rec, "WELL"), "Reservoir": v(rec, "RESERVOIR"), "Permeability (md)": v(rec, "PERM"),
                       "Net Pay (ft)": v(rec, "THICKNESS"), "KH (md·ft)": v(rec, "FLOW CAPACITY")},
                      src_row, rec, ["Well<-WELL", "Reservoir<-RESERVOIR", "Permeability (md)<-PERM",
                                     "Net Pay (ft)<-THICKNESS", "KH (md·ft)<-FLOW CAPACITY"],
                      f"Survey date {d} (no date column in template); SUBUNIT {v(rec, 'SUBUNIT')}")
            used_rows.add(src_row)

        # Well Test Monitoring: measured rates during survey
        if any(v(rec, n) is not None for n in ("WELL TEST LIQUID RATE", "WELL TEST OIL RATE", "WELL TEST WATER RATE")):
            applicable["Well Test Monitoring"] += 1
            write_row("Well Test Monitoring",
                      {"Entity": v(rec, "WELL"), "Date": d, "Gross (BFPD)": v(rec, "WELL TEST LIQUID RATE"),
                       "Oil (BOPD)": v(rec, "WELL TEST OIL RATE"), "Water (BWPD)": v(rec, "WELL TEST WATER RATE")},
                      src_row, rec, ["Entity<-WELL", "Date<-SURVEY DATE", "Gross<-LIQUID RATE", "Oil<-OIL RATE", "Water<-WATER RATE"],
                      f"Survey type {typ}; BSW {v(rec, 'WELL TEST BSW')} % (WC left blank)")
            used_rows.add(src_row)

    # ---------------------------------------------------------------- issues
    def issue(kind, sheet, out_row, src_row, rec, column, detail, action):
        issues.append([kind, sheet, out_row, src_row, v(rec, "WELL") if rec else "", v(rec, "SURVEY DATE") if rec else "",
                       column, detail, action])

    rec_by_row = dict(records)
    out_rows_by_src = defaultdict(list)
    for t in trace:
        out_rows_by_src[t[3]].append((t[0], t[1]))

    # suspected duplicates: same output sheet + well + date
    groups = defaultdict(list)
    for t in trace:
        key_date = t[7].date() if t[7] else None
        groups[(t[0], t[4], key_date)].append(t)
    for (sheet, well, dd), ts in groups.items():
        if len(ts) > 1:
            same_sub = len({t[6] for t in ts}) == 1
            for t in ts:
                issue("Suspected duplicate", sheet, t[1], t[3], rec_by_row[t[3]], "Well + Date",
                      f"{len(ts)} rows for {well} on {dd} in this sheet (source rows "
                      f"{', '.join(str(x[3]) for x in ts)}; times {', '.join(x[7].strftime('%H:%M') for x in ts)}; "
                      f"subunits {'identical' if same_sub else 'differ: ' + ', '.join(str(x[6]) for x in ts)})",
                      "Kept all rows. Confirm whether repeat surveys/zones or duplicate entries.")

    for src_row, rec in records:
        outs = out_rows_by_src.get(src_row, [])
        q = v(rec, "QUALITY")
        for sheet, orow in outs:
            if q == "BAD":
                issue("Data quality", sheet, orow, src_row, rec, "QUALITY", "Source survey QUALITY = BAD (no template column for quality).",
                      "Review before use; record kept.")
            elif q is None:
                issue("Data quality", sheet, orow, src_row, rec, "QUALITY", "Source QUALITY blank.", "Review.")
            if v(rec, "SUBUNIT") is None:
                issue("Missing field", sheet, orow, src_row, rec, "SUBUNIT", "SUBUNIT blank in source (zone not identified).", "Review.")
        liq, oil, wat = (v(rec, n) for n in ("WELL TEST LIQUID RATE", "WELL TEST OIL RATE", "WELL TEST WATER RATE"))
        if all(isinstance(x, (int, float)) for x in (liq, oil, wat)) and abs(liq - (oil + wat)) > max(1, 0.02 * liq):
            for sheet, orow in outs:
                if sheet == "Well Test Monitoring":
                    issue("Conflicting values", sheet, orow, src_row, rec, "Gross vs Oil + Water",
                          f"LIQUID {liq:.3f} != OIL {oil:.3f} + WATER {wat:.3f} = {oil + wat:.3f}", "Values loaded as-is; verify.")
        if v(rec, "TYPE") == "PBU" and (v(rec, "PBAR") is not None or v(rec, "P*") is not None or v(rec, "DATUM PRESS") is not None):
            issue("Ambiguous mapping", "Pressure Profile", "", src_row, rec, "Reservoir Pressure (psig)",
                  f"PBU record not loaded: DATUM PRESS={v(rec, 'DATUM PRESS')}, PBAR={v(rec, 'PBAR')}, P*={v(rec, 'P*')}, "
                  f"Pwf={v(rec, 'Pwf / BHFP')} — survey condition / depth reference unclear.", "Decide which PBU pressure (if any) is reservoir pressure.")
        if v(rec, "TYPE") == "FGS" and v(rec, "MID-PERF PRESS") is not None:
            issue("Ambiguous mapping", "BHP", next((o for s, o in outs if s == "BHP"), ""), src_row, rec, "Pwf - BHFP (psig)",
                  f"FGS mid-perf flowing pressure {v(rec, 'MID-PERF PRESS')} psig at {v(rec, 'MID-PERF TVD')} ft SS not loaded as Pwf "
                  "(not labelled Pwf; different depth reference).", "Confirm whether FGS mid-perf pressure should populate Pwf.")
        g = v(rec, "WELL TEST GAS RATE")
        if isinstance(g, (int, float)) and g > 100000:
            issue("Unit anomaly (unmapped column)", "", "", src_row, rec, "WELL TEST GAS RATE",
                  f"Gas rate {g} labelled Mscf/d (= {g / 1000:.1f} MMscf/d) — magnitude suggests scf/d.", "Not loaded; verify unit if gas is mapped later.")
        if not (datetime(1992, 1, 1) <= v(rec, "SURVEY DATE") <= datetime(2021, 12, 31, 23, 59)):
            issue("Out of period", "", "", src_row, rec, "SURVEY DATE", "Outside 1992-2021.", "Review.")

    for sheet, ex in excluded.items():
        for src_row, why in ex:
            issue("Excluded record", sheet, "", src_row, rec_by_row[src_row], "", why, "No target value; not loaded.")

    used_cols = {m[0] for m in MAP.values()} | {"TYPE", "QUALITY", "SUBUNIT", "DATUM TVD", "RESERVOIR"}
    for _, name, unit in SRC_COLS:
        if name not in used_cols:
            n = sum(1 for _, r in records if v(r, name) is not None)
            issue("Unmatched source column", "", "", "", None, name,
                  f"Unit '{unit}', {n} non-blank values — no template column with the same meaning/context.", "None (informational).")
    for col in ("TYPE", "QUALITY", "SUBUNIT", "DATUM TVD"):
        issue("Unmatched source column", "", "", "", None, col,
              "Used for record routing / traceability only; no template column.", "Informational.")

    # ---------------------------------------------------------------- attribute mapping log
    for osheet, info in tmpl_info.items():
        for c, label in info["cols"]:
            key = (osheet, label)
            if key in MAP:
                src, status, tx, why = MAP[key]
                attr_rows.append([info["file"], osheet, get_column_letter(c), label, SRC_SHEET, src, UNIT.get(src, ""),
                                  unit_of(label), status, tx, why])
            else:
                nm = NOT_MAPPED.get(key) or NOT_MAPPED.get((osheet, "*"))
                src, status, why = nm if nm else ("", NOSRC, "No source attribute with the same meaning.")
                attr_rows.append([info["file"], osheet, get_column_letter(c), label, SRC_SHEET if src else "", src, "",
                                  unit_of(label), status, "", why])
        for k, (src, status, why) in NOT_MAPPED.items():
            if k[0] == osheet and "#" in k[1]:
                attr_rows.append([info["file"], osheet, "", k[1].split("#")[0] + f" (candidate: {src})", SRC_SHEET, src, "",
                                  unit_of(k[1]), status, "", why])
        if info["file"] in OTHER_SHEETS_NOTE:
            attr_rows.append([info["file"], osheet, "", "(other sheets in template file)", "", "", "", "", "Note", "",
                              OTHER_SHEETS_NOTE[info["file"]]])

    # ---------------------------------------------------------------- summary
    for osheet, info in tmpl_info.items():
        pop = info["next_row"] - info["hrow"] - 1
        appl = applicable.get(osheet, 0)
        summary.append([osheet, info["file"], info["tsheet"], appl, pop, len(excluded.get(osheet, [])),
                        "OK" if appl == pop + len(excluded.get(osheet, [])) else "CHECK",
                        APPLICABLE_DEF.get(osheet, "No applicable source records (template left empty).")])

    OUT_DIR.mkdir(exist_ok=False)
    out.save(OUT_DATA)
    unused = [r for r, _ in records if r not in used_rows]
    write_log(attr_rows, trace, issues, summary, len(records), blank_rows, unused, rec_by_row, v)

    # ---------------------------------------------------------------- validation by reopening
    validate(applicable, excluded)
    assert all(sha(p) == h for p, h in before.items()), "an input file changed"
    print("inputs unchanged (sha256 verified)")


APPLICABLE_DEF = {
    "Pressure Profile": "TYPE = SGS (static gradient survey). Populated when DATUM PRESS present.",
    "BHP": "TYPE = FGS (flowing) or explicit Pwf/BHFP present. Populated when Pwf or FGS DATUM PRESS present.",
    "Well Reservoir Props": "Any of PERM / THICKNESS / FLOW CAPACITY present (PBU interpretations).",
    "Well Test Monitoring": "Any of LIQUID / OIL / WATER rate present.",
}


def unit_of(label):
    import re
    m = re.search(r"\(([^()]*)\)\s*$", label)
    return m.group(1) if m else ""


def style(ws, widths):
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="DDEBF7")
        cell.alignment = Alignment(wrap_text=True, vertical="top")
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def write_log(attr_rows, trace, issues, summary, n_rec, blank_rows, unused, rec_by_row, v):
    wb = Workbook()
    ws = wb.active
    ws.title = "Attribute Mapping"
    ws.append(["Template File", "Output Sheet", "Target Col", "Target Column", "Source Sheet", "Source Column",
               "Source Unit", "Target Unit", "Mapping Status", "Transformation", "Rationale"])
    for r in attr_rows:
        ws.append(r)
    style(ws, [34, 22, 9, 42, 12, 30, 12, 12, 26, 40, 90])
    for row in ws.iter_rows(min_row=2):
        for c in (row[10], row[9]):
            c.alignment = Alignment(wrap_text=True, vertical="top")

    ws = wb.create_sheet("Record Traceability")
    ws.append(["Output Sheet", "Output Row", "Source Sheet", "Source Row", "WELL", "RESERVOIR", "SUBUNIT",
               "SURVEY DATE (full, as in source)", "TYPE", "QUALITY", "Populated Columns <- Source", "Note"])
    for t in trace:
        ws.append(t)
    for row in ws.iter_rows(min_row=2, min_col=8, max_col=8):
        row[0].number_format = "yyyy-mm-dd hh:mm"
    style(ws, [22, 10, 12, 10, 20, 14, 30, 20, 7, 9, 70, 50])

    ws = wb.create_sheet("Issues")
    ws.append(["Issue Type", "Output Sheet", "Output Row", "Source Row", "WELL", "SURVEY DATE", "Column", "Detail", "Action / Review"])
    for i in issues:
        ws.append(i)
    for row in ws.iter_rows(min_row=2, min_col=6, max_col=6):
        row[0].number_format = "yyyy-mm-dd hh:mm"
    style(ws, [26, 20, 10, 10, 20, 17, 26, 90, 50])

    ws = wb.create_sheet("Summary")
    ws.append(["Output Sheet", "Template File", "Template Sheet Used", "Applicable Source Records", "Populated Rows",
               "Excluded Applicable Records", "Reconciles", "Applicable Record Definition"])
    for s in summary:
        ws.append(s)
    ws.append([])
    ws.append(["Source workbook", str(SRC.name)])
    ws.append(["Source sheet", SRC_SHEET])
    ws.append(["Source survey records (rows with WELL)", n_rec])
    ws.append(["Blank rows below data (ignored)", blank_rows])
    ws.append(["Records used in at least one output sheet", n_rec - len(unused)])
    ws.append(["Records not used in any output sheet", len(unused)])
    reasons = Counter()
    for r in unused:
        rec = rec_by_row[r]
        reasons[f"{v(rec, 'TYPE')}: " + ("no datum/Pwf pressure, rates or PBU properties"
                                         if v(rec, 'TYPE') != "PBU" else "PBU with no Pwf, PERM/THICKNESS/KH or rates (pressures ambiguous)")] += 1
    for k, n in sorted(reasons.items()):
        ws.append(["   " + k, n])
    issue_counts = Counter(i[0] for i in issues)
    ws.append([])
    ws.append(["Issue counts"])
    for k, n in issue_counts.most_common():
        ws.append(["   " + k, n])
    style(ws, [44, 38, 26, 14, 12, 14, 11, 80])
    wb.save(OUT_LOG)


def validate(applicable, excluded):
    data = load_workbook(OUT_DATA)
    assert len(data.sheetnames) == 15 and len(set(data.sheetnames)) == 15
    assert all(len(n) <= 31 for n in data.sheetnames)
    for fname, (tsheet, hrow, osheet) in TEMPLATES.items():
        tws, ro = load_template(TDIR / fname, tsheet)
        t = {(c.row, c.column): c.value for row in tws.iter_rows(min_row=1, max_row=hrow) for c in row if c.value not in (None, "")}
        o = {(c.row, c.column): c.value for row in data[osheet].iter_rows(min_row=1, max_row=hrow) for c in row if c.value not in (None, "")}
        assert t == o, f"header mismatch in {osheet}"
    log = load_workbook(OUT_LOG)
    trace = list(log["Record Traceability"].iter_rows(min_row=2, values_only=True))
    for osheet in data.sheetnames:
        hrow = next(h for _, (_, h, s) in TEMPLATES.items() if s == osheet)
        n_data = sum(1 for row in data[osheet].iter_rows(min_row=hrow + 1, values_only=True) if any(x is not None for x in row))
        n_trace = sum(1 for t in trace if t[0] == osheet)
        assert n_data == n_trace, f"{osheet}: {n_data} rows vs {n_trace} trace rows"
        if osheet in applicable:
            assert n_data + len(excluded.get(osheet, [])) == applicable[osheet], osheet
        print(f"  {osheet:22} rows={n_data:5}  applicable={applicable.get(osheet, 0):5}  excluded={len(excluded.get(osheet, [])):4}")
    print("validated: 15 sheets, headers identical to templates, row counts reconcile with traceability and applicable records")
    print(f"Saved:\n  {OUT_DATA}\n  {OUT_LOG}")


if __name__ == "__main__":
    main()
