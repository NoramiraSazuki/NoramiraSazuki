#!/usr/bin/env python3
"""Build the DST-CSV -> asset template attribute mapping workbook.

Reads (read-only):
  - DOCUMENT_DST/table_csv/*.csv  (extracted DULANG-7 DST tables)
  - asset_combined_templates.xlsx (template headers only)
Writes:
  - DOCUMENT_DST/attribute_mapping/DST_Template_Attribute_Mapping.xlsx

Mapping judgements are declared in SPECS below; evidence (pages, sample values)
is pulled from the CSVs automatically. No template is populated.
"""
import csv
import hashlib
import re
from collections import OrderedDict
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

ROOT = Path("/Users/farahyusof/Documents/GitHub/NoramiraSazuki/AinnyFarah/DOCUMENT_DST")
CSV_DIR = ROOT / "table_csv"
OUT_DIR = ROOT / "attribute_mapping"
OUT = OUT_DIR / "DST_Template_Attribute_Mapping.xlsx"
TEMPLATE = Path("/Users/farahyusof/Library/CloudStorage/OneDrive-MRCHURROSF&BSDN.BHD/Ainny Farah/"
                "07_STE/13_DULANG/01_EDAFY_Asset_Templates/asset_combined_templates.xlsx")

D, T, R, N = ("Direct 1-to-1", "Requires transformation", "Possible match — needs review", "No match")
RANK = {D: 0, T: 1, R: 2, N: 3}


def sha(p):
    return hashlib.sha256(p.read_bytes()).hexdigest()


# --------------------------------------------------------------------------- template attributes
BLOCK = [("Gross (BFPD)", "BFPD"), ("Oil (BOPD)", "BOPD"), ("WC (%)", "% (template sample values stored as fraction, e.g. 0.6895)"),
         ("Water (BWPD)", "BWPD"), ("Freq (Hz)", "Hz"), ("Ampere", "not stated (A)"), ("Voltage (Volt)", "Volt"),
         ("Pi (psig)", "psig"), ("Tm (deg-F)", "deg-F"), ("RH (hr)", "hr")]


def block_sheet(prefix, extra=()):
    return prefix + BLOCK + list(extra)


def unit_from_name(a):
    m = re.search(r"\(([^()]*)\)\s*$", a)
    return m.group(1) if m else ""


TEMPLATE_ATTRS = OrderedDict()
TEMPLATE_ATTRS["Summary"] = [("No.", ""), ("File Name", "text")]
TEMPLATE_ATTRS["Allocated Production"] = [(a, unit_from_name(a)) for a in [
    "Entity", "Date", "Oil Vol. (bbl)", "Water Vol. (bbl)", "Liq. Vol. (bbl)", "Gas Vol. (MMscf)",
    "Gas Inj. Vol. (MMscf) (MMscf/d)", "Oil PD Rate (bbls/d)", "Water PD Rate (bbls/d)", "Liq. PD Rate (bbls/d)",
    "Gas PD Rate (MMscf/d)", "Gas Inj. PD Rate (MMscf/d)", "Np (MMbbl)", "Wp (MMbbl)", "Cum. Water Inj (MMbbl)",
    "Cum. Gas Inj. (Bscf/d)", "WC (%)", "GOR (scf/bbl)", "Inst. GI/GP", "Cum. GI/GP"]]
TEMPLATE_ATTRS["BHP"] = [("Well", "code"), ("Survey Date", "date YYYY-MM-DD"), ("Pwf - BHFP (psig)", "psig"),
                         ("Datum Pressure (psig)", "psig")]
TEMPLATE_ATTRS["Chemical Usage"] = [(a, "") for a in ["No", "Date *", "Chemical *", "Category", "Stock *", "Received *",
                                                      "Dosage *", "Usage *", "Wash Tank *", "Texsteam 1 *"]]
TEMPLATE_ATTRS["ESP Brand"] = [("Brand", "text")]
TEMPLATE_ATTRS["ESP Monitoring"] = [
    ("Well (block header: WELL PB-xx)", "code"), ("Date", "date (Excel serial)"), ("Hours", "hr"),
    ("ESP Pump Online - Yes", "flag"), ("ESP Pump Online - No", "flag"), ("ESP Pump Freq", "Hz"),
    ("ESP Pump Ampere", "A"), ("ESP Pump Voltage", "V"), ("ESP Pump Pi", "psig"), ("ESP Pump Tm", "°C"),
    ("Wellhead WHP", "psig"), ("Wellhead WHT", "°F"), ("Header Manifold Press.", "psig"),
    ("Header Manifold Temp.", "°F"), ("Remarks", "text")]
TEMPLATE_ATTRS["Field Production"] = [("Date", "date YYYY-MM-DD"), ("Gross (bfpd)", "bfpd"), ("Oil Obs (bopd)", "bopd"),
                                      ("Water (bwpd)", "bwpd")]
TEMPLATE_ATTRS["Perforation Split Ratios"] = [(a, u) for a, u in [
    ("Well", "code e.g. DULG-A001W01-TS"), ("Well Type", "code e.g. OP"), ("Well Status", "text"),
    ("Perf. Code", "code e.g. DULG-A001W01-TS-SE1-E14"), ("Resv. Block Formation", "code"), ("Resv. Block", "code"),
    ("Resv. Formation", "code e.g. DULG-E"), ("Start Ratio Date", "date YYYY-MM-DD"), ("Oil %", "0-100"),
    ("Gas %", "0-100"), ("Cond %", "0-100"), ("Water %", "0-100"), ("Gas Inj. %", "0-100"), ("Water Inj. %", "0-100")]]
TEMPLATE_ATTRS["Perf. Split Ratio Instructions"] = []
TEMPLATE_ATTRS["Pressure Profile"] = [("Well", "code"), ("Date", "date YYYY-MM-DD"), ("Reservoir Pressure (psig)", "psig")]
TEMPLATE_ATTRS["Production Deferment"] = [(a, unit_from_name(a)) for a in [
    "Date", "Well", "Type", "Cause", "Sub-Cause", "Equipment", "Deferment (bbl/d)", "Remarks"]]
TEMPLATE_ATTRS["Spatial Data"] = [(a, unit_from_name(a)) for a in ["Well", "Well Name", "Easting (m)", "Northing (m)", "Source"]]
TEMPLATE_ATTRS["Stock Oil"] = [(a, "bbl" if a.startswith("T-") else "") for a in
                               ["Date", "Opening Stock", "Prod Obs", "Shipping Obs", "Closing Stock"] +
                               [f"T-{i:02d}" for i in range(1, 11)]]
TEMPLATE_ATTRS["Stock Oil Instructions"] = []
TEMPLATE_ATTRS["Tank"] = [(a, "text") for a in ["Tank Code", "Tank Name", "Field", "Status", "Description"]]
TEMPLATE_ATTRS["Fields"] = [("Field name (header-less list)", "text")]
TEMPLATE_ATTRS["Real Production PB#36"] = block_sheet([("Date", "date (Excel serial)"),
                                                       ("Well (block header: DATA WELL TEST PB-36)", "code")])
TEMPLATE_ATTRS["Water Injection Monitoring"] = [(a, u) for a, u in [
    ("No", ""), ("Date", "date"), ("Total Rate Inj", "Bpm"), ("Disch. Pressure WIP", "Psi"), ("WHP PB-08", "Psi"),
    ("Disch. Pressure to PB-15", "Psi"), ("WHP PB-15", "Psi"), ("Volume Inj.", "Bbl"), ("Volume Inj. PB-08", "Bbl"),
    ("Cumm Vol Inj. PB-08", "Bbl"), ("Volume Inj. PB-15", "Bbl"), ("Cumm Vol Inj. PB-15", "Bbl"),
    ("Cumm Total Vol Inj", "Bbl"), ("Remarks", "text")]]
TEMPLATE_ATTRS["Real Test PB#34"] = block_sheet([("Date", "date (Excel serial)"),
                                                 ("Well (block header: DATA WELL TEST PB-34)", "code")])
TEMPLATE_ATTRS["Daily Chemical Operation New"] = None   # flattened from the sheet below
TEMPLATE_ATTRS["Well Test Actual Due to PB-30"] = block_sheet(
    [("No", ""), ("Date", "date (Excel serial)"), ("Well (block header: DATA WELL TEST PB-xx, 30 wells)", "code")],
    [("Total (block)", "per block"), ("Online Wells", "count")])
TEMPLATE_ATTRS["Actual PB-16"] = block_sheet([("DATE", "date (Excel serial)"), ("Well (sheet title PB-16)", "code")])
TEMPLATE_ATTRS["Well Reservoir Props"] = [(a, u) for a, u in [
    ("Well", "code"), ("Reservoir", "code e.g. E70"), ("Permeability (md)", "md"), ("Net Pay (ft)", "ft"),
    ("KH (md·ft)", "md·ft"), ("Porosity (fraction)", "fraction"), ("Water Saturation (fraction)", "fraction"),
    ("Source", "text e.g. manual"), ("Sand", "text")]]
TEMPLATE_ATTRS["Well Test Monitoring"] = block_sheet(
    [("No", ""), ("Date", "date (Excel serial)"), ("Well (block header: DATA WELL TEST PB-xx)", "code")])

INSTRUCTION_SHEETS = {"Perf. Split Ratio Instructions", "Stock Oil Instructions"}
WELLTEST_BLOCK_SHEETS = ["Well Test Monitoring", "Well Test Actual Due to PB-30", "Real Test PB#34",
                         "Real Production PB#36", "Actual PB-16"]


def flatten_daily_chemical(wb):
    ws = wb["Daily Chemical Operation New"]
    rows = list(ws.iter_rows(min_row=6, max_row=10, max_col=49, values_only=True))
    g7 = g8 = ""
    attrs = []
    for c in range(3, 49):
        r6, r7, r8, r9, unit = (rows[i][c] for i in range(5))
        r7 = str(r7).strip() if r7 else ""
        r8 = str(r8).strip() if r8 else ""
        if r7:
            g7, g8 = r7, ""
        if r8:
            g8 = r8
        leaf = str(r9).strip() if r9 else (r8 or str(r6 or "").strip())
        if not leaf and not r6:
            continue
        name = str(r6).strip() if r6 else " / ".join(x for x in (g7, g8 if g8 != leaf else "", leaf) if x)
        attrs.append((name, str(unit).strip() if unit else ""))
    return attrs


# --------------------------------------------------------------------------- source attributes
FORM_ROW_CSVS = {"Field (as printed)": "Unit", "Parameter": "Unit", "Parameter (as printed)": None}

UNIT_HINTS = [
    (r"PSIG/PSI", "psig (some sheets printed psi)"), (r"PSIA", "psia"), (r"PSIG", "psig"),
    (r"\bSTB/D\b", "STB/D"), (r"MSCF/D", "Mscf/d"), (r"\bMSCF\b", "Mscf"), (r"\bB/D\b", "bbl/d"),
    (r"\bSTB\b", "STB"), (r"BBLS?\b", "bbl"), (r"°API", "°API"), (r"°F|degF|\(°F\)", "°F"),
    (r"AIR = 1", "sp. gr. (air = 1)"), (r"64TH", "64ths in (A = adjustable, P = positive/fixed)"),
    (r"IN H2O", "in H2O"), (r"PLATE IN", "in"), (r"Mcf/D", "Mcf/d"), (r"\(ppm\)|ppm", "ppm"),
    (r"Ωm", "ohm·m"), (r"%", "%"), (r"Δ TIME", "min"), (r"DEPTH", "m"), (r"GRADIENT", "psi/ft"),
    (r"PRESSURE \(psi\)", "psi (static, gauge depth)"),
]


def guess_unit(header):
    for pat, u in UNIT_HINTS:
        if re.search(pat, header):
            return u
    if re.search(r"DATE|Date|TIME|Time", header):
        return "date/time as written"
    return "not stated"


def load_csvs():
    index = {r["CSV File"]: r for r in csv.DictReader(open(CSV_DIR / "_table_index.csv", encoding="utf-8-sig"))}
    data = OrderedDict()
    for p in sorted(CSV_DIR.glob("T*.csv")):
        rows = list(csv.DictReader(open(p, encoding="utf-8-sig")))
        hdr = list(rows[0].keys()) if rows else []
        data[p.name] = {"rows": rows, "header": hdr, "index": index.get(p.name, {})}
    return data


def source_attributes(data):
    """Yield (csv, attribute, unit, kind) where kind = 'column' or 'row'."""
    out = []
    for name, d in data.items():
        hdr = d["header"]
        key = next((k for k in FORM_ROW_CSVS if k in hdr), None)
        if key:
            seen = set()
            ucol = FORM_ROW_CSVS[key]
            for r in d["rows"]:
                a = r[key]
                if a in seen:
                    continue
                seen.add(a)
                u = (r.get(ucol) or "") if ucol else (unit_from_name(a) or "see values")
                out.append((name, a, u or "not stated", "row"))
        else:
            for h in hdr:
                if h == "Source Page":
                    continue
                out.append((name, h, guess_unit(h), "column"))
    return out


def evidence(d, attr, kind, n=3):
    rows = d["rows"]
    if kind == "row":
        key = next(k for k in FORM_ROW_CSVS if k in d["header"])
        hits = [r for r in rows if r[key] == attr]
        pages = sorted({int(r["Source Page"]) for r in hits})
        vcols = [c for c in d["header"] if c not in ("Source Page", "Item No.", key, "Unit", "Section",
                                                       "Unit (as printed / hand-corrected)")]
        vals = []
        for r in hits[:2]:
            vals += [r[c] for c in vcols if r.get(c)]
        sample = "; ".join(v[:80] for v in vals[:n])
    else:
        pages = sorted({int(r["Source Page"]) for r in rows if r.get(attr) not in (None, "")})
        vals = [r[attr] for r in rows if r.get(attr) not in (None, "", "↓", '"', "-")]
        # Prefer value-like cells over margin notes copied into spanned rows.
        vals = sorted(vals, key=lambda v: not re.match(r"^[\d.,\s/@:-]+[A-Za-z%°]*", v))
        uniq = list(OrderedDict.fromkeys(vals))
        sample = "; ".join(v[:60] for v in uniq[:n])
    pg = (f"p.{pages[0]}" if len(pages) == 1 else f"pp.{pages[0]}-{pages[-1]}") if pages else "no values"
    return pg, sample


# --------------------------------------------------------------------------- mapping specifications
# (csv regex, source attribute, kind, template sheet, template attribute, status, reason, transformation)
D5 = r"^T(09|17|25|32)_D5_"
D7 = r"^T(10|18|26|33)_D7_"
D8 = r"^T(11|19|27|34)_D8_"
D9 = r"^T(12|20|28|35)_D9_"
D12 = r"^T(05|06)_D12_"
D13 = r"^T(15|22|30|37)_D13_"
D15 = r"^T(16|31)_D15_"
T1 = r"^T01_"
HZ = r"^T0[34]_"
D1 = r"^T(07|23)_D1_"

WELL_NOTE = ("Source gives the well as 'DULANG 7' / 'DULANG-7' (form header). Template uses its own well codes "
             "(examples: '01', 'PB-xx', 'DULG-A001W01-TS'); the code for Dulang-7 is not in any source table.")
WELL_TX = "Look up the template well code for DULANG-7 in the asset well master, then apply it to every row from this test."
PSIA_TX = ("psig = psia − atmospheric pressure (14.696 psi at sea level; confirm the value used by the asset). "
           "Separate the '@ depth' text from the number before converting.")

SPECS = []


def S(csvpat, attr, kind, sheet, tattr, status, reason, tx=""):
    SPECS.append((csvpat, attr, kind, sheet, tattr, status, reason, tx))


# ---- well identity rows (form headers)
for sheet, tattr in [("BHP", "Well"), ("Pressure Profile", "Well"), ("Well Reservoir Props", "Well"),
                     ("Perforation Split Ratios", "Well")]:
    S(D13, "Well", "row", sheet, tattr, R, WELL_NOTE, WELL_TX)
for sheet in WELLTEST_BLOCK_SHEETS:
    tattr = next(a for a, _ in TEMPLATE_ATTRS[sheet] if a.startswith("Well"))
    S(D13, "Well", "row", sheet, tattr, R, WELL_NOTE + " Template puts the well in a block header, not a column.", WELL_TX)
S(D13, "Well", "row", "ESP Monitoring", "Well (block header: WELL PB-xx)", R,
  WELL_NOTE + " ESP Monitoring is for ESP-lifted wells; DST well flowed naturally on a test string.", WELL_TX)

# ---- BHP sheet
S(D13, "Date", "row", "BHP", "Survey Date", T,
  "D-13 form date is the test summary date (written MM/DD/YY, e.g. 04/09/90 = 9 Apr 1990).",
  "Parse as MM/DD/YY → YYYY-MM-DD (04/09/90 → 1990-04-09). Confirm the test-summary date is the intended survey date.")
S(D12, "Date", "column", "BHP", "Survey Date", T,
  "Date of each gauge reading on the D-12 printout (DD-MMM-YY).",
  "Parse DD-MMM-YY → YYYY-MM-DD (08-APR-90 → 1990-04-08); use the date of the reading selected for Pwf.")
S(D13, "Final flowing pressure", "row", "BHP", "Pwf - BHFP (psig)", T,
  "Same meaning (final bottom-hole flowing pressure, one value per test). Source is psia at a stated gauge depth; template is psig with no depth field.",
  PSIA_TX + " Gauge depth (e.g. 1420 m) has no template field — record it elsewhere or flag the loss of detail.")
S(T1, "FBHP, PSIA", "column", "BHP", "Pwf - BHFP (psig)", T,
  "Same quantity as D-13 item 17 (competing source). Value text combines pressure and depth, e.g. '1546.9 @ 1420.0m'.",
  "Split '<value> @ <depth>m'; then " + PSIA_TX)
S(D12, "BHP (PSIA)", "column", "BHP", "Pwf - BHFP (psig)", R,
  "Gauge time series. Only rows in flow-period sections are flowing pressures; build-up and gradient rows are shut-in pressures. "
  "Template holds one value per survey date. Competes with D-13 item 17 and Table 1 FBHP.",
  "Filter Section to flow periods, choose the reading that represents Pwf (e.g. last flowing reading before shut-in), then " + PSIA_TX)
S(D5, "5 BOTTOMHOLE (CASING) PRESSURE PSIA", "column", "BHP", "Pwf - BHFP (psig)", R,
  "Handwritten copy of gauge pressures. Form header is 'CASING PRESSURE PSI', with 'CASING' struck through or 'BH' added on some sheets, so the unit (psi vs psia) is ambiguous. Duplicates D-12 and competes with it.",
  "Prefer the D-12 printout. If used: confirm psia, filter to flowing rows, convert psia → psig.")
S(D15, "Flowing Bottom-hole Pressure (psi)", "row", "BHP", "Pwf - BHFP (psig)", R,
  "Flowing BHP at sample time (written 'psia'), a single point at sampling, not necessarily final Pwf. Competing source.",
  "Strip unit text, " + PSIA_TX)
S(D13, "Initial reservoir pressure", "row", "BHP", "Datum Pressure (psig)", R,
  "Initial reservoir pressure at a measurement depth (e.g. 1992.3 @ 1420 m). Template 'Datum Pressure' implies a pressure corrected to an asset datum depth, which the source does not give.",
  "Obtain the datum depth; correct with the measured gradient (D-11: 0.31 psi/ft) if the datum differs; " + PSIA_TX)
S(T1, "Pi, PSIA", "column", "BHP", "Datum Pressure (psig)", R,
  "Initial pressure at mid-perforation depth (e.g. '1999.6 @ 1426.2m'), not stated as datum-corrected.",
  "Split value/depth; datum-correct; " + PSIA_TX)

# ---- Pressure Profile
S(D13, "Date", "row", "Pressure Profile", "Date", T, "Test summary date (MM/DD/YY).",
  "Parse MM/DD/YY → YYYY-MM-DD.")
S(D13, "Initial reservoir pressure", "row", "Pressure Profile", "Reservoir Pressure (psig)", T,
  "Same meaning (initial/static reservoir pressure for the tested zone, one value per test). Units differ; source carries a depth the template cannot hold. Test 2A form has no value ('- @ -').",
  PSIA_TX + " Keep the reference depth in a note; the template has no depth field.")
S(T1, "Pi, PSIA", "column", "Pressure Profile", "Reservoir Pressure (psig)", T,
  "Initial reservoir pressure per test (competes with D-13 item 16; values differ slightly, e.g. 1999.6 vs 1999.4).",
  "Split '<value> @ <depth>m'; " + PSIA_TX)
S(HZ, "Extrapolated pressure", "row", "Pressure Profile", "Reservoir Pressure (psig)", R,
  "Horner extrapolated pressure p* (1993.261 psia at gauge depth). p* is an analysis result, not necessarily the average reservoir pressure.",
  "Confirm the asset accepts p* as reservoir pressure; then " + PSIA_TX)
S(r"^T14_D11_", "PRESSURE (psi)", "column", "Pressure Profile", "Reservoir Pressure (psig)", R,
  "Static gradient-stop pressures at 5 depths (1385-1425 m) during shut-in; the template holds one pressure per well/date with no depth.",
  "Choose the reading at the reference depth (or datum-correct); confirm psi = psia; convert to psig.")

# ---- Well Reservoir Props
S(T1, "ZONE", "column", "Well Reservoir Props", "Reservoir", R,
  "Tested zone(s) e.g. 'E-34/36'; tests 1A/2A are commingled ('E-34/36 + E-27 + E-26'). Template example codes look like 'E70'.",
  "Map zone names to the asset reservoir codes (e.g. E-34/36 → ?); split commingled tests into one row per reservoir.")
S(HZ, "Effective permeability to oil (Ko)", "row", "Well Reservoir Props", "Permeability (md)", R,
  "Effective permeability to oil (Test 1, E-34/36 only) vs template 'Permeability' (absolute vs effective not stated). Competes with Table 4 k and Table 1 K.",
  "Confirm the template expects effective permeability; value is already in md.")
S(r"^T04_", "Permeability (k)", "row", "Well Reservoir Props", "Permeability (md)", R,
  "Type-curve permeability 58.282 md vs Horner Ko 52.809 md — competing analysis results for the same zone.",
  "Choose one analysis method (record which in 'Source').")
S(T1, "K, md", "column", "Well Reservoir Props", "Permeability (md)", T,
  "Table 1 permeability per test, e.g. '53 (11.4m)'; Test 1A lists three zone values ('45, 99, 84').",
  "Split number from the bracketed net pay; for 1A split into one value per zone (E-34/36 = 45, E-27 = 99, E-26 = 84, order as printed — verify).")
S(HZ, "Formation thickness (h)", "row", "Well Reservoir Props", "Net Pay (ft)", D,
  "Same meaning: the net pay thickness used in the analysis (37.4 ft = 11.4 m, consistent with p.11 'assuming a net pay of 11.4 m'). Same unit (ft). Test 1 / E-34/36 only.",
  "None (value 37.400 ft). Tables 3 and 4 carry the identical value.")
S(T1, "K, md", "column", "Well Reservoir Props", "Net Pay (ft)", T,
  "Net pay is embedded in the K cell, e.g. '(11.4m)', '(5.0m)'.",
  "Extract the bracketed metres; ft = m × 3.28084.")
S(r"^T03_", "Permeability-thickness (kh)", "row", "Well Reservoir Props", "KH (md·ft)", D,
  "Same meaning and unit (md.ft). Horner result 1975.068 md·ft (Test 1, E-34/36).",
  "None. Competes with Table 4 kh (type-curve 2179.739) — choose one method.")
S(r"^T04_", "Permeability-thickness (kh)", "row", "Well Reservoir Props", "KH (md·ft)", D,
  "Same meaning and unit; type-curve result 2179.739 md·ft. Competes with Table 3 kh.",
  "None, but only one of Table 3 / Table 4 can populate the attribute.")
S(HZ, "Average formation porosity (0)", "row", "Well Reservoir Props", "Porosity (fraction)", D,
  "Same meaning; fraction (0.2310). Analysis input for Test 1 zone E-34/36. Tables 3 and 4 identical.", "None.")
S(HZ, "Water saturation (Sw)", "row", "Well Reservoir Props", "Water Saturation (fraction)", D,
  "Same meaning; fraction (0.5220). Analysis input for Test 1 zone E-34/36. Tables 3 and 4 identical.", "None.")

# ---- Perforation Split Ratios
S(D5, "2 REMARKS", "column", "Perforation Split Ratios", "Oil %", R,
  "Test 1A D-5 remark (p.71) 'EST. CONTRIBUTIONS: 75% FROM E-34/36, 0% FROM E-27, 25% FROM E-26' — a PLT flow split during the test, not a production allocation ratio. Competes with D-13 Test 1A remarks (74.4 / 21.6 / 4 %).",
  "Parse percentages per zone from free text; build one row per perforation with Start Ratio Date; confirm PLT split is acceptable as an allocation ratio.")
S(D13, "Remarks", "row", "Perforation Split Ratios", "Oil %", R,
  "D-13 remarks hold PLT-based splits: Test 1A q split 74.4 % / 21.6 % / 4 % (E-34/36 / E-26 / E-27); Test 2A E-6 31.2 % / E-13 68.8 %. Test-period PLT split, not allocation.",
  "Parse zone percentages from text (values 0-100); verify flagged handwriting (see _review_flags.csv).")
S(T1, "ZONE", "column", "Perforation Split Ratios", "Resv. Formation", R,
  "Zone names (E-34/36, E-27, E-26, E-6, E-13) vs template codes like 'DULG-E'.",
  "Map to asset reservoir formation / block codes.")
S(T1, "INTERVAL M, RKB", "column", "Perforation Split Ratios", "Perf. Code", N,
  "Template Perf. Code is an identifier (DULG-<well>-<string>-<block>-<reservoir>); source gives perforation depths only.", "")
S(D13, "Date", "row", "Perforation Split Ratios", "Start Ratio Date", R,
  "Date of the test from which a PLT split was derived; whether this is the ratio start date is an asset decision.",
  "Parse MM/DD/YY → YYYY-MM-DD.")

# ---- Well-test block sheets
for sheet in WELLTEST_BLOCK_SHEETS:
    date_attr = next(a for a, _ in TEMPLATE_ATTRS[sheet] if a.upper() == "DATE")
    S(D5, "1 DATE TIME", "column", sheet, date_attr, T,
      "D-5 time stamps (hhmm) within a sheet date; template holds one row per day (Excel serial date).",
      "Combine 'Sheet Date (as written)' + time → datetime; aggregate readings to calendar day; write as date.")
    S(D5, "10 RATES OIL STB/D", "column", sheet, "Oil (BOPD)", T,
      "Measured well-test oil rate (stock-tank bbl/d) — same concept as template well-test Oil BOPD, but at 15-30 min resolution during a DST, not a daily test.",
      "Average flowing-period rates per day (exclude shut-in rows and '↓' ditto marks; resolve ditto to the value above only if confirmed). STB/D → BOPD 1:1.")
    S(D7, "12 CORRECTED RATE STB/D", "column", sheet, "Oil (BOPD)", T,
      "Same values as D-5 col 10 (D-7 is the calculation sheet). Competing source for the same attribute.",
      "Use either D-5 or D-7, not both; aggregate to daily average.")
    S(D13, "Average rate", "row", sheet, "Oil (BOPD)", T,
      "One average oil rate per test (e.g. 763.2 STB/D, last 2 h). Competes with D-5/D-7 time series.",
      "Strip bracketed note; assign to the test end date; STB/D → BOPD 1:1.")
    S(D5, "11 RATES WATER B/D", "column", sheet, "Water (BWPD)", T,
      "Measured water rate (bbl/d); mostly blank or '-' in Tests 1/1A.",
      "Daily average of non-blank values; leave blank (not 0) where not measured.")
    S(D5, "10 RATES OIL STB/D", "column", sheet, "Gross (BFPD)", T,
      "No gross-liquid column in source; gross = oil + water.",
      "Gross = col 10 oil rate + col 11 water rate (same timestamp); if water is blank, gross is unknown — do not assume 0.")
    S(D13, "Water cut", "row", sheet, "WC (%)", R,
      "Water cut % per test. Template header says % but its sample values are fractions (0.6895).",
      "Confirm template convention; if fraction, WC = value / 100. 'TRACE' (Test 1A) has no numeric value.")
    S(D9, "5 SHAKE OUT BS&W", "column", sheet, "WC (%)", R,
      "Shake-out BS&W includes sediment; not identical to water cut. Competes with D-13 water cut.",
      "Use only if asset treats BS&W ≈ WC; strip '%' and convert to fraction if required.")
    S(D12, "BHP (PSIA)", "column", sheet, "Pi (psig)", N,
      "Template Pi is ESP pump-intake pressure; DST gauge BHP is a different measurement point and lift context.", "")
    S(D12, "BHCT (degF)", "column", sheet, "Tm (deg-F)", N,
      "Template Tm is ESP motor temperature; source is bottom-hole circulating temperature.", "")
    S(D13, "Length of flow period", "row", sheet, "RH (hr)", R,
      "Template RH = daily running hours; source = total flow duration of the test (e.g. '17 hrs 55 min').",
      "Parse to decimal hours (17 h 55 min → 17.92 h) and split across calendar days if used.")

# ---- ESP Monitoring
S(D5, "1 DATE TIME", "column", "ESP Monitoring", "Date", T, "Time-stamped readings vs daily rows.",
  "Combine sheet date + time; aggregate per day.")
S(D13, "Length of flow period", "row", "ESP Monitoring", "Hours", R,
  "Test flow duration vs daily operating hours.", "Parse to decimal hours; split by day.")
S(D5, "3 WELLHEAD PRESSURE PSIG/PSI", "column", "ESP Monitoring", "Wellhead WHP", R,
  "Same measurement (wellhead pressure). Context differs: template sheet monitors ESP wells daily; DST readings are minutes apart on a temporary test string. Some sheets print 'PSI' without 'G'.",
  "Confirm gauge pressure (psig); daily average of flowing readings; decide whether DST data belongs on this sheet.")
S(D13, "Average flowing wellhead pressure", "row", "ESP Monitoring", "Wellhead WHP", R,
  "One average flowing WHP (psig) per test. Competes with D-5 col 3.", "Assign to test date; units already psig.")
S(D5, "4 WELLHEAD TEMPERATURE °F", "column", "ESP Monitoring", "Wellhead WHT", R,
  "Same measurement and unit (°F); same context concern as WHP.", "Daily average of flowing readings.")
S(D13, "Average flowing wellhead temperature", "row", "ESP Monitoring", "Wellhead WHT", R,
  "One average per test (°F). Competes with D-5 col 4.", "Assign to test date.")
S(D5, "2 REMARKS", "column", "ESP Monitoring", "Remarks", R,
  "Free-text operational remarks.", "Concatenate per day if used.")
S(D15, "Separator Pressure", "row", "ESP Monitoring", "Header Manifold Press.", N,
  "Test separator pressure is not a header manifold pressure.", "")
S(D12, "BHP (PSIA)", "column", "ESP Monitoring", "ESP Pump Pi", N,
  "Pump intake pressure (ESP) ≠ DST bottom-hole gauge pressure.", "")

# ---- explicit non-matches for related concepts (measured test vs allocated / field level)
S(D5, "10 RATES OIL STB/D", "column", "Allocated Production", "Oil PD Rate (bbls/d)", N,
  "Measured DST test rate is not an allocated producing-day rate.", "")
S(D7, "13 CORRECTED GOR SCF/STB", "column", "Allocated Production", "GOR (scf/bbl)", N,
  "Measured test GOR (separator, per reading) vs allocated GOR.", "")
S(D8, "11 RATE Q = .024 C' x √(hw pf) k (Mcf/D)", "column", "Allocated Production", "Gas PD Rate (MMscf/d)", N,
  "Measured test gas rate vs allocated producing-day rate (unit Mcf/d vs MMscf/d also differs).", "")
S(D7, "11 CORRECTED CUM. PROD STB", "column", "Allocated Production", "Np (MMbbl)", N,
  "Cumulative test production (STB, mostly blank) is not field/well cumulative allocated production.", "")
S(D13, "Average rate", "row", "Field Production", "Oil Obs (bopd)", N,
  "Well-level test rate vs field-level daily observed production.", "")
S(D13, "Cumulative production", "row", "Stock Oil", "Prod Obs", N,
  "Test cumulative volume vs daily field stock production.", "")


# --------------------------------------------------------------------------- build
def main():
    t_hash = sha(TEMPLATE)
    c_hash = {p.name: sha(p) for p in CSV_DIR.glob("*.csv")}

    twb = load_workbook(TEMPLATE, read_only=True, data_only=True)
    assert [s for s in TEMPLATE_ATTRS] == twb.sheetnames, f"sheet list differs: {twb.sheetnames}"
    TEMPLATE_ATTRS["Daily Chemical Operation New"] = flatten_daily_chemical(twb)
    twb.close()
    tunit = {(s, a): u for s, attrs in TEMPLATE_ATTRS.items() for a, u in attrs}

    data = load_csvs()
    src = source_attributes(data)
    src_keys = {(c, a) for c, a, _, _ in src}
    src_unit = {(c, a): u for c, a, u, _ in src}

    mapping = []
    for pat, attr, kind, sheet, tattr, status, reason, tx in SPECS:
        assert (sheet, tattr) in tunit, f"unknown template attribute {sheet} / {tattr}"
        # Form labels sometimes carry a bracketed qualifier, e.g.
        # 'Initial reservoir pressure (from Buildup)' on the Test 1 D-13 form.
        hits = []
        for c in data:
            if not re.search(pat, c):
                continue
            label = next((a for (cc, a) in src_keys if cc == c and (a == attr or a.startswith(attr + " ("))), None)
            if label:
                hits.append((c, label))
        assert hits, f"no source CSV for {pat} / {attr}"
        for c, attr_c in hits:
            d = data[c]
            pg, sample = evidence(d, attr_c, kind)
            idx = d["index"]
            mapping.append([c, f"{idx.get('Table Title (as printed)', '')} | Test {idx.get('Test', '')}",
                            attr_c, src_unit[(c, attr_c)], sheet, tattr, tunit[(sheet, tattr)], status, reason, tx,
                            f"Source {pg}; sample: {sample or '(none)'}"])

    # competing-source flags: same template attribute fed by >1 source attribute (status D/T/R)
    feeders = {}
    for m in mapping:
        if m[7] != N:
            feeders.setdefault((m[4], m[5]), set()).add(re.sub(r"^T\d+_", "", m[0]).split("_Test")[0] + ":" + m[2])
    for m in mapping:
        f = feeders.get((m[4], m[5]), set())
        if m[7] != N and len(f) > 1:
            m[8] += f" [COMPETING SOURCES for this template attribute: {len(f)} — " + "; ".join(sorted(f)) + "]"

    # best status per template attribute
    best = {}
    for m in mapping:
        k = (m[4], m[5])
        if k not in best or RANK[m[7]] < RANK[best[k]]:
            best[k] = m[7]

    gaps = []
    for s, attrs in TEMPLATE_ATTRS.items():
        if s in INSTRUCTION_SHEETS:
            gaps.append([s, "(instruction sheet — no data attributes)", "Sheet contains upload instructions only."])
            continue
        for a, _ in attrs:
            b = best.get((s, a))
            if b in (D, T):
                continue
            if b == R:
                why = "Only possible matches found (needs review) — see Attribute Mapping."
            elif b == N:
                why = "Related source data exists but is a different concept / level — see Attribute Mapping (No match rows)."
            else:
                why = gap_reason(s, a)
            gaps.append([s, a, why])

    mapped = {(m[0], m[2]) for m in mapping if m[7] != N}
    unmapped = []
    for c, a, u, kind in src:
        if (c, a) in mapped:
            continue
        unmapped.append([c, a, u, unmapped_note(c, a)])

    summary = []
    for s, attrs in TEMPLATE_ATTRS.items():
        st = [best.get((s, a), N) for a, _ in attrs]
        summary.append([s, len(attrs), st.count(D), st.count(T), st.count(R), st.count(N)])

    write_book(mapping, gaps, unmapped, summary)

    assert sha(TEMPLATE) == t_hash, "TEMPLATE CHANGED"
    assert all(sha(CSV_DIR / n) == h for n, h in c_hash.items()), "CSV CHANGED"
    print(f"mapping rows: {len(mapping)} | template gaps: {len(gaps)} | unmapped source attrs: {len(unmapped)}")
    print("template and CSVs unchanged (sha256 verified)")
    print(f"Saved: {OUT}")


def gap_reason(sheet, attr):
    reasons = {
        "Summary": "Index of template files — no data attribute to map.",
        "Chemical Usage": "No chemical injection data in a DST report.",
        "Daily Chemical Operation New": "No chemical injection data in a DST report.",
        "ESP Brand": "DST well flowed naturally; no ESP equipment.",
        "Spatial Data": "No well coordinates in the extracted tables.",
        "Stock Oil": "Field stock/tank inventory is not in a DST report.",
        "Tank": "Tank master data is not in a DST report (test gauge tank is temporary test equipment).",
        "Water Injection Monitoring": "No water injection in a production DST.",
        "Production Deferment": "No deferment records in a DST report.",
        "Allocated Production": "DST provides measured test data, not allocated production.",
        "Field Production": "DST data is well-level, not field-level.",
        "Fields": "Field list is reference data; 'DULANG' does not appear in the template list (Angsi … West Field) and is not a CSV attribute.",
    }
    if sheet in reasons:
        return reasons[sheet]
    if attr.startswith(("Freq", "Ampere", "Voltage", "ESP Pump")):
        return "ESP electrical/pump parameter — DST well had no ESP."
    if attr in ("No", "No.") :
        return "Template row counter — not a source attribute."
    if attr in ("Total (block)", "Online Wells"):
        return "Multi-well field aggregate; DST covers one well."
    if attr in ("Header Manifold Temp.",):
        return "No production header manifold measurement in DST data."
    specific = {
        ("Perforation Split Ratios", "Well Type"): "Well type code (e.g. OP) not stated in the tables.",
        ("Perforation Split Ratios", "Well Status"): "Well status not stated in the tables.",
        ("Perforation Split Ratios", "Resv. Block Formation"): "Reservoir block codes not in the source.",
        ("Perforation Split Ratios", "Resv. Block"): "Reservoir block codes not in the source.",
        ("Perforation Split Ratios", "Gas %"): "PLT splits in the source are oil-rate splits only.",
        ("Perforation Split Ratios", "Cond %"): "No condensate split in the source.",
        ("Perforation Split Ratios", "Water %"): "No water split in the source.",
        ("Perforation Split Ratios", "Gas Inj. %"): "Not applicable (production test).",
        ("Perforation Split Ratios", "Water Inj. %"): "Not applicable (production test).",
        ("Well Reservoir Props", "Source"): "Provenance label (e.g. 'DST Horner') would be a new value, not an extracted attribute — user decision.",
        ("Well Reservoir Props", "Sand"): "Sand/sub-unit not stated in the tables.",
    }
    return specific.get((sheet, attr), "No equivalent attribute in the extracted DST tables.")


def unmapped_note(c, a):
    if "_D10_" in c:
        return "Gas composition (C1-C5, CO2, H2S) and pour point — no template attribute for gas/fluid composition."
    if "_D8_" in c:
        return "Gas-meter orifice calculation input/result — no template attribute for test gas measurement (only allocated gas)."
    if "_D9_" in c:
        return "Liquid sample field analysis — no template attribute for sample properties (BS&W considered for WC)."
    if "_D12_" in c:
        return "Gauge printout context (section/time/event) — template has no high-frequency pressure-time table."
    if "_D1_" in c or "_D1A_" in c:
        return "Completion / perforation job detail — no template attribute for completion or perforation records."
    if "_D15_" in c:
        return "Separator sample record — no template attribute for PVT sample metadata."
    if "_D7_" in c:
        return "Oil rate calculation factor — intermediate value, no template attribute."
    if c.startswith("T02_"):
        return "Sample inventory per test — no template attribute."
    if c.startswith(("T03_", "T04_")):
        return "Pressure-transient analysis input/result with no template attribute."
    if "_D11_" in c:
        return "Gradient survey detail — depth/gradient have no template fields."
    if "_D5_" in c:
        return "Test data sheet field with no template counterpart (choke, cumulative, gravity, sheet date context)."
    if "_D13_" in c:
        return "Test summary item with no template counterpart."
    return "No template attribute with the same meaning."


def write_book(mapping, gaps, unmapped, summary):
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    sheets = [
        ("Attribute Mapping", ["Source CSV", "Source Table/Test", "Source Attribute", "Source Unit", "Template Sheet",
                               "Template Attribute", "Template Unit", "Match Status", "Reason",
                               "Required Transformation", "Evidence"], mapping,
         [34, 40, 34, 16, 24, 30, 18, 24, 70, 60, 50]),
        ("Template Gaps", ["Template Sheet", "Template Attribute", "Reason No Confirmed Match"], gaps, [28, 44, 90]),
        ("Unmapped Source Attributes", ["Source CSV", "Source Attribute", "Source Unit", "Notes"], unmapped,
         [42, 44, 22, 80]),
        ("Summary", ["Template Sheet", "Total Attributes", "Direct Matches", "Requires Transformation",
                     "Needs Review", "No Match"], summary, [32, 16, 16, 22, 16, 12]),
    ]
    fills = {D: "C6EFCE", T: "FFEB9C", R: "FCE4D6", N: "EDEDED"}
    first = True
    for title, hdr, rows, widths in sheets:
        ws = wb.active if first else wb.create_sheet()
        first = False
        ws.title = title
        ws.append(hdr)
        for r in rows:
            ws.append(r)
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill("solid", fgColor="DDEBF7")
        for i, w in enumerate(widths):
            ws.column_dimensions[chr(65 + i)].width = w
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(wrap_text=True, vertical="top")
            if title == "Attribute Mapping":
                row[7].fill = PatternFill("solid", fgColor=fills[row[7].value])
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
    s = wb["Summary"]
    n = s.max_row
    s.append(["TOTAL"] + [f"=SUM({c}2:{c}{n})" for c in "BCDEF"])
    for cell in s[s.max_row]:
        cell.font = Font(bold=True)
    notes = wb.create_sheet("Notes")
    for line in [
        "Scope: 37 CSVs extracted from 'DULANG-7_EPMI Well Test Report DST 1,1A,2,2A_ 01 Dec. 1990.pdf' vs every sheet of asset_combined_templates.xlsx.",
        "Templates were read for headers/structure only; nothing was populated. Original CSVs and the template are unchanged (sha256 verified).",
        "Summary counts use the best status per template attribute (Direct > Requires transformation > Needs review > No match); attributes with no candidate count as No match.",
        "Form-style CSVs (D-1, D-13, D-15, Tables 3/4) store one attribute per row; those row labels are treated as source attributes.",
        "'↓' / '\"' in source CSVs are ditto marks kept as written; they must be resolved before any value is loaded.",
        "Pressure conversion psia → psig requires an agreed atmospheric pressure (14.696 psi assumed in notes, to be confirmed).",
        "Values flagged in table_csv/_review_flags.csv should be verified before use.",
    ]:
        notes.append([line])
    notes.column_dimensions["A"].width = 150
    wb.save(OUT)


if __name__ == "__main__":
    main()
