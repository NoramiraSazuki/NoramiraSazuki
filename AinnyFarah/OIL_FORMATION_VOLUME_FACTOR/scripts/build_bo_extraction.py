#!/usr/bin/env python3
"""Build PVT_Bo_oil_volume_extraction.xlsx: Bo / oil volume records from the PVT Fluid Analysis Document folder.

Every number is read from its source cell at run time (local copies in _work_pvt_scan/src; originals are read only and
re-hashed). Record definitions below say WHERE to read and how the record is classified; they were set after reviewing
each file (cell text, text boxes, document text and OCR of embedded pictures - see _work_pvt_scan/hits_review.txt).
"""
import datetime as dt
import hashlib
import io
import json
import re
from pathlib import Path

import openpyxl
import xlrd
from openpyxl.styles import Alignment, Font, PatternFill

BASE = Path(__file__).resolve().parent.parent
SRC = BASE / "_work_pvt_scan" / "src"
ORIG = Path("/Volumes/homes/public/DulangOSS/20260925/DOCUMENTS/PVT Fluid Analysis Document")
OUT = BASE / "PVT_Bo_oil_volume_extraction.xlsx"

_cache = {}


def src_file(key):
    f = [p for p in SRC.iterdir() if key in p.name]
    assert len(f) == 1, (key, f)
    return f[0]


def cell(key, sheet, addr):
    """value of one cell (data_only) from the local copy of the file whose name contains `key`"""
    p = src_file(key)
    if p not in _cache:
        if p.read_bytes()[:2] == b"PK":
            _cache[p] = ("x", openpyxl.load_workbook(io.BytesIO(p.read_bytes()), data_only=True))
        else:
            _cache[p] = ("b", xlrd.open_workbook(str(p)))
    kind, wb = _cache[p]
    if kind == "x":
        return wb[sheet][addr].value
    s = wb.sheet_by_name(sheet)
    m = re.match(r"([A-Z]+)(\d+)", addr)
    col = 0
    for ch in m.group(1):
        col = col * 26 + ord(ch) - 64
    return s.cell_value(int(m.group(2)) - 1, col - 1)


def num(v):
    """source value -> (float, note). Handles typed spaces ('1. 242') and markers ('1780*')."""
    if isinstance(v, (int, float)):
        return float(v), ""
    s = str(v).strip()
    clean = re.sub(r"(?<=\d)\.\s+(?=\d)", ".", s).rstrip("*").replace("Pi=", "").replace("Pb=", "").replace("Pi =", "").replace("Pb =", "").strip()
    try:
        return float(clean), (f"source text '{s}' read as {clean}" if clean != s else "")
    except ValueError:
        return None, f"unreadable '{s}'"


ROWS = []


def add(**k):
    ROWS.append(k)


# --------------------------------------------------------------------------------------------- DIRECT Bo records
def flopetrol_dv(key, well, res, date, sheet, prow, bcol, rows, tres, psat, basis_note, review="", review_all=""):
    """Flopetrol differential-vaporisation table: oil volume factor per pressure step.
    `review` is flagged on the Pi/Pb rows only, `review_all` on every row."""
    for r in rows:
        p, pn = num(cell(key, sheet, f"B{r}"))
        b, bn = num(cell(key, sheet, f"{bcol}{r}"))
        label = str(cell(key, sheet, f"B{r}"))
        cond = "initial reservoir pressure (Pi)" if "Pi" in label else ("bubble point (Pb)" if "Pb" in label else "")
        add(file=key, well=well, res=res, date=date, bo=b, flag="DIRECT", p=p, t=tres,
            basis=f"Differential vaporisation at {tres}°F - oil volume factor per standard barrel of residual oil" + (f"; {cond}" if cond else ""),
            sheet=sheet, cells=f"B{r}, {bcol}{r}", bo_units="bbl/Std bbl (residual oil)",
            review="; ".join(x for x in (review if cond else "", review_all) if x),
            notes="; ".join(x for x in (basis_note, pn, bn) if x))


def sep_test(key, well, res, date, sheet, items, psat, tres, footcell, review="", extra=""):
    """separator / flash tests: items = [(bo_cell, stage_desc, cells_desc)]"""
    for bo_cell, stage, cells in items:
        b, bn = num(cell(key, sheet, bo_cell))
        add(file=key, well=well, res=res, date=date, bo=b, flag="DIRECT", p=psat, t=tres,
            basis=f"Separator test (Bofb): saturated oil at {psat:g} psig and {tres:g}°F per barrel of stock-tank oil at 60°F; separator {stage}",
            sheet=sheet, cells=cells, bo_units="bbl/STB", review=review,
            notes="; ".join(x for x in (f"definition: {cell(key, sheet, footcell)}" if footcell else "", bn, extra) if x))


def adjusted(key, well, res, date, sheet, pcol, bocol, rows, tres, basis, review_rows=None, skip=None, extra=""):
    for r in rows:
        if skip and r in skip:
            continue
        p, pn = num(cell(key, sheet, f"{pcol}{r}"))
        b, bn = num(cell(key, sheet, f"{bocol}{r}"))
        add(file=key, well=well, res=res, date=date, bo=b, flag="DIRECT", p=p, t=tres, basis=basis,
            sheet=sheet, cells=f"{pcol}{r}, {bocol}{r}", bo_units="bbl/STB",
            review=(review_rows or {}).get(r, ""), notes="; ".join(x for x in (pn, bn, extra) if x))


K004 = "82_L_004"
W004 = "DULANG 6G-1-2 DST 1"
R004 = "Producing zone 1272-1301 m (reservoir name not stated)"
D004 = "1983-01-03"
N004 = ("Date from report cover 'Samp. date: 03/01/1983' (visually checked; dd/mm). Typed sheet cells show 1983-03-01 / 1983-04-01 "
        "(day and month swapped on entry).")
flopetrol_dv(K004, W004, R004, D004, "TABLE 38", "B", "C", range(8, 25), 206, 1700,
             "Table heading 'Oil volume (bbl/Std bbl)'; SUMMARY RESULTS row 19 calls it 'oil volume factor'. " + N004,
             review="SUMMARY RESULTS D19 gives 1.245 at Pb but TABLE 38 C18 gives 1.254 (and 1.253 at Pi); 1.254 kept - check original report")
sep_test(K004, W004, R004, D004, "TABLE 41",
         [("G8", "200 psig / 100°F then tank", "B8:C8, G8"), ("G9", "150 psig / 100°F then tank", "B9:C9, G9"),
          ("G10", "100 psig / 100°F then tank", "B10:C10, G10"), ("G11", "50 psig / 100°F then tank", "B11:C11, G11"),
          ("G12", "0 psig / 100°F (single flash)", "B12:C12, G12")], 1700, 206, "C16",
         extra="Heading 'OIL VOLUME FACTOR (2)'. Saturation pressure 1700 psig at 206°F from SUMMARY RESULTS. " + N004)

K034 = "82_L_034"
W034 = "DULANG 6G-1-1B DST 4"
R034 = "Producing zone 1252-1259 m (reservoir name not stated)"
D034 = "1982-09-03"
N034 = ("Bottom-hole sample date typed as 1982-03-09 in 'SAMPLING COND.' D35; read as 03/09/1982 (dd/mm) like the other Flopetrol "
        "covers (report dated September 1982). Cover page has no sampling date.")
flopetrol_dv(K034, W034, R034, D034, "TABLE 13", "B", "C", range(7, 21), 203, 1660,
             "Heading 'Oil volume factor Bo (bbl/Std bbl)'; cylinder 15089/12 (FLO 350). " + N034,
             review_all="date interpreted (dd/mm) - verify")
sep_test(K034, W034, R034, D034, "TABLE 16",
         [("G8", "200 psig / 100°F then tank", "B8:C8, G8"), ("G9", "150 psig / 100°F then tank", "B9:C9, G9"),
          ("G10", "100 psig / 100°F then tank", "B10:C10, G10"), ("G11", "50 psig / 100°F then tank", "B11:C11, G11")],
         1660, 203, "B15", review="date interpreted (dd/mm) - verify", extra="Heading 'Formation Vol Factor (2)'. " + N034)

K005 = "83_L_005"
W005 = "DULANG 6G-1-2 DST 2"
R005 = "Zone 1252-1260 m (cover has handwritten 'E-1' - not confirmed as reservoir name)"
D005 = "1983-01-07"
N005 = ("Bottom-hole sample 20524/189 (used for complete PVT study) dated 1983-01-07 in 'BOTTOMHOLE SAMPLE (5)' C38; cover "
        "'Samp. date: 06/01/1983' (visually checked) refers to the surface samples. Differential-vaporisation table pages are not in this file; "
        "values come from the summary only.")
for col, cond, pcell in (("C22", "initial reservoir pressure (Pi)", "D14"), ("D22", "bubble point (Pb)", "D17")):
    b, bn = num(cell(K005, "SAMPLE SUMMARY", col))
    ptxt = str(cell(K005, "SAMPLE SUMMARY", pcell))
    p = float(re.match(r"[\d.]+", ptxt).group(0))
    add(file=K005, well=W005, res=R005, date=D005, bo=b, flag="DIRECT", p=p, t=209,
        basis=f"Differential vaporisation summary - oil volume factor at {cond}", sheet="SAMPLE SUMMARY",
        cells=f"{pcell}, {col}", bo_units="bbl/std bbl (residual oil)",
        review=("Summary Pi 1871.3 psig / Pb 1530 psig at 209°F differ from the sample sheets (static 1777 psig, Pb 1768 psig at 203°F); "
                "temperature typed '20 9 F' - check original"),
        notes="; ".join(x for x in (N005, bn, f"pressure text '{ptxt}'") if x))

KB12 = "91184_DULANG-B12_Reservoir Fluid Study_Apr 1992.xlsx"
WB12 = "DULANG B-12"
RB12 = "LOWER-D formation"
DB12 = "1991-09-21"
NB12 = ("Same values in the Word version (MFL 91184 ... .docx: separator-test table and adjustment pages) - counted once. "
        "Sample date from 'SAMPLE SUMMARY' C5. Reservoir 1727 psig at 200°F; laboratory bubble point 1367 psig adjusted to 1668 psig.")
sep_test(KB12, WB12, RB12, DB12, "Single Stage Separator Test",
         [("G14", "400 psig / 112°F then 0 psig / 112°F", "B13:C14, G14"), ("G18", "250 psig / 112°F then 0 psig / 112°F", "B17:C18, G18"),
          ("G22", "150 psig / 112°F then 0 psig / 112°F", "B21:C22, G22"), ("G26", "50 psig / 112°F then 0 psig / 112°F", "B25:C26, G26")],
         1668, 200, "B32", extra=NB12 + " Excel footnote says '1668 psig and 60°F'; Word version says 200°F (reservoir temperature) - 200°F used.")
for col, sep in (("C", "400 psig / 112°F"), ("D", "250 psig / 112°F"), ("E", "150 psig / 112°F"), ("F", "50 psig / 112°F")):
    adjusted(KB12, WB12, RB12, DB12, "Adjustment Oil", "B", col, range(16, 23), 200,
             f"Differential data adjusted to separator conditions ({sep} then 0 psig): Bo = Bod x (Bofb/Bodb), bbl at indicated pressure and 200°F per STB",
             extra=NB12 + " Row at 1668 psig equals Bofb (separator-test record) and is not repeated.")

K21 = "MFL 20120023"
W21 = "DULANG B21L"
R21 = ""
D21 = "2012-01-28"
N21 = ("Recombined reservoir fluid (cylinder TS 100002) from separator liquid 813785 and gas 2374, sampled 28 Jan 2012 ('Sample Info'). "
       "Reservoir name not stated in the report.")
b, bn = num(cell(K21, "A.1", "AR39"))
add(file=K21, well=W21, res=R21, date=D21, bo=b, flag="DIRECT", p=1100, t=206,
    basis="Multi-stage separator test Bofb: saturated fluid at 1100 psig and 206°F per STB at 60°F; stages 144 psig/115°F, 41 psig/138°F, 5 psig/135°F, 0 psig/60°F",
    sheet="A.1 (summary); G.1 H16", cells="AN39:AR39; G.1 H16", bo_units="bbl/STB",
    notes=N21 + f" G.1 H16 holds the unrounded value {cell(K21, 'G.1', 'H16')!r}; report shows 1.195.")
for r in list(range(16, 28)) + list(range(29, 35)):
    p, _ = num(cell(K21, "E.1", f"AQ{r}"))
    braw = cell(K21, "E.1", f"AU{r}")
    b = round(float(braw), 3)
    add(file=K21, well=W21, res=R21, date=D21, bo=b, flag="DIRECT", p=p, t=206,
        basis="Differential data converted to production separator conditions (Bo(1)), bbl at indicated pressure and 206°F per STB",
        sheet="E.1", cells=f"AQ{r}, AU{r}", bo_units="bbl/STB",
        notes=N21 + f" Cell value {braw!r} shown to 3 decimals as in the report. Row at 1100 psig (saturation) equals Bofb and is not repeated.")

K034b = "SFL 82034"
W1 = "DULANG-1 (Test 2)"
sep_test(K034b, W1, "", "", "SINGLE SEPARATOR TEST", [("G8", "265 psig / 148°F then 0 psig / 148°F", "B6:C8, G8")], 1825, 218,
         "", extra="Definition (MULTI SEPARATOR TEST B17): saturated oil at 1825 psig and 218°F per STB at 60°F. Bottom-hole sample 800116. Sampling date and reservoir not stated in the file.")
sep_test(K034b, W1, "", "", "MULTI SEPARATOR TEST", [("G10", "300 psig / 150°F, 100 psig / 140°F, 0 psig / 130°F", "B6:C10, G10")], 1825, 218,
         "B17", extra="Bottom-hole sample 800116. Sampling date and reservoir not stated in the file.")

K056 = "SFL 82056_"
W3 = "DULANG-3 (Test 1)"
sep_test(K056, W3, "E7 (per 'Page 1' J6)", "1982-03-10", "Page 7",
         [("G10", "300 psig / 140°F, 100 psig / 130°F, 0 psig / 120°F", "B6:C10, G10")], 1982, 224, "B17",
         review="date and reservoir need confirmation",
         extra=("'Date Sampled' typed 1982-10-03 (Page 1 I2) is after the report date (10 Sept 1982); read as 10/03/1982 (dd/mm). "
                "SFL 82056A (same test, same interval 4744-4760, 4772-4780 ft) names the formation E40/42."))

K142 = "SFL 82142"
W7 = "DULANG 6G-7-1 DST 4"
sep_test(K142, W7, "", "1982-07-05", "SEPARATOR TEST",
         [("H15", "200 psig / 100°F then 0 psig / 100°F", "C13:D15, H15"), ("H19", "150 psig / 100°F then 0 psig / 100°F", "C17:D19, H19"),
          ("H23", "100 psig / 100°F then 0 psig / 100°F", "C21:D23, H23"), ("H27", "50 psig / 100°F then 0 psig / 100°F", "C25:D27, H27")],
         1562, 181, "C32", review="sample date may be day/month swapped (5 Jul or 7 May 1982) - verify",
         extra="Bottom-hole sample 9024-37. 'Sample Date' cell B5 = 1982-07-05 (report 22 Dec 1982). Reservoir not stated.")

K143 = "SFL 82143"
sep_test(K143, "DULANG 6G-7-1 DST 3", "", "", "PAGE 9 of 13",
         [("G9", "200 psig / 100°F then 0 psig / 100°F", "B7:C9, G9"), ("G12", "150 psig / 100°F then 0 psig / 100°F", "B10:C12, G12"),
          ("G15", "100 psig / 100°F then 0 psig / 100°F", "B13:C15, G15"), ("G19", "50 psig / 100°F then 0 psig / 100°F", "B16:C19, G19")],
         1593, 203, "B25", review="footnote says '1593 psia' and 'stock tank at 69°F' (likely 60°F) - check original",
         extra="Separator samples (gas 1116-81 used). Sampling date and reservoir not stated in the file.")

for key, well, res, sheet_p, rows, tres, pb_row, pi_row, extra in (
        ("SFL 86024", "DULANG-5 (Test 1)", "E-14 (1264-1273 m)", "ADJ.DATA (PAGE 27)", range(5, 14), 211, 8, 5,
         "Pi 1807 psig (reported reservoir pressure) and Pb 1225 psig at 211°F."),
        ("SFL 86026", "DULANG-5 (Test 2)", "E-8/13 (1246.5-1255.5 m)", "ADJ.DATA (PAGE 27)", range(5, 15), 205, 8, 5,
         "Pi 1780 psig (reported reservoir pressure) and Pb 1370 psig at 205°F. Footnote 4 still says 211°F (copied from Test 1).")):
    adjusted(key, well, res, "", sheet_p, "B", "D", rows, tres,
             "Differential data adjusted to separator conditions 205 psig/160°F then 0 psig/140°F (Formation Vol Factor Bo): barrels at reservoir conditions per STB at 60°F",
             review_rows={pi_row: "", pb_row: ""},
             skip={8} if key == "SFL 86026" else None,
             extra=extra + " Sampling date not stated in the file." + (" Row at 1370 psig equals the multi-stage separator Bofb (separate record) and is not repeated." if key == "SFL 86026" else ""))
sep_test("SFL 86024", "DULANG-5 (Test 1)", "E-14 (1264-1273 m)", "", "ATMOSPHERIC RFS",
         [("F9", "flash from 5000 psig/120°F to 0 psig/140°F (as printed)", "B7:C9, F9")], 1225, 211, "B16",
         extra="Single atmospheric flash; differs from the adjusted Bo at Pb (1.170) because separator conditions differ. Sampling date not stated.")
sep_test("SFL 86026", "DULANG-5 (Test 2)", "E-8/13 (1246.5-1255.5 m)", "", "MULTI SEPARATOR TEST",
         [("G14", "205/160, 150/135, 145/150, 50/149 psig/°F then 0 psig/140°F", "B6:C14, G14")], 1370, 205, "B21",
         extra="Equals adjusted Bo at Pb (ADJ.DATA D8). Sampling date not stated.")
sep_test("SFL 86026", "DULANG-5 (Test 2)", "E-8/13 (1246.5-1255.5 m)", "", "ATMOSPHERIC RFS",
         [("F9", "flash from 5000 psig/120°F to 0 psig/140°F (as printed)", "B7:C9, F9")], 1370, 205, "B16",
         extra="Single atmospheric flash. Sampling date not stated.")

# --------------------------------------------------------------- records without Bo (REQUIRES CALCULATION)
INS = "REQUIRES CALCULATION – INSUFFICIENT DATA"
K006 = "83_L_006"
W006 = "DULANG 6G-1-2 DST 3"
R006 = "Producing zone 1222-1239 m (reservoir name not stated)"
for sheet, sc, cyl, d, note in (
        ("TABLE 2 & 3", "C41", "separator liquid 1116/55", "1983-01-09", "Separator 185 psig / 110°F."),
        ("TABLE 5 & 6", "C44", "separator liquid 1116/588", "1983-01-09", "Separator 185 psig / 111°F."),
        ("TABLE 8 & 9", "C46", "separator liquid 1116/508", "1983-01-09", "Separator 185 psig / 112°F; recombination table 28 labels it 1116/509.")):
    add(file=K006, well=W006, res=R006, date=d, bo=None, flag=INS, p=None, t=None,
        basis=f"Flash of {cyl} to stock-tank conditions: shrinkage factor {cell(K006, sheet, sc)}",
        sheet=sheet, cells=sc, bo_units="", volst="stock-tank volume given only as a ratio to separator-liquid volume (shrinkage factor, Std bbl/bbl)",
        notes=(f"{note} Missing: oil volume at reservoir pressure and temperature for this sample (separator conditions are not reservoir "
               "conditions), so Bo cannot be calculated. Date from cover 'Samp. date: 09/01/1983' (visually checked)."))
for sheet, cells_, cyl, d in (("TABLE 13 & 14", "C7:C8", "bottom-hole samples 9214/296 and 1116/33", "1983-01-10"),
                              ("TABLE 17 & 18", "C7:C8", "bottom-hole samples 80291/242 and 80291/5", "1983-01-10")):
    add(file=K006, well=W006, res=R006, date=d, bo=None, flag=INS, p=None, t=200, basis=f"Flash of {cyl} to stock-tank conditions (GOR and stock-tank gravity only)",
        sheet=sheet, cells=cells_, bo_units="",
        review="date interpreted (dd/mm) - verify",
        notes=("Missing: reservoir-condition oil volume and stock-tank oil volume (only GOR and gravity are reported). No differential "
               "or separator test in this report. Bottom-hole sample date typed 1983-10-01, read as 10/01/1983 (dd/mm)."))
for key, well, res, d, sheet, cells_, desc, rv in (
        ("SFL 82055", "DULANG-3 (Test 1)", "", "1982-03-10", "Page 4", "B11, E11", "CCE at 224°F, saturation pressure 2154 psig", "date typed 1982-10-03 is after the report (24 Aug 1982); read as dd/mm - verify"),
        ("SFL 82144", "DULANG 6G-7-1 DST 2", "", "", "PAGE 4 of 4", "B15, E15", "CCE at 206°F, bubble point 1492 psig", ""),
        ("SFL 82145", "DULANG 6G-7-1 DST 4", "", "", "PAGE 5 of 5", "PAGE 4 of 5 G5", "CCE at 196°F, bubble point 1410 psig", ""),
        ("SFL 82146", "DULANG 6G-7-1 DST 3", "", "", "PAGE 4 of 4", "B15, E15", "CCE at 206°F, bubble point 1173 psig", "")):
    add(file=key, well=well, res=res, date=d, bo=None, flag=INS, p=None, t=None, basis=f"Pressure-volume relation only ({desc}); relative volume V/Vsat",
        sheet=sheet, cells=cells_, bo_units="", review=rv,
        volres="relative volume V/Vsat only (barrels at indicated pressure per barrel at saturation pressure)",
        notes="Missing: stock-tank oil volume (no separator/flash test in this file), so Bo cannot be calculated. Sampling date and reservoir not stated unless shown.")


# ------------------------------------------------------------------------------------------------- write
def main():
    if OUT.exists():
        raise SystemExit(f"{OUT} exists - not overwritten")
    files = sorted(p for p in ORIG.iterdir() if p.is_file() and not p.name.startswith((".", "._")))
    before = {p.name: hashlib.sha256(p.read_bytes()).hexdigest() for p in files}
    for r in ROWS:
        r["fname"] = src_file(r["file"]).name
        if r["flag"] == "DIRECT":
            assert r["bo"] is not None and 0.9 < r["bo"] < 2.0, r
    # duplicate check: same file, same basis, same pressure, same Bo
    seen = set()
    for r in ROWS:
        k = (r["fname"], r["basis"], r.get("p"), r["bo"], r["sheet"])
        assert k not in seen, k
        seen.add(k)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Bo records"
    hdr = ["FILE NAME", "WELL NAME", "RESERVOIR", "DATE", "OIL VOLUME AT RESERVOIR CONDITIONS", "OIL VOLUME AT STOCK-TANK CONDITIONS", "Bo",
           "CALCULATION FLAG", "Volume units", "Bo units", "Pressure (psig)", "Temperature (°F)", "Test / Bo definition",
           "Source sheet / page", "Source cell(s)", "Calculation formula", "Review flag", "Notes"]
    ws.append(hdr)
    for r in ROWS:
        ws.append([r["fname"], r["well"], r["res"] or None, r["date"] or None, r.get("volres"), r.get("volst"), r["bo"], r["flag"],
                   ("ratio only - no absolute volumes reported" if r["flag"] == "DIRECT" else None), r["bo_units"] or None,
                   r.get("p"), r.get("t"), r["basis"], r["sheet"], r["cells"],
                   ("n/a - Bo reported directly" if r["flag"] == "DIRECT" else "Bo = V(reservoir) / V(stock tank) - not possible, see notes"),
                   r.get("review") or None, r["notes"] or None])
    fill = {"DIRECT": "C6EFCE", "CALCULATED": "BDD7EE", INS: "FFEB9C"}
    for row in ws.iter_rows(min_row=2):
        row[7].fill = PatternFill("solid", fgColor=fill[row[7].value])
        if row[16].value:
            row[16].fill = PatternFill("solid", fgColor="F8CBAD")
        row[6].number_format = "0.000"
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    for c in ws[1]:
        c.font = Font(bold=True)
        c.alignment = Alignment(wrap_text=True, vertical="top")
    for i, w in enumerate([46, 22, 26, 11, 24, 24, 8, 22, 18, 16, 10, 10, 50, 22, 16, 22, 34, 70], 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = ws.dimensions

    # files processed
    fs = wb.create_sheet("Files processed")
    fs.append(["#", "File name", "Records", "DIRECT", "REQUIRES CALCULATION", "Outcome"])
    why = {
        "82_L_033": "Cover page only - no data", "82_L_035": "Gas sample (dew point study); only separator-liquid shrinkage - no oil Bo",
        "CHEM 86021": "Water analysis", "SAMPLE INFORMATION": "Sample information form", "MFL 20100151_DULANG-A14_Reservoir": "Gas-condensate study (dew point) - no oil Bo",
        "MFL 20110028_DULANG-A25_Reservoir": "Stock-tank oil composition/properties only (no volumes)",
        "MFL 20110028_DULANG-B14_Reservoir": "Stock-tank oil composition/properties only (no volumes)",
        "MFL 20110028_DULANG-D26_Reservoir": "Stock-tank oil composition/properties only (no volumes)",
        "MFL 20130041": "Water analysis", "docx": "Same report as the .xlsx version - values cross-checked, counted once under the .xlsx file",
        "PRSS-LS-99-1": "Separator samples QC and compositions only", "PRSS-SR4-95-09": "QC and compositions only",
        "SFL 82033": "Separator gas only", "SFL 82035": "Compositions only", "SFL 82056A": "Composition + CCE identical to SFL 82056 Page 4 (duplicate) - no Bo",
        "SFL 82141": "Gas reservoir fluid (relative volume < 1 above dew point) - no oil Bo", "~$": "Office lock file - not a document"}
    for i, p in enumerate(files, 1):
        recs = [r for r in ROWS if r["fname"] == p.name]
        out = next((v for k, v in why.items() if k in p.name), "") if not recs else "Records extracted"
        fs.append([i, p.name, len(recs), sum(r["flag"] == "DIRECT" for r in recs), sum(r["flag"] != "DIRECT" for r in recs), out])
    for i, w in enumerate([4, 80, 9, 9, 22, 80], 1):
        fs.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    for c in fs[1]:
        c.font = Font(bold=True)

    rd = wb.create_sheet("Read me")
    for line in [
        "Source folder (read only): " + str(ORIG),
        "One row per distinct measurement/test condition. Values are read from the source cells; formatting artefacts such as '1. 242' or '1780*' are normalised and noted.",
        "DIRECT = Bo explicitly reported (Bo, Bofb, oil volume factor, formation volume factor). Every Bo in these reports is a ratio; no absolute oil volumes at reservoir and stock-tank conditions are reported, so the two volume columns are blank for DIRECT rows.",
        "CALCULATED = none: no file reports both absolute volumes for the same sample at documented reservoir and stock-tank conditions.",
        "REQUIRES CALCULATION – INSUFFICIENT DATA = oil volume information exists (stock-tank flash or CCE relative volume) but the matching reservoir/stock-tank volume is missing.",
        "Flopetrol differential-vaporisation 'oil volume factor' is per standard barrel of residual oil (differential basis), not per separator stock-tank barrel; Bo units column states the basis.",
        "Core Lab differential 'Relative Oil Volume (Bod)' values are not Bo and were not listed; their adjusted Bo tables are listed instead.",
        "Duplicates removed: summary values equal to table values, adjusted-table rows at saturation pressure equal to Bofb, and the Word copy of MFL 91184.",
        "Dates: Flopetrol typed sheets store dd/mm dates as mm/dd; covers were checked visually where available. Interpreted dates are flagged.",
        f"Built {dt.datetime.now().isoformat(timespec='seconds')} by scripts/build_bo_extraction.py",
    ]:
        rd.append([line])
    rd.column_dimensions["A"].width = 160
    after = {p.name: hashlib.sha256(p.read_bytes()).hexdigest() for p in files}
    assert before == after, "source changed"
    wb.save(OUT)
    n_direct = sum(r["flag"] == "DIRECT" for r in ROWS)
    print(OUT, "rows", len(ROWS), "direct", n_direct, "insufficient", len(ROWS) - n_direct,
          "review flags", sum(bool(r.get("review")) for r in ROWS), "sources unchanged")


if __name__ == "__main__":
    main()
