#!/usr/bin/env python3
"""Write the findings index for the PVT folder scan: where Bo, oil volume at reservoir conditions and oil volume at
stock-tank conditions appear in each file. Locations and labels only - no data values are copied.
Classification below was reviewed by hand from _work_pvt_scan/hits_review.txt (keyword hits on cell text, text boxes,
document text and OCR of pictures)."""
import json
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

BASE = Path(__file__).resolve().parent.parent
RES = json.loads((BASE / "_work_pvt_scan" / "scan_results.json").read_text())
OUT = BASE / "PVT_folder_Bo_oil_volume_findings.xlsx"

# file prefix -> (Bo, reservoir-condition oil volume, stock-tank oil volume, fluid/report note)
# each item: (status, where)  status: YES = stated directly; RELATED = only a related quantity; NO
F = {
    "82_L_004": (("YES", "SUMMARY RESULTS B19 'oil volume factor (bbl/Std bbl)'; TABLE 41 G6 'OIL VOLUME FACTOR (2)'"),
                 ("YES", "TABLE 41 C16 footnote: volume of reservoir fluid at saturation pressure (Bo basis); TABLE 37 relative volume (CCE)"),
                 ("YES", "TABLE 41 C15-C17 footnotes (per volume of stock tank oil at 60°F); flash to stock-tank conditions: TABLE 3, 13, 26, 29"),
                 "Flopetrol PVT study, DST 1"),
    "82_L_033": (("NO", ""), ("NO", ""), ("NO", ""), "Cover page only (one picture, OCR'd) - no tables in this file"),
    "82_L_034": (("YES", "SUMMARY RESULTS B19 'oil volume factor (bbl/Std bbl)'; TABLE 13 C6 'Oil volume factor Bo (bbl/Std bbl)'; TABLE 16 G6 'Formation Vol Factor'"),
                 ("YES", "TABLE 16 B15 footnote (volume of reservoir fluid at saturation pressure); TABLE 12 relative volume (CCE)"),
                 ("YES", "TABLE 16 B14-B16 footnotes (per volume of stock tank oil at 60°F); TABLE 6 & 7 flash to stock-tank conditions"),
                 "Flopetrol PVT study, DST 4"),
    "82_L_035": (("RELATED", "TABLE 8 B9 shrinkage factor only (separator liquid to stock tank) - no Bo"),
                 ("RELATED", "TABLE 7 relative volume (CCE) only"),
                 ("YES", "TABLE 8 / TABLE 9: flash of separator liquid to stock-tank conditions"),
                 "Flopetrol PVT study, DST 5"),
    "83_L_005": (("YES", "SAMPLE SUMMARY B22 'oil volume factor (bbl/std bbl)'"),
                 ("RELATED", "SAMPLE SUMMARY B13 'initial reservoir conditions' heading; bubble-point plots (pictures)"),
                 ("YES", "BUBBLE POINT PRESSURE 2387-8 C28-C32 and 20524-189 C30 (flash to stock-tank conditions); STO A12137 9214-124"),
                 "Flopetrol PVT study, DST 2"),
    "83_L_006": (("RELATED", "TABLE 2 & 3 B41 shrinkage factor only - no Bo found in cells or the 11 OCR'd pictures"),
                 ("NO", ""),
                 ("YES", "TABLE 2 & 3 B39-B43, TABLE 13 & 14, TABLE 26, TABLE 29 (flash to stock-tank conditions)"),
                 "Flopetrol PVT study, DST 3"),
    "CHEM 86021": (("NO", ""), ("NO", ""), ("NO", ""), "Water analysis sheet + cover picture only"),
    "MFL 20100151_DULANG-A14_Reservoir": (("RELATED", "D.1 F24 shrinkage factor (separator condensate); A.1 text box 'Shrinkage Measurement'"),
                 ("RELATED", "F.1 B62 relative volume to dew-point pressure (gas condensate, not oil)"),
                 ("RELATED", "D.1 F26 stocktank oil flow rate (well test data)"),
                 "Gas-condensate study (dew point) - no oil Bo"),
    "MFL 20100151_DULANG-A14_SAMPLE": (("NO", ""), ("NO", ""), ("NO", ""), "Sample information form"),
    "MFL 20110028_DULANG-A25_Reservoir": (("NO", ""), ("NO", ""),
                 ("RELATED", "Stock-tank oil samples: C.1 / C.3 compositions, D.1 basic properties (no volumes)"),
                 "Stock-tank oil composition/properties of B14 & D26S; same report content as the B14 and D26 files"),
    "MFL 20110028_DULANG-B14_Reservoir": (("NO", ""), ("NO", ""),
                 ("RELATED", "Stock-tank oil samples: C.1 / C.3 compositions, D.1 basic properties (no volumes)"),
                 "Same report content as the A25 and D26 files"),
    "MFL 20110028_DULANG-D26_Reservoir": (("NO", ""), ("NO", ""),
                 ("RELATED", "Stock-tank oil samples: C.1 / C.3 compositions, D.1 basic properties (no volumes)"),
                 "Same report content as the A25 and B14 files"),
    "MFL 20110028_DULANG-A25_SAMPLE": (("NO", ""), ("NO", ""), ("NO", "Sample Info C34 'STOCK TANK' = sample type only"), "Sample information form"),
    "MFL 20110028_DULANG-B14_SAMPLE": (("NO", ""), ("NO", ""), ("NO", "Sample Info C34 'STOCK TANK' = sample type only"), "Sample information form"),
    "MFL 20110028_DULANG-D26_SAMPLE": (("NO", ""), ("NO", ""), ("NO", "Sample Info C34 'STOCK TANK' = sample type only"), "Sample information form"),
    "MFL 20120023": (("YES", "A.1 AR35-AR36 'Volume Factor (bofb)'; E.1 AU14 'Bo(1)', E.1 AQ52 and E.1 text box (Bo definitions); G.1 H14 'Bofb (3)'"),
                 ("YES", "E.1 E14 'Bod(2)' + B63 footnote; G.1 B64-B65 footnotes (saturated oil at reservoir P,T); A.1 AN24, AN59; E.1 chart 'Relative Oil Volume' (picture 7)"),
                 ("YES", "G.1 B63-B65 and A.1 AN58-AN59 footnotes (per volume of stocktank oil at 60°F); G.1 U7 stocktank oil composition"),
                 "Core Lab RFS, B21, Oct 2012"),
    "MFL 20130041": (("NO", ""), ("NO", "B.1 E10 'Reservoir Conditions' = sampling P,T of water samples"), ("NO", ""),
                 "Water analysis; Bo text inside the file is only cached data of external links to other companies' jobs (not read)"),
    "MFL 91184_DULANG-B12_Reservoir Fluid Study_Apr 1992.docx": (
                 ("YES", "Summary section ('Bodb' and Formation Vol. Factor definition); single-stage separator test table 'Bofb'; nomenclature & equations page (Bo, Bob, Bofb, Bod, Bodb, BT); adjusted data pages 27-30 'Bo'; equation pictures image15/image18 (OCR)"),
                 ("YES", "Relative Volume at reservoir pressure; Relative Oil Volume (differential vaporisation) tables and charts; 'barrels of saturated oil at 1668 psig and 200°F' definition"),
                 ("YES", "Separator test footnotes 'per barrel of stock tank oil at 60°F'; hydrocarbon analysis of stock tank oil; image21 'STOCK TANK' (OCR)"),
                 "Core Lab RFS, B12, Apr 1992 (Word version)"),
    "MFL 91184_DULANG-B12_Reservoir Fluid Study_Apr 1992.xlsx": (
                 ("YES", "'Single Stage Separator Test' G12 'Formation Vol Factor (Bofb)' + B32; 'Adjustment Oil' C14 'Bo'; 'Adjustment Gas' B13-B14 (Bofb, Bodb)"),
                 ("YES", "'Single Stage Separator Test' B32 (saturated oil at 1668 psig); 'DVT at 200°F' D10 Relative Oil Volume + B25; 'Adjustment Oil' B8"),
                 ("YES", "'Single Stage Separator Test' B31-B33 footnotes; 'Stock Tank Oil CA' B7"),
                 "Core Lab RFS, B12, Apr 1992 (Excel version of the same report)"),
    "PRSS-LS-99-1": (("NO", ""), ("NO", ""), ("NO", "separator oil only (not stock tank)"), "Separator samples QC + compositions"),
    "PRSS-SR4-95-09": (("NO", ""), ("NO", ""), ("NO", ""), "QC + compositional tables only"),
    "SFL 82033": (("NO", ""), ("NO", "relative volume is for separator gas (PVR)"), ("NO", ""), "Partial study"),
    "SFL 82034": (("YES", "'SINGLE SEPARATOR TEST' G5 'Formation Vol Factor (3)'; 'MULTI SEPARATOR TEST' B17 definition"),
                 ("YES", "'MULTI SEPARATOR TEST' B17 (saturated oil at 1825 psig, 218°F); 'DVT at 218 °F' D5 Relative Oil Volume + B20; 'VOL. METRIC DATA OF RESV.' C18 relative volume"),
                 ("YES", "'MULTI SEPARATOR TEST' B16-B18 footnotes (per barrel of stock tank oil at 60°F)"),
                 "Dulang-1 Test 2"),
    "SFL 82035": (("NO", ""), ("NO", ""), ("NO", ""), "Compositional analysis only"),
    "SFL 82055": (("NO", ""), ("RELATED", "'Page 4' E6 relative volume (CCE) only"), ("NO", ""), "Composition + bubble point"),
    "SFL 82056A": (("NO", ""), ("RELATED", "'Page 6' F6 relative volume (CCE) only"),
                 ("RELATED", "'Page 1' B44 field stock tank liquid gravity (gravity, not volume)"), "Composition to C11+"),
    "SFL 82056_": (("YES", "'Page 7' G4 'Formation Volume Factor (3)' + B17 definition"),
                 ("YES", "'Page 7' B17 (saturated oil at 1982 psig, 224°F); 'Page 8' B24 relative oil volume (DV); 'Page 4' F6 relative volume (CCE)"),
                 ("YES", "'Page 7' B16-B18 footnotes (per barrel of stock tank oil at 60°F); 'Page 11' F6"),
                 "Dulang-3 Test 1"),
    "SFL 82141": (("NO", ""), ("RELATED", "'PAGE 4 of 4' E6 relative volume (CCE) only"), ("NO", ""), "Partial study DST 5"),
    "SFL 82142": (("YES", "'SEPARATOR TEST' C32 Formation Volume Factor definition"),
                 ("YES", "'SEPARATOR TEST' C32 (saturated oil at 1562 psig, 181°F); 'DVT 196°F' E10-F10 Relative Oil Volume + D27"),
                 ("YES", "'SEPARATOR TEST' C31-C33 footnotes (per barrel of stock tank oil at 60°F)"),
                 "Partial study DST 4 (22 Dec 1982)"),
    "SFL 82143": (("YES", "'PAGE 9 of 13' G5 'Formation Volume Factor (3)' + B25 definition"),
                 ("YES", "'PAGE 9 of 13' B25 (saturated oil at 1593 psia, 203°F); 'PAGE 4/6 of 13' relative volume (CCE)"),
                 ("YES", "'PAGE 9 of 13' B24-B26 footnotes (per barrel of stock tank oil)"),
                 "Partial study DST 3; footnote prints 'stock tank at 69°F' (likely typo for 60°F)"),
    "SFL 82144": (("NO", ""), ("RELATED", "'PAGE 4 of 4' E5 relative volume (CCE) only"), ("NO", ""), "Partial study DST 2"),
    "SFL 82145": (("NO", ""), ("RELATED", "'PAGE 5 of 5' E5 relative volume (CCE) only"), ("NO", ""), "Partial study DST 4 (3 Dec 1982)"),
    "SFL 82146": (("NO", ""), ("RELATED", "'PAGE 4 of 4' E5 relative volume (CCE) only"), ("NO", ""), "Partial study DST 3"),
    "SFL 86024": (("YES", "'ATMOSPHERIC RFS' F5 'FORMATION VOLUME FACTOR' + B16; 'ADJ.DATA (PAGE 27)' D4 'Formation Vol Factor (Bo)', E4 'Bt'"),
                 ("YES", "'ATMOSPHERIC RFS' B16 (saturated oil at 1225 psig, 211°F); 'ADJ.DATA (PAGE 27)' B22 'barrels at reservoir conditions'; 'DVT at 211 °F' D5 Relative Oil Volume"),
                 ("YES", "'ATMOSPHERIC RFS' B16; 'ADJ.DATA (PAGE 27)' B16, B21-B22"),
                 "Dulang-5 Test 1, E-14"),
    "SFL 86026": (("YES", "'MULTI SEPARATOR TEST' G5 'Formation Vol Factor (3)' + B21; 'ATMOSPHERIC RFS' F5 + B16; 'ADJ.DATA (PAGE 27)' D4 'Formation Vol Factor (Bo)', E4 'Bt'"),
                 ("YES", "'MULTI SEPARATOR TEST' B21 / 'ATMOSPHERIC RFS' B16 (saturated oil at 1370 psig, 205°F); 'ADJ.DATA (PAGE 27)' B23 'barrels at reservoir conditions'; 'DVT at 205 °F' D5"),
                 ("YES", "'MULTI SEPARATOR TEST' B20-B22; 'ATMOSPHERIC RFS' B16; 'ADJ.DATA (PAGE 27)' B17, B22-B23"),
                 "Dulang-5 Test 2, E8 & E13"),
}


def lookup(name):
    keys = [k for k in F if name.startswith(k) or k in name]
    assert len(keys) <= 1, (name, keys)
    return F[keys[0]] if keys else None


def main():
    if OUT.exists():
        raise SystemExit(f"{OUT} exists - not overwritten")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Findings"
    hdr = ["#", "File name", "Extension / actual format", "Scanned (cells, text boxes, OCR of pictures)", "Pictures OCR'd",
           "Bo (oil formation volume factor)", "Bo - where", "Oil volume at reservoir conditions", "Reservoir - where",
           "Oil volume at stock-tank conditions", "Stock-tank - where", "Note"]
    ws.append(hdr)
    fills = {"YES": "C6EFCE", "RELATED": "FFEB9C", "NO": "F2F2F2", "-": "F2F2F2"}
    for i, r in enumerate(RES, 1):
        f = lookup(r["file"])
        fmt = f"{r['type']} / {r.get('container') or '-'}"
        if r["file"].startswith("~$"):
            row = [i, r["file"], fmt, "no - Office lock file", "", "-", "", "-", "", "-", "", "Temporary owner file, not a document"]
        else:
            assert f, r["file"]
            (b, bw), (rv, rw), (s, sw), note = f
            row = [i, r["file"], fmt, "yes" if r["status"] == "scanned" else r["status"], r.get("n_pictures", 0),
                   b, bw, rv, rw, s, sw, note]
        ws.append(row)
        for col in (6, 8, 10):
            c = ws.cell(ws.max_row, col)
            c.fill = PatternFill("solid", fgColor=fills.get(c.value, "FFFFFF"))
    for c in ws[1]:
        c.font = Font(bold=True)
        c.alignment = Alignment(wrap_text=True, vertical="top")
    widths = [4, 60, 16, 16, 9, 12, 60, 12, 60, 12, 60, 40]
    for k, w in enumerate(widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(k)].width = w
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = ws.dimensions
    k = wb.create_sheet("Legend")
    for line in [
        ("YES", "The item is stated in the file (column heading, definition or footnote). Location given; values were not copied."),
        ("RELATED", "Only a related quantity is present (e.g. CCE relative volume V/Vsat, shrinkage factor, stock-tank sample gravity/composition) - not Bo or an oil volume itself."),
        ("NO", "Not found in cell text, text boxes, document text, or OCR of embedded pictures."),
        ("Scope", "Source folder: /Volumes/homes/public/DulangOSS/20260925/DOCUMENTS/PVT Fluid Analysis Document (36 files; read only, SHA-256 unchanged after the scan)."),
        ("Not read", "Cached values of external-workbook links (they belong to other workbooks/companies)."),
        ("Detail", "All keyword hits per location: _work_pvt_scan/hits_review.txt; OCR/text per file: _work_pvt_scan/text/*.json."),
    ]:
        k.append(line)
    k.column_dimensions["A"].width = 12
    k.column_dimensions["B"].width = 140
    wb.save(OUT)
    print(OUT)


if __name__ == "__main__":
    main()
