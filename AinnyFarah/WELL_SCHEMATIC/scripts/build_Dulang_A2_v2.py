#!/usr/bin/env python3
"""Dulang_A2.json from PDF page 3 (label 'WELL A-2') under the controlled-interpretation rules (page 3 only).

Reading: local Tesseract only — 6 variants of the page's native schematic (two stacked embedded images composited,
output_A/ocr_A2/img_x4_*), the earlier 600 dpi page renders (ocr_passes/p03_*), and the exact PDF text layer (typed boxes,
ocr_A2/text_layer.json). Every string used is asserted to occur verbatim (whitespace collapsed) in >= 2 OCR texts or in the
text layer. Depths: m as printed (equipment m MDTHF; casing / PBTD m MDDF — reference kept, never shifted) -> ft (x 3.28084,
3 dp); metres kept in *M keys. Each source note says DIRECT / CALCULATED / APPROX-TYPE / INTERPRETED / ASSUMED.
"""
import datetime as dt
import glob
import hashlib
import json
import os
import re
import shutil
import subprocess
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "output_A"
RUN = OUT / "run_A2_v2"
TPL = ROOT / "well-schematic-template-json.rtf"
XTPL = ROOT / "well-schematic-sample-template.xlsx"
FT = 3.28084
OCR_FILES = sorted(glob.glob(str(OUT / "ocr_A2" / "img_x4_*.txt"))) + sorted(glob.glob(str(OUT / "ocr_passes" / "p03_*.txt")))
TEXTS = [re.sub(r"[ \t]+", " ", Path(f).read_text(errors="ignore")) for f in OCR_FILES]
LAYER = [l["text"] for l in json.loads((OUT / "ocr_A2" / "text_layer.json").read_text())]
MDTHF = "m MDTHF (column header 'DEPTH m MDTHF')"
MDDF = "m MDDF (as printed in the header; different reference from the m MDTHF equipment depths — not shifted)"


def n_seen(s):
    if any(s in t for t in LAYER):
        return "text layer"
    return sum(1 for t in TEXTS if s in t)


def conf(s):
    k = n_seen(s)
    assert k == "text layer" or k >= 2, f"not confirmed: {s!r} ({k})"
    return k


def ft(m):
    return None if m is None else round(m * FT, 3)


def sha(p):
    fd = os.open(p, os.O_RDONLY); b = []
    while (x := os.read(fd, 1 << 20)):
        b.append(x)
    try:
        os.close(fd)
    except OSError:
        pass
    return hashlib.sha256(b"".join(b)).hexdigest()


S, L, SL = "Short", "Long", "Short & Long"
ID_NOTE = "ID from the 'MIN ID' column"
# (id, Component_Types name, name as written, string, top m, bottom m, ref, OD, ID, extra, approx-type, estimated fields+note, source note)
COMPS = [
    ("casing-conductor", "Casing", '18 5/8" X-56 COND.', None, None, 592.0, MDDF, 18.625, None, {"grade": "X-56", "weight": 87.5},
     None, (["outerDiameter"], "INTERPRETED: fraction of '18 _/_\"' illegible in all readings; 18-5/8 is the standard size for 87.5 ppf X-56 conductor"),
     "Header 'CASING:' line 1 — '18 ?/?\" X-56 COND. 87.5 ppf @ 592 m MDDF'; grade, weight, shoe depth DIRECT (11-14 readings); top not printed"),
    ("casing-surface", "Casing", '13 3/8" K-55 SURF.', None, None, 1172.0, MDDF, 13.375, None, {"grade": "K-55", "weight": 54.5},
     None, (["outerDiameter"], "INTERPRETED: fraction of '13 _/_\"' illegible; 13-3/8 is the standard size for 54.5 ppf K-55"),
     "Header 'CASING:' line 2 — '13 ?/?\" K-55 SURF., 54.5 ppf @ 1172 m MDDF'; grade, weight, shoe depth DIRECT; top not printed"),
    ("casing-production", "Casing", '9 5/8" N-80 PROD.', None, None, 2150.0, MDDF, 9.625, None, {"grade": "N-80", "weight": 40},
     None, (["outerDiameter"], "INTERPRETED: fraction of '9 _/_\"' illegible; 9-5/8 is the standard size for 40 ppf N-80"),
     "Header 'CASING:' line 3 — '9 ?/?\" N-80 PROD., 40 ppf @ 2150 m MDDF'; diagram bottom label '9 ?/?\" CASING SHOE' (no depth on the label); top not printed"),
    ("tubing", "Tubing", '2 7/8" 6.5 PPF N-80 NK3SB', SL, None, None, None, 2.875, None, {"grade": "N-80", "weight": 6.5},
     None, (["outerDiameter"], "INTERPRETED: fraction of '2 _/_\"' illegible; 2-7/8 is the standard size for 6.5 ppf N-80 tubing"),
     "Header 'TUBING :' — '6.5 PPF N-80 NK3SB' DIRECT (12-14 readings; 'N-60' once); one spec printed for the dual completion -> string 'Short & Long' "
     "(ASSUMED applies to both); no tubing depths printed"),
    ("xmas-tree", "Xmas Tree", "INGRAM CACTUS, DUAL, 3M", None, None, None, None, None, None, {}, None, None,
     "Header 'X-MAS TREE:' — DIRECT (12 readings); no depth"),
    # ---- short string (left columns: MIN ID | SHORT STRING | DEPTH m MDTHF)
    ("s-flow-coupling", "Flow Coupling", "FLOW COUPLING", S, 132.9, None, MDTHF, None, 2.441, {}, None, None, "Short string row 1 — DIRECT; " + ID_NOTE),
    ("s-scssv", "SCSSV", "TRDP-4A SCSSV", S, 133.7, None, MDTHF, None, 2.312, {}, None, None,
     "Short string row 2 — depth/ID DIRECT; name INTERPRETED: short-side readings 'TROP-4A' (4) / 'TRDOP-4A', long side 'TRDP-4A' (3) — same model assumed; "
     "typed box at the left of rows 2-3: '2.31” CSA Plug (Feb 2023)' (see s-csa-plug)"),
    ("s-wp1-nipple", "Landing Nipple", "WP-1 NIPPLE", S, 182.7, None, MDTHF, None, 2.312, {}, "WP-1 nipple -> landing-nipple", None,
     "Short string row 3 — DIRECT ('WP-1' 7 / 'VVP-1' 6 readings, VV = W); " + ID_NOTE),
    ("s-spm-1", "SPM (Gas Lift)", "KBMG-M SPM", S, 668.3, None, MDTHF, None, 2.371, {}, None, None,
     "Short string SPM 1 — DIRECT; typed box beside the short-string SPM rows: 'No clear GLV data'"),
    ("s-spm-2", "SPM (Gas Lift)", "KBMG-M SPM", S, 1183.6, None, MDTHF, None, 2.371, {}, None, None,
     "Short string SPM 2 — depth INTERPRETED: 1183.6 (6 readings) vs 1193.6 (2); 'No clear GLV data'"),
    ("s-spm-3", "SPM (Gas Lift)", "KBMG-M SPM", S, 1470.0, None, MDTHF, None, 2.371, {}, None, None, "Short string SPM 3 — DIRECT (8 readings); 'No clear GLV data'"),
    ("s-spm-4", "SPM (Gas Lift)", "KBMG-M SPM", S, 1623.6, None, MDTHF, None, 2.371, {}, None, (["innerDiameter"], "INTERPRETED: ID read '5.371'/'9 371' on this row; 2.371 as the other KBMG-M rows"),
     "Short string SPM 4 — depth 1623.6 read 2x ('16036' once, out-of-sequence alternatives none); 'No clear GLV data'"),
    ("s-spm-5", "SPM (Gas Lift)", "KBMG-M SPM", S, 1709.4, None, MDTHF, None, 2.371, {}, None, None, "Short string SPM 5 — DIRECT (6 readings); 'No clear GLV data'"),
    ("s-ssd", "Sliding Sleeve", "WPB-2DE SSD", S, 1882.1, None, MDTHF, None, 2.312, {}, None, None,
     "Short string — DIRECT; status column on image 'CLOSED'; typed box 'CLOSED' / '(March 2015)'"),
    ("dual-packer", "Packer", "DUAL HYD. PKR", SL, 1887.7, None, MDTHF, None, None, {}, None, None,
     "Shared dual packer (one component) — name/depth DIRECT (1887.7 4x vs 1987.7 2x, the latter out of sequence); model line above it read "
     "'GUIB UNI-XXVII' (2) / 'GUIB UNI-XXVI' (2) — unresolved, not added to the name; no ID read"),
    ("s-nogo-nipple", "No-Go Nipple", "'D' NOGO NIPPLE", S, None, None, MDTHF, None, 2.250, {}, None, None,
     "Short string below the dual packer — name INTERPRETED from ''D'NOGO NIPPLE' / ''D' MOGO MIPPLE' / ''D' NMO-GO'; ID 2.250 DIRECT; depth UNRESOLVED: "
     "read '1090.5' in 6 sources, which is out of sequence (between 1887.7 and 1890.9) — left empty; status on image 'Open'; typed box 'OPEN' / 'June 2012'"),
    ("s-wl-entry-guide", "Mule Shoe", "WL ENTRY GUIDE", S, 1890.9, None, MDTHF, None, 2.441, {}, "wireline entry guide (end of short string) -> mule-shoe", None,
     "Short string last item — DIRECT ('WL' 2 / 'VL' 2); " + ID_NOTE),
    # ---- long string (right columns: DEPTH m MDTHF | LONG STRING | MIN ID | STATUS)
    ("l-flow-coupling", "Flow Coupling", "FLOW COUPLING", L, 143.0, None, MDTHF, None, 2.441, {}, None, None, "Long string row 1 — DIRECT; " + ID_NOTE),
    ("l-scssv", "SCSSV", "TRDP-4A SCSSV", L, 143.7, None, MDTHF, None, 2.312, {}, None, None, "Long string row 2 — DIRECT"),
    ("l-wp1-nipple", "Landing Nipple", "WP-1 NIPPLE", L, 297.6, None, MDTHF, None, 2.312, {}, "WP-1 nipple -> landing-nipple", None, "Long string row 3 — DIRECT"),
    ("l-spm-1", "SPM (Gas Lift)", "KBMG-M SPM", L, 677.4, None, MDTHF, None, 2.371, {}, None, None,
     "Long string SPM 1 — DIRECT; typed note in the long-string status column, 1st of 4 (association by order/position, flagged): 'GLV 12/64” (Apr/08)'"),
    ("l-spm-2", "SPM (Gas Lift)", "KBMG-M SPM", L, 1173.8, None, MDTHF, None, 2.371, {}, None, (["innerDiameter"], "INTERPRETED: ID read 2.371 once ('2.971', '2.374' elsewhere)"),
     "Long string SPM 2 — depth INTERPRETED: 1173.8 (6) vs 1173.6 (2); typed note 2nd of 4: 'DV (Jun/06)'"),
    ("l-spm-3", "SPM (Gas Lift)", "KBMG-M SPM", L, 1470.4, None, MDTHF, None, 2.371, {}, None, (["innerDiameter"], "INTERPRETED: ID read 2.371 once"),
     "Long string SPM 3 — depth 1470.4 (3; '14704' 4 more, decimal point lost); typed note 3rd of 4: 'DV (Jun/06)'"),
    ("l-spm-4", "SPM (Gas Lift)", "KBMG-M SPM", L, 1613.8, None, MDTHF, None, 2.371, {}, None, (["innerDiameter"], "INTERPRETED: ID read '237' / '2.301' / '2a'; 2.371 as the other KBMG-M rows"),
     "Long string SPM 4 — depth DIRECT (4); typed note 4th of 4: 'GLOV 12/64” (Apr/08)'"),
    ("l-ssd-1", "Sliding Sleeve", "WPB-2DE SSD", L, 1881.6, None, MDTHF, None, 2.312, {}, None, (["innerDiameter"], "INTERPRETED: ID read '2312' / '9.312'"),
     "Long string above the dual packer — depth 1881.6 (4; '1981.8' once, out of sequence); status on image 'CLOSED'; typed box 'CLOSED'"),
    ("l-ssd-2", "Sliding Sleeve", "'WPB-2DE' SSD", L, 1899.4, None, MDTHF, None, 2.312, {}, None, None, "Long string below the dual packer — DIRECT (4)"),
    ("l-blast-joint-1", "Blast Joint", "BLAST JOINT", L, 1902.6, 1914.8, MDTHF, None, None, {}, None, None,
     "Long string across SAND: E-2 & E-7 — label DIRECT; interval INTERPRETED from the depth-column sequence 1902.6 (6) / BLAST JOINT / 1914.8 (4); no ID read; "
     "1914.8 - 1925.4 has no label (not added)"),
    ("l-blast-joint-2", "Blast Joint", "BLAST JOINT", L, 1925.4, 1947.7, MDTHF, None, 2.441, {}, None, None,
     "Long string across SAND: E-2 & E-7 — label DIRECT; interval INTERPRETED from 1925.4 (3) / 1947.7 (2 + '19477', 'joa7 7') on the 'BLAST JOINT 2.441' row"),
    ("l-wp1-nipple-2", "Landing Nipple", "WP1 NIPPLE", L, 1970.0, None, MDTHF, None, 2.312, {}, "WP1 nipple -> landing-nipple", None, "Long string — DIRECT (8)"),
    ("l-single-packer", "Packer", "GUIB UNI-VII SINGLE. HYD. PKR", L, 1976.4, None, MDTHF, None, 2.441, {}, None, None,
     "Label on the left-hand column ('GUIB UNI-VII' line above 'SINGLE. HYD. PKR', joined) — DIRECT; string 'Long' INTERPRETED: depth is below the short-string "
     "end (1890.9)"),
    ("l-ssd-3", "Sliding Sleeve", "'WPB-2DE' SSD", L, 1982.6, None, MDTHF, None, 2.312, {}, None, None,
     "Long string — DIRECT (5); typed box in the status column on this row: 'OPEN' / 'Oct 2016'"),
    ("l-blast-joint-3", "Blast Joint", "BLAST JOINT", L, 1987.9, 2002.2, MDTHF, None, 2.441, {}, None, (["innerDiameter"], "INTERPRETED: ID read '2444' / '2441' / '5444'"),
     "Long string across SAND: E-12/13 — label and interval DIRECT (1987.9 4, 2002.2 7)"),
    ("l-wp1-nipple-3", "Landing Nipple", "WP1 NIPPLE", L, 2003.2, None, MDTHF, None, 2.312, {}, "WP1 nipple -> landing-nipple", None, "Long string — DIRECT (6; '9003.2' once)"),
    ("l-centralizer", "Pup Joint", "CENTRALIZER", L, 2006.8, None, MDTHF, None, 2.441, {}, "tubing centralizer -> pup-joint (no centralizer Type_Key; template 'centralizers' section has no record structure)", None,
     "Long string — label DIRECT; depth INTERPRETED: 2006.8 (4) vs 2008.8 (3)"),
    ("l-seal-assy", "Seal Assembly", "LOC. SEAL ASSY.", L, 2008.3, None, MDTHF, None, 2.441, {}, None, None, "Long string — DIRECT (7-8)"),
    ("l-perm-packer", "Packer", "GUIB MAGNUM HT PERM. PKR", L, 2008.9, None, MDTHF, None, 4.000, {}, None, None,
     "Label on the left-hand column ('GUIB MAGNUM' line above \"'HT PERM. PKR\", joined; leading quote dropped) — DIRECT; ID 4.000 from the left MIN ID "
     "column; string 'Long' INTERPRETED (below the short-string end)"),
    ("l-seal-bore-ext", "PBR", "SEAL BORE EXT.", L, None, None, MDTHF, None, 4.000, {}, "seal bore extension -> pbr", None,
     "Long string below the permanent packer — DIRECT; no depth printed"),
    ("l-nogo-nipple", "No-Go Nipple", "'D' NO-GO NIPPLE", L, 2012.6, None, MDTHF, None, 2.250, {}, None, (["innerDiameter"], "INTERPRETED: ID read '2250' / '9950'"),
     "Long string — depth 2012.6 (5) vs 2012.8 (1); typed box in the status column on this row: 'CLOSE' / 'Oct 2016'; see l-ca-plug"),
    ("l-weg", "Mule Shoe", "WEG", L, 2013.2, None, MDTHF, None, 2.441, {}, "WEG (wireline entry guide) -> mule-shoe", None, "Long string last item — DIRECT (7)"),
    ("pbtd", "PBTD", "PBTD: 2124 m MDDF", None, 2124.0, None, MDDF, None, None, {}, None, None, "Header — DIRECT (8)"),
    # ---- perforations (typed text layer, left of the diagram; 'mdTHF' read as m MDTHF)
    ("perf-e2", "Perforation", "E2 - 1906.4 – 1910.4 mdTHF", L, 1906.4, 1910.4, MDTHF, None, None, {"zone": "E2", "interval": "1906.4 – 1910.4"}, None, None,
     "Typed text layer beside 'SAND: E-2 & E-7' — DIRECT; string 'Long' INTERPRETED (below the short-string end)"),
    ("perf-e7", "Perforation", "E7 - 1931 - 1944", L, 1931.0, 1944.0, MDTHF, None, None, {"zone": "E7", "interval": "1931 - 1944"}, None, None,
     "Typed text layer — DIRECT; unit not printed on this line, m MDTHF ASSUMED as the E2 line"),
    ("perf-e12-13-a", "Perforation", "E12/13 – 1990 – 1995", L, 1990.0, 1995.0, MDTHF, None, None, {"zone": "E12/13", "interval": "1990 – 1995"}, None, None,
     "Typed text layer beside 'SAND: E-12/13' — DIRECT; unit ASSUMED m MDTHF; kept separate from the 1998 - 2001 interval"),
    ("perf-e12-13-b", "Perforation", "1998 - 2001", L, 1998.0, 2001.0, MDTHF, None, None, {"zone": "E12/13", "interval": "1998 - 2001"}, None, None,
     "Typed text layer directly below the E12/13 line — DIRECT; zone E12/13 INTERPRETED by position; unit ASSUMED m MDTHF"),
    ("perf-e14", "Perforation", "E14 – 2015 – 2025", L, 2015.0, 2025.0, MDTHF, None, None, {"zone": "E14", "interval": "2015 – 2025"}, None, None,
     "Typed text layer beside 'SAND: E-14' — DIRECT; unit ASSUMED m MDTHF; below the plugged long-string No-Go (see l-ca-plug)"),
    # ---- plugs (typed boxes)
    ("s-csa-plug", "Mechanical Plug", "2.31” CSA Plug (Feb 2023)", S, None, None, None, None, None, {}, "CSA plug -> mechanical-plug", None,
     "Typed box in the short-string status column between rows 2-3 (SCSSV 133.7 / WP-1 nipple 182.7) — DIRECT name and partial date; string 'Short' "
     "INTERPRETED by position; depth not printed and the row is ambiguous — left empty"),
    ("l-ca-plug", "Mechanical Plug", "2.25” CA Plug @ No-Go", L, 2012.6, None, MDTHF, None, None, {}, "CA plug -> mechanical-plug", None,
     "Typed box under the long-string depth column — DIRECT name; depth ASSUMED = long-string 'D' NO-GO NIPPLE 2012.6 (box says '@ No-Go', sits beside that row; "
     "the short string also has a 2.250 No-Go, position decides)"),
]
CHECK = ["X-56 COND.", "87.5 ppf", "592 m MDDF", "K-55 SURF.,", "54.5 ppf", "1172 m MDDF", "N-80 PROD.,", "40 ppf", "2150 m MDDF", "CASING SHOE",
         "6.5 PPF", "N-80 NK3SB", "INGRAM CACTUS, DUAL, 3M", "FLOW COUPLING", "132.9", "TRDP-4A SCSSV", "TROP-4A SCSSV", "133.7", "WP-1 NIPPLE", "182.7",
         "KBMG-M SPM", "668.3", "1183.6", "1470.0", "1623.6", "1709.4", "WPB-2DE SSD", "1882.1", "DUAL HYD. PKR", "1887.7", "GUIB UNI-XXVII", "1090.5",
         "2.250", "WL ENTRY GUIDE", "1890.9", "143.0", "143.7", "297.6", "677.4", "1173.8", "1470.4", "1613.8", "1881.6", "1899.4", "1902.6", "BLAST JOINT",
         "1914.8", "1925.4", "1947.7", "WP1 NIPPLE", "1970.0", "GUIB UNI-VII", "SINGLE. HYD. PKR", "1976.4", "1982.6", "1987.9", "2002.2", "2003.2",
         "CENTRALIZER", "2006.8", "LOC. SEAL ASSY.", "2008.3", "GUIB MAGNUM", "HT PERM. PKR", "2008.9", "4.000", "SEAL BORE EXT.", "'D' NO-GO NIPPLE",
         "2012.6", "WEG", "2013.2", "2.441", "2.312", "2.371", "2124 m MDDF", "SAND: E-2 & E-7", "SAND: E-12/13", "m MDTHF",
         "1906.4", "1910.4", "mdTHF", "E7", "1931", "1944", "E12/13", "1990", "1995", "1998", "2001", "E14", "2015", "2025",
         "2.31” CSA Plug", "(Feb 2023)", "2.25” CA Plug", "@ No-Go", "No clear", "GLV data", "CLOSED", "(March 2015)", "OPEN", "June 2012", "Oct 2016", "CLOSE",
         "GLV 12/64” (Apr/08)", "DV (Jun/06)", "GLOV 12/64” (Apr/08)", "WELL A-2", "DUAL OIL PRODUCER WITH LONG STRING SELECTIVE (TYPE 6)",
         "6 MARCH 1991", "TIOMAN 1", "9.0 PPG3% KCI + NaCl + 0.05% NALCO", "3900 BIOCIDE & CORR. INHIBITOR", "62.3 DEG", "14.13 m", "DOGLEG (DEPTH)"]


def main():
    hb = [l.split("  ", 1) for l in (RUN / "sha256_before.txt").read_text().splitlines()]
    confirm = {s: conf(s) for s in CHECK}
    tpl = json.loads(subprocess.run(["textutil", "-convert", "txt", "-stdout", str(TPL)], capture_output=True, text=True, check=True).stdout)
    keymap = {r[0]: r[2] for r in load_workbook(XTPL)["Component_Types"].iter_rows(min_row=2, values_only=True)}
    wi = {k: v if v in ("-", 0) else (None if isinstance(v, (int, float)) and not isinstance(v, bool) else "") for k, v in tpl["wellInfo"].items()}
    wi.update(name="Dulang_A2", completionDate="6 MARCH 1991", completionType="DUAL OIL PRODUCER WITH LONG STRING SELECTIVE (TYPE 6)",
              maxDeviation="62.3 DEG", packerFluid="9.0 PPG3% KCI + NaCl + 0.05% NALCO 3900 BIOCIDE & CORR. INHIBITOR", rig="TIOMAN 1",
              surfaceCasing='13 3/8" K-55 SURF., 54.5 ppf @ 1172 m MDDF', tubing='2 7/8" 6.5 PPF N-80 NK3SB', wellhead="INGRAM CACTUS, DUAL, 3M")
    comps = []
    for cid, ctype, name, string, top, bot, ref, od, idd, extra, approx, est, note in COMPS:
        d = {"id": cid, "type": keymap[ctype], "name": name, "topMd": ft(top), "bottomMd": ft(bot), "outerDiameter": od, "innerDiameter": idd}
        d.update(extra)
        fields, notes = ([], [])
        if approx:
            fields.append("type"); notes.append(f"APPROX-TYPE: no exact Type_Key; {approx}")
        if est:
            fields += est[0]; notes.append(est[1])
        if fields:
            d.update(estimated=True, estimatedFields=fields, estimationNote="; ".join(notes))
        if string:
            d["string"] = string
        conv = f"; depth reference {ref}; CALCULATED: ft = m x 3.28084 (3 dp)" if top is not None or bot is not None else ""
        d["source"] = f"PDF page 3 (label 'WELL A-2'): {note}{conv}"
        d.update(actualType=name, topMdM=top, bottomMdM=bot)
        comps.append(d)
    doc = {k: {"wellInfo": wi, "components": comps}.get(k, []) for k in tpl}
    out = OUT / "Dulang_A2.json"
    if out.exists():                       # naming collision: keep the earlier file, never overwrite silently
        bak = OUT / "_previous_runs" / f"Dulang_A2_before_v2_{dt.datetime.now():%Y%m%d_%H%M%S}.json"
        shutil.copy2(out, bak)
        print("existing Dulang_A2.json backed up ->", bak.name)
    out.write_text(json.dumps(doc, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")

    back = json.loads(out.read_text())
    kind = lambda v: "number" if isinstance(v, (int, float)) and not isinstance(v, bool) else type(v).__name__
    cs = back["components"]
    approx_ids = {c[0] for c in COMPS if c[10]}
    v = [("valid JSON", True),
         ("top-level sections and order = template", list(back) == list(tpl)),
         ("wellInfo keys = template", list(back["wellInfo"]) == list(tpl["wellInfo"])),
         ("wellInfo types = template (where populated)", all(x in ("", None) or kind(x) == kind(tpl["wellInfo"][k]) for k, x in back["wellInfo"].items())),
         ("no template sample records", back["surveys"] == [] and back["formationTops"] == [] and not {c["id"] for c in cs} & {t["id"] for t in tpl["components"]}),
         ("all types are Component_Types keys", all(c["type"] in set(keymap.values()) for c in cs)),
         ("component ids unique", len({c["id"] for c in cs}) == len(cs)),
         ("ft = m x 3.28084 (3 dp)", all((c[a] is None and c[b] is None) or c[a] == round(c[b] * FT, 3) for c in cs for a, b in (("topMd", "topMdM"), ("bottomMd", "bottomMdM")))),
         ("every component has source traceability", all(c.get("source", "").startswith("PDF page 3") for c in cs)),
         ("approximate types flagged (estimated + 'type' in estimatedFields)", all(c.get("estimated") and "type" in c["estimatedFields"] for c in cs if c["id"] in approx_ids)),
         ("non-approximate types not flagged as type estimates", all("type" not in c.get("estimatedFields", []) for c in cs if c["id"] not in approx_ids)),
         ("string values in {Short, Long, Short & Long}", all(c.get("string", S) in (S, L, SL) for c in cs)),
         ("exactly one shared dual packer", sum(1 for c in cs if c["type"] == "packer" and c.get("string") == SL) == 1),
         ("standardized name 'Dulang_A2' = filename", back["wellInfo"]["name"] == out.stem),
         ("source PDF unchanged", sha(hb[0][1]) == hb[0][0]),
         ("template unchanged", sha(str(ROOT / hb[1][1])) == hb[1][0]),
         ("Component_Types reference unchanged", sha(str(ROOT / hb[2][1])) == hb[2][0])]
    for a, b in v:
        print("PASS" if b else "FAIL", "-", a)
    print("components:", len(cs), "| strings confirmed:", len(confirm), "(", sum(1 for x in confirm.values() if x == "text layer"), "from text layer )")
    (RUN / "confirmations.json").write_text(json.dumps(confirm, indent=1, ensure_ascii=False))


if __name__ == "__main__":
    main()
