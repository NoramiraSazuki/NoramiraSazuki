#!/usr/bin/env python3
"""Re-OCR build of Dulang_A1.json (PDF page 1, label 'DLA-01') — exact-match rules.

Source text: six local OCR variants of the page's native schematic image (output_A/ocr_A1/img_x4_*.txt) plus the exact PDF
text layer (typed update boxes). Every label / number below is written only if it occurs verbatim (whitespace collapsed)
in >= 2 OCR variants (or exactly in the text layer); the script asserts this before writing. Depths are as printed
(m MDRKB on this page; WLTHF where the text says so) — no conversion. Values whose readings conflict are left empty.
"""
import glob
import json
import re
import subprocess
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "output_A"
OCR = sorted(glob.glob(str(OUT / "ocr_A1" / "img_x4_*.txt")))
LAYER = json.loads((OUT / "ocr_passes" / "text_layer.json").read_text())["p01"]
TPL = ROOT / "well-schematic-template-json.rtf"
XTPL = ROOT / "well-schematic-sample-template.xlsx"
TEXTS = [re.sub(r"[ \t]+", " ", Path(f).read_text(errors="ignore")) for f in OCR]
LAYER_TXT = [l["text"] for l in LAYER]


def seen(s):
    """number of OCR variants containing s verbatim; text-layer strings count as exact."""
    if any(s in t for t in LAYER_TXT):
        return 99
    return sum(1 for t in TEXTS if s in t)


def ok(s, n=2):
    assert seen(s) >= n, f"not confirmed in >= {n} sources: {s!r} ({seen(s)})"
    return s


def ok_ci(s, n=2):
    k = sum(1 for t in TEXTS if s.lower() in t.lower())
    assert k >= n, f"not confirmed (case-insensitive) in >= {n} sources: {s!r} ({k})"
    return s


T = "PDF p.1 schematic image (local OCR, >=2 identical variants)"
L = "PDF p.1 typed text box (exact text layer)"
# (type, name, topMd, bottomMd, OD, ID, extra, estimated-note, source note) — every string checked with ok()
C = [
    ("Casing", ok_ci("CASING"), None, None, 7, None, {"grade": ok("N-80")},
     None, "casing table: size '7', grade 'N-80', thread 'BUTTRESS'; weight illegible (not written); word 'CASING' confirmed case-insensitively only (letter case not verifiable)"),
    ("Tubing", ok("TUBING"), None, None, None, None, {}, None,
     "tubing table row: size, weight, grade and thread not read consistently (e.g. 276\"/278\"/27e\", AMS-28/AMS 28) — not written"),
    ("Flow Coupling", ok("FLOW COUPLING"), None, None, None, None, {}, None,
     "depth readings conflict (446.6 x2, 146.6, 148.6; 446.6 is out of sequence before 150.6) — not written"),
    ("SCSSV", ok("CAMCO TRM-4A TRSCSSV"), ok("150.6"), None, float(ok("4.700")), float(ok("2.313")), {}, "TRSCSSV -> scssv", ""),
    ("Nipple (X)", ok("Cam X- NIPPLE"), ok("302.7"), None, float(ok("3.500")), float(ok("2.313")), {}, None, ""),
    ("SPM (Gas Lift)", ok("SPM (w/ BK1 8/64, Ptro 815 psi)"), ok("490.9"), None, 4.7, float(ok("2.410")), {}, None,
     "text box on this row: 'SPM ( Dummy valve) August 2024'"),
    ("SPM (Gas Lift)", ok("SPM (w/ dummy)"), 653.4, None, 4.7, 2.41, {}, None,
     "depth digits '653 4' in 2 variants and '653.4' in 1 (decimal point lost); text box on this row: 'SPM ( Dummy valve) August 2024'"),
    ("SPM (Gas Lift)", ok("SPM (w/ BK1 12/64, Ptro 875 psi)"), ok("824.7"), None, 4.7, 2.41, {}, None,
     "label read 'Ptro' x3 / 'Piro' x2; text box on this row: 'SPM ( Dummy valve) August 2024'"),
    ("SPM (Gas Lift)", ok("SPM (w/ BK1 12/64, Ptro 900 psi)"), ok("1119.6"), None, None, None, {}, None,
     "OD/ID not read on this row; text box on this row: 'SPM ( Dummy valve) August 2024'"),
    ("SPM (Gas Lift)", ok("SPM (w/ GLOV 16/64)"), ok("1348.1"), None, None, None, {}, None,
     "OD/ID not read on this row; text box on this row: 'SPM ( GLOV 10/64) August 2024'"),
    ("Nipple (X)", ok("X- NIPPLE"), ok("1378.4"), None, None, None, {}, None, ""),
    ("Packer", ok('7" x 4.5" x 3.88" BWD GP Packer'), ok("1384.1"), None, None, None, {}, None,
     "depth read 1384.1 x2, 1284.1 x1, 1304.1 x1 — 1384.1 written (2 identical readings, in sequence)"),
    ("Sliding Sleeve", ok("SLIMLINE SSD (3"), ok("1408.4"), None, None, None, {}, "SSD -> sliding-sleeve",
     "full label read as 'SLIMLINE SSD (3,22” OD) X-PROFILE' x3 / '(3.22\" OO)' x2 / '(3.22\" OD)' x1 — OD digits not confirmed, name kept to the confirmed part"),
    ("Pup Joint", ok('4 1/2" BLANK PIPE'), None, None, None, None, {}, "blank pipe -> pup-joint", "no depth printed"),
    ("Screen", ok('4" W.W.SCREEN'), None, None, None, None, {}, None, "no depth printed"),
    ("Seal Assembly", ok('4.5" COLLET SEAL ASSY.'), None, None, None, None, {}, None, "no depth printed"),
    ("Packer", ok('7" x 4.5" x 3.88" BWD Sump Packer'), ok("1435.2"), None, None, None, {}, None, ""),
    ("Sliding Sleeve", ok("SLIMLINE SSD (3.22” OD) X-PROFILE"), ok("1451.7"), None, 3.22, None, {}, "SSD -> sliding-sleeve",
     "OD 3.22 as printed in the label; text box: '2.31” PXX plug set at SSD#2 (Sep 2018)'"),
    ("Blast Joint", ok("BLAST JOINT (1457.8 - 1485.8 m)"), 1457.8, 1485.8, float(ok('3.20"')[:-1]), None, {}, None, "'(MAX OD 3.20\")'"),
    ("Packer", ok('7" x 4.5" x 3.88" BWD Sump Packer'), ok("1533.3"), None, None, None, {}, None, ""),
    ("No-Go Nipple", ok("NOGO NIPPLE"), ok("1540.9"), None, None, None, {}, None, ""),
    ("Mule Shoe", ok("CONICAL"), ok("1541.7"), None, None, None, {}, "conical -> mule-shoe", ""),
    ("PBTD", ok("PBTD @ 1602 M"), 1602, None, None, None, {}, None, ""),
    ("Perforation", ok("E14 (OIL PRODUCER)"), 1415, 1424, None, None, {"zone": "E14", "interval": ok("1415 - 1424")}, None,
     "status on schematic image: 'CLOSE'; text box over it: 'OPEN' / 'Sep 2018'"),
    ("Perforation", ok("E23 Lower (OIL PRODUCER)"), 1485.0, 1492.5, None, None, {"zone": "E23 Lower", "interval": ok("1485.0 - 1492.5")}, None,
     "status on schematic image: 'OPEN'; text boxes over it: 'CLOSED' / 'Sep 2018' and 'E22'"),
    ("Perforation", ok("E36 (OIL PRODUCER)"), 1557.0, 1568.0, None, None, {"zone": "E36", "interval": ok("1557.0 - 1568.0")}, None,
     "status on schematic image: 'OPEN'; text box over it: 'CLOSED' / 'Sep 2018'"),
    ("Mechanical Plug", ok("2.31” PXX plug set at SSD#2 (Sep 2018)"), None, None, None, None, {}, "PXX plug -> mechanical-plug",
     "typed text box; no depth number printed"),
    ("Packer", ok("Baker tubing packoff (635 – 641m WLTHF)"), 635, 641, None, None, {}, "tubing packoff -> packer",
     "typed text box; depths are WLTHF as printed"),
]


def main():
    assert seen("1415 - 1424") >= 2 and seen("1557.0 - 1568.0") >= 2 and seen("653 4") >= 2 and seen("653.4") >= 1
    assert seen("26-Oct-01") >= 2 and seen("25-Oct-01") >= 2          # conflicting date readings -> not written
    tpl = json.loads(subprocess.run(["textutil", "-convert", "txt", "-stdout", str(TPL)], capture_output=True, text=True, check=True).stdout)
    keymap = {r[0]: r[2] for r in load_workbook(XTPL)["Component_Types"].iter_rows(min_row=2, values_only=True)}
    wi = {k: (None if isinstance(v, (int, float)) and not isinstance(v, bool) else "") for k, v in tpl["wellInfo"].items()}
    wi["name"] = "Dulang_A1"
    wi["completionType"] = ok("SINGLE STRING SELECTIVE OIL PRODUCER")
    wi["completionFluid"] = ok("8.6 ppg FILTERED SEAWATER WITH 3% KCI", 4)
    comps = []
    for n, (ctype, name, top, bot, od, idd, extra, est, note) in enumerate(C, 1):
        num = lambda v: (float(v) if "." in v else int(v)) if isinstance(v, str) else v
        d = {"id": f"dulang_a1-{n:02d}", "type": keymap[ctype], "name": name, "topMd": num(top), "bottomMd": num(bot),
             "outerDiameter": od, "innerDiameter": idd}
        d.update(extra)
        if est:
            d.update(estimated=True, estimatedFields=["type"], estimationNote=f"label is not literally a template type: {est}")
        d["source"] = (L if seen(name) == 99 else T) + (f"; {note}" if note else "") + "; depth units as printed (m MDRKB unless stated)"
        comps.append(d)
    doc = {k: {"wellInfo": wi, "components": comps}.get(k, []) for k in tpl}
    (OUT / "Dulang_A1.json").write_text(json.dumps(doc, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    print("Dulang_A1.json written:", len(comps), "components;", {k: v for k, v in wi.items() if v not in ("", None)})


if __name__ == "__main__":
    main()
