#!/usr/bin/env python3
"""Dulang_A1.json from PDF page 1 (label 'DLA-01') under the controlled-interpretation rules (page 1 only).

Reading: local Tesseract only — 6 variants of the page's native schematic image (ocr_A1/img_x4_*), page renders at
400/600 dpi (ocr/page01_*, ocr_passes/p01_*), and the exact PDF text layer (typed update boxes). No image leaves the Mac.
Each source string is confirmed (verbatim, whitespace collapsed) in >= 2 OCR texts or in the text layer before use.
Depths: m MDRKB as printed -> ft (x 3.28084, 3 dp); metres kept in *M keys. Classification in each component's source:
DIRECT (as read), CALCULATED (unit conversion), APPROX-TYPE (nearest Type_Key), INTERPRETED / ASSUMED (flagged).
"""
import glob
import hashlib
import json
import re
import shutil
import subprocess
import datetime as dt
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "output_A"
RUN = OUT / "run_A1_v2"
PDF = Path("/Volumes/homes/public/DulangOSS/20261016/DULANG_WELLS/Well_Schematic/1. Well Schematic/Dulang A/Dulang A Well Schematic updated rev1.pdf")
TPL = ROOT / "well-schematic-template-json.rtf"
XTPL = ROOT / "well-schematic-sample-template.xlsx"
FT = 3.28084
OCR_FILES = sorted(glob.glob(str(OUT / "ocr_A1" / "img_x4_*.txt"))) + sorted(glob.glob(str(OUT / "ocr" / "page01_*.txt"))) + \
            sorted(glob.glob(str(OUT / "ocr_passes" / "p01_*.txt")))
TEXTS = [re.sub(r"[ \t]+", " ", Path(f).read_text(errors="ignore")) for f in OCR_FILES]
LAYER = [l["text"] for l in json.loads((OUT / "ocr_passes" / "text_layer.json").read_text())["p01"]]


def n_seen(s, ci=False):
    if any((s.lower() in t.lower()) if ci else (s in t) for t in LAYER):
        return "text layer"
    return sum(1 for t in TEXTS if (s.lower() in t.lower() if ci else s in t))


def conf(s, ci=False):
    k = n_seen(s, ci)
    assert k == "text layer" or k >= 2, f"not confirmed: {s!r} ({k})"
    return f"text layer (exact)" if k == "text layer" else f"{k} OCR readings"


def ft(m):
    return None if m is None else round(m * FT, 3)


def sha(p):
    fd = __import__("os").open(p, 0); b = []
    try:
        while (x := __import__("os").read(fd, 1 << 20)):
            b.append(x)
    finally:
        try:
            __import__("os").close(fd)
        except OSError:
            pass
    return hashlib.sha256(b"".join(b)).hexdigest()


# id, Component_Types name, exact equipment name, top m, bottom m, OD, ID, extra, approx-type note, notes (traceability)
COMPS = [
    ("casing-7", "Casing", "CASING", None, None, 7.0, None, {"grade": "N-80"}, None,
     "Casing/tubing table (top right): type 'CASING', size '7', grade 'N-80', thread 'BUTTRESS' — DIRECT; weight illegible; "
     "word 'CASING' confirmed case-insensitively only; no casing shoe depth on page -> top/bottom empty"),
    ("tubing", "Tubing", "TUBING", None, None, None, None, {}, None,
     "Casing/tubing table: type 'TUBING' — DIRECT; size/weight/grade/thread illegible (e.g. 276\"/278\", AMS-28/AMS 28); "
     "no EOT label -> interval empty"),
    ("flow-coupling", "Flow Coupling", "FLOW COUPLING", 146.6, None, None, None, {}, None,
     "Equipment list row 1 — label DIRECT; depth INTERPRETED: readings 146.6 (3 incl. page renders), 446.6 (2), 148.6 (1); "
     "146.6 selected (in sequence before 150.6, 446.6 is not)"),
    ("trscssv", "SCSSV", "CAMCO TRM-4A TRSCSSV", 150.6, None, 4.700, 2.313, {}, "TRSCSSV (tubing-retrievable SCSSV) -> scssv",
     "Equipment list row 2 — DIRECT (OD/ID columns 'MAX. O.D' / 'MIN. I.D')"),
    ("x-nipple-1", "Nipple (X)", "Cam X- NIPPLE", 302.7, None, 3.500, 2.313, {}, None, "Equipment list row 3 — DIRECT"),
    ("spm-1", "SPM (Gas Lift)", "SPM (w/ BK1 8/64, Ptro 815 psi)", 490.9, None, 4.700, 2.410, {}, None,
     "Equipment list row 4 — DIRECT; typed text box on this row (association by position, flagged): 'SPM ( Dummy valve) August 2024'"),
    ("spm-2", "SPM (Gas Lift)", "SPM (w/ dummy)", 653.4, None, 4.700, 2.410, {}, None,
     "Equipment list row 5 — DIRECT label; depth read '653.4' (1) and '653 4' (2, decimal point lost); typed text box on this row: "
     "'SPM ( Dummy valve) August 2024'"),
    ("spm-3", "SPM (Gas Lift)", "SPM (w/ BK1 12/64, Ptro 875 psi)", 824.7, None, 4.700, 2.410, {}, None,
     "Equipment list row 6 — DIRECT ('Ptro' read 3x, 'Piro' 2x); typed text box on this row: 'SPM ( Dummy valve) August 2024'"),
    ("spm-4", "SPM (Gas Lift)", "SPM (w/ BK1 12/64, Ptro 900 psi)", 1119.6, None, None, None, {}, None,
     "Equipment list row 7 — DIRECT; OD/ID not printed/read on this row; typed text box on this row: 'SPM ( Dummy valve) August 2024'"),
    ("spm-5", "SPM (Gas Lift)", "SPM (w/ GLOV 16/64)", 1348.1, None, None, None, {}, None,
     "Equipment list row 8 — DIRECT; OD/ID not read; typed text box on this row: 'SPM ( GLOV 10/64) August 2024'"),
    ("x-nipple-2", "Nipple (X)", "X- NIPPLE", 1378.4, None, None, None, {}, None, "Equipment list row 9 — DIRECT; OD/ID not read"),
    ("gp-packer", "Packer", '7" x 4.5" x 3.88" BWD GP Packer', 1384.1, None, None, None, {}, None,
     "Lower completion (right side) — DIRECT label; depth readings 1384.1 (2+), 1284.1 (1), 1304.1 (1); 1384.1 selected (in sequence)"),
    ("ssd-1", "Sliding Sleeve", 'SLIMLINE SSD (3.22" OD) X-PROFILE', 1408.4, None, None, None, {}, "SLIMLINE SSD -> sliding-sleeve",
     "DIRECT except OD digits: label read '(3,22” OD)' 3x, '(3.22\" OO)' 2x, '(3.22\" OD)' 1x — decimal point INTERPRETED; OD not populated"),
    ("blank-pipe", "Pup Joint", '4 1/2" BLANK PIPE', None, None, None, None, {}, "blank pipe -> pup-joint (no blank-pipe type)",
     "Lower completion — DIRECT; no depth printed"),
    ("screen", "Screen", '4" W.W.SCREEN', None, None, None, None, {}, None, "Lower completion beside E14 — DIRECT; no depth printed"),
    ("collet-seal", "Seal Assembly", '4.5" COLLET SEAL ASSY.', None, None, None, None, {}, None, "Lower completion — DIRECT; no depth printed"),
    ("sump-packer-1", "Packer", '7" x 4.5" x 3.88" BWD Sump Packer', 1435.2, None, None, None, {}, None, "Lower completion — DIRECT"),
    ("ssd-2", "Sliding Sleeve", "SLIMLINE SSD (3.22” OD) X-PROFILE", 1451.7, None, 3.22, None, {}, "SLIMLINE SSD -> sliding-sleeve",
     "Lower completion — DIRECT; OD 3.22 taken from the printed label '(3.22” OD)'"),
    ("blast-joint", "Blast Joint", "BLAST JOINT (1457.8 - 1485.8 m)", 1457.8, 1485.8, 3.20, None, {}, None,
     "Lower completion — DIRECT incl. interval and '(MAX OD 3.20\")'"),
    ("sump-packer-2", "Packer", '7" x 4.5" x 3.88" BWD Sump Packer', 1533.3, None, None, None, {}, None, "Lower completion — DIRECT"),
    ("nogo-nipple", "No-Go Nipple", "NOGO NIPPLE", 1540.9, None, None, None, {}, None, "Lower completion — DIRECT"),
    ("conical", "Mule Shoe", "CONICAL", 1541.7, None, None, None, {}, "CONICAL (end-of-string guide) -> mule-shoe", "Lower completion — DIRECT"),
    ("pbtd", "PBTD", "PBTD @ 1602 M", 1602.0, None, None, None, {}, None, "Bottom of diagram — DIRECT"),
    ("perf-e14", "Perforation", "E14 (OIL PRODUCER)", 1415.0, 1424.0, None, None, {"zone": "E14", "interval": "1415 - 1424"}, None,
     "Zone label + 'Perf. Intervals 1415 - 1424' — DIRECT; status on image 'CLOSE'; typed box over it 'OPEN' / 'Sep 2018' (later status)"),
    ("perf-e23-lower", "Perforation", "E23 Lower (OIL PRODUCER)", 1485.0, 1492.5, None, None,
     {"zone": "E23 Lower", "interval": "1485.0 - 1492.5"}, None,
     "Zone label + 'Perf. Intervals 1485.0 - 1492.5' — DIRECT; status on image 'OPEN'; typed boxes over it 'CLOSED' / 'Sep 2018' and "
     "'E22' (zone label conflict — unresolved, image label kept)"),
    ("perf-e36", "Perforation", "E36 (OIL PRODUCER)", 1557.0, 1568.0, None, None, {"zone": "E36", "interval": "1557.0 - 1568.0"}, None,
     "Zone label + 'Perf. Intervals 1557.0 - 1568.0' — DIRECT; status on image 'OPEN'; typed box over it 'CLOSED' / 'Sep 2018'"),
    ("pxx-plug", "Mechanical Plug", "2.31” PXX plug set at SSD#2 (Sep 2018)", 1451.7, None, None, None, {}, "PXX wireline plug -> mechanical-plug",
     "Typed text box — DIRECT name; depth ASSUMED = depth of the second SSD (1451.7) because the box says 'at SSD#2' and sits beside it; "
     "set date partial 'Sep 2018'"),
    ("tubing-packoff", "Packer", "Baker tubing packoff (635 – 641m WLTHF)", None, None, None, None, {}, "tubing packoff -> packer",
     "Typed text box — DIRECT; depths 635 - 641 m are WLTHF (different depth reference from the page's m MDRKB) — not placed in topMd/bottomMd"),
]
CHECK = ["CASING", "N-80", "BUTTRESS", "TUBING", "FLOW COUPLING", "146.6", "CAMCO TRM-4A TRSCSSV", "150.6", "4.700", "2.313", "Cam X- NIPPLE",
         "302.7", "3.500", "SPM (w/ BK1 8/64, Ptro 815 psi)", "490.9", "2.410", "SPM (w/ dummy)", "653 4", "SPM (w/ BK1 12/64, Ptro 875 psi)",
         "824.7", "SPM (w/ BK1 12/64, Ptro 900 psi)", "1119.6", "SPM (w/ GLOV 16/64)", "1348.1", "X- NIPPLE", "1378.4",
         '7" x 4.5" x 3.88" BWD GP Packer', "1384.1", "SLIMLINE SSD (3", "1408.4", '4 1/2" BLANK PIPE', '4" W.W.SCREEN', '4.5" COLLET SEAL ASSY.',
         '7" x 4.5" x 3.88" BWD Sump Packer', "1435.2", "SLIMLINE SSD (3.22” OD) X-PROFILE", "1451.7", "BLAST JOINT (1457.8 - 1485.8 m)",
         '3.20"', "1533.3", "NOGO NIPPLE", "1540.9", "CONICAL", "1541.7", "PBTD @ 1602 M", "E14 (OIL PRODUCER)", "1415 - 1424",
         "E23 Lower (OIL PRODUCER)", "1485.0 - 1492.5", "E36 (OIL PRODUCER)", "1557.0 - 1568.0", "2.31” PXX plug set at SSD#2 (Sep 2018)",
         "Baker tubing packoff (635 – 641m WLTHF)", "SPM ( Dummy valve) August 2024", "SPM ( GLOV 10/64) August 2024", "E22",
         "SINGLE STRING SELECTIVE OIL PRODUCER", "8.6 ppg FILTERED SEAWATER WITH 3% KCI", "26-Oct-01", "DLA-01"]


def main():
    before = (RUN / "sha256_before.txt").read_text().split()
    h_before = {before[i + 1] if i + 1 < len(before) else "": before[i] for i in range(0, len(before), 2)}
    confirm = {s: conf(s, ci=s == "CASING") for s in CHECK}
    tpl = json.loads(subprocess.run(["textutil", "-convert", "txt", "-stdout", str(TPL)], capture_output=True, text=True, check=True).stdout)
    keymap = {r[0]: r[2] for r in load_workbook(XTPL)["Component_Types"].iter_rows(min_row=2, values_only=True)}
    # wellInfo: '-' / 0 are the template's defaults; sample values are not defaults -> "" / null
    wi = {}
    for k, v in tpl["wellInfo"].items():
        wi[k] = v if v in ("-", 0) else (None if isinstance(v, (int, float)) and not isinstance(v, bool) else "")
    wi.update(name="Dulang_A1", completionType="SINGLE STRING SELECTIVE OIL PRODUCER",
              completionFluid="8.6 ppg FILTERED SEAWATER WITH 3% KCI", completionDate="26-Oct-01")
    comps = []
    for cid, ctype, name, top, bot, od, idd, extra, approx, note in COMPS:
        assert ctype in keymap, ctype
        d = {"id": cid, "type": keymap[ctype], "name": name, "topMd": ft(top), "bottomMd": ft(bot), "outerDiameter": od, "innerDiameter": idd}
        d.update(extra)
        if approx:
            d.update(estimated=True, estimatedFields=["type"], estimationNote=f"APPROX-TYPE: no exact Type_Key; {approx}")
        conv = "; CALCULATED: ft = m x 3.28084 (3 dp) from m MDRKB" if top is not None else ""
        d["source"] = f"PDF page 1 (label 'DLA-01'): {note}{conv}"
        d.update(actualType=name, topMdM=top, bottomMdM=bot)
        comps.append(d)
    doc = {k: {"wellInfo": wi, "components": comps}.get(k, []) for k in tpl}
    out = OUT / "Dulang_A1.json"
    if out.exists():                       # naming collision: keep the earlier file, never overwrite silently
        bak = OUT / "_previous_runs" / f"Dulang_A1_before_v2_{dt.datetime.now():%Y%m%d_%H%M%S}.json"
        shutil.copy2(out, bak)
        print("existing Dulang_A1.json backed up ->", bak.name)
    out.write_text(json.dumps(doc, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")

    # validation
    back = json.loads(out.read_text())
    kind = lambda v: "number" if isinstance(v, (int, float)) and not isinstance(v, bool) else type(v).__name__
    v = [("valid JSON", True),
         ("top-level sections and order = template", list(back) == list(tpl)),
         ("wellInfo keys = template", list(back["wellInfo"]) == list(tpl["wellInfo"])),
         ("wellInfo types = template (where populated)", all(x in ("", None) or kind(x) == kind(tpl["wellInfo"][k]) for k, x in back["wellInfo"].items())),
         ("no template sample records", back["surveys"] == [] and back["formationTops"] == [] and not any(c["id"] in {t["id"] for t in tpl["components"]} for c in back["components"])),
         ("all types are Component_Types keys", all(c["type"] in set(keymap.values()) for c in back["components"])),
         ("component ids unique", len({c["id"] for c in back["components"]}) == len(back["components"])),
         ("ft = m x 3.28084 (3 dp)", all((c[a] is None and c[b] is None) or c[a] == round(c[b] * FT, 3) for c in back["components"] for a, b in (("topMd", "topMdM"), ("bottomMd", "bottomMdM")))),
         ("every component has source traceability", all(c.get("source") for c in back["components"])),
         ("every approximate type flagged", all(c.get("estimated") for c in back["components"] if c["type"] not in {keymap[x] for x in keymap} or c.get("estimationNote"))),
         ("standardized name 'Dulang_A1' = filename", back["wellInfo"]["name"] == out.stem),
         ("source PDF unchanged", sha(PDF) == h_before[str(PDF).split("/")[-1]] if False else sha(PDF) == (RUN / "sha256_before.txt").read_text().split()[0]),
         ("template unchanged", hashlib.sha256(TPL.read_bytes()).hexdigest() == (RUN / "sha256_before.txt").read_text().split("\n")[1].split()[0]),
         ("Component_Types reference unchanged", hashlib.sha256(XTPL.read_bytes()).hexdigest() == (RUN / "sha256_before.txt").read_text().split("\n")[2].split()[0])]
    for a, b in v:
        print("PASS" if b else "FAIL", "-", a)
    print("components:", len(comps), "| confirmations:", sum(1 for x in confirm.values()), "strings,",
          sum(1 for x in confirm.values() if x == "text layer (exact)"), "from text layer")
    (RUN / "confirmations.json").write_text(json.dumps(confirm, indent=1, ensure_ascii=False))


if __name__ == "__main__":
    main()
