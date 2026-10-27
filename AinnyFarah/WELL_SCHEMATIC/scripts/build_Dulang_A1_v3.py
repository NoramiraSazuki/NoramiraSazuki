#!/usr/bin/env python3
"""Dulang_A1.json (v3 re-extraction) from PDF page 1 (label 'DLA-01') under the controlled-interpretation rules (page 1 only).

Reading: local Tesseract only — 22 OCR texts: 6 variants of the page's native schematic image re-run with the page-position
pipeline (ocr_A1_v3/img_x4_*), the earlier 6 native-image variants (ocr_A1/), and the 400/600 dpi page renders (ocr/page01_*,
ocr_passes/p01_*) — plus the exact PDF text layer with positions (ocr_A1_v3/text_layer.json), used to tie each typed update
box to the row it sits on. Every string used is asserted to occur verbatim in >= 2 OCR texts or in the text layer.
Depths: m MDRKB as printed -> ft (x 3.28084, 3 dp); metres kept in *M keys. Each source note says DIRECT / CALCULATED /
APPROX-TYPE / INTERPRETED / ASSUMED. Single-string well: no 'string' key.
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
RUN = OUT / "run_A1_v3"
TPL = ROOT / "well-schematic-template-json.rtf"
XTPL = ROOT / "well-schematic-sample-template.xlsx"
FT = 3.28084
OCR_FILES = sorted(glob.glob(str(OUT / "ocr_A1_v3" / "img_x4_*.txt"))) + sorted(glob.glob(str(OUT / "ocr_A1" / "img_x4_*.txt"))) + \
            sorted(glob.glob(str(OUT / "ocr" / "page01_*.txt"))) + sorted(glob.glob(str(OUT / "ocr_passes" / "p01_*.txt")))
TEXTS = [re.sub(r"[ \t]+", " ", Path(f).read_text(errors="ignore")) for f in OCR_FILES]
LAYER = [l["text"] for l in json.loads((OUT / "ocr_A1_v3" / "text_layer.json").read_text())]
REF = "m MDRKB (column header 'DEPTH (m MDRKB)')"
POS = "typed text box at the same page height as this row (x108, page-point y within 1-2 pt of the row) — association by position"


def n_seen(s, ci=False):
    if any((s.lower() in t.lower()) if ci else (s in t) for t in LAYER):
        return "text layer"
    return sum(1 for t in TEXTS if (s.lower() in t.lower() if ci else s in t))


def conf(s, ci=False):
    k = n_seen(s, ci)
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


# (id, Component_Types name, name as written, top m, bottom m, OD, ID, extra, approx-type, (estimated fields, note), source note)
COMPS = [
    ("casing-7", "Casing", "CASING", None, 1602.0, 7.0, None, {"grade": "N-80"}, None,
     (["bottomMd"], "ASSUMED (user instruction): casing bottom depth taken as the PBTD depth 'PBTD @ 1602 M' — no casing shoe depth is printed"),
     "Casing/tubing table (top right): TYPE 'CASING' | SIZE '7' | WT/FT illegible | GRADE 'N-80' (10) | THREAD 'BUTTRESS' (7) — DIRECT; "
     "word 'CASING' confirmed case-insensitively only; bottom = PBTD (see estimate); top not printed"),
    ("tubing", "Tubing", "TUBING", None, 1541.7, 2.875, None, {}, None,
     (["outerDiameter", "bottomMd"], "INTERPRETED: SIZE read '278\"' (2) / '276\"' (2) / '27e\"' (3) — taken as 2 7/8\" with the fraction glyph misread; "
      "ASSUMED (user instruction): tubing bottom depth taken as the CONICAL depth 1541.7 (end of string) — no EOT depth is printed"),
     "Casing/tubing table: TYPE 'TUBING' | SIZE see estimate | WT/FT and GRADE illegible | THREAD 'AMS 28' (4; 'AMS-28' 3) — DIRECT; "
     "bottom = CONICAL depth (see estimate); top not printed"),
    ("flow-coupling", "Flow Coupling", "FLOW COUPLING", 146.6, None, None, None, {}, None, None,
     "Equipment list row 1 — label DIRECT; depth 146.6 (9 readings) vs 446.6 (2) / 148.6 (2); 146.6 is also the only one in sequence before 150.6"),
    ("trscssv", "SCSSV", "CAMCO TRM-4A TRSCSSV", 150.6, None, 4.700, 2.313, {}, "TRSCSSV (tubing-retrievable SCSSV) -> scssv", None,
     "Equipment list row 2 — DIRECT (OD / ID columns 'MAX OD' / 'MIN. I.D (IN)')"),
    ("x-nipple-1", "Nipple (X)", "Cam X- NIPPLE", 302.7, None, 3.500, 2.313, {}, None, None, "Equipment list row 3 — DIRECT"),
    ("spm-1", "SPM (Gas Lift)", "SPM (w/ BK1 8/64, Ptro 815 psi)", 490.9, None, 4.700, 2.410, {}, None, None,
     f"Equipment list row 4 — DIRECT; {POS}: 'SPM ( Dummy valve) August 2024' (later valve change)"),
    ("spm-2", "SPM (Gas Lift)", "SPM (w/ dummy)", 653.4, None, 4.700, 2.410, {}, None, None,
     f"Equipment list row 5 — DIRECT label; depth '653.4' (4) / '653 4' (8, decimal point lost); {POS}: 'SPM ( Dummy valve) August 2024'"),
    ("spm-3", "SPM (Gas Lift)", "SPM (w/ BK1 12/64, Ptro 875 psi)", 824.7, None, 4.700, 2.410, {}, None, None,
     f"Equipment list row 6 — DIRECT; {POS}: 'SPM ( Dummy valve) August 2024'"),
    ("spm-4", "SPM (Gas Lift)", "SPM (w/ BK1 12/64, Ptro 900 psi)", 1119.6, None, None, None, {}, None, None,
     f"Equipment list row 7 — DIRECT; OD/ID columns blank on this row in every reading; {POS}: 'SPM ( Dummy valve) August 2024'"),
    ("spm-5", "SPM (Gas Lift)", "SPM (w/ GLOV 16/64)", 1348.1, None, None, None, {}, None, None,
     f"Equipment list row 8 — DIRECT; OD/ID columns blank; {POS}: 'SPM ( GLOV 10/64) August 2024'"),
    ("x-nipple-2", "Nipple (X)", "X- NIPPLE", 1378.4, None, None, None, {}, None, None, "Equipment list row 9 — DIRECT; OD/ID columns blank"),
    ("gp-packer", "Packer", '7" x 4.5" x 3.88" BWD GP Packer', 1384.1, None, None, None, {}, None, None,
     "Lower completion (right-hand labels) — label DIRECT; depth 1384.1 (7 readings) vs 1304.1 (4) / 1984.1 (2) / 1284.1 (1); 1384.1 is in sequence"),
    ("ssd-1", "Sliding Sleeve", 'SLIMLINE SSD (3.22" OD) X-PROFILE', 1408.4, None, 3.22, None, {}, "SLIMLINE SSD -> sliding-sleeve", None,
     "Equipment list (left labels) — DIRECT; '(3.22\" OD)' now confirmed (10 readings; '3,22' 6); OD 3.22 from the label"),
    ("blank-pipe", "Pup Joint", '4 1/2" BLANK PIPE', None, None, None, None, {}, "blank pipe -> pup-joint (no blank-pipe type)", None,
     "Lower completion — DIRECT; no depth printed"),
    ("screen", "Screen", '4" W.W.SCREEN', None, None, None, None, {}, None, None, "Lower completion beside E14 — DIRECT; no depth printed"),
    ("collet-seal", "Seal Assembly", '4.5" COLLET SEAL ASSY.', None, None, None, None, {}, None, None, "Lower completion — DIRECT; no depth printed"),
    ("sump-packer-1", "Packer", '7" x 4.5" x 3.88" BWD Sump Packer', 1435.2, None, None, None, {}, None, None, "Lower completion — DIRECT (11)"),
    ("ssd-2", "Sliding Sleeve", "SLIMLINE SSD (3.22” OD) X-PROFILE", 1451.7, None, 3.22, None, {}, "SLIMLINE SSD -> sliding-sleeve", None,
     "Lower completion — DIRECT; OD 3.22 from the label; see pxx-plug ('SSD#2')"),
    ("blast-joint", "Blast Joint", "BLAST JOINT (1457.8 - 1485.8 m)", 1457.8, 1485.8, 3.20, None, {}, None, None,
     "Lower completion — DIRECT incl. interval and '(MAX OD 3.20\")'"),
    ("sump-packer-2", "Packer", '7" x 4.5" x 3.88" BWD Sump Packer', 1533.3, None, None, None, {}, None, None, "Lower completion — DIRECT"),
    ("nogo-nipple", "No-Go Nipple", "NOGO NIPPLE", 1540.9, None, None, None, {}, None, None, "Lower completion — DIRECT"),
    ("conical", "Mule Shoe", "CONICAL", 1541.7, None, None, None, {}, "CONICAL (end-of-string guide) -> mule-shoe", None, "Lower completion — DIRECT"),
    ("pbtd", "PBTD", "PBTD @ 1602 M", 1602.0, None, None, None, {}, None, None, "Bottom of diagram — DIRECT (14)"),
    ("perf-e14", "Perforation", "E14 (OIL PRODUCER)", 1415.0, 1424.0, None, None, {"zone": "E14", "interval": "1415 - 1424"}, None, None,
     "Zone label + 'Perf. Intervals 1415 - 1424' — DIRECT; status column on image 'CLOSE'; typed box directly over that status: 'OPEN' / 'Sep 2018'"),
    ("perf-e23-lower", "Perforation", "E23 Lower (OIL PRODUCER)", 1485.0, 1492.5, None, None, {"zone": "E23 Lower", "interval": "1485.0 - 1492.5"}, None, None,
     "Zone label + 'Perf. Intervals 1485.0 - 1492.5' — DIRECT; status on image 'OPEN'; typed box over it 'CLOSED' / 'Sep 2018'; ZONE CONFLICT (unresolved): "
     "a typed 'E22' box sits directly on top of the word 'E23' in the label (likely a later correction) — image label kept, both recorded"),
    ("perf-e36", "Perforation", "E36 (OIL PRODUCER)", 1557.0, 1568.0, None, None, {"zone": "E36", "interval": "1557.0 - 1568.0"}, None, None,
     "Zone label + 'Perf. Intervals 1557.0 - 1568.0' — DIRECT; status on image 'OPEN'; typed box over it 'CLOSED' / 'Sep 2018'"),
    ("pxx-plug", "Mechanical Plug", "2.31” PXX plug set at SSD#2 (Sep 2018)", 1451.7, None, None, None, {}, "PXX wireline plug -> mechanical-plug", None,
     "Typed text box (right side, between the 1435.2 sump packer and the 1451.7 SSD) — DIRECT name; depth ASSUMED = second SSD 1451.7 ('at SSD#2'); "
     "partial date 'Sep 2018'"),
    ("tubing-packoff", "Packer", "Baker tubing packoff (635 – 641m WLTHF)", None, None, None, None, {}, "tubing packoff -> packer", None,
     "Typed text box beside the tubing between the 490.9 and 653.4 rows — DIRECT; 635 - 641 m are WLTHF (different depth reference from m MDRKB) — "
     "not placed in topMd/bottomMd"),
]
CHECK = ["CASING", "N-80", "BUTTRESS", "TUBING", "278\"", "27e\"", "AMS 28", "FLOW COUPLING", "146.6", "CAMCO TRM-4A TRSCSSV", "150.6", "4.700",
         "2.313", "Cam X- NIPPLE", "302.7", "3.500", "SPM (w/ BK1 8/64, Ptro 815 psi)", "490.9", "2.410", "SPM (w/ dummy)", "653.4",
         "SPM (w/ BK1 12/64, Ptro 875 psi)", "824.7", "SPM (w/ BK1 12/64, Ptro 900 psi)", "1119.6", "SPM (w/ GLOV 16/64)", "1348.1", "X- NIPPLE",
         "1378.4", '7" x 4.5" x 3.88" BWD GP Packer', "1384.1", 'SLIMLINE SSD (3.22" OD) X-PROFILE', "1408.4", '4 1/2" BLANK PIPE', '4" W.W.SCREEN',
         '4.5" COLLET SEAL ASSY.', '7" x 4.5" x 3.88" BWD Sump Packer', "1435.2", "SLIMLINE SSD (3.22” OD) X-PROFILE", "1451.7",
         "BLAST JOINT (1457.8 - 1485.8 m)", '3.20"', "1533.3", "NOGO NIPPLE", "1540.9", "CONICAL", "1541.7", "PBTD @ 1602 M", "MDRKB",
         "E14 (OIL PRODUCER)", "1415 - 1424", "E23 Lower (OIL PRODUCER)", "1485.0 - 1492.5", "E36 (OIL PRODUCER)", "1557.0 - 1568.0", "CLOSE", "OPEN",
         "2.31” PXX plug set at SSD#2 (Sep 2018)", "Baker tubing packoff (635 – 641m WLTHF)", "SPM ( Dummy valve) August 2024",
         "SPM ( GLOV 10/64) August 2024", "E22", "Sep 2018", "CLOSED", "SINGLE STRING SELECTIVE OIL PRODUCER", "8.6 ppg FILTERED SEAWATER WITH 3% KCI",
         "26-Oct-01", "DLA-01 (WORKOVER", "2.29” GR & 2” LIB tag HUD at 1430m WLTHF.", "Impression suspected scale (Nov 2022)"]


def main():
    hb = [l.split("  ", 1) for l in (RUN / "sha256_before.txt").read_text().splitlines()]
    confirm = {s: conf(s, ci=s == "CASING") for s in CHECK}
    tpl = json.loads(subprocess.run(["textutil", "-convert", "txt", "-stdout", str(TPL)], capture_output=True, text=True, check=True).stdout)
    keymap = {r[0]: r[2] for r in load_workbook(XTPL)["Component_Types"].iter_rows(min_row=2, values_only=True)}
    wi = {k: v if v in ("-", 0) else (None if isinstance(v, (int, float)) and not isinstance(v, bool) else "") for k, v in tpl["wellInfo"].items()}
    wi.update(name="Dulang_A1", completionType="SINGLE STRING SELECTIVE OIL PRODUCER", completionFluid="8.6 ppg FILTERED SEAWATER WITH 3% KCI",
              completionDate="26-Oct-01", tubing='TUBING 2 7/8" AMS 28')
    comps = []
    for cid, ctype, name, top, bot, od, idd, extra, approx, est, note in COMPS:
        d = {"id": cid, "type": keymap[ctype], "name": name, "topMd": ft(top), "bottomMd": ft(bot), "outerDiameter": od, "innerDiameter": idd}
        d.update(extra)
        fields, notes = [], []
        if approx:
            fields.append("type"); notes.append(f"APPROX-TYPE: no exact Type_Key; {approx}")
        if est:
            fields += est[0]; notes.append(est[1])
        if fields:
            d.update(estimated=True, estimatedFields=fields, estimationNote="; ".join(notes))
        conv = f"; depth reference {REF}; CALCULATED: ft = m x 3.28084 (3 dp)" if top is not None or bot is not None else ""
        d["source"] = f"PDF page 1 (label 'DLA-01'): {note}{conv}"
        d.update(actualType=name, topMdM=top, bottomMdM=bot)
        comps.append(d)
    doc = {k: {"wellInfo": wi, "components": comps}.get(k, []) for k in tpl}
    out = OUT / "Dulang_A1.json"
    if out.exists():                       # naming collision: keep the earlier file, never overwrite silently
        bak = OUT / "_previous_runs" / f"Dulang_A1_before_v3_{dt.datetime.now():%Y%m%d_%H%M%S}.json"
        shutil.copy2(out, bak)
        print("existing Dulang_A1.json backed up ->", bak.name)
    out.write_text(json.dumps(doc, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")

    back = json.loads(out.read_text())
    kind = lambda v: "number" if isinstance(v, (int, float)) and not isinstance(v, bool) else type(v).__name__
    cs = back["components"]
    approx_ids = {c[0] for c in COMPS if c[8]}
    v = [("valid JSON", True),
         ("top-level sections and order = template", list(back) == list(tpl)),
         ("wellInfo keys = template", list(back["wellInfo"]) == list(tpl["wellInfo"])),
         ("wellInfo types = template (where populated)", all(x in ("", None) or kind(x) == kind(tpl["wellInfo"][k]) for k, x in back["wellInfo"].items())),
         ("no template sample records", back["surveys"] == [] and back["formationTops"] == [] and not {c["id"] for c in cs} & {t["id"] for t in tpl["components"]}),
         ("all types are Component_Types keys", all(c["type"] in set(keymap.values()) for c in cs)),
         ("component ids unique", len({c["id"] for c in cs}) == len(cs)),
         ("ft = m x 3.28084 (3 dp)", all((c[a] is None and c[b] is None) or c[a] == round(c[b] * FT, 3) for c in cs for a, b in (("topMd", "topMdM"), ("bottomMd", "bottomMdM")))),
         ("every component has source traceability", all(c.get("source", "").startswith("PDF page 1") for c in cs)),
         ("approximate types flagged (estimated + 'type' in estimatedFields)", all(c.get("estimated") and "type" in c["estimatedFields"] for c in cs if c["id"] in approx_ids)),
         ("non-approximate types not flagged as type estimates", all("type" not in c.get("estimatedFields", []) for c in cs if c["id"] not in approx_ids)),
         ("standardized name 'Dulang_A1' = filename", back["wellInfo"]["name"] == out.stem),
         ("source PDF unchanged", sha(hb[0][1]) == hb[0][0]),
         ("template unchanged", sha(str(ROOT / hb[1][1])) == hb[1][0]),
         ("Component_Types reference unchanged", sha(str(ROOT / hb[2][1])) == hb[2][0])]
    for a, b in v:
        print("PASS" if b else "FAIL", "-", a)
    print("components:", len(cs), "| strings confirmed:", len(confirm), "(", sum(1 for x in confirm.values() if x == "text layer"), "from text layer )")
    (RUN / "confirmations.json").write_text(json.dumps(confirm, indent=1, ensure_ascii=False))


if __name__ == "__main__":
    main()
