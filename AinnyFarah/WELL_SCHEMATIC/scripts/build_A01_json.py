#!/usr/bin/env python3
"""DLA-01 (Dulang A Well Schematic updated rev1.pdf, page 1) -> well-schematic-template-json structure.

Source is read-only on the NAS. Page 1 = well DLA-01 (OCR header 'COMPLETED WELL COMPLETION DIAGRAM : DLA-01').
The schematic is an embedded raster image (OCR'd locally with Tesseract, 3x upscaled); later updates are typed
text boxes on top of it (exact PDF text layer). Values below are transcribed from those two sources and each carries
its source; illegible items are left empty and flagged. Conventions as for the A11ST2 sample JSON: absent data empty
(null / ""), depths ft (m x 3.28084) with meters kept in *M keys, nearest Component_Types key with actualType kept.
"""
import datetime as dt
import hashlib
import json
import os
import subprocess
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "output_A01"
PDF = Path("/Volumes/homes/public/DulangOSS/20261016/DULANG_WELLS/Well_Schematic/1. Well Schematic/Dulang A/"
           "Dulang A Well Schematic updated rev1.pdf")
TPL = ROOT / "well-schematic-template-json.rtf"
XTPL = ROOT / "well-schematic-sample-template.xlsx"
FT = 3.28084
ft = lambda m: None if m is None else round(m * FT, 3)
IMG, TXT = "page 1 schematic image (OCR)", "page 1 text box (PDF text layer)"


def read(p):
    fd = os.open(p, os.O_RDONLY)
    b = []
    try:
        while (x := os.read(fd, 1 << 20)):
            b.append(x)
    finally:
        try:
            os.close(fd)
        except OSError:
            pass
    return b"".join(b)


# (id, Component_Type, name as on schematic, top m, bottom m, OD in, ID in, actualType, source, extra, estimate-note)
COMP = [
    ("casing-7in", "Casing", 'CASING 7" N-80 BUTTRESS', None, None, 7.0, None, "CASING", IMG,
     {"grade": "N-80", "description": "Thread: BUTTRESS; weight illegible in scan"}, None),
    ("tubing", "Tubing", 'TUBING 2-7/8" (as read) AMS-28', None, None, 2.875, None, "TUBING", IMG,
     {"description": "Size read as '2 7/8\"' from a partly illegible table (consistent with the 2.313\" X-nipples); weight and grade illegible; thread AMS-28"},
     "size partly legible"),
    ("flow-coupling-1", "Flow Coupling", "FLOW COUPLING", 146.6, None, None, None, "FLOW COUPLING", IMG, {}, None),
    ("trscssv", "SCSSV", "CAMCO TRM-4A TRSCSSV", 150.6, None, 4.700, 2.313, "TRSCSSV", IMG,
     {"manufacturer": "CAMCO", "model": "TRM-4A"}, None),
    ("x-nipple-1", "Nipple (X)", "X- NIPPLE", 302.7, None, 3.500, 2.313, "X-NIPPLE", IMG, {"manufacturer": "Camco (as 'Cam')"}, None),
    ("spm-1", "SPM (Gas Lift)", "SPM ( Dummy valve) August 2024", 490.9, None, 4.700, 2.410, "SPM", TXT,
     {"description": "August 2024 update (text box): dummy valve. Original schematic: 'SPM (w/ BK1 8/64, Ptro 815 psi)'"}, None),
    ("spm-2", "SPM (Gas Lift)", "SPM ( Dummy valve) August 2024", 653.4, None, 4.700, 2.410, "SPM", TXT,
     {"description": "August 2024 update (text box): dummy valve. Original schematic: 'SPM (w/ dummy)'"}, None),
    ("spm-3", "SPM (Gas Lift)", "SPM ( Dummy valve) August 2024", 824.7, None, 4.700, 2.410, "SPM", TXT,
     {"description": "August 2024 update (text box): dummy valve. Original schematic: 'SPM (w/ BK1 12/64, Ptro 875 psi)'"}, None),
    ("spm-4", "SPM (Gas Lift)", "SPM ( Dummy valve) August 2024", 1119.6, None, None, None, "SPM", TXT,
     {"description": "August 2024 update (text box): dummy valve. Original schematic: 'SPM (w/ BK1 12/64, Ptro 900 psi)'. OD/ID not legible"}, None),
    ("spm-5", "SPM (Gas Lift)", "SPM ( GLOV 10/64) August 2024", 1348.1, None, None, None, "SPM", TXT,
     {"hasValve": True, "description": "August 2024 update (text box): GLOV 10/64. Original schematic: 'SPM (w/ GLOV 16/64)'. OD/ID not legible"}, None),
    ("x-nipple-2", "Nipple (X)", "X- NIPPLE", 1378.4, None, None, None, "X-NIPPLE", IMG, {}, None),
    ("gp-packer", "Packer", '7" x 4.5" x 3.88" BWD GP Packer', 1384.1, None, None, 3.88, "GRAVEL PACK PACKER (BWD)", IMG,
     {"model": "BWD", "description": "Depth OCR'd as 1384.1 (one pass read 1284.1; 1384.1 fits between X-nipple 1378.4 and SSD 1408.4)"}, None),
    ("ssd-1", "Sliding Sleeve", 'SLIMLINE SSD (3.22" OD) X-PROFILE', 1408.4, None, 3.22, None, "SLIMLINE SSD (X-PROFILE)", IMG,
     {"description": "SSD#1"}, None),
    ("blank-pipe", "Pup Joint", '4 1/2" BLANK PIPE', None, None, 4.5, None, "BLANK PIPE", IMG,
     {"description": "Gravel-pack blank pipe below SSD#1; no depth on schematic"}, "no 'blank pipe' type"),
    ("screen-e14", "Sand Screen", '4" W.W.SCREEN', None, None, 4.0, None, "WIRE-WRAP SCREEN", IMG,
     {"description": "Across E14 perforations (1415 - 1424 m); screen depths not stated"}, None),
    ("collet-seal", "Seal Assembly", '4.5" COLLET SEAL ASSY.', None, None, 4.5, None, "COLLET SEAL ASSEMBLY", IMG, {}, None),
    ("sump-packer-1", "Packer", '7" x 4.5" x 3.88" BWD Sump Packer', 1435.2, None, None, 3.88, "SUMP PACKER (BWD)", IMG,
     {"model": "BWD"}, None),
    ("ssd-2", "Sliding Sleeve", 'SLIMLINE SSD (3.22" OD) X-PROFILE', 1451.7, None, 3.22, None, "SLIMLINE SSD (X-PROFILE)", IMG,
     {"description": "SSD#2. Text box: '2.31” PXX plug set at SSD#2 (Sep 2018)'"}, None),
    ("pxx-plug", "Mechanical Plug", "2.31” PXX plug set at SSD#2 (Sep 2018)", 1451.7, None, 2.31, None, "PXX PLUG", TXT,
     {"description": "Set Sep 2018 in SSD#2 (depth of SSD#2)"}, "PXX wireline plug"),
    ("blast-joint", "Blast Joint", "BLAST JOINT (1457.8 - 1485.8 m)", 1457.8, 1485.8, 3.20, None, "BLAST JOINT", IMG,
     {"description": "MAX OD 3.20\""}, None),
    ("sump-packer-2", "Packer", '7" x 4.5" x 3.88" BWD Sump Packer', 1533.3, None, None, 3.88, "SUMP PACKER (BWD)", IMG, {"model": "BWD"}, None),
    ("nogo-nipple", "No-Go Nipple", "NOGO NIPPLE", 1540.9, None, None, None, "NO-GO NIPPLE", IMG, {}, None),
    ("conical", "Mule Shoe", "CONICAL", 1541.7, None, None, None, "CONICAL (end-of-string guide)", IMG, {}, "conical guide -> mule shoe"),
    ("tubing-packoff", "Packer", "Baker tubing packoff (635 – 641m WLTHF)", None, None, None, None, "TUBING PACKOFF", TXT,
     {"manufacturer": "Baker", "description": "Depth 635 - 641 m WLTHF (wireline, from tubing-hanger flange) — not converted: RKB-to-THF offset not stated"},
     "packoff -> packer"),
    ("perf-e14", "Perforation", "E14 (OIL PRODUCER) Perf. Intervals 1415-1424", 1415.0, 1424.0, None, None, "PERF INTERVAL", IMG,
     {"description": "Status: OPEN (text box 'OPEN Sep 2018' over the schematic's 'CLOSE')"}, None),
    ("perf-e22", "Perforation", "E23 Lower (OIL PRODUCER) Perf. Intervals 1485.0 - 1492.5", 1485.0, 1492.5, None, None, "PERF INTERVAL", IMG,
     {"description": "Status: CLOSED (text box 'CLOSED Sep 2018' over the schematic's 'OPEN'). Text box labels this zone 'E22'"}, None),
    ("perf-e36", "Perforation", "E36 (OIL PRODUCER) Perf. Intervals 1557.0 - 1568.0", 1557.0, 1568.0, None, None, "PERF INTERVAL", IMG,
     {"description": "Status: CLOSED (text box 'CLOSED Sep 2018' over the schematic's 'OPEN')"}, None),
    ("pbtd", "PBTD", "PBTD @ 1602 M", 1602.0, None, None, None, "PBTD", IMG, {}, None),
]
FLAGS = [
    ("Well", "Page 1 = DLA-01 from the OCR header 'COMPLETED WELL COMPLETION DIAGRAM : DLA-01' / 'WELL : DLA-01 (WORKOVER…' (rest of the line illegible)."),
    ("Two layers", "Schematic is a scanned image; later updates are typed text boxes. Current status (text boxes) written; original image labels kept in description."),
    ("Zone label", "Text box 'E22' sits on the zone the image labels 'E23 Lower' (1485.0 - 1492.5 m) — both kept; please confirm."),
    ("Casing / tubing table", "Weights and tubing grade illegible in the scan; casing 7\" N-80 Buttress and tubing 2-7/8\" (partly legible), thread AMS-28 written; rest empty."),
    ("Depth reference", "Equipment depths are m MDRKB; packoff (635 - 641 m) and HUD (1430 m) are WLTHF and are not converted."),
    ("Observation", "'2.29” GR & 2” LIB tag HUD at 1430m WLTHF. Impression suspected scale (Nov 2022)' — no template field; recorded here only."),
    ("Footer", "Text box 'Open' at the bottom-left of the page has no clear target — not mapped."),
    ("Not stated", "KB / ground elevation, total depth, max deviation, rig, packer fluid, wellhead, trajectory — left empty."),
]


def main():
    OUT.mkdir(exist_ok=True)
    pdf_b = read(PDF)
    h_pdf, h_tpl = hashlib.sha256(pdf_b).hexdigest(), hashlib.sha256(TPL.read_bytes()).hexdigest()
    tpl = json.loads(subprocess.run(["textutil", "-convert", "txt", "-stdout", str(TPL)], capture_output=True, text=True, check=True).stdout)
    keymap = {r[0]: r[2] for r in load_workbook(XTPL)["Component_Types"].iter_rows(min_row=2, values_only=True)}
    stated = {"completionDate": "2001-10-26", "completionFluid": "8.6 ppg FILTERED SEAWATER WITH 3% KCl",
              "completionType": "SINGLE STRING SELECTIVE OIL PRODUCER", "name": "DLA-01",
              "tubing": 'TUBING 2-7/8" (as read), thread AMS-28'}
    wi = {k: stated.get(k, None if isinstance(v, (int, float)) and not isinstance(v, bool) else "") for k, v in tpl["wellInfo"].items()}
    comps = []
    for cid, ctype, name, top, bot, od, idd, actual, src, extra, est in COMP:
        d = {"id": cid, "type": keymap[ctype], "name": name, "topMd": ft(top), "bottomMd": ft(bot), "outerDiameter": od, "innerDiameter": idd}
        d.update(extra)
        if est:
            d.update(estimated=True, estimatedFields=["type"], estimationNote=f"nearest type '{ctype}' ({est})")
        d.update(source=f"{src}: {name}", actualType=actual, topMdM=top, bottomMdM=bot)
        comps.append(d)
    tops = [dict(name=n, md=ft(m), tvd=None, mdM=m, tvdM=None) for n, m in (("E14", 1415.0), ("E23 Lower (E22)", 1485.0), ("E36", 1557.0))]
    sections = {"wellInfo": wi, "surveys": [], "components": comps, "cement": [], "plugBackCement": [], "centralizers": [],
                "floatEquipment": [], "formations": [], "formationTops": tops, "fluidContacts": []}
    doc = {k: sections[k] for k in tpl}
    out = OUT / "DLA-01_well-schematic.json"
    k = 2
    while out.exists():
        out = OUT / f"DLA-01_well-schematic_v{k}.json"
        k += 1
    out.write_text(json.dumps(doc, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")

    back = json.loads(out.read_text())
    val = [("Valid JSON; sections and order identical to template", list(back) == list(tpl)),
           ("wellInfo keys identical to template, in order", list(back["wellInfo"]) == list(tpl["wellInfo"])),
           ("Template sample records removed", back["surveys"] == [] and not any(c["id"] in {t["id"] for t in tpl["components"]} for c in back["components"])),
           ("Every component type is a Component_Types key", all(c["type"] in set(keymap.values()) for c in back["components"])),
           ("Every ft depth = m × 3.28084", all((c["topMd"] is None) == (c["topMdM"] is None) and (c["topMd"] is None or abs(c["topMd"] - round(c["topMdM"] * FT, 3)) < 1e-9) for c in back["components"])),
           ("Source PDF unchanged", hashlib.sha256(read(PDF)).hexdigest() == h_pdf),
           ("JSON template unchanged", hashlib.sha256(TPL.read_bytes()).hexdigest() == h_tpl)]
    # review workbook
    wb = Workbook()
    s = wb.active
    s.title = "Components"
    hdr = ["id", "type", "name", "topMdM", "bottomMdM", "outerDiameter", "innerDiameter", "actualType", "source", "description", "estimated"]
    s.append(hdr)
    for c in back["components"]:
        s.append([c.get(h) if not isinstance(c.get(h), list) else ", ".join(c.get(h)) for h in hdr])
    for title, rows, h in (("Flags", FLAGS, ["Item", "Detail"]), ("Validation", [(a, "PASS" if b else "FAIL") for a, b in val], ["Check", "Result"])):
        x = wb.create_sheet(title)
        x.append(h)
        for r in rows:
            x.append(list(r))
    for x in wb.worksheets:
        for c in x[1]:
            c.font, c.fill = Font(bold=True), PatternFill("solid", fgColor="DDEBF7")
        for col in x.columns:
            x.column_dimensions[col[0].column_letter].width = min(90, max(10, max(len(str(v.value or "")) for v in col) + 2))
    rev = OUT / out.name.replace("_well-schematic", "_mapping_review").replace(".json", ".xlsx")
    wb.save(rev)
    for a, b in val:
        print("PASS" if b else "FAIL", "-", a)
    print(out.name, rev.name, "| components", len(comps), "| formationTops", len(tops))


if __name__ == "__main__":
    main()
