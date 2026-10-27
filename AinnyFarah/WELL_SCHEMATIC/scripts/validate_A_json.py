#!/usr/bin/env python3
"""Validate every output_A/Dulang_A*.json against well-schematic-template-json and write the QC report workbook."""
import glob, hashlib, json, os, re, subprocess
from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
ROOT = Path(__file__).resolve().parent.parent; OUT = ROOT / "output_A"; PASS = OUT / "ocr_passes"
TPL = ROOT / "well-schematic-template-json.rtf"
PDF = "/Volumes/homes/public/DulangOSS/20261016/DULANG_WELLS/Well_Schematic/1. Well Schematic/Dulang A/Dulang A Well Schematic updated rev1.pdf"
tpl = json.loads(subprocess.run(["textutil", "-convert", "txt", "-stdout", str(TPL)], capture_output=True, text=True, check=True).stdout)
types = {r[2] for r in load_workbook(ROOT / "well-schematic-sample-template.xlsx")["Component_Types"].iter_rows(min_row=2, values_only=True)}
comp_keys = set().union(*(c.keys() for c in tpl["components"]))      # keys used by the template's own components
def kind(v): return "number" if isinstance(v, (int, float)) and not isinstance(v, bool) else type(v).__name__
fd = os.open(PDF, os.O_RDONLY); b = b""
while (x := os.read(fd, 1 << 20)): b += x
try: os.close(fd)
except OSError: pass
src_ok = hashlib.sha256(b).hexdigest() == (PASS / "source_sha256.txt").read_text().strip()
rows, comps = [], []
files = sorted(glob.glob(str(OUT / "Dulang_A*.json")), key=lambda f: (len(f), f))
for f in files:
    d = json.load(open(f)); errs = []
    if list(d) != list(tpl): errs.append("top-level sections/order differ")
    if list(d["wellInfo"]) != list(tpl["wellInfo"]): errs.append("wellInfo keys differ")
    for k, v in d["wellInfo"].items():
        if v not in ("", None) and kind(v) != kind(tpl["wellInfo"][k]): errs.append(f"wellInfo.{k} type {kind(v)} != {kind(tpl['wellInfo'][k])}")
    if d["wellInfo"]["name"] != Path(f).stem: errs.append("name != filename")
    if not re.fullmatch(r"Dulang_A[1-9]\d*(ST\d?)?", Path(f).stem): errs.append("filename format")
    for sec in tpl:
        if sec != "wellInfo" and not isinstance(d[sec], list): errs.append(f"{sec} not a list")
    for c in d["components"]:
        extra = set(c) - comp_keys
        if extra: errs.append(f"component {c['id']} keys not in template: {sorted(extra)}")
        if c["type"] not in types: errs.append(f"component {c['id']} type '{c['type']}' not a template type")
        for k in ("topMd", "bottomMd", "outerDiameter", "innerDiameter"):
            if c.get(k) is not None and kind(c[k]) != "number": errs.append(f"{c['id']}.{k} not numeric")
        comps.append([Path(f).stem, c["id"], c["type"], c["name"], c["topMd"], c.get("bottomMd"), c["outerDiameter"], c["innerDiameter"], c.get("estimated", False), c["source"]])
    rows.append([Path(f).stem, "PASS" if not errs else "FAIL", "; ".join(errs), len(d["components"]),
                 ", ".join(f"{k}={v}" for k, v in d["wellInfo"].items() if v not in ("", None) and k != "name")])
rep = json.loads((PASS / "extraction_report.json").read_text())
wb = Workbook(); s = wb.active; s.title = "Pages"
s.append(["PDF page", "Standardized well", "Status", "Detail", "Rejected readings (OCR misreads, not written)"])
for r in rep: s.append([r["page"], r["well"], r["status"], r["detail"], json.dumps(r.get("rejected", {}), ensure_ascii=False) if r.get("rejected") else ""])
v = wb.create_sheet("Schema Validation"); v.append(["JSON file", "Result", "Errors", "Components", "wellInfo values written"]); [v.append(r) for r in rows]
c = wb.create_sheet("Components"); c.append(["Well", "id", "type", "name (as read)", "topMd (as printed)", "bottomMd", "OD", "ID", "type estimated", "source"]); [c.append(r) for r in comps]
n = wb.create_sheet("Notes")
for line in [("Source PDF unchanged (sha256)", "PASS" if src_ok else "FAIL"),
             ("Method", "Local Tesseract only (no AI vision). 3 page passes at 600 dpi + 400 dpi pass + title-line variants; typed text boxes read exactly from the PDF text layer."),
             ("Acceptance", "A value is written only if read identically in >= 2 sources (>= 3 of 4 passes for free text), and passes validity checks (known completion terms, day 1-31, deviation <= 90). Rejected readings are listed, never corrected."),
             ("Depth units", "As printed on each page (m MDDF / m MDRKB / m MDTHF); no conversion."),
             ("source_well_name", "Not a field of the template, so not added; the exact PDF label is in the 'Pages' sheet."),
             ("Limitation", "The schematics are low-resolution raster scans; identical OCR misreads can repeat across passes. Components and values not meeting the rules are left empty — the JSON files are incomplete by design and need review against the PDF.")]:
    n.append(list(line))
for ws in wb.worksheets:
    for cell in ws[1]: cell.font, cell.fill = Font(bold=True), PatternFill("solid", fgColor="DDEBF7")
    for col in ws.columns: ws.column_dimensions[col[0].column_letter].width = min(80, max(12, max(len(str(x.value or "")) for x in col) + 2))
    ws.freeze_panes = "A2"
out = PASS / "Dulang_A_extraction_QC_report.xlsx"; wb.save(out)
print("files:", len(files), "| schema PASS:", sum(r[1] == "PASS" for r in rows), "| FAIL:", [r[0] for r in rows if r[1] != "PASS"], "| source unchanged:", src_ok)
print("report:", out)
