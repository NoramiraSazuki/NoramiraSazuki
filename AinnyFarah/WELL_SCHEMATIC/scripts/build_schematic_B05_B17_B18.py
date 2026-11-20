#!/usr/bin/env python3
"""Map the Dulang B05 / B17 / B18 ESP workover schematics (NAS, read-only) into the JSON template (29 Oct 2026).

Output folder WELL_SCHEMATIC/output_B/: Dulang_<well>.json, Dulang_<well>_mapping_report.md, Dulang_<well>_traceability.csv;
_work/ holds working copies (src/), cell + drawing dumps (dump/), Quick Look renders (render/ql/), extracted media and OCR.
User decisions: names Dulang_B5 / Dulang_B17 / Dulang_B18; deliverables as for the A17ST2/A23ST2/A33ST3 extraction.
Rules applied:
  * depth_ft = depth_m x 3.28084 (3 dp); original metres kept in topMdM / bottomMdM; MD reference kept as printed (m-MDTHF,
    tubing hanger flange) — no datum shift, no MD->TVD
  * hidden rows 4-14 (header + casing table) are identical in all three workbooks and describe another well
    ('Erb West / EWDP-A', 'EW-119 / Slot #9', 'XX XX 2017', depths deeper than TD) -> not mapped, flagged
  * Type_Keys from well-schematic-sample-template.xlsx > Component_Types; nearest-type mappings flagged (estimatedFields
    'type', actualType = source name); items with no suitable type are listed as unresolved, not forced
  * user-defined size rule (from the user's earlier corrections): a size in a component label with no ID/OD designation
    and no MIN ID / MAX OD column value is recorded as innerDiameter
"""
import csv
import datetime as dt
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent.parent
NAS = Path("/Volumes/homes/public/DulangOSS/20261016/DULANG_WELLS/Well_Schematic/1. Well Schematic/Dulang B")
TPL = ROOT / "well-schematic-template-json.rtf"
XTPL = ROOT / "well-schematic-sample-template.xlsx"
OUT = ROOT / "output_B"
WORK = OUT / "_work"
FT_PER_M = 3.28084
SOURCES = {"B5": "Dulang B05 ESP Workover Schematic.xlsx", "B17": "Dulang B17 ESP Workover Schematic.xlsx",
           "B18": "Dulang B18 ESP Workover Schematic Updated .xlsx"}
SHEETS = {"B5": "B-05 v1", "B17": "New B-17 (12.12.23)", "B18": "B18 WO (Jan25)"}
LABELS = {"B5": "COMPLETION SCHEMATIC : DULANG B-05", "B17": "COMPLETION SCHEMATIC : DULANG B-17",
          "B18": "COMPLETION SCHEMATIC : DULANG B-18"}
DATUM = "m-MDTHF (MD from tubing hanger flange; column headers J16/Q16)"


def ft(m):
    return None if m is None else round(m * FT_PER_M, 3)


def sha(p):
    h = hashlib.sha256()
    fd = os.open(p, os.O_RDONLY)
    while (b := os.read(fd, 1 << 20)):
        h.update(b)
    try:
        os.close(fd)
    except OSError:  # SMB (NAS) handles can report EBADF on close after a full read
        pass
    return h.hexdigest()


class Book:
    def __init__(self, workbook, sheet, cells, drawings):
        self.workbook, self.sheet = workbook, sheet
        self.c = {c["ref"]: c for c in cells}
        self.draw = {o["name"]: o for objs in drawings.values() for o in objs if o.get("text")}

    def v(self, ref):
        assert ref in self.c, f"{self.workbook}: {ref} is empty"
        return self.c[ref]["value"]

    def txt(self, ref, line=None):
        v = self.v(ref)
        assert isinstance(v, str), (ref, v)
        lines = [x.strip() for x in v.split("\n") if x.strip()]
        return " ".join(lines) if line is None else lines[line]

    def num(self, ref):
        v = self.v(ref)
        assert isinstance(v, (int, float)) and not isinstance(v, bool), (ref, v)
        return float(v)

    def formula(self, ref):
        return self.c.get(ref, {}).get("formula")


# ------------------------------------------------------------------ specs
def C(cid, typ, name, top=None, bot=None, od=None, idd=None, ln=None, wt=None, gr=None, string=None, closest=None,
      note=None, zone=None, size_id=None):
    """name: list of (cell, line-selector|None) or ('draw', shape name) or ('ocr', text, where).
    top/bot: cell ref (number in m), ('txt', m, cell, text-evidence) or ('draw', m, shape).
    od/idd: cell refs (MIN ID / MAX OD columns, inches). size_id: (inches, cell, text) for the user-defined size rule."""
    return dict(cid=cid, typ=typ, name=name, top=top, bot=bot, od=od, idd=idd, ln=ln, wt=wt, gr=gr, string=string,
                closest=closest, note=note, zone=zone, size_id=size_id)


GLM = "'KBMG GLM' = gas-lift mandrel (side-pocket) -> SPM (Gas Lift) is the nearest type"
STUB = "tubing stub = top of the tubing left in hole -> Tubing (top = stub depth; bottom not printed)"
WEG = "wireline entry guide (WEG) at the tubing end -> Mule Shoe is the nearest type"
WP1 = "'WP-1 NIPPLE' (nipple profile, no no-go stated) -> Landing Nipple is the nearest type"
SANDGUARD = "ESP intake sand guard (filter) -> Sand Screen is the nearest type"
TRMAXX = "TRMAXX-5E is a tubing-retrievable safety valve model (label does not say 'TRSV') -> TRSV"
SBE = "sealbore extension -> PBR is the nearest type"
TELESCOPIC = "telescopic joint -> Expansion Joint is the nearest type"
CONICAL = "conical (guide) shoe at the tubing end -> Mule Shoe is the nearest type"
BLANK = "blank (unperforated) pipe in the screen assembly -> Tubing is the nearest type"
ENDESP = "'End of ESP Assy' = deepest point of the completion string -> EOT is the nearest type"
LEFT, RIGHT = "left equipment column (B-J)", "right equipment column (Q-Y)"


def R(cid, typ, row, extra_name_rows=(), closest=None, depth=True, ln=True, idd=True, od=True, **kw):
    """right-hand equipment row: name T, depth Q, length S, MIN ID X, MAX OD Y"""
    return C(cid, typ, [(f"T{row}", None)] + [(f"T{r}", None) for r in extra_name_rows], top=f"Q{row}" if depth else None,
             ln=f"S{row}" if ln else None, idd=f"X{row}" if idd else None, od=f"Y{row}" if od else None, closest=closest,
             note=kw.pop("note", RIGHT), **kw)


def L(cid, typ, row, extra_name_rows=(), closest=None, depth=True, idd=False, line=None, **kw):
    """left-hand equipment row: name D, depth J, MIN ID B"""
    return C(cid, typ, [(f"D{row}", line)] + [(f"D{r}", None) for r in extra_name_rows], top=f"J{row}" if depth else None,
             idd=f"B{row}" if idd else None, closest=closest, note=kw.pop("note", LEFT), **kw)


def perf(cid, zone, cell, top, bot, text):
    return C(cid, "Perforation", [("perf", f"{zone} perforation {top} m - {bot} m" if zone else f"Perforation {top} m - {bot} m", cell)],
             top=("txt", top, cell, text), bot=("txt", bot, cell, text), zone=zone)


def spec_b5():
    return [
        R("flow-coupling-1", "Flow Coupling", 20), R("trsv", "TRSV", 22), R("flow-coupling-2", "Flow Coupling", 24),
        R("glm-1", "SPM (Gas Lift)", 26, closest=GLM), R("glm-2", "SPM (Gas Lift)", 29, closest=GLM),
        R("glm-3", "SPM (Gas Lift)", 32, closest=GLM), R("ssd-1", "Sliding Sleeve", 34), R("esp-packer", "Packer", 36),
        R("ssd-2", "Sliding Sleeve", 39), R("crossover-1", "Crossover", 43, (44,)),
        R("sand-guard", "Sand Screen", 46, closest=SANDGUARD, idd=False), R("crossover-2", "Crossover", 48, (49,)),
        R("esp", "ESP", 52, idd=False),
        L("tubing-stub-left", "Tubing", 56, closest=STUB), R("tubing-stub-right", "Tubing", 56, closest=STUB, ln=False, idd=False, od=False),
        R("guib-packer", "Packer", 58, ln=False, idd=False, od=False),
        C("bridge-plug-1", "Bridge Plug", [("draw", "Rectangle 605")], top=("draw", 2037.0, "Rectangle 605"),
          note="drawing text box anchored D65:G67 (left column, between the E-7 perforations and the dual packer)"),
        L("dual-hyd-packer", "Packer", 68, (69,)),
        C("bridge-plug-2", "Bridge Plug", [("draw", "Rectangle 606")], top=("draw", 2050.0, "Rectangle 606"),
          note="drawing text box anchored S74:W76 (right column, above the WPB-2DE' SSD)"),
        L("no-go-nipple-left", "No-Go Nipple", 72), L("wl-entry-guide-left", "Mule Shoe", 73, closest=WEG),
        R("ssd-3", "Sliding Sleeve", 77, ln=False, od=False), R("blast-joint", "Blast Joint", 79, ln=False, od=False),
        R("wp1-nipple", "Landing Nipple", 82, closest=WP1, ln=False, od=False),
        R("locator-seal-assembly", "Seal Assembly", 86, ln=False, od=False),
        L("retainer-packer", "Packer", 87, (88,)),
        R("no-go-nipple-right", "No-Go Nipple", 90, ln=False, od=False),
        R("wl-entry-guide-right", "Mule Shoe", 92, closest=WEG, ln=False, od=False),
        L("perm-packer", "Packer", 91, (92,)),
        perf("perf-e7", "E-7 (FB S3, WIR)", "C61", 1974.74, 1994.24, "PERFT : 1974.74 m - 1994.24 m"),
        perf("perf-e12-13-1", "E-12/13 (FB S3, WIR)", "C77", 2067.24, 2075.74, "E-12/13 PERF : 2067.24 m - 2075.74 m"),
        perf("perf-e12-13-2", "E-12/13 (FB S3, WIR)", "C77", 2077.74, 2083.74, ": 2077.74 m - 2083.74 m"),
        perf("perf-e14-1", "E-14 (FB S3, WIR)", "C77", 2087.74, 2103.74, "E-14 PERF       : 2087.74 m - 2103.74 m"),
        perf("perf-e14-2", "E-14 (FB S3, WIR)", "C77", 2106.24, 2135.74, "2106.24 m - 2135.74 m"),
        perf("perf-c89", None, "C89", 2180.24, 2186.24, "PERF : 2180.24 m - 2186.24 m"),
        *[perf(f"perf-e20-32-{i}", "E-20/22 & E-23/32 (FB S3, NDR)", "C94", a, b, t) for i, (a, b, t) in enumerate([
            (2215.24, 2217.24, "PERF : 2215.24 m - 2217.24 m"), (2223.74, 2228.12, "2223.74 m - 2228.12 m"),
            (2243.24, 2249.24, "2243.24 m - 2249.24 m"), (2258.24, 2275.24, "2258.24 m - 2275.24 m"),
            (2276.24, 2279.74, "2276.24 m - 2279.74 m")], 1)],
    ]


def spec_b17():
    return [
        R("flow-coupling-1", "Flow Coupling", 20, depth=False), R("trsv", "TRSV", 22), R("flow-coupling-2", "Flow Coupling", 24, depth=False),
        R("glm-1", "SPM (Gas Lift)", 26, closest=GLM), R("glm-2", "SPM (Gas Lift)", 28, closest=GLM),
        R("ssd-1", "Sliding Sleeve", 30), R("esp-packer", "Packer", 33), R("ssd-2", "Sliding Sleeve", 37),
        R("crossover-1", "Crossover", 41, (42,), depth=False), R("sand-guard", "Sand Screen", 44, closest=SANDGUARD, idd=False),
        R("crossover-2", "Crossover", 46, (47,), depth=False), R("esp", "ESP", 50, idd=False),
        R("eot", "EOT", 51, ln=False, idd=False, od=False),
        L("ss-stub", "Tubing", 54, closest=STUB, string="Short", note=LEFT + "; 'SS' read as short string"),
        R("ls-stub", "Tubing", 54, closest=STUB, string="Long", ln=False, idd=False, od=False, note=RIGHT + "; 'LS' read as long string"),
        L("dual-hyd-packer", "Packer", 55, string="Short & Long", note=LEFT + "; one 'DUAL' packer drawn across both strings"),
        L("xn-no-go-nipple-left", "No-Go Nipple", 58),
        C("liner-7", "Liner", [("T61", None), ("T93", None)], top="Q61", bot="Q93", size_id=(7.0, "T61", '7"'),
          note="top = diagram label T61 '7\" TOP OF LINER'; bottom = diagram label T93 '7\" LINER SHOE'; name = both labels joined"),
        C("casing-9-5-8", "Casing", [("T64", None)], bot="Q64", size_id=(9.625, "T64", '9-5/8"'),
          note="bottom = diagram label T64 '9-5/8\" Casing Shoe'; casing top not printed"),
        R("ssd-3", "Sliding Sleeve", 65, ln=False, od=False),
        L("bwd-packer-1", "Packer", 67, (68,), idd=True),
        R("blank-pipe-1", "Tubing", 68, closest=BLANK, depth=False, ln=False, od=False),
        R("screen-1", "Screen", 71, depth=False, ln=False, idd=False, od=False, size_id=(4.0, "T71", '4"')),
        R("seal-bore-assembly-1", "Seal Assembly", 74, depth=False, ln=False, idd=False, od=False, size_id=(3.25, "T74", '3.25"')),
        L("bwd-packer-2", "Packer", 76, (77,), idd=True),
        R("ssd-4", "Sliding Sleeve", 80, ln=False, od=False),
        R("blank-pipe-2", "Tubing", 82, closest=BLANK, depth=False, ln=False, idd=False, od=False),
        R("screen-2", "Screen", 83, depth=False, ln=False, idd=False, od=False, size_id=(4.0, "T83", '4"')),
        R("seal-bore-assembly-2", "Seal Assembly", 85, depth=False, ln=False, idd=False, od=False, size_id=(3.25, "T85", '3.25"')),
        L("bwd-packer-3", "Packer", 86, (87,), idd=True),
        R("xn-no-go-nipple-right", "No-Go Nipple", 89, ln=False, od=False),
        R("conical-shoe", "Mule Shoe", 90, closest=CONICAL, ln=False, idd=False, od=False),
        C("pbtd", "PBTD", [("L95", None)], top=("txt", 2171.7, "L95", "PBTD @ 2171.7 m MDTHF"), note="diagram label L95"),
        perf("perf-e12-13", "E12/13 (FB S2; WIR)", "C71", 1734.7, 1763.7, "PERFT: 1734.7 m  - 1763.7 m MDTHF"),
        perf("perf-e14-1", "E14 (FB S2; WIR)", "C80", 1796.7, 1816.7, "PERFT: 1796.7 m  - 1816.7 m MDTHF"),
        perf("perf-e14-2", "E14 (FB S2; WIR)", "C80", 1824.7, 1832.7, "1824.7 m  - 1832.7 m MDTHF"),
        perf("perf-e34-36-1", "E34/36 (FB S2; WIR)", "C90", 2012.7, 2042.7, "PERFT: 2012.7 m  - 2042.7 m MDTHF"),
        perf("perf-e34-36-2", "E34/36 (FB S2; WIR)", "C90", 2052.7, 2067.7, "2052.7 m  - 2067.7 m MDTHF"),
        perf("perf-e40-42", "E40/42", "C90", 2102.7, 2120.7, "PERFT : 2102.7 m  - 2120.7 m MDTHF"),
    ]


def spec_b18():
    return [
        C("christmas-tree", "Xmas Tree", [("ocr", "ASSY, XMAS TREE, SINGLE, 4-1/16\" 5K, MODEL 'SAE' SLAB GV", "Picture 1 (xl/media/image4.png) equipment table item 1")],
          note="OCR of the embedded wellhead drawing (Picture 1, AD17:BL48), visually verified; no depth"),
        R("flow-coupling-1", "Flow Coupling", 20, ln=False), R("trsv", "TRSV", 22, closest=TRMAXX, ln=False),
        R("flow-coupling-2", "Flow Coupling", 24, ln=False),
        R("spm-1", "SPM (Gas Lift)", 27, ln=False), R("spm-2", "SPM (Gas Lift)", 29, ln=False), R("spm-3", "SPM (Gas Lift)", 32, ln=False),
        R("sliding-sleeve-1", "Sliding Sleeve", 34, (35,), ln=False), R("esp-packer", "Packer", 38, ln=False),
        R("sliding-sleeve-2", "Sliding Sleeve", 41, (42,), ln=False),
        R("esp-pump", "ESP", 49, ln=False, idd=False, od=False),
        R("end-of-esp-assy", "EOT", 51, closest=ENDESP, ln=False, idd=False, od=False),
        L("ls-tubing-stub", "Tubing", 51, closest=STUB, string="Long", wt=("6.5", "D51"), size_id=(2.875, "D51", '2-7/8"')),
        L("ss-tubing-stub", "Tubing", 52, closest=STUB, string="Short", wt=("6.5", "D52"), size_id=(2.875, "D52", '2-7/8"')),
        L("guib-packer-1", "Packer", 54),
        L("telescopic-joint", "Expansion Joint", 56, closest=TELESCOPIC),
        L("ssd-left", "Sliding Sleeve", 57, idd=True), R("ssd-right", "Sliding Sleeve", 57, ln=False, od=False),
        R("blast-joint-1", "Blast Joint", 59, ln=False, od=False),
        L("wp1-nipple-left", "Landing Nipple", 61, closest=WP1, idd=True), R("wp1-nipple-right", "Landing Nipple", 61, closest=WP1, ln=False, od=False),
        L("guib-packer-2", "Packer", 63, idd=True, note=LEFT + "; D64 repeats 'PERM. PKR' (stray text, not a separate item)"),
        L("no-go-nipple-left", "No-Go Nipple", 66, idd=True, line=0, note=LEFT + "; second line of D66 'PERM. PKR' not part of the name (stray)"),
        L("wl-entry-guide-left", "Mule Shoe", 67, closest=WEG, idd=True),
        R("ssd-right-2", "Sliding Sleeve", 74, ln=False, od=False),
        R("blast-joint-2", "Blast Joint", 76, depth=False, ln=False, od=False),
        R("wp1-nipple-right-2", "Landing Nipple", 78, closest=WP1, ln=False, od=False),
        R("locator-seal-assembly", "Seal Assembly", 84, ln=False, od=False),
        L("guib-packer-3", "Packer", 84, idd=True, note=LEFT + "; D85 repeats 'PERM. PKR' (stray text)"),
        R("seal-bore-extension", "PBR", 85, closest=SBE, ln=False, od=False),
        R("no-go-nipple-right", "No-Go Nipple", 88, ln=False, od=False),
        R("weg-right", "Mule Shoe", 89, closest=WEG, ln=False, od=False),
        C("casing-9-5-8", "Casing", [("D92", 0), ("D93", None)], wt=("40", "D93"), gr=("N80", "D93"), size_id=(9.625, "D92", '9 5/8"'),
          note="left column label D92 '9 5/8\" CASING SHOE' (no depth printed in J92) + D93 '9-5/8\" 40PPF, N80'; second line of D92 'PERM. PKR' is stray"),
        perf("perf-e2", "E-2 (FB S2; NDR)", "C58", 1639.74, 1643.74, "PERFT: 1639.74 m- 1643.74 m"),
        perf("perf-e10-11-1", "E-10/11 (FB S2; WIR)", "C76", 1701.24, 1715.24, "PERFT : 1701.24 m - 1715.24 m"),
        perf("perf-e10-11-2", "E-10/11 (FB S2; WIR)", "C76", 1717.24, 1722.24, "1717.24 m- 1722.24 m"),
        perf("perf-e10-11-3", "E-10/11 (FB S2; WIR)", "C76", 1725.24, 1735.24, "1725,24 m - 1735.24 m"),
        perf("perf-e12-13", "E-12/13 (FB S2; WIR)", "C88", 1744.4, 1760.24, "PERFT : 1744,4 m - 1760,24 m"),
    ]


# ------------------------------------------------------------------ build
class Trace:
    def __init__(self, well, B):
        self.well, self.B, self.rows, self.conv = well, B, [], []

    def add(self, path, value, ref, method, note="", unit="", original=None):
        B = self.B
        f = B.formula(ref) if ref and ref in B.c else None
        orig = original if original is not None else (repr(B.v(ref)) if ref and ref in B.c else "")
        self.rows.append(dict(well=self.well, json_path=path, mapped_value=value, workbook=B.workbook, worksheet=B.sheet,
                              cell_or_object=ref or "", original_value=orig, formula=f or "", cached_value=repr(B.v(ref)) if f else "",
                              original_unit=unit, method=method, note=note))


def build_components(B, specs, keymap, tr):
    comps, review = [], []
    for s in specs:
        nm = s["name"]
        if nm[0][0] == "draw":
            name, nref = B.draw[nm[0][1]]["text"][0], nm[0][1]
            assert re.fullmatch(r"Permament Bridge plug @ \d+m-MDTHF", name)
            nmethod = "DIRECT (drawing text box)"
        elif nm[0][0] == "ocr":
            name, nref, nmethod = nm[0][1], nm[0][2], "OCR (tesseract psm 6 on x4 upscale) + visual check"
        elif nm[0][0] == "perf":
            name, nref, nmethod = nm[0][1], nm[0][2], "DERIVED name (zone label + interval from the perforation text)"
        else:
            name = " ".join(B.txt(r, ln) for r, ln in nm)
            nref, nmethod = "+".join(r for r, _ in nm), "DIRECT (cell text" + (", lines joined" if len(nm) > 1 else "") + ")"
        d = {"id": s["cid"], "type": keymap[s["typ"]], "name": name, "topMd": None, "bottomMd": None,
             "outerDiameter": None, "innerDiameter": None}
        p = f"components[{s['cid']}]"
        tr.add(f"{p}.name", name, nref if nref in B.c else None, nmethod, original=None if nref in B.c else repr(nref))
        tr.add(f"{p}.type", d["type"], None, "CLOSEST TYPE" if s["closest"] else "EXACT TYPE",
               f"Component_Type '{s['typ']}'" + (f" — {s['closest']}" if s["closest"] else ""), original="")
        est, enote, src = [], [], [f"name {nref}"]
        if s["closest"]:
            est.append("type"); enote.append(f"type: {s['closest']}")
        for key, field in (("top", "topMd"), ("bot", "bottomMd")):
            x = s[key]
            if x is None:
                continue
            if isinstance(x, tuple) and x[0] == "txt":
                m, ref, ev = x[1], x[2], x[3]
                txt = B.v(ref)
                norm = txt.replace(",", ".") if "," in ev else txt
                assert ev in txt and (f"{m:g}" in norm or f"{m}" in norm), (ref, ev, m)
                unit = "m (in text" + ("; MDTHF stated" if "MDTHF" in ev else "; datum not stated in the label, sheet depths are m-MDTHF") + ")"
                how = f"value parsed from text '{ev}'" + ("; comma read as decimal point" if "," in ev else "")
                orig = repr(ev)
            elif isinstance(x, tuple) and x[0] == "draw":
                m, ref = x[1], x[2]
                ev = B.draw[ref]["text"][0]
                assert f"@ {m:g}m-MDTHF" in ev
                unit, how, orig = "m-MDTHF (in drawing text)", f"value parsed from drawing text '{ev}'", repr(ev)
            else:
                ref = x
                m = B.num(ref)
                hdr = "J16" if ref[0] == "J" else "Q16"
                assert B.txt(hdr) == "m-MDTHF", hdr
                unit, how, orig = f"m-MDTHF (column header {hdr})", "DIRECT", None
            d[field] = ft(m)
            d[field + "M"] = m
            tr.add(f"{p}.{field}", d[field], ref if ref in B.c else None, "CONVERTED",
                   f"{how}; {m} m x 3.28084 = {d[field]} ft (3 dp); datum unchanged", unit, original=orig)
            tr.add(f"{p}.{field}M", m, ref if ref in B.c else None, "DIRECT (original metres)", how, unit, original=orig)
            tr.conv.append(dict(field=f"{s['cid']}.{field}", ref=ref, m=m, unit=unit, ft=d[field]))
            src.append(f"{field} {ref}={m} m")
        for key, field, hdrs in (("od", "outerDiameter", {"C": "C15", "Y": "Y15"}), ("idd", "innerDiameter", {"B": "B15", "X": "X15"})):
            ref = s[key]
            if not ref or ref not in B.c:
                continue
            col = re.match(r"[A-Z]+", ref).group(0)
            label = B.txt(hdrs[col])
            assert label in ("MIN ID", "MAX OD") and B.txt(hdrs[col][0] + "16") == "IN"
            d[field] = B.num(ref)
            tr.add(f"{p}.{field}", d[field], ref, "DIRECT", f"column '{label}' ({hdrs[col]}), unit 'IN' ({hdrs[col][0]}16)", "in")
            src.append(f"{field} {ref}")
        if s["size_id"]:
            size, ref, t = s["size_id"]
            assert t in B.v(ref) and d["innerDiameter"] is None and d["outerDiameter"] is None, (s["cid"], ref)
            d["innerDiameter"] = size
            est.append("innerDiameter")
            enote.append(f"innerDiameter: size {t} in label {ref} has no ID/OD designation -> recorded as ID (user-defined mapping rule)")
            tr.add(f"{p}.innerDiameter", size, ref, "USER-DEFINED RULE",
                   f"size {t} in the label, no ID/OD designation and no MIN ID/MAX OD value -> innerDiameter; fraction written as decimal inches", "in")
            src.append(f"ID {ref} '{t}' (user-defined rule)")
        if s["ln"] and s["ln"] in B.c:
            assert B.txt("S16") == "m"
            d["interval"] = f"{B.num(s['ln'])!r} m"
            tr.add(f"{p}.interval", d["interval"], s["ln"], "DIRECT", "equipment LENGTH column S15/S16 'm' -> interval (tool length, not a depth); unit kept", "m")
            src.append(f"length {s['ln']}")
        if s["wt"]:
            val, ref = s["wt"]
            assert re.search(rf"{val}\s*ppf", B.v(ref), re.I)
            d["weight"] = float(val)
            tr.add(f"{p}.weight", d["weight"], ref, "DIRECT (parsed from label)", f"'{val}ppf' in the label", "lb/ft (ppf)")
        if s["gr"]:
            val, ref = s["gr"]
            assert val in B.v(ref)
            d["grade"] = val
            tr.add(f"{p}.grade", val, ref, "DIRECT (parsed from label)")
        if s["zone"]:
            d["zone"] = s["zone"]
            tr.add(f"{p}.zone", s["zone"], nref, "DIRECT", "sand label of the perforation text box ('SAND : …')")
        if s["string"]:
            d["string"] = s["string"]
            tr.add(f"{p}.string", s["string"], nref.split("+")[0] if nref in B.c or "+" in nref else None, "INTERPRETED",
                   "from the label ('Long String' / 'Short String' / 'LS' / 'SS' / 'DUAL')", original="")
        if est:
            d.update(estimated=True, estimatedFields=est, estimationNote="; ".join(enote))
            if "type" in est:
                d["actualType"] = name
        d["source"] = f"{B.workbook} > sheet '{B.sheet}' > " + "; ".join(src) + f"; depths {DATUM}" + (f"; {s['note']}" if s["note"] else "")
        tr.add(f"{p}.source", "(traceability text)", None, "GENERATED", original="")
        if d["topMd"] is None and d["bottomMd"] is None and s["typ"] != "Xmas Tree":
            review.append(("Missing", nref, f"{s['cid']}: no depth printed for '{name[:70]}' -> topMd/bottomMd empty"))
        comps.append(d)
    return comps, review


def well_info(tpl, entries, tr):
    wi = {}
    for k, tv in tpl["wellInfo"].items():
        if k in entries:
            val, ref, method, note = entries[k]
            wi[k] = val
            tr.add(f"wellInfo.{k}", val, ref if ref in tr.B.c else None, method, note, original=None if ref in tr.B.c else repr(ref or ""))
        else:
            wi[k] = "-" if isinstance(tv, str) else None
            tr.add(f"wellInfo.{k}", wi[k], None, "EMPTY", "not supported by the visible source -> template empty value ('-' text / null number)", original="")
    return wi


# ------------------------------------------------------------------ notes per well
HIDDEN = ("Conflict", "rows 4-14 (hidden)",
          "Hidden header rows 4-14 (B4-B14, casing & tubing table N4:Y11, fluids N13/N14) are identical in B05, B17 and B18 and name "
          "another well: 'Field / Platform : Erb West / EWDP-A Platform', 'Well ID / SLOT NO : EW-119 / Slot #9', spud/completion "
          "'XX XX 2017', rig 'XXX'. Casing depths there (e.g. 7\" liner 7150-9959 ft-MDTHF, max deviation @ 10092 ft) exceed the "
          "well's TD. Treated as a copied template block: NOT mapped (wellInfo fields left empty, no casing components from it).")
DATUM_NOTE = ("Unclear", "J16 / Q16",
              "All depths are m-MDTHF (from the tubing hanger flange), not drill floor / RKB. Converted to ft only; the reference was not "
              "shifted (the DF-to-THF offset '46 ft' exists only in the hidden copied block). Consumers expecting MD-RKB must add the offset.")
DEPTH_POINT = ("Unclear", "column Q / J 'DEPTH'", "The sheet does not say whether equipment depths are to the top of the item; mapped to topMd. "
               "Item lengths (column S, m) are kept in 'interval' and not used to derive bottoms.")


def notes(well, B):
    common_unmapped = [
        ("rows 4-14 (hidden)", "Erb West / EW-119 header, casing & tubing table, completion/packer fluid, max deviation", "copied block from another well — see Conflict"),
        ("K84" if well != "B5" else "K88", f"{B.v('K84' if well != 'B5' else 'K88')} (no label, no unit)", "meaning unknown"),
        ("T94-T95" if well != "B5" else "T98-T99", f"{B.txt('T95' if well != 'B5' else 'T99')}", "document control (updated by / date)"),
    ]
    if well == "B5":
        review = [HIDDEN, DATUM_NOTE, DEPTH_POINT,
                  ("Unclear", "D56/J56 vs T56/Q56", "Two 'TUBING STUB' labels (left 1952.7 m, right 1951.7 m) — read as two strings of the old dual "
                   "completion; the sheet does not label which is short/long, so 'string' is not set."),
                  ("Unclear", "Rectangle 605 / 606", "Bridge plugs placed from drawing text boxes; string association by position only (605 left column, 606 right column)."),
                  ("Unclear", "C61", "'PERFT' (spelling as printed) read as perforation."),
                  ("Unclear", "C89", "Perforation 2180.24-2186.24 m has no sand label -> zone not set."),
                  ("Unclear", "C94", "Five intervals listed under 'E-20/22 & E-23/32' without per-sand split -> each carries the combined zone label."),
                  ("Unclear", "perforation datum", "Perforation texts give 'm' without a datum; sheet depths are m-MDTHF (assumed same reference)."),
                  ("Missing", "TD only", "No casing, liner or tubing intervals are printed in the visible schematic; no tubing component was created "
                   "(no tubing label/size/interval). PBTD not printed."),
                  ("Missing", "wellInfo", "Completion date, rig, fluids, elevations, wellhead not in the visible sheet."),
                  ("Anomaly", "J72/J73 vs J68", "Left-column depths: dual packer 2047.2 m, no-go 2051.1 m, entry guide 2051.8 m below bridge plug 2037 m — kept as printed.")]
        unmapped = common_unmapped + [
            ("T41 / Q41 / X41 / Y41", f"{B.txt('T41')} at {B.v('Q41')} m-MDTHF ({ft(B.num('Q41'))} ft), length {B.v('S41')} m, ID {B.v('X41')} in, OD {B.v('Y41')} in",
             "UNRESOLVED TYPE: IPV1 valve — no suitable Component_Type"),
            ("T84 / Q84 / X84", f"{B.txt('T84')} at {B.v('Q84')} m-MDTHF ({ft(B.num('Q84'))} ft), ID {B.v('X84')} in",
             "tubing centralizer: no Component_Type; template 'centralizers' are per casing (Casing_ID, MD) -> not created"),
            ("B3", B.txt("B3"), "used for completionType"),
            ("Picture 57 (L35:P37)", "small equipment graphic (xl/media/image1.png)", "image without data"),
        ]
    elif well == "B17":
        review = [HIDDEN, DATUM_NOTE, DEPTH_POINT,
                  ("Unclear", "D54 / T54", "'SS STUB' / 'LS STUB' read as short / long string tubing stubs."),
                  ("Unclear", "D55", "'9-5/8\" DUAL HYD. PKR.' drawn in the left column only; recorded once with string 'Short & Long'."),
                  ("Unclear", "sheet 'MLE Space out'", "ESP space-out item list (lengths, no depths, no well name; identical in B18) — not mapped (well association not stated)."),
                  ("Missing", "T20, T24, T41, T46", "Flow couplings and crossovers have no depth printed."),
                  ("Missing", "T68, T71, T74, T82, T83, T85", "Blank pipe / screens / seal bore assemblies have no depth printed."),
                  ("Missing", "casing-9-5-8 top", "Casing top not printed -> empty (not assumed 0)."),
                  ("Missing", "TD", "No TD label in B17 (PBTD 2171.7 m only) -> totalDepth empty."),
                  ("Unclear", "perforation datum", "Perforation texts state 'm MDTHF' (same reference as the depth columns).")]
        unmapped = common_unmapped + [
            ("T39 / Q39", f"{B.txt('T39')} at {B.v('Q39')} m-MDTHF ({ft(B.num('Q39'))} ft)", "UNRESOLVED TYPE: IPV1 valve — no suitable Component_Type"),
            ("D60-D62", f"{B.txt('D60')} {B.txt('D61')} {B.txt('D62')}", "UNRESOLVED: combined item (hydro sub + perforated pup + blanking plug), no depth printed"),
            ("sheet 'MLE Space out'", "ESP space-out: 15 numbered items (packer, IPV1, crossovers, sand guard, pump 538-5800, gas handler, gas separators, seals, motor, "
             "sensor ACE Xtreme, motor guide), total length F62 = 70.86 m (formula =SUM(F14:F61)), pump setting depth to packer F64 = 55.46 m (=SUM(F16:F49))",
             "no depths / no well name; lengths only"),
            ("B3", B.txt("B3"), "used for completionType"),
            ("images", "Picture 57 equipment graphic, Graphic 754 tubing graphic, MLE sheet ESP sketch + 'UAL — An Uzma Company' logo", "images without data"),
        ]
    else:
        review = [HIDDEN, DATUM_NOTE, DEPTH_POINT,
                  ("Unclear", "D63/D64, D66, D84/D85, D92", "Several left-column cells repeat a second line 'PERM. PKR' (visible on the sheet). Treated as stray "
                   "text: D63/D84 keep 'GUIB MAGNUM 'HT' PERM. PKR'; D66 and D92 use only their first line. Confirm."),
                  ("Unclear", "C76 / C88", "Perforation depths written with comma decimals ('1725,24', '1744,4', '1760,24') -> read as decimal points."),
                  ("Unclear", "T51", "'End of ESP Assy' mapped as EOT (nearest type)."),
                  ("Unclear", "maxDeviation / trajectoryType", "Derived from the ANGLE column R (max 50.8 deg at Q25 701.51 m); rule: any printed inclination > 10 deg -> 'Deviated'."),
                  ("Unclear", "Picture 1 (AD17:BL48)", "Embedded wellhead drawing is a cropped screenshot of a PDF viewer: equipment table items 1-2 legible, items 3-9 cut off. "
                   "Read by OCR + visual check."),
                  ("Unclear", "sheet 'MLE Space out' (hidden)", "Identical to B17's sheet (pump 538-5800 differs from B18's pump label T49 '538-7400') -> not mapped."),
                  ("Missing", "casing-9-5-8", "Casing shoe label D92 has no depth (J92 empty) -> top/bottom empty."),
                  ("Missing", "T76", "Second blast joint has no depth."),
                  ("Unclear", "perforation datum", "Perforation texts give 'm' without a datum; sheet depths are m-MDTHF (assumed same reference).")]
        unmapped = common_unmapped + [
            ("T25 / Q25 / R25", f"{B.txt('T25')} at {B.v('Q25')} m-MDTHF ({ft(B.num('Q25'))} ft), angle {B.v('R25')} deg", "control-line end point: no Component_Type"),
            ("T45 / Q45", f"{B.txt('T45')} at {B.v('Q45')} m-MDTHF ({ft(B.num('Q45'))} ft), ID {B.v('X45')} in, OD {B.v('Y45')} in", "UNRESOLVED TYPE: IPV1 valve"),
            ("T50 / Q50", f"{B.txt('T50')} at {B.v('Q50')} m-MDTHF ({ft(B.num('Q50'))} ft)", "UNRESOLVED TYPE: downhole temperature/pressure sensor"),
            ("T82 / Q82", f"{B.txt('T82')} at {B.v('Q82')} m-MDTHF ({ft(B.num('Q82'))} ft), ID {B.v('X82')} in", "tubing centralizer: no Component_Type / casing-centralizer section"),
            ("column R", "inclinations " + ", ".join(f"{B.v('Q'+str(r))} m: {B.v('R'+str(r))} deg" for r in (20, 25, 27, 29, 32, 38, 41, 45, 49, 51)),
             "no azimuth -> no survey stations"),
            ("Picture 1 item 2", "ASSY, TUBING HEAD ADAPTER, 4-1/16\" 5K RX-39 ROTATING FLANGE TOP X 13-5/8\" BX-160 SEGMENTED CLAMP BTM W/1 ESP PENETRATOR FEED THRU HOLE",
             "OCR; no depth; tubing head adapter has no Component_Type"),
            ("K37", repr(B.v("K37")), "stray character"),
            ("B3", B.txt("B3"), "used for completionType"),
        ]
    return review, unmapped


def info_for(well, B, comps):
    td = {"B5": ("L99", "TD @ 2424.8 m MDTHF", 2424.8), "B18": ("L95", "TD @ 1852 m MDTHF", 1852.0)}.get(well)
    e = {"name": (f"Dulang_{well}", "B2", "STANDARDISED NAME", f"internal label B2 '{B.txt('B2')}' -> Dulang_{well} (user-approved naming)"),
         "completionType": (B.txt("B3"), "B3", "DIRECT", "visible title row 3 ('SINGLE OIL PRODUCER'); ESP completion")}
    if td:
        assert B.txt(td[0]) == td[1]
        e["totalDepth"] = (ft(td[2]), td[0], "CONVERTED", f"diagram label '{td[1]}': {td[2]} m x 3.28084; datum MDTHF kept")
    if well == "B18":
        e["wellhead"] = ("SOLAR ALERT SDN. BHD. — WELL B18 (OEM: SOLAR/ICC); 9-5/8\" CSG. X 4-1/2\" TBG (SINGLE) W/ESP; 5,000 PSI WP",
                         "Picture 1 (xl/media/image4.png) title block", "OCR + visual check", "embedded wellhead drawing, title block")
        e["tubing"] = ('4-1/2" TBG (SINGLE) W/ESP', "Picture 1 (xl/media/image4.png) title block", "OCR + visual check", "casing programme line of the wellhead drawing")
        angles = [(B.num(f"R{r}"), r) for r in range(17, 94) if f"R{r}" in B.c and isinstance(B.v(f"R{r}"), (int, float))]
        mx, row = max(angles)
        e["maxDeviation"] = (f"{mx:g} deg @ {ft(B.num(f'Q{row}'))} ft MD (THF)", f"R{row}", "DERIVED",
                             f"maximum of the ANGLE column R (R15/R16 '( ° )'), at Q{row} {B.v(f'Q{row}')} m-MDTHF")
        e["trajectoryType"] = ("Deviated", f"R{row}", "DERIVED", f"rule: maximum printed inclination {mx:g} deg > 10 deg -> 'Deviated'")
    return e, td


# ------------------------------------------------------------------ report
def md_table(rows, cols):
    esc = lambda x: "" if x is None else str(x).replace("|", "\\|").replace("\n", " ")
    return "\n".join(["| " + " | ".join(cols) + " |", "|" + "---|" * len(cols)] + ["| " + " | ".join(esc(r.get(c)) for c in cols) + " |" for r in rows])


def write_report(path, name, well, B, doc, tr, review, unmapped, checks, integrity, sheets_seen):
    cs = doc["components"]
    L = [f"# {name} — mapping and quality-check report", "",
         f"- Source workbook: `{NAS / SOURCES[well]}` (read-only; worked on a copy in `_work/src/`)",
         f"- Internal well label: `{B.txt('B2')}` (B2) — agrees with the file name; output name `{name}` (user-approved)",
         f"- Worksheet mapped: `{B.sheet}`; worksheets inspected: {sheets_seen}",
         f"- Reading: cells + formulas + cached values + drawing layer (shapes / text boxes / pictures) from the .xlsx XML; rendered with macOS "
         f"Quick Look (`_work/render/ql/`) and visually compared; embedded images extracted to `_work/media/{well}/` and OCR'd where they hold text",
         f"- Template: `{TPL}`; Type_Keys from `{XTPL.name}` > Component_Types",
         f"- Output JSON: `{name}.json` — {len(cs)} components; surveys / cement / plugBackCement / centralizers / floatEquipment / formations / "
         "formationTops / fluidContacts left empty (no supported source data)",
         f"- Generated: {dt.datetime.now():%Y-%m-%d %H:%M}", "",
         "## Depth units, reference and conversion", "",
         f"- Source depths are **{DATUM}**. JSON depths are **ft = m × 3.28084**, 3 decimals; original metres kept in `topMdM` / `bottomMdM`.",
         "- The MD reference (THF) was not shifted to drill floor/RKB and no MD→TVD conversion was made.", "",
         md_table(tr.conv, ["field", "ref", "m", "unit", "ft"]), "",
         "## wellInfo", "", md_table([r for r in tr.rows if r["json_path"].startswith("wellInfo")],
                                     ["json_path", "mapped_value", "cell_or_object", "original_value", "method", "note"]), "",
         "## Components", "",
         md_table([dict(id=c["id"], type=c["type"], name=c["name"], string=c.get("string"), topMd_ft=c["topMd"], bottomMd_ft=c["bottomMd"],
                        topMd_m=c.get("topMdM"), bottomMd_m=c.get("bottomMdM"), OD_in=c["outerDiameter"], ID_in=c["innerDiameter"],
                        interval=c.get("interval"), zone=c.get("zone"), estimated=",".join(c.get("estimatedFields", []))) for c in cs],
                  ["id", "type", "name", "string", "topMd_ft", "bottomMd_ft", "topMd_m", "bottomMd_m", "OD_in", "ID_in", "interval", "zone", "estimated"]), "",
         f"Cell-by-cell traceability: `{name}_traceability.csv` (every populated JSON path, source cell/object, original value, formula + cached value, unit, method).", "",
         "## Approximate type mappings and assumptions", "",
         md_table([dict(id=c["id"], type=c["type"], actualType=c.get("actualType"), note=c.get("estimationNote")) for c in cs if c.get("estimated")],
                  ["id", "type", "actualType", "note"]), "",
         "## Items requiring review", ""]
    for cat in ("Identity", "Conflict", "Missing", "Unclear", "Anomaly"):
        items = [r for r in review if r[0] == cat]
        if items:
            L += [f"### {cat} ({len(items)})", "", md_table([dict(location=r[1], detail=r[2]) for r in items], ["location", "detail"]), ""]
    L += ["## Unmapped information", "", md_table([dict(location=u[0], content=u[1], reason=u[2]) for u in unmapped], ["location", "content", "reason"]), "",
          "## Quality checks", "", md_table([dict(check=c, result="PASS" if ok else "FAIL") for c, ok in checks], ["check", "result"]), "",
          "## Source / template integrity (SHA-256)", "", md_table(integrity, ["file", "sha256 before", "sha256 after", "unchanged"]), ""]
    path.write_text("\n".join(L), encoding="utf-8")


# ------------------------------------------------------------------ main
def main():
    py = sys.executable
    ref_files = [TPL, XTPL] + [NAS / f for f in SOURCES.values()]
    before = {p: sha(p) for p in ref_files}
    tpl = json.loads(subprocess.run(["textutil", "-convert", "txt", "-stdout", str(TPL)], capture_output=True, text=True, check=True).stdout)
    keymap = {r[0]: r[2] for r in load_workbook(XTPL, read_only=True)["Component_Types"].iter_rows(min_row=2, values_only=True)}
    for f in OUT.glob("Dulang_*"):
        raise SystemExit(f"{f} exists — not overwriting")
    allok = True
    for well in ("B5", "B17", "B18"):
        name = f"Dulang_{well}"
        src = WORK / "src" / SOURCES[well]
        if not src.exists():
            shutil.copyfile(NAS / SOURCES[well], src)
        assert sha(src) == before[NAS / SOURCES[well]], "working copy differs from the NAS source"
        dump = WORK / "dump" / (SOURCES[well][:-5] + ".json")
        subprocess.run([py, str(ROOT / "scripts" / "dump_xlsx_schematic.py"), str(src), str(dump)], check=True, capture_output=True)
        d = json.loads(dump.read_text())
        s0 = next(s for s in d["sheets"] if s["title"] == SHEETS[well])
        main_drawing = {k: v for k, v in d["drawings"].items() if k == "xl/drawings/drawing1.xml"}
        B = Book(SOURCES[well], s0["title"], s0["cells"], main_drawing)
        assert B.txt("B2") == LABELS[well] and B.txt("J16") == "m-MDTHF" and B.txt("Q16") == "m-MDTHF"
        assert B.txt("B5") == "Field / Platform : Erb West / EWDP-A Platform" and B.txt("B4") == "Well ID / SLOT NO : EW-119 / Slot #9"
        sheets_seen = ", ".join(f"'{s['title']}' ({s['state']})" for s in d["sheets"])

        tr = Trace(name, B)
        specs = {"B5": spec_b5, "B17": spec_b17, "B18": spec_b18}[well]()
        comps, flags = build_components(B, specs, keymap, tr)
        info, td = info_for(well, B, comps)
        wi = well_info(tpl, info, tr)
        if td:
            tr.conv.append(dict(field="wellInfo.totalDepth", ref=td[0], m=td[2], unit="m-MDTHF (in text)", ft=ft(td[2])))
        review, unmapped = notes(well, B)
        review += flags
        doc = {k: {"wellInfo": wi, "components": comps}.get(k, []) for k in tpl}
        jpath = OUT / f"{name}.json"
        jpath.write_text(json.dumps(doc, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")

        # ---------------- checks (independent re-read of the working copy with openpyxl)
        back = json.loads(jpath.read_text())
        cs = back["components"]
        ws = load_workbook(src, data_only=True)[SHEETS[well]]
        kind = lambda v: "number" if isinstance(v, (int, float)) and not isinstance(v, bool) else type(v).__name__
        cell_depth_ok = all(round(float(ws[c["ref"]].value) * FT_PER_M, 3) == c["ft"] for c in tr.conv if c["ref"] in B.c and kind(ws[c["ref"]].value) == "number")
        conv_ok = all(c["ft"] == round(c["m"] * FT_PER_M, 3) for c in tr.conv)
        json_depths = {f"{c['id']}.{f}": c[f] for c in cs for f in ("topMd", "bottomMd")}
        allowed_extra = {"topMdM", "bottomMdM", "actualType", "string"}
        tpl_keys = {k for c in tpl["components"] for k in c}
        checks = [
            ("valid JSON (re-parsed)", True),
            ("top-level sections and order = template", list(back) == list(tpl)),
            ("wellInfo keys and order = template", list(back["wellInfo"]) == list(tpl["wellInfo"])),
            ("wellInfo types = template type or empty value", all(v is None or v == "-" or kind(v) == kind(tpl["wellInfo"][k]) for k, v in back["wellInfo"].items())),
            ("component keys = template keys + supplementary (topMdM, bottomMdM, actualType, string)", all(set(c) <= tpl_keys | allowed_extra for c in cs)),
            ("numeric fields numeric or null", all(all(c.get(f) is None or kind(c[f]) == "number" for f in ("topMd", "bottomMd", "outerDiameter", "innerDiameter", "topMdM", "bottomMdM", "weight")) for c in cs)),
            ("every type is a Component_Types Type_Key", all(c["type"] in set(keymap.values()) for c in cs)),
            ("component ids unique", len({c["id"] for c in cs}) == len(cs)),
            ("no template sample records (surveys, components, formationTops)", back["surveys"] == [] and back["formationTops"] == [] and not {c["id"] for c in cs} & {t["id"] for t in tpl["components"]}),
            ("every depth ft = m x 3.28084 (3 dp)", conv_ok),
            ("cell depths re-read from the working copy match", cell_depth_ok),
            ("JSON depths = conversion table; metres kept", all(json_depths[c["field"]] == c["ft"] for c in tr.conv if c["field"] in json_depths)
             and all((c.get("topMd") is None) == (c.get("topMdM") is None) and (c.get("bottomMd") is None) == (c.get("bottomMdM") is None) for c in cs)),
            ("closest-type mappings flagged with estimated / estimatedFields 'type' / actualType", all(("type" in c.get("estimatedFields", [])) == bool(sp["closest"]) and (not sp["closest"] or c.get("actualType") == c["name"]) for c, sp in zip(cs, specs))),
            ("cell-text names equal the source cells", all(c["name"] == " ".join(B.txt(r, ln) for r, ln in sp["name"]) for c, sp in zip(cs, specs) if sp["name"][0][0] not in ("draw", "ocr", "perf"))),
            ("no depth placeholder / default 100 ft", all(v not in (100, 0) for v in json_depths.values() if v is not None)),
            ("well identity: B2 label matches the output name", well.replace("B", "B-0" if len(well) == 2 else "B-") in B.txt("B2")),
        ]
        integrity = [dict(file=str(p), **{"sha256 before": before[p], "sha256 after": sha(p), "unchanged": "YES" if before[p] == sha(p) else "NO"}) for p in ref_files]
        checks.append(("sources, template and component reference unchanged (SHA-256)", all(i["unchanged"] == "YES" for i in integrity)))
        write_report(OUT / f"{name}_mapping_report.md", name, well, B, back, tr, review, unmapped, checks, integrity, sheets_seen)
        with open(OUT / f"{name}_traceability.csv", "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=list(tr.rows[0]))
            w.writeheader(); w.writerows(tr.rows)
        print(f"=== {name}: {len(cs)} components ({sum(1 for c in cs if c.get('estimated'))} flagged), {len(review)} review items, {len(unmapped)} unmapped")
        for c, ok in checks:
            allok &= ok
            print("  PASS" if ok else "  FAIL", "-", c)
    (WORK / "sha256.txt").write_text("\n".join(f"{before[p]}  {sha(p)}  {p}" for p in ref_files) + "\n")
    print("ALL PASS" if allok else "SOME CHECKS FAILED")


if __name__ == "__main__":
    main()
