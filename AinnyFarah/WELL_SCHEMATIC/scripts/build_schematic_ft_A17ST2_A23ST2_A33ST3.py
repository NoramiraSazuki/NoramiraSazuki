#!/usr/bin/env python3
"""Map the Dulang A17ST2 / A23ST2 / A33ST3 well-schematic workbooks (NAS, read-only) into the JSON template.

One folder per well (WELL_SCHEMATIC/output_<well>/): <well>.json, <well>_mapping_report.md, <well>_traceability.csv and
_work/ (working copy of the source, cell dumps, SHA-256 before/after). Rules applied:
  * depths reported in ft = m / 0.3048 (3 dp); original metres, unit evidence and cell kept in the report / CSV
  * MD reference and datum kept as printed (m-MDDF); no MD<->TVD or datum conversion; TVD has no component field -> unmapped
  * a value is mapped only when its cell is read and asserted; nothing is assumed. Values that follow from an explicit
    source statement (e.g. 'From Tubing Hanger to TRSV', a length formula naming its end cells) are flagged estimated/DERIVED
  * component types come from the Component_Types list of well-schematic-sample-template.xlsx; a closest-type mapping is
    flagged (estimatedFields 'type'); items with no supported type are left out of the JSON and listed as unmapped
  * conflicts are recorded with every value and location; the conflicting field is left empty
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

import xlrd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from xlrd import compdoc

ROOT = Path(__file__).resolve().parent.parent
NAS = Path("/Volumes/homes/public/DulangOSS/20261016/DULANG_WELLS/Well_Schematic/1. Well Schematic/Dulang A")
TPL = ROOT / "well-schematic-template-json.rtf"
XTPL = ROOT / "well-schematic-sample-template.xlsx"
PREV = ROOT / "output_A" / "run_xls_A17ST2_A23ST2_A33ST3"
M_PER_FT = 0.3048
SOURCES = {"A17ST2": "Dulang A17ST2 Well Schematic updated .xlsx", "A23ST2": "Dulang A23ST2 Well Schematic.xlsx",
           "A33ST3": "Dulang A33ST3 Well Schematic.xls"}


def ft(m):
    return None if m is None else round(m / M_PER_FT, 3)


def sha(p):
    h = hashlib.sha256()
    fd = os.open(p, os.O_RDONLY)
    while (b := os.read(fd, 1 << 20)):
        h.update(b)
    os.close(fd)
    return h.hexdigest()


# ------------------------------------------------------------------ source access
class Book:
    """Cell values of one worksheet (cached values for formulas) + unit evidence per column."""

    def __init__(self, workbook, sheet, cells, depth_units, dia_units):
        self.workbook, self.sheet = workbook, sheet
        self.c = {c["ref"]: c for c in cells}
        self.depth_units, self.dia_units = depth_units, dia_units

    def v(self, ref):
        assert ref in self.c, f"{self.workbook}: {ref} is empty"
        return self.c[ref]["value"]

    def txt(self, ref):
        v = self.v(ref)
        assert isinstance(v, str), (ref, v)
        return v.strip()

    def num(self, ref):
        v = self.v(ref)
        if isinstance(v, str):
            m = re.fullmatch(r"\s*([\d,]+(?:\.\d+)?)\s*(m\s*MDDF|PPF|)\s*", v)
            assert m, (self.workbook, ref, v)
            return float(m.group(1).replace(",", ""))
        assert isinstance(v, (int, float)) and not isinstance(v, bool), (ref, v)
        return float(v)

    def size(self, ref):
        """Nominal size from text such as 13-3/8" or 7" Liner, in inches."""
        m = re.search(r'(\d+)(?:-(\d+)/(\d+))?\s*"', self.txt(ref))
        assert m, (ref, self.v(ref))
        return int(m.group(1)) + (int(m.group(2)) / int(m.group(3)) if m.group(2) else 0)

    def formula(self, ref):
        return self.c[ref].get("formula")

    def orig(self, ref):
        s = repr(self.v(ref))
        return s + (f" (formula {self.formula(ref)}, cached value)" if self.formula(ref) else "")

    def depth_unit(self, ref):
        v = self.v(ref)
        if isinstance(v, str) and re.search(r"m\s*MDDF", v):
            return f"m MDDF (in cell text {ref})"
        f = self.formula(ref)
        if f and re.fullmatch(r"=[A-Z]+\d+", f):
            return self.depth_unit(f[1:]) + f" via formula {f}"
        col = re.match(r"[A-Z]+", ref).group(0)
        hdr = self.depth_units[col]
        t = self.txt(hdr)
        assert "m-MDDF" in t or "m MDDF" in t, (hdr, t)
        return f"m-MDDF (column header {hdr} '{t}')"

    def dia_unit(self, ref):
        col = re.match(r"[A-Z]+", ref).group(0)
        hdr = self.dia_units[col]
        assert self.txt(hdr) == "IN", (hdr, self.v(hdr))
        return f"in (column header {hdr} 'IN')"


def dump_xls(path):
    b = xlrd.open_workbook(path, formatting_info=True)
    out = {"sheets": []}
    for s in b.sheets():
        cells = []
        for r in range(s.nrows):
            for c in range(s.ncols):
                ct = s.cell_type(r, c)
                if ct in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                    continue
                cells.append({"ref": f"{get_column_letter(c + 1)}{r + 1}", "value": s.cell_value(r, c), "ctype": ct})
        merged = [f"{get_column_letter(c0 + 1)}{r0 + 1}:{get_column_letter(c1)}{r1}" for r0, r1, c0, c1 in s.merged_cells]
        out["sheets"].append({"title": s.name, "visibility": s.visibility, "cells": cells, "merged": merged})
    return out


def workbook_stream(path):
    return compdoc.CompDoc(open(path, "rb").read()).get_named_stream("Workbook")


# ------------------------------------------------------------------ specs
def C(cid, ctype, name, top=None, bot=None, od=None, idd=None, wt=None, gr=None, ln=None, typ=None, note=None):
    """top/bot: 'S30' (direct) or ('der', 'S30', why). od/idd: 'Z29' (number) or ('size', 'V9'). wt/gr/ln: cell refs.
    typ: rationale when the Component_Type is the closest available one (flagged)."""
    return dict(cid=cid, ctype=ctype, name=name, top=top, bot=bot, od=od, idd=idd, wt=wt, gr=gr, ln=ln, typ=typ, note=note)


CLOSING = "closing sleeve (shifts to open/close ports) -> Sliding Sleeve is the closest type"
POP = "pump-out plug (plug removed by pressure) -> Mechanical Plug is the closest type"
SEALBORE = "sealbore extension (polished seal bore) -> PBR is the closest type"
FCS = "'FRACTURE / CIRCULATION SLEEVE' -> Circulating Valve (closest by 'circulation'; Sliding Sleeve also possible)"
DIAL = "'AUTO SPM DIAL' is a side-pocket mandrel variant -> SPM (Gas Lift)"
PERFPUP = "perforated pup -> Pup Joint (a pup joint with holes)"
MULE_DROP = "'Muleshoe Guide / Drop Sub' (combined item) -> Mule Shoe"
CASING_DEPTH = "casing-table 'DEPTH' value mapped to bottomMd as the setting (shoe) depth — interpretation, review"


def spec_a17():
    t = "header tubing row"
    return [
        C("conductor", "Casing", "P9", bot="AB9", od=("size", "V9"), wt="X9", gr="Y9", note=CASING_DEPTH),
        C("int-casing", "Casing", "P10", bot="AB10", od=("size", "V10"), wt="X10", gr="Y10", note=CASING_DEPTH),
        C("casing", "Casing", "P11", bot="AB11", od=("size", "V11"), wt="X11", gr="Y11", note=CASING_DEPTH),
        C("liner-7", "Liner", "P13", top="K34", bot="AB13", od=("size", "V13"), wt="X13", gr="Y13",
          note="top = diagram label F34 ' TOL (7\" LINER)' / K34; bottom = header AB13 (formula =K110, i.e. set equal to TD) — " + CASING_DEPTH),
        C("tubing-upper", "Tubing", "U26", bot=("der", "AB14", "header row 14 DEPTH = formula =S30 (TRSV) and label U25 'From Tubing Hanger to TRSV'"),
          od=("size", "V14"), wt="X14", gr="Y14", note=f"{t} 14; top (tubing hanger depth) not printed"),
        C("tubing-lower", "Tubing", ["V15", "X15", "Y15", "Z15"], top=("der", "S30", "upper tubing ends at TRSV (U25); header AB15 formula =ROUNDUP(S71-S30,0) starts at S30"),
          od=("size", "V15"), wt="X15", gr="Y15", note=f"{t} 15 (no diagram label; name = header cells joined); bottom NOT mapped — conflicting candidates, see report"),
        C("tubing-hanger", "Tubing Hanger", "U21"),
        C("flow-coupling-1", "Flow Coupling", "U29", od="AB29", idd="Z29"),
        C("trsv", "TRSV", "U30", top="S30", od="AB30", idd="Z30"),
        C("flow-coupling-2", "Flow Coupling", "U31", od="AB31", idd="Z31"),
        C("spm-1", "SPM (Gas Lift)", "U42", top="S42", od="AB42", idd="Z42"),
        C("spm-dial-1", "SPM (Gas Lift)", "U44", top="S44", od="AB44", idd="Z44", typ=DIAL),
        C("spm-2", "SPM (Gas Lift)", "U46", top="S46", od="AB46", idd="Z46"),
        C("spm-dial-2", "SPM (Gas Lift)", "U48", top="S48", od="AB48", idd="Z48", typ=DIAL),
        C("spm-3", "SPM (Gas Lift)", ["U50", "U51"], top="S50", od="AB50", idd="Z50", note="name continues on U51 (same sentence)"),
        C("ahr-packer", "Packer", "U59", top="S59", od="AB59", idd="Z59"),
        C("landing-nipple", "Landing Nipple", "U62", top="S62", od="AB62", idd="Z62"),
        C("mule-shoe", "Mule Shoe", "U68", top="S68", od="AB68", idd="Z68"),
        C("vch-packer", "Packer", "U71", top="S71", od="AB71", idd="Z71"),
        C("mcsz-sleeve-1", "Sliding Sleeve", "U73", top="S73", od="AB73", idd="Z73", typ=CLOSING),
        C("screen-1-top-joint", "Screen", "U75", top="S75", od="AB75", idd="Z75"),
        C("screen-1-production-joint", "Screen", "U77", top="S77", od="AB77", idd="Z77"),
        C("fcs-1", "Circulating Valve", "U79", top="S79", od="AB79", idd="Z79", typ=FCS),
        C("tnt-isolation-packer", "Packer", "U83", top="S83", od="AB83", idd="Z83"),
        C("mcsz-sleeve-2", "Sliding Sleeve", "U88", top="S88", od="AB88", idd="Z88", typ=CLOSING),
        C("screen-2-top-joint", "Screen", "U90", top="S90", od="AB90", idd="Z90"),
        C("screen-2-production-joint", "Screen", "U92", top="S92", od="AB92", idd="Z92"),
        C("fcs-2", "Circulating Valve", "U94", top="S94", od="AB94", idd="Z94", typ=FCS),
        C("sump-seal-assembly", "Seal Assembly", "U98", top="S98", od="AB98", idd="Z98"),
        C("awv-sump-packer", "Packer", "U99", top="S99", od="AB99", idd="Z99"),
        C("pbtd", "PBTD", "I108", top="K108"),
    ]


def spec_a23():
    S, L = "short string (columns B/D/F/K)", "long string (columns U/S/Z/AB)"
    grade_conflict = "grade NOT mapped: header Y15 '13CR-L80' conflicts with diagram label 'HP2 13Cr95'"
    return [
        C("conductor", "Casing", "P9", od=("size", "V9"), note="depth AB9 'xx m MDDF' is a placeholder -> not mapped"),
        C("int-casing", "Casing", "P11", od=("size", "V11"), wt="X11", gr="Y11", note="depth AB11 'xx m MDDF' placeholder -> not mapped"),
        C("casing", "Casing", "P12", od=("size", "V12"), wt="X12", gr="Y12", note="depth AB12 'xx m MDDF' placeholder -> not mapped"),
        C("liner", "Liner", "P13", od=("size", "V13"), wt="X13", gr="Y13", note="AB13 is a length (=ROUNDUP(S127-S85,0)), not a depth -> depths not mapped"),
        C("short-tubing-upper", "Short String", "F26", bot=("der", "K32", "label F27 'From Tubing Hanger to TRSV'; short-string TRSV at K32"),
          od=("size", "V14"), wt="X14", gr="Y14", note=f"{S}; header row 14 (applies to both strings); top (hanger) not printed"),
        C("short-tubing-lower", "Short String", "F34", top=("der", "K32", "header AB15 formula =ROUNDUP((K73-K32)+(S119-S30),0): short tubing from K32 (TRSV)"),
          bot=("der", "K73", "same AB15 formula: to K73 (label F73 'EOT')"), od=("size", "V15"), wt="X15", note=f"{S}; header row 15; {grade_conflict}"),
        C("long-tubing-upper", "Long String", "U26", bot=("der", "S30", "label U27 'From Tubing Hanger to TRSV'; long-string TRSV at S30"),
          od=("size", "V14"), wt="X14", gr="Y14", note=f"{L}; header row 14; top (hanger) not printed"),
        C("long-tubing-lower", "Long String", "U32", top=("der", "S30", "header AB15 formula: long tubing from S30 (TRSV)"),
          bot=("der", "S119", "same AB15 formula: to S119 (SELF ALIGNING MULE SHOE)"), od=("size", "V15"), wt="X15", note=f"{L}; header row 15; {grade_conflict}"),
        # short string
        C("short-flow-coupling-1", "Flow Coupling", "F31", od="B31", idd="D31", note=S),
        C("short-trsv", "TRSV", "F32", top="K32", od="B32", idd="D32", note=S),
        C("short-flow-coupling-2", "Flow Coupling", "F33", od="B33", idd="D33", note=S),
        C("short-spm-1", "SPM (Gas Lift)", "F40", top="K40", od="B40", idd="D40", note=S),
        C("short-spm-2", "SPM (Gas Lift)", "F43", top="K43", od="B43", idd="D43", note=S),
        C("short-spm-3", "SPM (Gas Lift)", "F46", top="K46", od="B46", idd="D46", note=S),
        C("short-spm-4", "SPM (Gas Lift)", "F49", top="K49", od="B49", idd="D49", note=S),
        C("short-sliding-sleeve", "Sliding Sleeve", "F54", top="K54", od="B54", idd="D54", note=S),
        C("short-landing-nipple", "Landing Nipple", "F67", top="K67", od="B67", idd="D67", note=S),
        C("short-pump-out-plug", "Mechanical Plug", "F69", top="K69", od="B69", idd="D69", typ=POP, note=S),
        C("short-perforated-pup", "Pup Joint", "F70", od="B70", idd="D70", typ=PERFPUP, note=S),
        C("short-eot", "EOT", "F73", top="K73", note=S),
        C("dual-packer", "Packer", "U60", top="S60", od="AB60", idd="Z60", note="dual-string packer ('DUAL STRING' in name), values in the long-string columns"),
        # long string
        C("long-flow-coupling-1", "Flow Coupling", "U29", od="AB29", idd="Z29", note=L),
        C("long-trsv", "TRSV", "U30", top="S30", od="AB30", idd="Z30", note=L),
        C("long-flow-coupling-2", "Flow Coupling", "U31", od="AB31", idd="Z31", note=L),
        C("long-spm-1", "SPM (Gas Lift)", "U39", top="S39", od="AB39", idd="Z39", note=L),
        C("long-spm-2", "SPM (Gas Lift)", "U42", top="S42", od="AB42", idd="Z42", note=L),
        C("long-spm-3", "SPM (Gas Lift)", "U45", top="S45", od="AB45", idd="Z45", note=L),
        C("long-spm-4", "SPM (Gas Lift)", "U48", top="S48", od="AB48", idd="Z48", note=L),
        C("long-sliding-sleeve", "Sliding Sleeve", "U80", top="S80", od="AB80", idd="Z80", note=L),
        C("gp-packer", "Packer", "U85", top="S85", od="AB85", idd="Z85", note=L),
        C("mcs-sleeve-1", "Sliding Sleeve", "U86", top="S86", od="AB86", idd="Z86", typ=CLOSING, note=L),
        C("sealbore-extension-1", "PBR", "U87", top="S87", od="AB87", idd="Z87", typ=SEALBORE, note=L),
        C("screen-1", "Screen", "U95", od="AB95", idd="Z95", note=L),
        C("isolation-packer-1", "Packer", "U102", top="S102", od="AB102", idd="Z102", note=L),
        C("mcs-sleeve-2", "Sliding Sleeve", "U103", top="S103", od="AB103", idd="Z103", typ=CLOSING, note=L),
        C("sealbore-extension-2", "PBR", "U104", top="S104", od="AB104", idd="Z104", typ=SEALBORE, note=L),
        C("isolation-packer-2", "Packer", "U111", top="S111", od="AB111", idd="Z111", note=L),
        C("mcs-sleeve-3", "Sliding Sleeve", "U112", top="S112", od="AB112", idd="Z112", typ=CLOSING, note=L),
        C("sealbore-extension-3", "PBR", "U113", top="S113", od="AB113", idd="Z113", typ=SEALBORE, note=L),
        C("long-seal-assembly", "Seal Assembly", "U115", top="S115", od="AB115", idd="Z115", note=L),
        C("long-landing-nipple", "Landing Nipple", "U116", top="S116", od="AB116", idd="Z116", note=L),
        C("long-pump-out-plug", "Mechanical Plug", "U117", top="S117", od="AB117", idd="Z117", typ=POP, note=L),
        C("long-perforated-pup", "Pup Joint", "U118", top="S118", od="AB118", idd="Z118", typ=PERFPUP, note=L),
        C("long-mule-shoe", "Mule Shoe", "U119", top="S119", od="AB119", idd="Z119", note=L),
        C("screen-2", "Screen", "U121", od="AB121", idd="Z121", note=L),
        C("collet-seal-assembly", "Seal Assembly", "U127", top="S127", od="AB127", idd="Z127", note=L),
        C("awg-sump-packer", "Packer", "U129", top="S129", od="AB129", idd="Z129", note=L),
        C("pbtd", "PBTD", "I137", top="K137"),
    ]


def spec_a33():
    cas = lambda cid, r: C(cid, "Casing", f"I{r}", bot=f"T{r}", od=("size", f"I{r}"), wt=f"P{r}", gr=f"Q{r}",
                           note=f"casing scheme row {r}: TOP S{r} 'Surface' (no depth value) -> topMd not mapped")
    return [
        cas("conductor", 6), cas("casing-16", 7), cas("casing-13-3-8", 8), cas("casing-9-5-8", 9),
        C("liner-7", "Liner", "I10", top="F40", od=("size", "I10"),
          note="header row 10 weight/grade/thread/top/bottom all 'TBC'; top from diagram label B40 '7\"Liner (TOL)' / F40; bottom not printed"),
        C("tubing-gre", "Tubing", "O48", od=("size", "I11"), wt="P11", gr="Q11",
          note="diagram label O48; weight/grade from header row 11 'Tubing 3-1/2\" GRE lined' (matches 'GRE'/'FOX'); interval not printed"),
        C("tubing-15cr", "Tubing", "O59", od=("size", "I12"), wt="P12", gr="Q12",
          note="diagram label O59; weight/grade from header row 12 'Tubing 3-1/2\" - Below 7\" AHC' (grade 15CR-125 matches 'UHP 15Cr-125'); interval not printed"),
        C("xmas-tree", "Xmas Tree", "B13", note="no depth printed"),
        C("tubing-hanger", "Tubing Hanger", "O18", idd="S18", ln="M18", note="MAX OD T18 'TBC'; depth not printed (K17 15.8 is 'ELEVATION (DF to THF)', see report)"),
        C("crossover-1", "Crossover", "O19", od="T19", idd="S19", ln="M19"),
        C("pup-joint-12ft", "Pup Joint", "O20", od="T20", idd="S20", ln="M20"),
        C("pup-joint-6ft-1", "Pup Joint", "O23", od="T23", idd="S23", ln="M23"),
        C("flow-coupling-1", "Flow Coupling", "O24", od="T24", idd="S24", ln="M24"),
        C("trsv", "TRSV", "O25", top="K25", od="T25", idd="S25", ln="M25"),
        C("flow-coupling-2", "Flow Coupling", "O26", od="T26", idd="S26", ln="M26"),
        C("pup-joint-4ft", "Pup Joint", "O27", od="T27", idd="S27", ln="M27"),
        C("pup-joint-6ft-2", "Pup Joint", "O50", od="T50", idd="S50", ln="M50"),
        C("crossover-2", "Crossover", "O51", od="T51", idd="S51", ln="M51"),
        C("bluepack-packer", "Packer", "O52", top="K52", od="T52", idd="S52", ln="M52"),
        C("pup-joint-5ft", "Pup Joint", "O53", od="T53", idd="S53", ln="M53"),
        C("crossover-3", "Crossover", "O54", od="T54", idd="S54", ln="M54"),
        C("landing-nipple", "Landing Nipple", "O55", top="K55", od="T55", idd="S55", ln="M55"),
        C("crossover-4", "Crossover", "O56", od="T56", idd="S56", ln="M56"),
        C("pup-joint-6ft-3", "Pup Joint", "O57", od="T57", idd="S57", ln="M57"),
        C("muleshoe", "Mule Shoe", "O61", top="K61", od="T61", idd="S61", typ=MULE_DROP),
        C("pbtd", "PBTD", "B74", top="F74"),
    ]


# ------------------------------------------------------------------ builders
class Trace:
    def __init__(self, well, B):
        self.well, self.B, self.rows, self.conv = well, B, [], []

    def add(self, path, mapped, ref, unit, note, label=None):
        self.rows.append(dict(well=self.well, json_path=path, mapped_value=mapped, workbook=self.B.workbook, worksheet=self.B.sheet,
                              cell=ref, source_label=label or "", original_value="; ".join(self.B.orig(r) for r in ref.split("+")) if ref else "", original_unit=unit, note=note))


def build_components(B, specs, keymap, tr):
    comps, flags = [], []
    for s in specs:
        refs = s["name"] if isinstance(s["name"], list) else [s["name"]]
        name = " ".join(B.txt(r) for r in refs)
        d = {"id": s["cid"], "type": keymap[s["ctype"]], "name": name, "topMd": None, "bottomMd": None,
             "outerDiameter": None, "innerDiameter": None}
        p = f"components[{s['cid']}]"
        tr.add(f"{p}.name", name, "+".join(refs), "", "DIRECT (cell text)")
        tr.add(f"{p}.type", d["type"], refs[0], "", f"Component_Type '{s['ctype']}'" + (f" — CLOSEST TYPE: {s['typ']}" if s["typ"] else " (matches source term)"))
        est, enote, src = [], [], [f"name {'+'.join(refs)}"]
        if s["typ"]:
            est.append("type"); enote.append(f"type is the closest available: {s['typ']}")
        for key, field in (("top", "topMd"), ("bot", "bottomMd")):
            x = s[key]
            if x is None:
                continue
            derived = isinstance(x, tuple)
            ref = x[1] if derived else x
            m = B.num(ref)
            unit = B.depth_unit(ref)
            d[field] = ft(m)
            how = f"{m} m -> {d[field]} ft (m / 0.3048, 3 dp)"
            tr.add(f"{p}.{field}", d[field], ref, unit, ("DERIVED: " + x[2] + "; " if derived else "DIRECT; ") + how)
            tr.conv.append(dict(field=f"{s['cid']}.{field}", cell=ref, m=m, unit=unit, ft=d[field], kind="derived" if derived else "direct"))
            src.append(f"{field} {ref}={m} m MDDF")
            if derived:
                est.append(field); enote.append(f"{field} DERIVED: {x[2]}")
        for key, field in (("od", "outerDiameter"), ("idd", "innerDiameter")):
            x = s[key]
            if x is None:
                continue
            if isinstance(x, tuple):
                d[field] = B.size(x[1])
                tr.add(f"{p}.{field}", d[field], x[1], "in (\" in text)", "DIRECT; nominal size text, fraction written as decimal inches")
                src.append(f"{field} {x[1]} '{B.txt(x[1])}'")
            else:
                d[field] = B.num(x)
                tr.add(f"{p}.{field}", d[field], x, B.dia_unit(x), "DIRECT")
                src.append(f"{field} {x}")
        if s["wt"]:
            w = B.v(s["wt"])
            unit = "ppf (in cell text)" if isinstance(w, str) and "PPF" in w else ("lb/ft (column header P4 '#')" if B.sheet == "SCHEM-1" else "unit not printed")
            d["weight"] = B.num(s["wt"])
            tr.add(f"{p}.weight", d["weight"], s["wt"], unit, "DIRECT")
            src.append(f"weight {s['wt']}")
            if unit == "unit not printed":
                flags.append(("Unclear", f"{s['wt']}", f"{s['cid']}: weight {w!r} printed without a unit (other rows say PPF) — value kept, unit to confirm"))
        if s["gr"]:
            d["grade"] = B.txt(s["gr"])
            tr.add(f"{p}.grade", d["grade"], s["gr"], "", "DIRECT")
            src.append(f"grade {s['gr']}")
        if s["ln"]:
            d["interval"] = f"{B.num(s['ln'])!r} m"
            tr.add(f"{p}.interval", d["interval"], s["ln"], "m (column header M16 'm')", "DIRECT; equipment LENGTH -> 'interval' (template uses it for a tool length, e.g. '50 ft'); not a depth, unit kept")
            src.append(f"length {s['ln']}")
        if est:
            d.update(estimated=True, estimatedFields=est, estimationNote="; ".join(enote))
        d["source"] = f"{B.workbook} > sheet '{B.sheet}' > " + "; ".join(src) + (f"; {s['note']}" if s["note"] else "")
        if d["topMd"] is None and d["bottomMd"] is None and s["ctype"] not in ("Xmas Tree",):
            flags.append(("Missing", "+".join(refs), f"{s['cid']}: no depth printed for '{name[:60]}' -> topMd/bottomMd null"))
        comps.append(d)
    return comps, flags


def well_info(tpl, entries, tr):
    """entries: key -> (value, refs, note). Keys not given take the template's empty value ('-' for text, null for numbers)."""
    wi = {}
    for k, tv in tpl["wellInfo"].items():
        if k in entries:
            val, refs, note = entries[k]
            wi[k] = val
            for r in refs or [None]:
                tr.add(f"wellInfo.{k}", val, r, "", note)
        else:
            wi[k] = "-" if isinstance(tv, str) else None
            tr.add(f"wellInfo.{k}", wi[k], None, "", "NOT IN SOURCE -> template empty value; flagged")
    return wi


# ------------------------------------------------------------------ report
def md_table(rows, cols):
    esc = lambda x: "" if x is None else str(x).replace("|", "\\|").replace("\n", " ")
    return "\n".join(["| " + " | ".join(cols) + " |", "|" + "---|" * len(cols)] + ["| " + " | ".join(esc(r.get(c)) for c in cols) + " |" for r in rows])


def write_report(path, well, B, doc, tr, review, unmapped, checks, integrity, ref_note, type_rows):
    cs = doc["components"]
    L = [f"# {well} — mapping and quality-check report", "",
         f"- Source workbook: `{NAS / SOURCES[well.split('_')[1]]}` (read-only; worked on a copy in `_work/src/`)",
         f"- Worksheet mapped: `{B.sheet}`",
         f"- Template: `{TPL}` (read-only); component Type_Keys from `{XTPL.name}` > Component_Types",
         f"- Output JSON: `{well}.json` — {len(cs)} components; surveys / cement / plugBackCement / centralizers / floatEquipment / formations / "
         "formationTops / fluidContacts left empty (no supported source data)",
         f"- Generated: {dt.datetime.now():%Y-%m-%d %H:%M}", "",
         "## Depth units, reference and conversion", "", ref_note, "",
         "- Every depth in the JSON is in **ft = m / 0.3048**, rounded to 3 decimals. Metres, unit evidence and cell are listed below.",
         "- MD reference kept as printed (MDDF). No MD↔TVD or datum conversion was made. TVD values have no component field and are listed as unmapped.", "",
         md_table([dict(field=c["field"], cell=c["cell"], m=c["m"], unit=c["unit"], ft=c["ft"], kind=c["kind"]) for c in tr.conv],
                  ["field", "cell", "m", "unit", "ft", "kind"]), "",
         "## wellInfo mapping", "",
         md_table([r for r in tr.rows if r["json_path"].startswith("wellInfo")], ["json_path", "mapped_value", "cell", "original_value", "note"]), "",
         "## Components", "",
         md_table([dict(id=c["id"], type=c["type"], name=c["name"], topMd_ft=c["topMd"], bottomMd_ft=c["bottomMd"], OD_in=c["outerDiameter"],
                        ID_in=c["innerDiameter"], weight=c.get("weight"), grade=c.get("grade"), interval=c.get("interval"),
                        estimated=",".join(c.get("estimatedFields", []))) for c in cs],
                  ["id", "type", "name", "topMd_ft", "bottomMd_ft", "OD_in", "ID_in", "weight", "grade", "interval", "estimated"]), "",
         "Full cell-by-cell traceability (every mapped field, original value, unit, interpretation): "
         f"`{well}_traceability.csv`; each component's `source` string in the JSON names the cells too.", "",
         "## Terminology → template type", "",
         md_table(type_rows, ["source term", "cell(s)", "template type", "basis"]), "",
         "## Items requiring review", ""]
    for cat in ("Identity", "Conflict", "Missing", "Unclear", "Anomaly"):
        items = [r for r in review if r[0] == cat]
        if items:
            L += [f"### {cat} ({len(items)})", "", md_table([dict(location=r[1], detail=r[2]) for r in items], ["location", "detail"]), ""]
    L += ["## Unmapped information", "", "Relevant source content with no suitable template field (or not mapped because its meaning/association is not "
          "established by the source):", "", md_table([dict(location=u[0], content=u[1], reason=u[2]) for u in unmapped], ["location", "content", "reason"]), "",
          "## Quality checks", "", md_table([dict(check=c, result="PASS" if ok else "FAIL") for c, ok in checks], ["check", "result"]), "",
          "## Source / template integrity", "", md_table(integrity, ["file", "sha256 before", "sha256 after", "unchanged"]), ""]
    path.write_text("\n".join(L), encoding="utf-8")


# ------------------------------------------------------------------ per-well review / unmapped content
def q(B, ref):
    return f"{ref} {B.v(ref)!r}"


def notes_a17(B):
    review = [
        ("Conflict", "AB15; S71; S68", f"Lower tubing bottom not mapped: header {q(B,'AB15')} formula {B.formula('AB15')} is a LENGTH ending at S71 "
         f"(VCH packer, {B.num('S71')} m), while the string physically ends at the mule shoe {q(B,'S68')}. Two candidates -> left empty."),
        ("Conflict", "AB15 vs AB8", f"Header column AB8 'DEPTH' holds a length in AB15 ({B.v('AB15')}, formula {B.formula('AB15')}) -> not used as a depth."),
        ("Conflict", "B18 vs K20/S20", f"Note {q(B,'B18')} says depths are from RKB; the depth columns say 'm-MDDF' (drill floor). Kept as m-MDDF; datum not changed."),
        ("Conflict", "B76 vs B77", f"B76 formula {B.formula('B76')} uses 1115.9 psi, label {q(B,'B77')} says 1115.30 psi (both unmapped)."),
        ("Unclear", "AB13; H109", f"Liner bottom = header AB13 {B.v('AB13')} (formula {B.formula('AB13')} = TD). Diagram label {q(B,'H109')} has no depth. Mapped, review."),
        ("Unclear", "AB9-AB11, AB13, AB14", CASING_DEPTH + ". Note B18 says depths are 'to top of equipment'; for casings the DEPTH column is read as the shoe depth."),
        ("Unclear", "kbElevation G14", f"{q(B,'G14')} (DRILL FLOOR HEIGHT) mapped to kbElevation (closest field) and converted to ft — confirm DF = KB for this rig."),
        ("Missing", "casing tops", "Top depths of conductor / intermediate casing / 9-5/8\" casing not printed -> topMd null."),
        ("Missing", "U21; AB14", "Tubing hanger depth and the upper tubing top are not printed."),
        ("Missing", "groundElevation, surfaceCasing", "No ground elevation; no row labelled 'surface casing' (rows are CONDUCTOR / INT. CASING / CASING) -> template empty values."),
        ("Anomaly", "S67/S68 vs S71", f"Depth order: no-go locator {q(B,'S67')} and mule shoe {q(B,'S68')} are deeper than the VCH packer {q(B,'S71')} listed below them "
         "(typical when the locator lands inside the packer). Kept as printed."),
        ("Anomaly", "S98 vs S99", f"{q(B,'S99')} (sump packer) is shallower than the seal assembly above it {q(B,'S98')}. Kept as printed."),
    ]
    unmapped = [
        ("B3", B.txt("B3"), "slot number — no template field"), ("G13", f"WATER DEPTH {B.v('G13')}", "no template field (groundElevation not stated)"),
        ("B18", B.txt("B18"), "depth convention note (see conflict)"),
        ("Z9-Z15", "threads: " + ", ".join(f"{r} {B.txt(r)}" for r in ("Z9", "Z10", "Z11", "Z13", "Z14", "Z15")), "no thread field in template"),
        ("F28", B.txt("F28"), "tie-back liner label without depth or size data"),
        ("U67 / S67 / Z67 / AB67", f"{B.txt('U67')} at {B.v('S67')} m-MDDF ({ft(B.num('S67'))} ft), ID {B.v('Z67')} in, OD {B.v('AB67')} in",
         "no-go locator: no Component_Type with this meaning (ST Slot Locator / No-Go Nipple are different equipment)"),
        ("U45, U47, U49, U52", ", ".join(f"{r} {str(B.v(r))[:10]}" for r in ("U45", "U47", "U49", "U52")), "dates printed under SPM rows (meaning not stated) — no field"),
        ("U54-U56", f"{B.txt('U54')} / {B.txt('U55')} / {str(B.v('U56'))[:10]}", "well-intervention note; HUD 1830 m is m-MDTHF (different datum) — not converted to MDDF"),
        ("B73-B77", f"Zone {B.txt('B73')}: {B.txt('B75')} ({ft(1944.0)}-{ft(1951.0)} ft MDDF); {B.txt('B74')}; B76 {B.v('B76'):.3f} (EMW formula); {B.txt('B77')}",
         "reservoir zone interval: not stated as perforation/formation top; association with the screens is by position only -> not mapped"),
        ("B91-B95", f"Zone {B.txt('B91')}: {B.txt('B93')} ({ft(1981.0)}-{ft(1988.0)} ft MDDF); {B.txt('B92')}; B94 {B.v('B94'):.3f} (EMW formula); {B.txt('B95')}",
         "as above"),
        ("J75, J93, J109", f"deviation {B.v('J75')} / {B.v('J93')} / {B.v('J109')} deg (column J 'Deviation Degree')",
         "no MD/azimuth pair -> cannot form survey stations"),
        ("B84", B.txt("B84"), "no stated start/end"),
        ("K72, K78, K90, K96", ", ".join(f"{r} {B.v(r)}" for r in ("K72", "K78", "K90", "K96")) + " m-MDDF", "depths with no label in the row -> meaning unknown"),
        ("U110-Z112", f"{B.txt('U110')} {B.txt('Z110')}; {B.txt('Z111')}; {B.txt('U112')}; {B.txt('Z112')}", "document control — no template field"),
        ("drawing", "packer picture (xl/media image) anchored beside the AHR packer row", "image only, no data"),
    ]
    return review, unmapped


def notes_a23(B):
    review = [
        ("Identity", "file name / B2 / sheet title", f"File says A23ST2, {q(B,'B2')}, sheet title {B.sheet!r} says 'A23St1'. "
         "Output named Dulang_A23ST2 per the requested folder — confirm which wellbore this schematic is."),
        ("Conflict", "Y15 vs F34 / U32", f"Lower tubing grade: header {q(B,'Y15')} vs diagram labels {q(B,'F34')} and {q(B,'U32')} -> grade left empty on both lower tubing components."),
        ("Conflict", "AB13 / AB14 / AB15", f"Header 'DEPTH' column holds lengths/sums: AB13 {B.v('AB13')} ({B.formula('AB13')}), AB14 {B.v('AB14'):.2f} ({B.formula('AB14')}), "
         f"AB15 {B.v('AB15')} ({B.formula('AB15')}) -> none used as depths."),
        ("Conflict", "G17 vs U144", f"{q(B,'G17')} vs {q(B,'U144')} — revision date differs."),
        ("Conflict", "B18 vs K20/S20", f"Note {q(B,'B18')} (RKB) vs depth columns 'm-MDDF'. Kept as m-MDDF."),
        ("Unclear", "G12", f"maxDeviation {B.v('G12')} is a typed constant ({B.formula('G12')}); the deviation column shows at most "
         f"{max(B.num(r) for r in ('J95', 'J121', 'J140'))} (J95/J121/J140). Mapped as printed in G12."),
        ("Unclear", "V9", f"Conductor size {q(B,'V9')}: meaning of 'DP' not stated; OD 26 in mapped."),
        ("Unclear", "U60 / U61", f"{q(B,'U61')} below the dual packer may be part of the packer description or a separate 3 ft item -> not mapped."),
        ("Unclear", "kbElevation G14", f"{q(B,'G14')} (DRILL FLOOR HEIGHT) mapped to kbElevation (closest field), converted to ft."),
        ("Unclear", "short/long tubing depths", "Tubing intervals derived from the AB15 length formula and the 'From Tubing Hanger to TRSV' labels (flagged estimated)."),
        ("Missing", "AB9, AB11, AB12", "Casing depths printed as placeholder 'xx m MDDF' -> null."),
        ("Missing", "P10", "SURF. CASING row has no size/weight/grade/depth -> no component; wellInfo.surfaceCasing '-'."),
        ("Missing", "liner", "No liner top (TOL) or shoe depth printed -> null."),
        ("Missing", "groundElevation", "Not in source -> null."),
        ("Anomaly", "S103 vs S104", f"{q(B,'S104')} shallower than {q(B,'S103')} listed above it. Kept as printed."),
        ("Anomaly", "S127 vs S129", f"Sump packer {q(B,'S129')} shallower than collet seal {q(B,'S127')} above it. Kept as printed."),
        ("Anomaly", "D71", f"Ball catcher MIN ID {q(B,'D71')} (0.000)."),
    ]
    unmapped = [
        ("B3", B.txt("B3"), "slot number"), ("G13", f"WATER DEPTH {B.v('G13')}", "no template field"),
        ("B20 / U20", f"{B.txt('B19')} {B.txt('B20')}; {B.txt('U19')} {B.txt('U20')}", "well function — no template field"),
        ("Z11-Z15", "threads: " + ", ".join(f"{r} {B.txt(r)}" for r in ("Z11", "Z12", "Z13", "Z14", "Z15")), "no thread field"),
        ("F41, F44, F47, F50", B.txt("F41") + " (same text under each short-string SPM)", "valve status — no field"),
        ("U40, U43, U46", ", ".join(B.txt(r) for r in ("U40", "U43", "U46")), "gas-lift valve TRO settings — no field"),
        ("I56 / K56", f"{B.txt('I56')} at {B.v('K56')} m-MDDF ({ft(B.num('K56'))} ft)", "depth marker — no Component_Type"),
        ("F71 / K71 / B71 / D71", f"{B.txt('F71')} at {B.v('K71')} m-MDDF ({ft(B.num('K71'))} ft), OD {B.v('B71')} in, ID {B.v('D71')} in", "no Component_Type for a ball catcher"),
        ("U82 / S82", f"{B.txt('U82')} at {B.v('S82')} m-MDDF ({ft(B.num('S82'))} ft), ID {B.v('Z82')} in, OD {B.v('AB82')} in", "no-go locator: no matching Component_Type"),
        ("U61", B.txt("U61"), "see Unclear"),
        ("drawing K79:O81", "text box 'CLOSED (Nov 2024)' beside the long-string sliding sleeve U80 (row 80)", "status associated by position only"),
        ("B93-B97", f"Zone {B.txt('B93')}: {B.txt('B95')} ({ft(1416.0)}-{ft(1423.0)} ft MDDF); {B.txt('B94')}; B96 {B.v('B96'):.3f} (EMW formula); {B.txt('B97')}; J95 {B.v('J95')} deg",
         "reservoir zone interval: not stated as perforation/formation top; association with screens by position only"),
        ("B119-B123", f"Zone {B.txt('B119')}: {B.txt('B121')} ({ft(1517.0)}-{ft(1524.0)} ft MDDF); {B.txt('B120')}; B122 {B.v('B122'):.3f} (EMW formula); {B.txt('B123')}; J121 {B.v('J121')} deg",
         "as above"),
        ("K92, K98, K118, K124", ", ".join(f"{r} {B.v(r)}" for r in ("K92", "K98", "K118", "K124")) + " m-MDDF", "depths with no label in the row"),
        ("I139", B.txt("I139"), "casing shoe label without a depth"),
        ("J140", f"deviation {B.v('J140')} deg at TD row", "no azimuth -> no survey station"),
        ("G17, U142, U144", f"{B.txt('G17')}; {B.txt('U142')}; {B.txt('U144')}", "document control"),
    ]
    return review, unmapped


def notes_a33(B, tb, sheets):
    review = [
        ("Identity", "B2 vs B4", f"{q(B,'B2')} vs {q(B,'B4')} — wellbore suffix differs (ST3 vs ST2). Output named Dulang_A33ST3 (title + file name); confirm."),
        ("Identity", "sheet 'OCTG' C1", f"Sheet OCTG cell C1 = {sheets['OCTG']['C1']!r} — refers to another well; OCTG sheets not used."),
        ("Conflict", "I12 vs O52", f"Header {q(B,'I12')} names an AHC packer; the diagram packer is {B.txt('O52')[:40]!r}... Recorded; tubing grade/weight still from row 12."),
        ("Conflict", "B14 vs I13", f"{q(B,'B14')} vs {q(B,'I13')} — I13 mapped to completionFluid (more detailed), B14 listed as unmapped."),
        ("Unclear", "kbElevation B9", f"{q(B,'B9')} mapped to kbElevation (closest field), converted to ft."),
        ("Unclear", "maxDeviation B11", f"{q(B,'B11')} — kept as text with the depth converted to ft."),
        ("Unclear", "completionType", "'Single string' from wellhead text B12 '…Single' and one tubing string on the diagram — confirm."),
        ("Missing", "S6-S9", "Casing TOP printed as 'Surface' (no number); not converted to 0 because MDDF is measured from the drill floor -> topMd null."),
        ("Missing", "row 10", "7\" liner header weight / grade / thread / top / bottom all 'TBC'; only the diagram TOL depth F40 is mapped."),
        ("Missing", "O48, O59", f"Tubing intervals not printed. Candidates (not mapped): GRE tubing between the TRSV assembly (K25 {B.v('K25')}) and packer assembly "
         f"(K52 {B.v('K52')}); 15Cr tubing below the packer to the muleshoe (K61 {B.v('K61')})."),
        ("Missing", "O18", f"Tubing hanger depth not printed; row above says {B.txt('O17')} = {B.v('K17')} m (THF depth below DF?) — not mapped."),
        ("Missing", "T18", "Tubing hanger MAX OD 'TBC'."),
        ("Missing", "groundElevation, surfaceCasing", "Not in source -> template empty values."),
        ("Unclear", "B67", f"B67 = {B.v('B67')} (no unit; .xls formula not readable) — unmapped."),
    ]
    unmapped = [
        ("B4 / B5", f"{B.txt('B4')}; {B.txt('B5')}", "slot / field — see identity"),
        ("B7, B8", f"{B.txt('B7')}; {B.txt('B8')}", "injection design data — no field"), ("B14", B.txt("B14"), "see conflict"),
        ("R6-R12", "threads: " + ", ".join(f"{r} {B.txt(r)}" for r in ("R6", "R7", "R8", "R9", "R11", "R12")), "no thread field"),
        ("O17 / K17 / M17", f"{B.txt('O17')} {B.v('K17')} m", "elevation, not equipment"),
        ("N52", f"angle {B.v('N52')} deg at packer K52 {B.v('K52')} m-MDDF ({ft(B.num('K52'))} ft)", "no azimuth -> no survey station"),
        ("B11", f"{B.txt('B11')}", "also only one survey value, no azimuth"),
        ("O62", B.txt("O62"), "no depth, no Component_Type"),
        ("text box", f"{tb!r} (1858 m WLTHF = {ft(1858.0)} ft WLTHF; wireline datum, not converted to MDDF)", "intervention note; sheet anchor not resolvable from the .xls"),
        ("B63-B66", f"Zone {B.txt('B63')}: {B.txt('B65')} ({ft(1846.0)}-{ft(1856.0)} ft MDDF); {B.txt('B64')}; {B.txt('B66')}",
         "reservoir zone interval: not stated as perforation/formation top"),
        ("B67", f"{B.v('B67')}", "no unit / label"), ("O75", B.txt("O75"), "document control"),
        ("other sheets", "OCTG, OCTG Movement, Summary OCTG Movement (tubular inventory/procurement); '   ' (symbol legend); 'Sheet1', 'XX2R&' (very hidden) empty",
         "not schematic data"),
    ]
    return review, unmapped


# ------------------------------------------------------------------ main
def main():
    py = sys.executable
    tpl_before = sha(TPL); xtpl_before = sha(XTPL)
    tpl = json.loads(subprocess.run(["textutil", "-convert", "txt", "-stdout", str(TPL)], capture_output=True, text=True, check=True).stdout)
    keymap = {r[0]: r[2] for r in load_workbook(XTPL, read_only=True)["Component_Types"].iter_rows(min_row=2, values_only=True)}
    tpl_comp_keys = {k for c in tpl["components"] for k in c}
    allok = True

    for well in ("A17ST2", "A23ST2", "A33ST3"):
        out = ROOT / f"output_{well}"
        name = f"Dulang_{well}"
        out.mkdir(exist_ok=True)
        if any(p for p in out.iterdir() if not p.name.startswith(".")):
            raise SystemExit(f"{out} is not empty — not overwriting")
        work = out / "_work"; (work / "src").mkdir(parents=True)
        nas = NAS / SOURCES[well]
        h_before = sha(nas)
        copy = work / "src" / SOURCES[well]
        shutil.copyfile(nas, copy)
        assert sha(copy) == h_before

        if well == "A33ST3":
            d = dump_xls(copy)
            (work / "dump_cells.json").write_text(json.dumps(d, indent=1, ensure_ascii=False))
            sh = {s["title"]: s for s in d["sheets"]}
            cells = sh["SCHEM-1"]["cells"]
            assert "K17:L17" in sh["SCHEM-1"]["merged"] and "K25:L25" in sh["SCHEM-1"]["merged"]
            B = Book(SOURCES[well], "SCHEM-1", cells, {"K": "L16", "F": "F16", "T": "S4"}, {"S": "S16", "T": "T16"})
            # text boxes: same Workbook stream as the copy already parsed in output_A/run_xls_... -> reuse that text-box dump
            assert any(workbook_stream(copy) == workbook_stream(p) for p in (PREV / "src").glob("A33ST3*.xls"))
            tb = json.loads((PREV / "dump_A33ST3_textboxes.json").read_text())[0]["text"]
            assert tb.startswith('2" LIB tagged HUD at 1858m WLTHF')
            other = {s["title"]: {c["ref"]: c["value"] for c in s["cells"]} for s in d["sheets"]}
        else:
            subprocess.run([py, str(ROOT / "scripts" / "dump_xlsx_schematic.py"), str(copy), str(work / "dump.json")], check=True, capture_output=True)
            d = json.loads((work / "dump.json").read_text())
            assert len(d["sheets"]) == 1 and d["sheets"][0]["state"] == "visible"
            s0 = d["sheets"][0]
            B = Book(SOURCES[well], s0["title"], s0["cells"], {"K": "K20", "S": "S20"},
                     {"B": "B20" if well == "A17ST2" else "B23", "D": "D20" if well == "A17ST2" else "D23",
                      "Z": "Z20" if well == "A17ST2" else "Z23", "AB": "AB20" if well == "A17ST2" else "AB23"})
            assert "K19:L19" in s0["merged"] and "S19:T19" in s0["merged"]

        tr = Trace(name, B)
        if well == "A17ST2":
            assert B.txt("B2") == "WELL NAME: DULANG A17-ST2 OP5"
            specs = spec_a17(); review, unmapped = notes_a17(B)
            assert B.formula("AB14") == "=S30" and B.formula("AB13") == "=K110" and B.formula("G10") == "=K110" and B.formula("G11") == "=K108"
            info = dict(
                completionDate=("2020-05-30", ["G8"], "DIRECT (date cell, label B8 'DATE OF COMPLETION'), written ISO yyyy-mm-dd"),
                completionFluid=(B.txt("V16"), ["V16"], "DIRECT (label P16)"),
                completionType=("Single string", ["U21"], "CLOSEST: U21 'SINGLE STRING TUBING HANGER' and sheet title 'OP5 single …'"),
                kbElevation=(ft(33.0), ["G14"], "CLOSEST FIELD: DRILL FLOOR HEIGHT '33 m-AMSL' -> 33 / 0.3048 ft"),
                maxDeviation=(f"{B.num('G12'):g} deg", ["G12", "J20"], "DIRECT (formula =MAX(J24:J142); unit J20 'Degree')"),
                name=(name, ["B2"], "DIRECT: 'WELL NAME: DULANG A17-ST2 OP5' written in the output naming form"),
                packerFluid=(B.txt("V17") + " " + B.txt("V18"), ["V17", "V18"], "DIRECT (two lines joined)"),
                rig=(B.txt("G9"), ["G9"], "DIRECT"),
                totalDepth=(ft(B.num("G10")), ["G10"], f"DIRECT: TD {B.num('G10')} m-MDDF (formula =K110; K20 'm-MDDF') -> ft"),
                trajectoryType=("Deviated", ["G12"], "supported by maximum deviation 80.3 deg"),
                tubing=('3-1/2" 9.2 PPF HP2-13CR-95 JFEBEAR; 3-1/2" 9.2 PPF 13CR-L80 JFEBEAR', ["V14", "X14", "Y14", "Z14", "V15", "X15", "Y15", "Z15"], "DIRECT: header rows 14-15 joined"),
                wellhead=(B.txt("G15"), ["G15"], "DIRECT"))
            assert B.txt("G14") == "33 m-AMSL" and B.v("G8") == "2020-05-30T00:00:00"
            ref_note = (f"Depth columns K19/K20 and S19/S20 'DEPTH m-MDDF' (merged K:L, S:T); casing table AB9-AB11 carry 'm MDDF' in the text. "
                        f"Note {B.txt('B18')!r} — see conflict.")
        elif well == "A23ST2":
            assert B.txt("B2") == "WELL NAME: DULANG A23 OP2"
            specs = spec_a23(); review, unmapped = notes_a23(B)
            assert all(B.txt(r) == "xx m MDDF" for r in ("AB9", "AB11", "AB12")) and all(B.txt(r) == "From Tubing Hanger to TRSV" for r in ("F27", "U27"))
            assert B.formula("AB15") == "=ROUNDUP((K73-K32)+(S119-S30),0)" and B.txt("F73") == "EOT"
            info = dict(
                completionDate=("2020-03-17", ["G8"], "DIRECT (label B8), ISO"),
                completionFluid=(B.txt("V16"), ["V16"], "DIRECT"),
                completionType=("Dual string", ["B19", "U19", "U60"], "CLOSEST: SHORT STRING / LONG STRING columns and 'DUAL STRING' packer"),
                kbElevation=(ft(33.0), ["G14"], "CLOSEST FIELD: DRILL FLOOR HEIGHT '33 m-AMSL' -> ft"),
                maxDeviation=(f"{B.num('G12'):g} deg", ["G12"], "DIRECT (G12 typed value; unit from label 'MAXIMUM DEVIATION' / column J 'Degree')"),
                name=(name, ["B2"], "as requested (output_A23ST2); B2 says 'DULANG A23 OP2' — see identity item"),
                packerFluid=(B.txt("V17") + " " + B.txt("V18"), ["V17", "V18"], "DIRECT (two lines joined)"),
                rig=(B.txt("G9"), ["G9"], "DIRECT"),
                totalDepth=(ft(B.num("G10")), ["G10"], f"DIRECT: TD {B.num('G10')} m-MDDF (formula =K140) -> ft"),
                trajectoryType=("Deviated", ["G12"], "supported by maximum deviation 44.69 deg"),
                tubing=('3-1/2" 9.2 PPF HP2-13CR-95 JFEBEAR; 3-1/2" 9.2 PPF 13CR-L80 JFEBEAR',
                        ["V14", "X14", "Y14", "Z14", "V15", "X15", "Y15", "Z15"], "DIRECT: header rows 14-15 joined (row 15 grade conflicts with diagram labels — see report)"),
                wellhead=(B.txt("G15"), ["G15"], "DIRECT"))
            assert B.txt("G14") == "33 m-AMSL" and B.v("G8") == "2020-03-17T00:00:00" and B.formula("G10") == "=K140"
            ref_note = f"Depth columns K19/K20 and S19/S20 'DEPTH m-MDDF'. Note {B.txt('B18')!r} — see conflict."
        else:
            assert B.txt("B2") == "DULANG A33-ST3"
            specs = spec_a33(); review, unmapped = notes_a33(B, tb, other)
            assert all(B.txt(f"S{r}") == "Surface" for r in (6, 7, 8, 9)) and B.txt("S4") == "DEPTH (m-MDDF)" and B.txt("T5") == "BOTTOM"
            assert B.txt("L16") == "m-MDDF" and B.txt("F16") == "m-MDDF" and B.txt("M16") == "m" and B.txt("P4") == "#"
            assert B.txt("B11") == "Maximum Deviation: 58deg @ 1,800m MDDF" and B.txt("B10") == "Completion Date : 3rd April 2020"
            strip = lambda r: B.txt(r).split(":", 1)[1].strip()
            info = dict(
                completionDate=("2020-04-03", ["B10"], "DIRECT: '3rd April 2020' written ISO"),
                completionFluid=(strip("I13"), ["I13"], "DIRECT (label removed)"),
                completionType=("Single string", ["B12"], "CLOSEST: wellhead text '…Single' and one tubing string — review"),
                kbElevation=(ft(30.1), ["B9"], "CLOSEST FIELD: 'Drill floor height: 30.1m-AMSL' -> ft"),
                maxDeviation=(f"58 deg @ {ft(1800.0)} ft MDDF", ["B11"], "DIRECT text; depth 1,800 m MDDF -> ft"),
                name=(name, ["B2"], "DIRECT: B2 'DULANG A33-ST3' (B4 says A33ST2 — see identity)"),
                packerFluid=(strip("I14"), ["I14"], "DIRECT (label removed)"),
                rig=(strip("B6"), ["B6"], "DIRECT"),
                totalDepth=(ft(B.num("F75")), ["F75"], f"DIRECT: TD (label B75) {B.num('F75')} m-MDDF (F16) -> ft"),
                trajectoryType=("Deviated", ["B11"], "supported by maximum deviation 58 deg"),
                tubing=('Tubing 3-1/2" GRE lined 9.2 L80 JFEFOX; Tubing 3-1/2" - Below 7" AHC 9.2 15CR-125 JFE BEAR',
                        ["I11", "P11", "Q11", "R11", "I12", "P12", "Q12", "R12"], "DIRECT: header rows 11-12 joined"),
                wellhead=(strip("B12"), ["B12"], "DIRECT (label removed)"))
            ref_note = "Depths: casing scheme S4 'DEPTH (m-MDDF)' (TOP/BOTTOM); equipment depths in K:L under L15/L16 'DEPTH m-MDDF'; diagram depths F40/F74/F75 under F15/F16 'DEPTH m-MDDF'."
            tr.conv.append(dict(field="wellInfo.maxDeviation (depth in text)", cell="B11", m=1800.0, unit="m MDDF (in text)", ft=ft(1800.0), kind="direct"))

        comps, flags = build_components(B, specs, keymap, tr)
        wi = well_info(tpl, info, tr)
        for k, (v, refs, note) in info.items():
            if k in ("totalDepth", "kbElevation"):
                m = {"kbElevation": {"A17ST2": 33.0, "A23ST2": 33.0, "A33ST3": 30.1}[well]}.get(k) or (B.num(refs[0]))
                tr.conv.append(dict(field=f"wellInfo.{k}", cell=refs[0], m=m, unit="m-AMSL (elevation)" if k == "kbElevation" else B.depth_unit(refs[0]) if well != "A33ST3" else "m-MDDF (F16)",
                                    ft=v, kind="direct"))
        review = review + flags
        doc = {k: {"wellInfo": wi, "components": comps}.get(k, []) for k in tpl}

        jpath = out / f"{name}.json"
        jpath.write_text(json.dumps(doc, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")

        # ---------------- checks (independent re-read of the copy)
        back = json.loads(jpath.read_text())
        cs = back["components"]
        if well == "A33ST3":
            s = xlrd.open_workbook(copy).sheet_by_name("SCHEM-1")
            def raw(ref):
                m = re.fullmatch(r"([A-Z]+)(\d+)", ref)
                col = sum((ord(ch) - 64) * 26 ** i for i, ch in enumerate(reversed(m.group(1)))) - 1
                return s.cell_value(int(m.group(2)) - 1, col)
        else:
            ws = load_workbook(copy, data_only=True).worksheets[0]
            raw = lambda ref: ws[ref].value
        num = lambda v: float(re.search(r"[\d,]+(?:\.\d+)?", v).group(0).replace(",", "")) if isinstance(v, str) else float(v)
        kind = lambda v: "number" if isinstance(v, (int, float)) and not isinstance(v, bool) else type(v).__name__
        depth_ok = all(c["ft"] == round(num(raw(c["cell"])) / M_PER_FT, 3) for c in tr.conv if not c["field"].startswith("wellInfo.maxDev"))
        json_depths = {f"{c['id']}.{f}": c[f] for c in cs for f in ("topMd", "bottomMd")}
        names_ok = all(r["mapped_value"] == " ".join(str(raw(x)).strip() for x in r["cell"].split("+")) for r in tr.rows if r["json_path"].startswith("components") and r["json_path"].endswith(".name"))
        checks = [
            ("valid JSON (re-parsed)", True),
            ("top-level sections and order = template", list(back) == list(tpl)),
            ("wellInfo keys and order = template (no extra keys)", list(back["wellInfo"]) == list(tpl["wellInfo"])),
            ("wellInfo types = template type, or template empty value ('-' / null)",
             all(v is None or v == "-" or kind(v) == kind(tpl["wellInfo"][k]) for k, v in back["wellInfo"].items())),
            ("component keys are template component keys only", all(set(c) <= tpl_comp_keys for c in cs)),
            ("component value types (numbers/null for depths & diameters, text for names, bool flag)",
             all(all(c[f] is None or kind(c[f]) == "number" for f in ("topMd", "bottomMd", "outerDiameter", "innerDiameter")) and isinstance(c["name"], str)
                 and isinstance(c["type"], str) and c.get("estimated", True) is True and set(c.get("estimatedFields", [])) <= set(c) for c in cs)),
            ("every type is a Component_Types Type_Key", all(c["type"] in set(keymap.values()) for c in cs)),
            ("component ids unique", len({c["id"] for c in cs}) == len(cs)),
            ("no template sample records copied", not {c["id"] for c in cs} & {t["id"] for t in tpl["components"]} and back["surveys"] == [] and back["formationTops"] == []),
            ("every depth: ft = source m / 0.3048 (re-read from the copy)", depth_ok),
            ("JSON depths = converted values in the conversion table",
             all(json_depths[c["field"]] == c["ft"] for c in tr.conv if c["field"] in json_depths)
             and all(v is None or any(c["field"] == k for c in tr.conv) for k, v in json_depths.items())),
            ("every depth has unit evidence (m-MDDF header or text)", all("MDDF" in c["unit"] or "AMSL" in c["unit"] for c in tr.conv)),
            ("component names = source cell text", names_ok),
            ("derived / closest-type values flagged estimated", all(("type" in c.get("estimatedFields", [])) == bool(sp["typ"]) for c, sp in zip(cs, specs))),
            ("every component source names workbook > sheet > cells", all(c["source"].startswith(f"{B.workbook} > sheet '{B.sheet}' > name ") for c in cs)),
            ("well identity: B2 names this well", {"A17ST2": "A17", "A23ST2": "A23", "A33ST3": "A33"}[well] in B.txt("B2").replace(" ", "").replace("-", "")),
        ]
        h_after = sha(nas)
        integrity = [dict(file=str(nas), **{"sha256 before": h_before, "sha256 after": h_after, "unchanged": "YES" if h_before == h_after else "NO"}),
                     dict(file=str(TPL), **{"sha256 before": tpl_before, "sha256 after": sha(TPL), "unchanged": "YES" if tpl_before == sha(TPL) else "NO"}),
                     dict(file=str(XTPL), **{"sha256 before": xtpl_before, "sha256 after": sha(XTPL), "unchanged": "YES" if xtpl_before == sha(XTPL) else "NO"})]
        checks.append(("source workbook and template unchanged (SHA-256)", all(i["unchanged"] == "YES" for i in integrity)))
        (work / "sha256.txt").write_text("\n".join(f"{i['sha256 before']}  {i['sha256 after']}  {i['file']}" for i in integrity) + "\n")

        type_rows = []
        for sp, c in zip(specs, cs):
            type_rows.append({"source term": c["name"][:70], "cell(s)": sp["name"] if isinstance(sp["name"], str) else "+".join(sp["name"]),
                              "template type": c["type"], "basis": ("CLOSEST — " + sp["typ"]) if sp["typ"] else "matches source term"})
        write_report(out / f"{name}_mapping_report.md", name, B, back, tr, review, unmapped, checks, integrity, ref_note, type_rows)
        with open(out / f"{name}_traceability.csv", "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=list(tr.rows[0]))
            w.writeheader(); w.writerows(tr.rows)

        print(f"=== {name}: {len(cs)} components, {sum(1 for c in cs if c.get('estimated'))} flagged estimated, "
              f"{len(review)} review items, {len(unmapped)} unmapped entries")
        for c, ok in checks:
            allok &= ok
            print("  PASS" if ok else "  FAIL", "-", c)
    print("ALL PASS" if allok else "SOME CHECKS FAILED")


if __name__ == "__main__":
    main()
