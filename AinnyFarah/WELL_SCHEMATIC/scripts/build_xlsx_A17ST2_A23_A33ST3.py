#!/usr/bin/env python3
"""Map three Dulang A well-schematic workbooks (read-only; working copies in output_A/run_xls_A17ST2_A23ST2_A33ST3/src)
into the JSON template, one JSON per well.

Inputs are the dumps made by dump_xlsx_schematic.py (.xlsx: cells + formulas + cached values + drawing text) and by the
xlrd/BIFF readers (.xls: cached cell values, text box). Every value is read by cell reference and asserted to be what the
mapping expects, so each component's 'source' names workbook > sheet > cell (formula + cached value where relevant).
Depths m-MDDF -> ft (x 3.28084, 3 dp), metres kept in *M keys; TVD only from explicit 'm TVD' text into topTvd/bottomTvd.
OD/ID stay in inches (as printed and as the template uses). Notes say DIRECT / CALCULATED / DERIVED / ASSUMED / APPROX-TYPE.
"""
import hashlib
import json
import os
import re
import subprocess
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "output_A"
RUN = OUT / "run_xls_A17ST2_A23ST2_A33ST3"
TPL = ROOT / "well-schematic-template-json.rtf"
XTPL = ROOT / "well-schematic-sample-template.xlsx"
FT = 3.28084
S, L, SL = "Short", "Long", "Short & Long"


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


class Book:
    def __init__(self, label, sheet, cells, formulas_readable=True):
        self.label, self.sheet, self.formulas_readable = label, sheet, formulas_readable
        self.c = {c["ref"]: c for c in cells}

    def v(self, ref):
        assert ref in self.c, f"{self.label}: {ref} is empty"
        return self.c[ref]["value"]

    def txt(self, ref):
        v = self.v(ref)
        assert isinstance(v, str), (ref, v)
        return v.strip()

    def num(self, ref):
        v = self.v(ref)
        if isinstance(v, str):
            m = re.fullmatch(r"\s*([\d,]+(?:\.\d+)?)\s*(?:m\b.*|PPF|\"|)\s*", v)
            assert m, (self.label, ref, v)
            return float(m.group(1).replace(",", ""))
        return float(v)

    def tr(self, ref):
        c = self.c[ref]
        s = f"{ref}={c['value']!r}"
        if "formula" in c:
            s += f" (formula {c['formula']}, cached value used)"
        elif not self.formulas_readable and isinstance(c["value"], float):
            s += " (cached value; .xls formulas not readable with xlrd)"
        return s


def rng(text):
    m = re.search(r"([\d,]+(?:\.\d+)?)\s*m?\s*-\s*([\d,]+(?:\.\d+)?)", text)
    assert m, text
    return float(m.group(1).replace(",", "")), float(m.group(2).replace(",", ""))


def load_books():
    a17 = json.loads((RUN / "dump_A17ST2.json").read_text())["sheets"][0]
    a23 = json.loads((RUN / "dump_A23ST2.json").read_text())["sheets"][0]
    a33 = [c for c in json.loads((RUN / "dump_A33ST3_cells.json").read_text()) if c["sheet"] == "SCHEM-1"]
    return (Book("Dulang A17ST2 Well Schematic updated .xlsx", a17["title"], a17["cells"]),
            Book("Dulang A23ST2 Well Schematic.xlsx", a23["title"], a23["cells"]),
            Book("Dulang A33ST3 Well Schematic.xls", "SCHEM-1", a33, formulas_readable=False))


# ---------------------------------------------------------------- component spec helpers
def comp(cid, ctype, name, top=None, bot=None, od=None, idd=None, string=None, extra=None, approx=None, est=None, note="", tvd=None):
    """name: ref or list of refs (joined); top/bot/od/idd: ref (read) or ('val', number, why) for derived/assumed values."""
    return dict(cid=cid, ctype=ctype, name=name, top=top, bot=bot, od=od, idd=idd, string=string, extra=extra or {},
                approx=approx, est=est, note=note, tvd=tvd)


def build_components(B, specs, keymap, ref_label):
    out = []
    for s in specs:
        refs = s["name"] if isinstance(s["name"], list) else [s["name"]]
        name = " ".join(B.txt(r) for r in refs)
        trace = [f"name {'+'.join(refs)} '{name}'"]
        vals = {}
        for k in ("top", "bot", "od", "idd"):
            x = s[k]
            if x is None:
                vals[k] = None
            elif isinstance(x, tuple):
                vals[k] = x[1]
                trace.append(f"{k} {x[1]} ({x[2]})")
            else:
                vals[k] = B.num(x)
                trace.append(f"{k} {B.tr(x)}")
        d = {"id": s["cid"], "type": keymap[s["ctype"]], "name": name, "topMd": ft(vals["top"]), "bottomMd": ft(vals["bot"]),
             "outerDiameter": vals["od"], "innerDiameter": vals["idd"]}
        for k, v in s["extra"].items():
            if isinstance(v, tuple) and v[0] == "ref":
                val = B.num(v[1]) if v[2] == "num" else B.txt(v[1])
                d[k] = val
                trace.append(f"{k} {B.tr(v[1])}")
            else:
                d[k] = v
        if s["tvd"]:
            t0, t1 = rng(B.txt(s["tvd"]))
            d.update(topTvd=ft(t0), bottomTvd=ft(t1))
            trace.append(f"TVD {B.tr(s['tvd'])} (explicit 'm TVD' text)")
        fields, notes = [], []
        if s["approx"]:
            fields.append("type"); notes.append(f"APPROX-TYPE: no exact Type_Key; {s['approx']}")
        if s["est"]:
            fields += s["est"][0]; notes.append(s["est"][1])
        if fields:
            d.update(estimated=True, estimatedFields=fields, estimationNote="; ".join(notes))
        if s["string"]:
            d["string"] = s["string"]
        conv = f"; depths {ref_label}; CALCULATED: ft = m x 3.28084 (3 dp)" if vals["top"] is not None or vals["bot"] is not None else ""
        conv += "; TVD m -> ft x 3.28084 (TVD kept separate from MD)" if s["tvd"] else ""
        d["source"] = f"{B.label} > sheet '{B.sheet}' > " + "; ".join(trace) + (f"; {s['note']}" if s["note"] else "") + conv
        d.update(actualType=name, topMdM=vals["top"], bottomMdM=vals["bot"])
        if s["tvd"]:
            d.update(topTvdM=t0, bottomTvdM=t1)
        out.append(d)
    return out


# ---------------------------------------------------------------- A17ST2 (single string, columns U name / S depth / Z ID / AB OD)
def spec_a17(B):
    hdr = "header casing/tubing table (P label, V size, X weight, Y grade, Z thread, AB depth)"
    w = lambda r: ("ref", r, "num")
    g = lambda r: ("ref", r, "txt")
    return [
        comp("conductor", "Casing", "P9", bot="AB9", od=("val", 26.0, 'V9 \'26"\' DIRECT'), extra={"weight": w("X9"), "grade": g("Y9")},
             note=f"{hdr}; thread Z9 'WELDED'; weight X9 printed without unit; top not printed"),
        comp("int-casing", "Casing", "P10", bot="AB10", od=("val", 13.375, "V10 '13-3/8\"' DIRECT (fraction converted to decimal inches)"),
             extra={"weight": w("X10"), "grade": g("Y10")}, note=f"{hdr}; thread Z10 'BTC'; top not printed"),
        comp("casing", "Casing", "P11", bot="AB11", od=("val", 9.625, "V11 '9-5/8\"' (decimal inches)"), extra={"weight": w("X11"), "grade": g("Y11")},
             note=f"{hdr}; thread Z11 'BTC'; top not printed"),
        comp("liner-7", "Liner", "P13", top="K34", bot="AB13", od=("val", 7.0, "V13 '7\"'"), extra={"weight": w("X13"), "grade": g("Y13")},
             est=(["bottomMd"], "Liner bottom from header AB13 whose formula =K110 points at the TD cell; the diagram label H109 '7\" LINER SHOE' "
                  "carries no depth — header value used as documented, flagged because it equals TD by formula"),
             note="top from diagram F34 ' TOL (7\" LINER)' / K34 (explicit TOL label); thread Z13 'VAM TOP/BEAR'; diagram also shows F28 'Tie Back 7\" Liner' "
                  "(no depth); J109 '58' printed beside the liner-shoe label in the Deviation column — meaning unclear, not used"),
        comp("tubing-upper", "Tubing", "U26", bot="AB14", od=("val", 3.5, "V14 '3-1/2\"'"), extra={"weight": w("X14"), "grade": g("Y14")},
             note="U25 'From Tubing Hanger to TRSV'; header row 14 (9.2 PPF HP2-13CR-95 JFEBEAR) DIRECT; bottom AB14 = formula =S30 (TRSV depth); "
                  "top (hanger) depth not printed"),
        comp("tubing-lower", "Tubing", ["V15", "X15", "Y15", "Z15"], top=("val", 250.738, "DERIVED: TRSV depth S30, where the upper tubing ends"),
             bot=("val", 1911.54, "DERIVED: SELF ALIGN MULE SHOE S68 (end of string)"), od=("val", 3.5, "V15 '3-1/2\"'"),
             extra={"weight": w("X15"), "grade": g("Y15")}, est=(["topMd", "bottomMd"], "DERIVED interval TRSV -> mule shoe; header AB15 1660 is a LENGTH "
                                                          "(formula =ROUNDUP(S71-S30,0)), not a depth, and is not used as one"),
             note="name = header row 15 cells joined (no diagram label for this section)"),
        comp("tubing-hanger", "Tubing Hanger", "U21", note="no depth printed"),
        comp("flow-coupling-1", "Flow Coupling", "U29", od="AB29", idd="Z29", note="no depth printed (S29 empty)"),
        comp("trsv", "TRSV", "U30", top="S30", od="AB30", idd="Z30", note="TRMAXX-5E is the TRSV named in U25 / header AB14 formula"),
        comp("flow-coupling-2", "Flow Coupling", "U31", od="AB31", idd="Z31", note="no depth printed"),
        comp("spm-1", "SPM (Gas Lift)", "U42", top="S42", od="AB42", idd="Z42"),
        comp("spm-dial-1", "SPM (Gas Lift)", "U44", top="S44", od="AB44", idd="Z44", approx="AUTO SPM DIAL -> spm",
             note="date line below it U45 = 2026-05-02 (displayed '2-May-26')"),
        comp("spm-2", "SPM (Gas Lift)", "U46", top="S46", od="AB46", idd="Z46", note="date line below it U47 = 2026-05-14 (displayed '14-May-26')"),
        comp("spm-dial-2", "SPM (Gas Lift)", "U48", top="S48", od="AB48", idd="Z48", approx="AUTO SPM DIAL -> spm",
             note="date line below it U49 = 2026-05-02 (displayed '2-May-26')"),
        comp("spm-3", "SPM (Gas Lift)", ["U50", "U51"], top="S50", od="AB50", idd="Z50",
             note="date U52 = 2026-05-14 ('14-May-26'); notes directly below (U54-U56, association by position and the 'SPM#3' label — counting SPMs "
                  "without the DIAL mandrels this is SPM#3; uncertain): 'Run MLT from HUD at 1830m-MDTHF' / 'Small leak (from SPM#3) observed by MLT "
                  "Condition#2 run' / U56 = 2026-05-01 ('1-May-26'); HUD depth is m-MDTHF (different reference), not mapped"),
        comp("ahr-packer", "Packer", "U59", top="S59", od="AB59", idd="Z59", note="packer picture (image2.png) anchored N58:Q61 beside this row"),
        comp("landing-nipple", "Landing Nipple", "U62", top="S62", od="AB62", idd="Z62"),
        comp("no-go-locator", "ST Slot Locator", "U67", top="S67", od="AB67", idd="Z67", approx="no-go locator -> st-slot-locator"),
        comp("mule-shoe", "Mule Shoe", "U68", top="S68", od="AB68", idd="Z68", note="end of the tubing string"),
        comp("vch-packer", "Packer", "U71", top="S71", od="AB71", idd="Z71", note="top of the lower completion"),
        comp("mcsz-1", "Sliding Sleeve", "U73", top="S73", od="AB73", idd="Z73", approx="MCSZ closing sleeve -> sliding-sleeve"),
        comp("screen-1-top", "Screen", "U75", top="S75", od="AB75", idd="Z75", extra={"zone": "E22A"}, note="zone E22A by position (label B73 beside it)"),
        comp("screen-1-prod", "Screen", "U77", top="S77", od="AB77", idd="Z77", extra={"zone": "E22A"}, note="zone E22A by position"),
        comp("fcs-1", "Circulating Valve", "U79", top="S79", od="AB79", idd="Z79", approx="FCS fracture / circulation sleeve -> circulating-valve"),
        comp("tnt-packer", "Packer", "U83", top="S83", od="AB83", idd="Z83", note="B84 'Distance 30 m' printed on the left near this row (no destination)"),
        comp("mcsz-2", "Sliding Sleeve", "U88", top="S88", od="AB88", idd="Z88", approx="MCSZ closing sleeve -> sliding-sleeve"),
        comp("screen-2-top", "Screen", "U90", top="S90", od="AB90", idd="Z90", extra={"zone": "E22B"}, note="zone E22B by position (label B91)"),
        comp("screen-2-prod", "Screen", "U92", top="S92", od="AB92", idd="Z92", extra={"zone": "E22B"}, note="zone E22B by position"),
        comp("fcs-2", "Circulating Valve", "U94", top="S94", od="AB94", idd="Z94", approx="FCS fracture / circulation sleeve -> circulating-valve"),
        comp("sump-seal", "Seal Assembly", "U98", top="S98", od="AB98", idd="Z98"),
        comp("sump-packer", "Packer", "U99", top="S99", od="AB99", idd="Z99", note="S99 1992 printed shallower than the seal assembly S98 1992.621 — as printed"),
        comp("pbtd", "PBTD", "I108", top="K108", note="header G11 = formula =K108, same value"),
        comp("zone-e22a", "Perforation", "B73", top=("val", 1944.0, "B75 '1944 - 1951 m MDDF (7 m)'"), bot=("val", 1951.0, "B75"),
             extra={"zone": "E22A", "interval": ("ref", "B75", "txt")}, tvd="B74",
             approx="reservoir zone interval shown beside the screens; not stated as perforations and no zone Type_Key — perforation is the nearest interval type",
             note="B76 = 5.011 (formula =1115.9/(0.0521*1302.63*3.281), an EMW in ppg; the formula uses 1115.9 psi while B77 says '1115.30 psi'); "
                  "B77 '1115.30 psi, 177.8 deg F'; J75 80.3 (Degree)"),
        comp("zone-e22b", "Perforation", "B91", top=("val", 1981.0, "B93 '1981 - 1988 m MDDF (7 m)'"), bot=("val", 1988.0, "B93"),
             extra={"zone": "E22B", "interval": ("ref", "B93", "txt")}, tvd="B92",
             approx="reservoir zone interval beside the screens; perforation is the nearest interval type",
             note="B94 = 4.871 (formula =1089.8/(0.0521*1308.89*3.281), EMW ppg); B95 '1089.80 psi, 177.8 deg F'; J93 80.1 (Degree)"),
    ]


# ---------------------------------------------------------------- A23 (dual: short F/K/D/B, long U/S/Z/AB)
def spec_a23(B):
    w = lambda r: ("ref", r, "num")
    g = lambda r: ("ref", r, "txt")
    return [
        comp("conductor", "Casing", "P9", od=("val", 26.0, "V9 '26\" DP'"), note="depth AB9 'xx m MDDF' = placeholder, treated as missing"),
        comp("int-casing", "Casing", "P11", od=("val", 13.375, "V11 '13-3/8\"'"), extra={"weight": w("X11"), "grade": g("Y11")},
             note="thread Z11 'BTC'; depth AB11 'xx m MDDF' = placeholder, missing"),
        comp("casing", "Casing", "P12", od=("val", 9.625, "V12 '9-5/8\"'"), extra={"weight": w("X12"), "grade": g("Y12")},
             note="thread Z12 'VAM TOP'; header depth AB12 'xx m MDDF' (missing); diagram label I139 '9 5/8\" Casing Shoe' has no depth"),
        comp("liner", "Liner", "P13", od=("val", 5.5, "V13 '5-1/2\"'"), extra={"weight": w("X13"), "grade": g("Y13")},
             note="thread Z13 'JFE BEAR'; header AB13 = 144 is a LENGTH (formula =ROUNDUP(S127-S85,0); displayed '144.00 m MDDF') — not used as a depth; "
                  "no TOL label and no liner shoe depth -> top/bottom empty"),
        comp("s-tubing-upper", "Tubing", "F26", bot="K32", od=("val", 3.5, "3-1/2\""), string=S,
             note="F27 'From Tubing Hanger to TRSV'; bottom = short-string TRSV depth K32 (DERIVED from that label); top not printed; "
                  "header AB14 478.71 = formula =K32+S30 (sum of both strings) — not a depth"),
        comp("s-tubing-lower", "Tubing", "F34", top=("val", 245.6, "DERIVED: short TRSV K32"), bot="K73", od=("val", 3.5, "3-1/2\""), string=S,
             est=(["topMd"], "DERIVED: top = short-string TRSV depth"),
             note="bottom = 'EOT' label F73 / K73 (explicit); CONFLICT: header row 15 gives grade '13CR-L80' but the diagram label F34 says 'HP2 13Cr95' — "
                  "diagram wording kept; header AB15 2413 = formula =ROUNDUP((K73-K32)+(S119-S30),0), a summed length, not a depth"),
        comp("l-tubing-upper", "Tubing", "U26", bot="S30", od=("val", 3.5, "3-1/2\""), string=L,
             note="U27 'From Tubing Hanger to TRSV'; bottom = long-string TRSV depth S30 (DERIVED from that label); top not printed"),
        comp("l-tubing-lower", "Tubing", "U32", top=("val", 233.11, "DERIVED: long TRSV S30"), bot=("val", 1521.9, "DERIVED: long-string mule shoe S119"),
             od=("val", 3.5, "3-1/2\""), string=L, est=(["topMd", "bottomMd"], "DERIVED interval TRSV -> mule shoe (no EOT label on the long string)"),
             note="same header/diagram grade conflict as the short string"),
        # short string
        comp("s-flow-coupling-1", "Flow Coupling", "F31", od="B31", idd="D31", string=S, note="no depth printed"),
        comp("s-trsv", "TRSV", "F32", top="K32", od="B32", idd="D32", string=S, note="TRMAXX-5E = the TRSV named in F27"),
        comp("s-flow-coupling-2", "Flow Coupling", "F33", od="B33", idd="D33", string=S, note="no depth printed"),
        comp("s-spm-1", "SPM (Gas Lift)", "F40", top="K40", od="B40", idd="D40", string=S, note="F41 'DV (in DROID, set on 21  Feb 2022)'"),
        comp("s-spm-2", "SPM (Gas Lift)", "F43", top="K43", od="B43", idd="D43", string=S, note="F44 'DV (in DROID, set on 21  Feb 2022)'"),
        comp("s-spm-3", "SPM (Gas Lift)", "F46", top="K46", od="B46", idd="D46", string=S, note="F47 'DV (in DROID, set on 21  Feb 2022)'"),
        comp("s-spm-4", "SPM (Gas Lift)", "F49", top="K49", od="B49", idd="D49", string=S, note="F50 'DV (in DROID, set on 21  Feb 2022)'"),
        comp("s-ssd", "Sliding Sleeve", "F54", top="K54", od="B54", idd="D54", string=S),
        comp("dual-packer", "Packer", ["U60", "U61"], top="S60", od="AB60", idd="Z60", string=SL, note="one shared dual-string packer (label 'DUAL STRING')"),
        comp("s-landing-nipple", "Landing Nipple", "F67", top="K67", od="B67", idd="D67", string=S),
        comp("s-pop", "Mechanical Plug", "F69", top="K69", od="B69", idd="D69", string=S, approx="pump-out plug -> mechanical-plug"),
        comp("s-perf-pup", "Pup Joint", "F70", od="B70", idd="D70", string=S, approx="perforated pup -> pup-joint", note="no depth printed"),
        comp("s-ball-catcher", "Pup Joint", "F71", top="K71", od="B71", idd="D71", string=S, approx="ball catcher sub -> pup-joint (nearest short tubular)",
             note="MIN ID printed 0.000 (D71) — kept as printed"),
        comp("s-eot", "EOT", "F73", top="K73", string=S),
        # long string
        comp("l-flow-coupling-1", "Flow Coupling", "U29", od="AB29", idd="Z29", string=L, note="no depth printed"),
        comp("l-trsv", "TRSV", "U30", top="S30", od="AB30", idd="Z30", string=L),
        comp("l-flow-coupling-2", "Flow Coupling", "U31", od="AB31", idd="Z31", string=L, note="no depth printed"),
        comp("l-spm-1", "SPM (Gas Lift)", "U39", top="S39", od="AB39", idd="Z39", string=L, note="U40 'TRO : 650'"),
        comp("l-spm-2", "SPM (Gas Lift)", "U42", top="S42", od="AB42", idd="Z42", string=L, note="U43 'TRO : 715'"),
        comp("l-spm-3", "SPM (Gas Lift)", "U45", top="S45", od="AB45", idd="Z45", string=L, note="U46 'TRO : 770'"),
        comp("l-spm-4", "SPM (Gas Lift)", "U48", top="S48", od="AB48", idd="Z48", string=L),
        comp("l-ssd", "Sliding Sleeve", "U80", top="S80", od="AB80", idd="Z80", string=L,
             note="drawing text box 'Rectangle 8' anchored K79:O81 (same rows as this SSD, row 80): 'CLOSED (Nov 2024)' — association by position, flagged"),
        comp("l-no-go-locator", "ST Slot Locator", "U82", top="S82", od="AB82", idd="Z82", string=L, approx="no-go locator -> st-slot-locator"),
        comp("gp-packer", "Packer", "U85", top="S85", od="AB85", idd="Z85", string=L, note="lower completion (long-string column)"),
        comp("mcs-1", "Sliding Sleeve", "U86", top="S86", od="AB86", idd="Z86", string=L, approx="MCS closing sleeve -> sliding-sleeve"),
        comp("sealbore-1", "PBR", "U87", top="S87", od="AB87", idd="Z87", string=L, approx="sealbore extension -> pbr"),
        comp("screen-e14b", "Screen", "U95", od="AB95", idd="Z95", string=L, extra={"zone": "E14B"}, note="no depth printed; zone E14B by position (B93)"),
        comp("iso-packer-1", "Packer", "U102", top="S102", od="AB102", idd="Z102", string=L, note="name/ID/OD cells are formulas copying row 111 (=U111, =Z111, =AB111)"),
        comp("mcs-2", "Sliding Sleeve", "U103", top="S103", od="AB103", idd="Z103", string=L, approx="MCS closing sleeve -> sliding-sleeve", note="cells are formulas (=U112 ...)"),
        comp("sealbore-2", "PBR", "U104", top="S104", od="AB104", idd="Z104", string=L, approx="sealbore extension -> pbr",
             note="cells are formulas (=U113 ...); S104 1427.34 printed shallower than S103 1427.83 — as printed"),
        comp("iso-packer-2", "Packer", "U111", top="S111", od="AB111", idd="Z111", string=L),
        comp("mcs-3", "Sliding Sleeve", "U112", top="S112", od="AB112", idd="Z112", string=L, approx="MCS closing sleeve -> sliding-sleeve", note="cells are formulas (=U86 ...)"),
        comp("sealbore-3", "PBR", "U113", top="S113", od="AB113", idd="Z113", string=L, approx="sealbore extension -> pbr", note="cells are formulas (=U87 ...)"),
        comp("l-seal-assy", "Seal Assembly", "U115", top="S115", od="AB115", idd="Z115", string=L),
        comp("l-landing-nipple", "Landing Nipple", "U116", top="S116", od="AB116", idd="Z116", string=L, note="cells are formulas (=F67, =D67, =B67)"),
        comp("l-pop", "Mechanical Plug", "U117", top="S117", od="AB117", idd="Z117", string=L, approx="pump-out plug -> mechanical-plug", note="cells are formulas (=F69 ...)"),
        comp("l-perf-pup", "Pup Joint", "U118", top="S118", od="AB118", idd="Z118", string=L, approx="perforated pup -> pup-joint", note="cells are formulas (=F70 ...)"),
        comp("l-mule-shoe", "Mule Shoe", "U119", top="S119", od="AB119", idd="Z119", string=L, note="end of the long string"),
        comp("screen-e34a", "Screen", "U121", od="AB121", idd="Z121", string=L, extra={"zone": "E34A"}, note="no depth printed; cells are formulas (=U95 ...); zone E34A by position (B119)"),
        comp("collet-seal", "Seal Assembly", "U127", top="S127", od="AB127", idd="Z127", string=L),
        comp("sump-packer", "Packer", "U129", top="S129", od="AB129", idd="Z129", string=L, note="S129 1528 printed shallower than the collet seal S127 1530.01 — as printed"),
        comp("pbtd", "PBTD", "I137", top="K137", note="header G11 = formula =K137, same value"),
        comp("zone-e14b", "Perforation", "B93", top=("val", 1416.0, "B95 '1416m - 1423 m MDDF (7 m)'"), bot=("val", 1423.0, "B95"), string=L,
             extra={"zone": "E14B", "interval": ("ref", "B95", "txt")}, tvd="B94", approx="reservoir zone interval beside the screens; perforation is the nearest interval type",
             note="B96 = 6.120 (formula =1317.16/(0.0521*1259.11*3.281), EMW ppg); B97 '1317.16 psi, 194 deg F'; J95 33 (Degree); string Long by position"),
        comp("zone-e34a", "Perforation", "B119", top=("val", 1517.0, "B121 '1517 m - 1524 m MDDF (7m)'"), bot=("val", 1524.0, "B121"), string=L,
             extra={"zone": "E34A", "interval": ("ref", "B121", "txt")}, tvd="B120", approx="reservoir zone interval beside the screens; perforation is the nearest interval type",
             note="B122 = 6.307 (formula =1448.81/(0.0521*1343.93*3.281), EMW ppg); B123 '1448.81 psi, 206 deg F'; J121 32.8 (Degree); string Long by position"),
    ]


# ---------------------------------------------------------------- A33ST3 (single injector; O name / K depth / M length / S ID / T OD)
def spec_a33(B):
    w = lambda r: ("ref", r, "num")
    g = lambda r: ("ref", r, "txt")
    surf = ("val", 0.0, "ASSUMED: header TOP 'Surface' -> 0")
    cas = lambda cid, nm, od, r: comp(cid, "Casing", nm, top=surf, bot=f"T{r}", od=("val", od, f"{nm} size"), extra={"weight": w(f"P{r}"), "grade": g(f"Q{r}")},
                                      est=(["topMd"], f"ASSUMED: S{r} 'Surface' mapped to topMd 0 (surface in the m-MDDF reference)"),
                                      note=f"casing scheme row {r}: thread R{r}; weight P{r} printed without unit")
    L_ = lambda r: f"length M{r} {B.num(f'M{r}'):.3f} m"
    return [
        cas("conductor", "I6", 26.0, 6), cas("casing-16", "I7", 16.0, 7), cas("casing-13-3-8", "I8", 13.375, 8), cas("casing-9-5-8", "I9", 9.625, 9),
        comp("liner-7", "Liner", "I10", top="F40", od=("val", 7.0, "I10 '7\" Liner'"),
             note="CONFLICT: header row 10 weight/grade/thread/top/bottom all 'TBC'; diagram B40 '7\"Liner (TOL)' F40 = 902.73 -> diagram value used for the top; "
                  "no liner shoe depth -> bottom empty"),
        comp("tubing-gre", "Tubing", "O48", od=("val", 3.5, "3-1/2\""), extra={"weight": w("P11"), "grade": g("Q11")},
             note="diagram label O48; header I11 'Tubing 3-1/2\" GRE lined' (P11 9.2, Q11 L80, R11 JFEFOX); interval not printed (component lengths in column M "
                  "do not close between the anchored depths) -> top/bottom empty"),
        comp("tubing-15cr", "Tubing", "O59", bot=("val", 1836.02, "DERIVED: Muleshoe Guide / Drop Sub K61 (end of string)"), od=("val", 3.5, "3-1/2\""),
             extra={"weight": w("P12"), "grade": g("Q12")}, est=(["bottomMd"], "DERIVED: bottom = muleshoe guide depth K61"),
             note="header I12 'Tubing 3-1/2\" - Below 7\" AHC' (R12 JFE BEAR) — the packer on the diagram is a BLUEPACK MAX (header says AHC; recorded); top not printed"),
        comp("xmas-tree", "Xmas Tree", "B13", note="label text incl. 'X-Mas Tree :' kept as written"),
        comp("tubing-hanger", "Tubing Hanger", "O18", top="K17", idd="S18",
             est=(["topMd"], "ASSUMED: hanger top = K17 15.8 m, the 'ELEVATION (DF to THF)' row directly above it"),
             note=f"{L_(18)}; MAX OD T18 'TBC' -> empty"),
        comp("xover-1", "Crossover", "O19", od="T19", idd="S19", note=f"{L_(19)}; no depth printed"),
        comp("pup-12ft", "Pup Joint", "O20", od="T20", idd="S20", note=f"{L_(20)}; no depth printed"),
        comp("pup-6ft-1", "Pup Joint", "O23", od="T23", idd="S23", note=f"{L_(23)}; no depth printed"),
        comp("flow-coupling-1", "Flow Coupling", "O24", od="T24", idd="S24", note=f"{L_(24)}; no depth printed"),
        comp("trsv", "TRSV", "O25", top="K25", od="T25", idd="S25", note=L_(25)),
        comp("flow-coupling-2", "Flow Coupling", "O26", od="T26", idd="S26", note=f"{L_(26)}; no depth printed"),
        comp("pup-4ft", "Pup Joint", "O27", od="T27", idd="S27", note=f"{L_(27)}; no depth printed"),
        comp("pup-6ft-2", "Pup Joint", "O50", od="T50", idd="S50", note=f"{L_(50)}; no depth printed"),
        comp("xover-2", "Crossover", "O51", od="T51", idd="S51", note=f"{L_(51)}; no depth printed"),
        comp("packer", "Packer", "O52", top="K52", od="T52", idd="S52", note=f"{L_(52)}; angle N52 58 (°)"),
        comp("pup-5ft", "Pup Joint", "O53", od="T53", idd="S53", note=f"{L_(53)}; no depth printed"),
        comp("xover-3", "Crossover", "O54", od="T54", idd="S54", note=f"{L_(54)}; no depth printed"),
        comp("landing-nipple", "Landing Nipple", "O55", top="K55", od="T55", idd="S55", note=L_(55)),
        comp("xover-4", "Crossover", "O56", od="T56", idd="S56", note=f"{L_(56)}; no depth printed"),
        comp("pup-6ft-3", "Pup Joint", "O57", od="T57", idd="S57", note=f"{L_(57)}; no depth printed"),
        comp("muleshoe", "Mule Shoe", "O61", top="K61", od="T61", idd="S61", note="end of the tubing string"),
        comp("pbtd", "PBTD", "B74", top="F74"),
        comp("zone-e34a", "Perforation", "B63", top=("val", 1846.0, "B65 '1,846-1,856m MDDF'"), bot=("val", 1856.0, "B65"),
             extra={"zone": "E34A", "interval": ("ref", "B65", "txt")}, tvd="B64",
             approx="reservoir zone interval below the muleshoe; not stated as perforations — perforation is the nearest interval type",
             note="B66 'Reservoir Pressure: 1,503 psi, 93deg C'; B67 6.47 (cached value, no unit; formula not readable from .xls with xlrd)"),
    ]


def wellinfo(tpl, B, d):
    wi = {k: v if v in ("-", 0) else (None if isinstance(v, (int, float)) and not isinstance(v, bool) else "") for k, v in tpl["wellInfo"].items()}
    wi.update(d)
    return wi


def main():
    hb = [l.split("  ", 1) for l in (RUN / "sha256_before.txt").read_text().splitlines()]
    tpl = json.loads(subprocess.run(["textutil", "-convert", "txt", "-stdout", str(TPL)], capture_output=True, text=True, check=True).stdout)
    keymap = {r[0]: r[2] for r in load_workbook(XTPL)["Component_Types"].iter_rows(min_row=2, values_only=True)}
    a17, a23, a33 = load_books()
    tb23 = [o for o in json.loads((RUN / "dump_A23ST2.json").read_text())["drawings"]["xl/drawings/drawing1.xml"] if o["text"]]
    assert [o["text"] for o in tb23] == [["CLOSED (Nov 2024)"]] and tb23[0]["from"] == "K79"
    tb33 = json.loads((RUN / "dump_A33ST3_textboxes.json").read_text())
    assert tb33[0]["sheet"] == 2 and tb33[0]["text"].startswith('2" LIB tagged HUD at 1858m WLTHF')
    assert a17.txt("B2") == "WELL NAME: DULANG A17-ST2 OP5" and a23.txt("B2") == "WELL NAME: DULANG A23 OP2" and a33.txt("B2") == "DULANG A33-ST3"

    wells = [
        ("Dulang_A17ST2", OUT / "Dulang_A17ST2.json", a17, spec_a17(a17), "m-MDDF (columns K20/S20 'm-MDDF'; note B18 says 'from RKB')", dict(
            name="Dulang_A17ST2", completionDate="2020-05-30", rig=a17.txt("G9"), totalDepth=ft(a17.num("G10")), maxDeviation="80.3 deg",
            kbElevation=ft(33.0), wellhead=a17.txt("G15"), completionFluid=a17.txt("V16"), packerFluid=a17.txt("V17") + " " + a17.txt("V18"),
            tubing='3-1/2" 9.2 PPF HP2-13CR-95 JFEBEAR; 3-1/2" 9.2 PPF 13CR-L80 JFEBEAR', trajectoryType="Deviated")),
        ("Dulang_A23", OUT / "from_workbook_A23" / "Dulang_A23.json", a23, spec_a23(a23), "m-MDDF (columns K20/S20 'm-MDDF'; note B18 says 'from RKB')", dict(
            name="Dulang_A23", completionDate="2020-03-17", rig=a23.txt("G9"), totalDepth=ft(a23.num("G10")), maxDeviation="44.69 deg",
            kbElevation=ft(33.0), wellhead=a23.txt("G15"), completionFluid=a23.txt("V16"), packerFluid=a23.txt("V17") + " " + a23.txt("V18"),
            tubing='3-1/2" 9.2 PPF HP2-13CR-95 JFEBEAR; 3-1/2" 9.2 PPF HP2 13Cr95 JFE BEAR (header row 15 says 13CR-L80)', trajectoryType="Deviated")),
        ("Dulang_A33ST3", OUT / "Dulang_A33ST3.json", a33, spec_a33(a33), "m-MDDF (column headers F16/L16 'm-MDDF')", dict(
            name="Dulang_A33ST3", completionDate="3rd April 2020", rig="Sapura Berani", totalDepth=ft(a33.num("F75")),
            maxDeviation="58deg @ 1,800m MDDF", kbElevation=ft(30.1), wellhead="Surface Unihead, Solar Alert, Surface, Standard, Single",
            completionFluid="8.5 ppg NaCl filtered to 30 NTU",
            packerFluid="8.5 ppg NaCl with 2% Corrosion Inhibitor + 2% Oxygen Scavenger + 0.2% Biocide",
            tubing='Tubing 3-1/2" GRE lined 9.2 L80 JFEFOX; Tubing 3-1/2" - Below 7" AHC 9.2 15CR-125 JFE BEAR', trajectoryType="Deviated")),
    ]
    assert a33.txt("B10") == "Completion Date : 3rd April 2020" and a33.txt("B6").startswith("Rig Name: Sapura Berani")
    assert a33.txt("B9") == "Drill floor height: 30.1m-AMSL" and a17.txt("G14") == "33 m-AMSL" and a23.txt("G14") == "33 m-AMSL"
    assert a33.txt("I13") == "Completion Fluid : 8.5 ppg NaCl filtered to 30 NTU"
    assert a33.txt("I14") == "Packer Fluid: 8.5 ppg NaCl with 2% Corrosion Inhibitor + 2% Oxygen Scavenger + 0.2% Biocide"
    assert a17.v("G8") == "2020-05-30T00:00:00" and a23.v("G8") == "2020-03-17T00:00:00"
    assert a17.num("G12") == 80.3 and a23.num("G12") == 44.69

    allok = True
    for name, path, B, specs, ref, wi_vals in wells:
        comps = build_components(B, specs, keymap, ref)
        doc = {k: {"wellInfo": wellinfo(tpl, B, wi_vals), "components": comps}.get(k, []) for k in tpl}
        if path.exists():
            raise SystemExit(f"COLLISION: {path} exists — not overwritten")
        path.parent.mkdir(exist_ok=True)
        path.write_text(json.dumps(doc, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
        back = json.loads(path.read_text())
        cs = back["components"]
        kind = lambda v: "number" if isinstance(v, (int, float)) and not isinstance(v, bool) else type(v).__name__
        approx_ids = {s["cid"] for s in specs if s["approx"]}
        checks = [
            ("valid JSON", True),
            ("top-level sections and order = template", list(back) == list(tpl)),
            ("wellInfo keys = template (no extra keys)", list(back["wellInfo"]) == list(tpl["wellInfo"])),
            ("wellInfo types = template where populated", all(x in ("", None) or kind(x) == kind(tpl["wellInfo"][k]) for k, x in back["wellInfo"].items())),
            ("no template sample records", back["surveys"] == [] and back["formationTops"] == [] and not {c["id"] for c in cs} & {t["id"] for t in tpl["components"]}),
            ("all types are Component_Types keys", all(c["type"] in set(keymap.values()) for c in cs)),
            ("component ids unique", len({c["id"] for c in cs}) == len(cs)),
            ("MD ft = m x 3.28084 (3 dp)", all((c[a] is None and c[b] is None) or c[a] == round(c[b] * FT, 3) for c in cs for a, b in (("topMd", "topMdM"), ("bottomMd", "bottomMdM")))),
            ("TVD ft = m x 3.28084 and only from 'm TVD' text", all(c["topTvd"] == round(c["topTvdM"] * FT, 3) and c["bottomTvd"] == round(c["bottomTvdM"] * FT, 3) and "m TVD" in c["source"] for c in cs if "topTvd" in c)),
            ("every component has workbook/sheet/cell traceability", all(c["source"].startswith(B.label + " > sheet") for c in cs)),
            ("approximate types flagged", all(c.get("estimated") and "type" in c["estimatedFields"] for c in cs if c["id"] in approx_ids)),
            ("no unflagged type estimates", all("type" not in c.get("estimatedFields", []) for c in cs if c["id"] not in approx_ids)),
            ("string values valid", all(c.get("string", S) in (S, L, SL) for c in cs)),
            ("name field = filename", back["wellInfo"]["name"] == path.stem),
        ]
        print(f"=== {path.relative_to(OUT)}: {len(cs)} components")
        for a, b in checks:
            allok &= b
            print("  PASS" if b else "  FAIL", "-", a)
    for h, p in hb:
        same = sha(p if p.startswith("/") else str(ROOT / p)) == h
        allok &= same
        print("PASS" if same else "FAIL", "- unchanged:", p.split("/")[-1])
    print("ALL PASS" if allok else "SOME CHECKS FAILED")


if __name__ == "__main__":
    main()
