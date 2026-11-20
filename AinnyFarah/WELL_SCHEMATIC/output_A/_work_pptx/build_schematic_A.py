#!/usr/bin/env python3
"""Map 'Dulang A Well Schematic updated.pptx' (NAS, read-only) into well-schematic JSON for the wells in specs/ (11 Nov 2026).

Copied from output_D/_work/build_schematic_D.py — same rules; adapted: source file, DULANG-A<no> naming, A specs only,
upper/lower picture (p1/p2) locations, summary written to _work_pptx/_completion_summary_A_pptx.md.

The Dulang A schematics are screenshot pictures (upper + lower part per slide). Values were
transcribed per well into specs/<well>.txt (format in specs/README.txt): EMF native text where available (D2ST1, D10ST1,
D12ST1, D13ST1, D20ST1), otherwise visual reading of x3-upscaled crops (media_png/) cross-checked with tesseract OCR.
This script turns the specs into JSON with full traceability, adds every slide text box / table / off-slide picture as
unmapped content, validates each file and writes _work_pptx/_completion_summary_A_pptx.md. Existing outputs are never overwritten.

Rules: ft = m x 3.28084 (3 dp), original metres in ...M keys; datums kept as printed (equipment column datum per well;
values on another datum flagged, never shifted); WLTHF depths not used as MD; header sizes without ID/OD designation -> ID
(user-defined rule, flagged); explicit MIN ID / MAX OD columns kept; diagram values preferred over header values (both
recorded); tubular tops of 0 only when the depth reference is the tubing hanger flange (flagged); no TD/PBTD substitution.
"""
import hashlib
import json
import os
import re
import subprocess
import sys
from pathlib import Path

from openpyxl import load_workbook

W = Path(__file__).resolve().parent
sys.path.insert(0, str(W))
from d_rules import FT_PER_M, SRC_NAME, classify, ft  # noqa: E402

OUT = W.parent
ROOT = OUT.parent
NAS = Path("/Volumes/homes/public/DulangOSS/20261016/DULANG_WELLS/Well_Schematic/1. Well Schematic/Dulang A/Dulang A Well Schematic updated.pptx")
NAS_SHA_AT_COPY = "a5ca9a3b0076122633de31730fc97b5b8e3185d813c93bdcfdb3ad56bef2861b"  # NAS file and local copy, hashed 11 Nov 2026 before processing
TPL = ROOT / "well-schematic-template-json.rtf"
XTPL = ROOT / "well-schematic-sample-template.xlsx"
OFFSLIDE = {}  # no off-slide pictures on the A slides processed here


def sha(p):
    h = hashlib.sha256()
    fd = os.open(p, os.O_RDONLY)
    while (b := os.read(fd, 1 << 20)):
        h.update(b)
    try:
        os.close(fd)
    except OSError:
        pass
    return h.hexdigest()


def load_template():
    txt = subprocess.run(["textutil", "-convert", "txt", "-stdout", str(TPL)], capture_output=True, text=True, check=True).stdout
    return json.loads(txt)


def type_keys():
    ws = load_workbook(XTPL, read_only=True)["Component_Types"]
    return {r[2] for r in ws.iter_rows(min_row=2, values_only=True) if r[2]}


def num(s):
    """'1234.5' -> (1234.5, False); '?1386.2' -> (1386.2, True); '' / 'TBA' -> (None, False)"""
    s = (s or "").strip()
    unclear = s.startswith("?")
    m = re.fullmatch(r"\??\s*(\d+(?:\.\d+)?)", s.replace(",", ""))
    return (float(m.group(1)), unclear) if m else (None, False)


def family(d):
    d = (d or "").upper()
    if "WLTHF" in d:
        return "WLTHF"
    if re.search(r"THF|MDTH\b|MDTH$|\bTHF\b", d):
        return "THF"
    if "RKB" in d:
        return "RKB"
    if re.search(r"DF|MDDF", d):
        return "DF"
    return "?"


MONTHS = {m: i + 1 for i, m in enumerate(["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"])}


def iso_date(text):
    t = text.upper().replace(",", " ").replace(".", " ")
    note = ""
    if "0CTOBER" in t:
        t, note = t.replace("0CTOBER", "OCTOBER"), "'0CTOBER' read as OCTOBER (digit zero printed for letter O); "

    def yy(y):
        return (("19" + y) if int(y) > 50 else ("20" + y)) if len(y) == 2 else y
    m = re.search(r"\b(\d{1,2})\s*[-/ ]\s*([A-Z]{3})[A-Z]*\s*[-/ ]?\s*(\d{2,4})\b", t)
    if m and m.group(2) in MONTHS:
        return f"{int(yy(m.group(3))):04d}-{MONTHS[m.group(2)]:02d}-{int(m.group(1)):02d}", note + ("two-digit year expanded" if len(m.group(3)) == 2 else "")
    m = re.fullmatch(r"\s*(\d{1,2})/(\d{1,2})/(\d{4})\s*", t)
    if m:
        return f"{m.group(3)}-{int(m.group(2)):02d}-{int(m.group(1)):02d}", note + "dd/mm/yyyy"
    m = re.search(r"\b([A-Z]{3})[A-Z]*\s*-?\s*(\d{2,4})\b", t)
    if m and m.group(1) in MONTHS:
        return f"{int(yy(m.group(2))):04d}-{MONTHS[m.group(1)]:02d}", note + "month and year only"
    return None, "date not parseable"


def parse_spec(path):
    S = {k: [] for k in ("INFO", "CASING", "TUBING", "ROW", "ZONE", "PBTD", "TD", "SHOE", "INT", "NOTE", "ISSUE", "CEMENT", "FLOAT", "SRC")}
    for line in path.read_text().splitlines():
        if not line.strip():
            continue
        p = [x.strip() for x in re.sub(r"\s+\|\s*$", "", line).split(" | ")]
        if p[0] == "WELL":
            S["key"], S["slides"], S["media"], S["label"] = p[1], [int(x) for x in p[2].split(",")], p[3], p[4]
        elif p[0] == "DATUM":
            S["datum"] = p[1]
        else:
            S[p[0]].append(p[1:])
    return S


class Well:
    def __init__(self, S):
        self.S = S
        self.key, self.name, self.slide, self.label = S["key"], f"DULANG-{S['key']}", S["slides"][0], S["label"]
        self.comps, self.trace, self.issues, self.unmapped, self.ids = [], [], [], [], {}
        self.info, self.float_equipment, self.plug_cement = {}, [], []
        self.prim = family(S["datum"])
        src = S["SRC"][0]
        self.media, self.method = src[1], src[2]

    def cid(self, base):
        base = re.sub(r"[^a-z0-9]+", "-", base.lower()).strip("-")
        n = self.ids.get(base, 0) + 1
        self.ids[base] = n
        return f"{base}-{n}"

    def tr(self, path, value, where, original, method):
        self.trace.append(dict(path=path, value=value, where=where, original=original, method=method))

    def issue(self, kind, where, text):
        self.issues.append((kind, where, text))

    def where(self, part):
        part = part or ""
        if part.startswith("y="):
            return f"slide {self.slide} {self.media} EMF text record {part}"
        imgs = [m.strip() for m in re.split(r"\+|\(", self.media)[0:2] if m.strip().endswith((".png", ".emf", ".jpeg"))]
        if part in ("p1", "p2") and len(imgs) >= 2:
            return f"slide {self.slide} {imgs[int(part[1]) - 1]} ({'upper' if part == 'p1' else 'lower'} picture)"
        if part[:1] in "ab" and part[1:].isdigit() and len(imgs) >= 2:
            return f"slide {self.slide} {imgs[0] if part[0] == 'a' else imgs[1]} {'upper' if part[1] == '1' else 'lower'} half"
        return f"slide {self.slide} {imgs[0] if imgs else self.media} {part}".strip()


def add_comp(w, *, cid, typ, name, where, string=None, top=None, bot=None, idd=None, od=None, weight=None, grade=None,
             zone=None, closest=None, unresolved=False, est=(), notes=(), extra=()):
    """top/bot: (m, original, unit) ; idd/od: (inches, original, method)"""
    p = f"components[{cid}]"
    d = {"id": cid, "type": typ, "name": name, "topMd": None, "bottomMd": None, "outerDiameter": None, "innerDiameter": None,
         "topMdM": None, "bottomMdM": None}
    w.tr(f"{p}.name", name, where, name, w.method)
    w.tr(f"{p}.type", typ, where, name, ("UNRESOLVED / " if unresolved else "") + ("NEAREST TYPE (flagged)" if closest else "EXACT TYPE (Component_Types)"))
    src = [f"{SRC_NAME} > {where} name '{name}' ({w.method})"]
    estf, enote = list(est), list(notes)
    if closest:
        estf.insert(0, "type")
        enote.insert(0, ("UNRESOLVED type: " if unresolved else "type: ") + closest)
    for val, field in ((top, "topMd"), (bot, "bottomMd")):
        if not val or val[0] is None:
            continue
        m, orig, unit = val
        d[field], d[field + "M"] = ft(m), float(m)
        w.tr(f"{p}.{field}", d[field], where, orig, f"{m} m x {FT_PER_M} = {d[field]} ft (3 dp); {unit}")
        w.tr(f"{p}.{field}M", float(m), where, orig, f"original metres; {unit}")
        src.append(f"{field} '{orig}' = {m} m -> {d[field]} ft ({unit})")
    for val, field in ((idd, "innerDiameter"), (od, "outerDiameter")):
        if not val or val[0] is None:
            continue
        d[field] = val[0]
        w.tr(f"{p}.{field}", val[0], where, val[1], val[2])
        src.append(f"{field} '{val[1]}' ({val[2]})")
    if weight is not None:
        d["weight"] = weight
        w.tr(f"{p}.weight", weight, where, name, "parsed from the label (ppf = lb/ft)")
    if grade:
        d["grade"] = grade
        w.tr(f"{p}.grade", grade, where, name, "parsed from the label")
    if zone:
        d["zone"] = zone
        w.tr(f"{p}.zone", zone, where, zone, "zone label as printed")
    if string:
        d["string"] = string
        w.tr(f"{p}.string", string, where, "", "string column of the schematic (short / long / dual packer) as printed")
    src += [e for e in extra if e]
    if estf:
        d.update(estimated=True, estimatedFields=list(dict.fromkeys(estf)), estimationNote="; ".join(enote))
        if "type" in estf:
            d["actualType"] = name
    elif enote:
        src.append("NOTE: " + "; ".join(enote))
    d["source"] = "; ".join(src)
    w.comps.append(d)
    return d


STRING = {"S": "Short", "L": "Long", "D": "Short & Long", "X": None, "Lw": None}
PREFIX = {"S": "s-", "L": "l-", "D": "", "X": "", "Lw": "lc-"}


def size_in(text):
    t = text.replace("½", " 1/2").replace("¼", " 1/4")
    m = re.search(r'(?<![\d.])(\d{1,2})(?:\s*[- ]\s*(\d)/(\d))?\s*("|"")', t) or re.search(r'(?<![\d.])(\d{1,2})[- ](\d)/(\d)\b', t)
    if not m:
        return None
    v = float(m.group(1))
    if m.group(2):
        v += int(m.group(2)) / int(m.group(3))
    return round(v, 4)


def build_rows(w):
    S, rows_out = w.S, []
    for r in S["ROW"]:
        side, name, top, bot, idv, odv, status, note, part = (r + [""] * 9)[:9]
        where = w.where(part)
        typ, closest, unres = classify(name)
        if typ is None:
            w.unmapped.append((where, name, "no Type_Key could be assigned"))
            w.issue("Unresolved", where, f"'{name}': no Type_Key")
            continue
        est, notes, extra = [], [], []
        if status:
            extra.append(f"status/remark '{status.lstrip('?')}'" + (" (position-matched, see issues)" if status.startswith("?") else ""))
        if note:
            extra.append(f"transcription note: {note}")
        tv, tu = num(top)
        bv, bu = num(bot)
        unit = f"{S['datum']} (equipment depth column as printed)"
        if tu or bu:
            est += (["topMd"] if tu else []) + (["bottomMd"] if bu else [])
            notes.append("UNCLEAR READ: depth digits of low legibility (marked '?' in the transcription)")
        if bv is not None:
            if typ in ("blast-joint",):
                notes.append("INTERPRETED: depths printed at both ends of the blast-joint bracket taken as top/bottom")
                est += ["topMd", "bottomMd"]
            elif typ == "packer":
                notes.append("depth range printed for the swell packer element (top - bottom)")
        iv, iu = num(idv)
        ov, ou = num(odv)
        idd = (iv, idv.lstrip("?"), "column 'MIN ID' = explicit ID") if iv else None
        odd = (ov, odv.lstrip("?"), "column 'MAX OD' = explicit OD") if ov else None
        if iu or ou:
            est += (["innerDiameter"] if iu else []) + (["outerDiameter"] if ou else [])
            notes.append("UNCLEAR: MIN ID / MAX OD value not aligned with the row or of low legibility (marked '?')")
        if idd is None and odd is None:
            m = re.match(r'\s*(\d\.\d{2})\s*"', name)
            if m and typ not in ("tubing", "screen"):
                idd = (float(m.group(1)), name, "USER-DEFINED RULE: label size without ID/OD designation -> innerDiameter")
                est.append("innerDiameter")
                notes.append(f"innerDiameter: label size {m.group(1)}\" has no ID/OD designation -> recorded as ID (user-defined rule)")
        if side == "Lw":
            notes.append("lower-completion item (gravel-pack / screen / swell-packer assembly, no string assigned)") if not est and not closest else extra.append("lower-completion item (no string assigned)")
        d = add_comp(w, cid=w.cid(PREFIX[side] + typ.replace(" (", "-").replace(")", "")), typ=typ, name=name, where=where,
                     string=STRING[side], top=(tv, top.lstrip("?"), unit) if tv is not None else None,
                     bot=(bv, bot.lstrip("?"), unit) if bv is not None else None, idd=idd, od=odd, closest=closest,
                     unresolved=unres, est=est, notes=notes, extra=extra + [f"depth datum: {S['datum']}"])
        rows_out.append((side, d, note))
    w.row_count = len(rows_out)
    # depth sequence per string (rows are transcribed top-down as printed)
    last = {}
    for side, d, _ in rows_out:
        if d["topMdM"] is None or side == "Lw":
            continue
        for k in ({"D": ("S", "L")}.get(side) or (side,)):
            p = last.get(k)
            if p and d["topMdM"] < p["topMdM"] - 0.05:
                w.issue("Conflict", d["id"], f"'{d['name']}' at {d['topMdM']} m is shallower than the item listed above it ({p['id']} at {p['topMdM']} m) — kept as printed")
                d["source"] += f"; CONFLICT: out of depth sequence (item above {p['id']} at {p['topMdM']} m)"
            last[k] = d
    return rows_out


def end_item(rows, sides, inner):
    cands = [(s, d, n) for s, d, n in rows if s in sides and d["topMdM"] is not None]
    tagged = [d for s, d, n in cands if "end of" in n.lower() and (("inner" in n.lower()) == inner)]
    if tagged:
        return tagged[-1], "row marked as the end of the string in the transcription"
    pool = [d for s, d, n in cands if bool("inner" in d["name"].lower() or re.match(r'\s*3\.2\d"', d["name"])) == inner]
    eot = [d for d in pool if d["type"] == "eot"]
    if eot:
        return eot[-1], "documented EOT / END OF TUBING label"
    shoe = [d for d in pool if d["type"] == "mule-shoe"]
    if shoe:
        return shoe[-1], "mule shoe / entry guide at the string end"
    return None, None


def build_tubulars(w, rows):
    S = w.S
    sizes = {}
    # casing / liner
    for c in S["CASING"]:
        name, size, wt, grade, shoe, datum = (c + [""] * 6)[:6]
        tol, tolsrc = (c[6], c[7]) if len(c) > 7 else ("", "")
        where = w.where("header")
        sv, _ = num(size)
        wv, _ = num(wt)
        bv, bu = num(shoe)
        typ = "liner" if re.search(r"(?i)LINER", name) and not re.search(r"(?i)CASING/LINER", name) else "casing"
        est, notes = (["innerDiameter"] if sv else []), (["innerDiameter: header size has no ID/OD designation -> recorded as ID (user-defined rule)"] if sv else [])
        bot = None
        if bv is not None:
            fam = family(datum)
            bot = (bv, f"{name}", f"m {datum} as printed")
            if fam != w.prim:
                est.append("bottomMd")
                notes.append(f"DATUM: shoe depth printed as '{datum}', equipment datum {S['datum']}; not shifted")
            if bu:
                est.append("bottomMd")
                notes.append("UNCLEAR READ: shoe depth of low legibility")
        elif re.search(r"\bTD\b", datum):
            w.issue("Missing", where, f"'{name}': depth column says 'TD' — no numeric shoe depth")
        else:
            w.issue("Missing", where, f"'{name}': setting depth not printed")
        top = None
        if tol:
            tv, _ = num(tol)
            top = (tv, tolsrc, "TOL from the diagram label (m MDTHF)")
            if typ == "casing":
                typ = "liner"
        base = "conductor" if re.search(r"(?i)COND|DRIVE", name) else "surface-casing" if re.search(r"(?i)SURF|INT\. CASING", name) else \
            ("liner" if typ == "liner" else "production-casing")
        d = add_comp(w, cid=w.cid(base), typ=typ, name=name, where=where, top=top, bot=bot,
                     idd=(sv, size, "USER-DEFINED RULE: header size without ID/OD designation -> innerDiameter") if sv else None,
                     weight=wv, grade=grade or None, est=est, notes=notes)
        if sv:
            sizes.setdefault(sv, []).append(d)
    # diagram shoes: diagram value preferred over header
    for sh in S["SHOE"]:
        name, depth, datum, tvd, orig = (sh + [""] * 5)[:5]
        dv, _ = num(depth)
        where = w.where("diagram")
        if dv is None:
            w.unmapped.append((where, name, f"casing-shoe label without its own depth ({orig})"))
            continue
        sv = size_in(name)
        match = sizes.get(sv, []) if sv else []
        if match:
            c = match[0]
            old = (c["bottomMdM"], c["source"])
            c["bottomMd"], c["bottomMdM"] = ft(dv), dv
            c["source"] += f"; bottomMd from the diagram label '{orig}' = {dv} m ({datum}) -> {ft(dv)} ft — diagram value preferred over the header" + \
                (f" (header shoe {old[0]} m recorded)" if old[0] is not None else "")
            if "estimatedFields" in c and "bottomMd" in c["estimatedFields"]:
                c["estimatedFields"].remove("bottomMd")
                c["estimationNote"] += "; (header shoe datum note superseded by the diagram value)"
                if not c["estimatedFields"]:
                    for k in ("estimated", "estimatedFields", "estimationNote"):
                        c.pop(k, None)
            if family(datum) != w.prim:
                c.setdefault("estimatedFields", []).append("bottomMd")
                c["estimated"] = True
                c["estimationNote"] = (c.get("estimationNote", "") + f"; DATUM: diagram shoe printed in {datum}, not shifted").lstrip("; ")
            w.tr(f"components[{c['id']}].bottomMd", ft(dv), where, orig, f"{dv} m x {FT_PER_M}; diagram value preferred over header ({old[0]} m)")
            w.tr(f"components[{c['id']}].bottomMdM", dv, where, orig, f"original metres; {datum}")
            if old[0] is not None and abs(old[0] - dv) > 0.05:
                w.issue("Conflict", where, f"'{c['name']}': header shoe {old[0]} m vs diagram '{orig}' {dv} m — diagram value used, header recorded")
        else:
            est, notes = ["bottomMd"] if family(datum) != w.prim else [], ([f"DATUM: shoe printed in {datum}; not shifted"] if family(datum) != w.prim else [])
            idd = None
            if sv:
                idd = (sv, name, "USER-DEFINED RULE: size without ID/OD designation -> innerDiameter")
                est.append("innerDiameter")
                notes.append("innerDiameter: size without ID/OD designation -> ID (user-defined rule)")
            else:
                notes.append("casing size not printed with the shoe")
            d = add_comp(w, cid=w.cid("casing-shoe"), typ="casing", name=name, where=where, bot=(dv, orig, f"m {datum} as printed"),
                         idd=idd, est=est, notes=notes)
            tv, _ = num(tvd)
            if tv is not None:
                d["bottomTvd"], d["bottomTvdM"] = ft(tv), tv
                d["source"] += f"; bottomTvd {tv} m TVD DF x {FT_PER_M} = {ft(tv)} ft (TVD as printed)"
                w.tr(f"components[{d['id']}].bottomTvd", ft(tv), where, orig, f"{tv} m TVD DF x {FT_PER_M}")
    # tubing
    tubs = S["TUBING"]
    tailpipe = any(t[4] == "tailpipe" for t in tubs)
    given = [t for t in tubs if len(t) > 5 and t[5]]
    prev_bot = None
    for t in tubs:
        name, size, wt, grade, string, bottom = (t + [""] * 6)[:6]
        where = w.where("header")
        sv, _ = num(size)
        wv, _ = num(wt)
        st = {"Short": "Short", "Long": "Long"}.get(string)
        est, notes = [], []
        idd = None
        if sv:
            idd = (sv, name, "USER-DEFINED RULE: header size without ID/OD designation -> innerDiameter")
            est.append("innerDiameter")
            notes.append("innerDiameter: header size has no ID/OD designation -> recorded as ID (user-defined rule)")
        else:
            w.issue("Missing", where, f"tubing '{name}': size not printed -> diameter empty")
        top = bot = None
        if bottom:
            bm = re.match(r"([\d.]+)\s*(.*)", bottom)
            bv = float(bm.group(1))
            bot = (bv, f"{name} {bottom}", f"m {bm.group(2) or 'datum not printed'} (tubing table depth)")
            if family(bm.group(2)) != w.prim:
                est.append("bottomMd")
                notes.append(f"DATUM: tubing table depth '{bottom}' not on the equipment datum ({S['datum']}); not shifted")
            if prev_bot is not None and string != "inner":
                top = (prev_bot[0], prev_bot[1], prev_bot[2] + "; DERIVED top = bottom of the tubing section above")
                est.append("topMd")
                notes.append("DERIVED: top = bottom of the tubing section listed above")
            prev_bot = bot
        else:
            if string == "tailpipe":
                d, how = end_item(rows, ("X",), False)
            elif string == "inner":
                d, how = end_item(rows, ("X", "L"), True)
            elif tailpipe:
                d, how = None, None
                notes.append("main tubing ends at the tailpipe crossover; accessory depth not used as a tubular endpoint")
            else:
                d, how = end_item(rows, ("S",) if st == "Short" else ("L",) if st == "Long" else ("X",), False)
            if d is not None:
                bot = (d["topMdM"], d["name"], f"{S['datum']}; depth of '{d['name']}' ({d['id']})")
                est.append("bottomMd")
                notes.append(f"DERIVED: bottom = {how} '{d['name']}' ({d['id']})")
            else:
                w.issue("Missing", where, f"tubing '{name}' ({string}): no documented end-of-string depth -> bottomMd empty")
        if top is None and string not in ("inner", "tailpipe") and (not given or t is given[0]):
            if w.prim == "THF":
                top = (0.0, "depth reference m MDTHF / THF", "ASSUMED 0 m (tubing hangs from the tubing hanger flange)")
                est.append("topMd")
                notes.append("ASSUMED: top = 0 m on the THF depth reference")
            else:
                notes.append(f"top not set: depth reference is {S['datum']}, not the tubing hanger (0 would not be supported)")
        if string in ("inner", "tailpipe"):
            notes.append(f"{string} top not printed -> empty")
        add_comp(w, cid=w.cid(("" if not st else st[0].lower() + "-") + ("inner-string-tubing" if string == "inner" else "tailpipe" if string == "tailpipe" else "tubing")),
                 typ="tubing", name=name, where=where, string=st, top=top, bot=bot, idd=idd, weight=wv, grade=grade or None, est=est, notes=notes)


def build_rest(w, dump_by_slide):
    S = w.S
    for z in S["ZONE"]:
        zone, top, bot, datum, orig, part, kind = (z + [""] * 7)[:7]
        where = w.where(part)
        tv, _ = num(top)
        bv, _ = num(bot)
        est, notes = [], []
        unit = f"m {datum} as printed"
        if family(datum) != w.prim:
            est += ["topMd", "bottomMd"]
            notes.append(f"DATUM: interval printed in {datum}; equipment datum {S['datum']}; not shifted")
        if bv is not None and tv is not None and bv < tv:
            w.issue("Conflict", where, f"zone '{orig}': bottom {bv} < top {tv} — bottom left empty (not corrected)")
            bv = None
        closest = None if kind == "perf" else "reservoir interval printed without the word perforation (interval opposite screens / sleeves); Perforation is the nearest interval type"
        add_comp(w, cid=w.cid(f"zone-{zone}"), typ="perforation", name=f"{zone} {top}-{bot} m {datum}", where=where, zone=zone,
                 top=(tv, orig, unit) if tv is not None else None, bot=(bv, orig, unit) if bv is not None else None,
                 closest=closest, est=est, notes=notes, extra=[f"original text: {orig}"])
    for p in S["PBTD"]:
        val, datum, orig = (p + [""] * 3)[:3]
        v, u = num(val)
        est = ["topMd"] if family(datum) != w.prim or u else []
        add_comp(w, cid=w.cid("pbtd"), typ="pbtd", name=orig, where=w.where("diagram"), top=(v, orig, f"m {datum or 'datum not printed'}"),
                 est=est, notes=[f"DATUM: PBTD printed as '{datum or 'no datum'}', equipment datum {S['datum']}; not shifted"] if est else [])
    if not S["PBTD"]:
        w.issue("Missing", f"slide {w.slide}", "PBTD not printed")
    for t in S["TD"]:
        val, datum, orig = (t + [""] * 3)[:3]
        v, u = num(val)
        w.info["totalDepth"] = (ft(v), w.where("header/diagram"), orig, f"{v} m x {FT_PER_M} = {ft(v)} ft; datum {datum} as printed (not shifted)")
        if family(datum) != w.prim:
            w.issue("Unclear", w.where("header/diagram"), f"TD {v} m printed in {datum}; equipment datum {S['datum']} — no datum shift applied")
    if not S["TD"]:
        w.issue("Missing", f"slide {w.slide}", "TD not printed")
    for c in S["CEMENT"]:
        name, top, bot, datum, orig = (c + [""] * 5)[:5]
        tv, _ = num(top)
        bv, _ = num(bot)
        w.plug_cement.append({"name": name, "topMd": ft(tv), "bottomMd": ft(bv) if bv is not None else None})
        w.tr(f"plugBackCement[{len(w.plug_cement) - 1}].topMd", ft(tv), w.where("diagram"), orig, f"{tv} m x {FT_PER_M} ({datum}); bottom not printed — not derived")
        w.issue("Unclear", w.where("diagram"), f"'{orig}': cement top only; bottom not printed (not derived from TD/PBTD); keys from the template sheet Plug_Back_Cement")
    for f in S["FLOAT"]:
        name, md, datum, orig = (f + [""] * 4)[:4]
        mv, _ = num(md)
        w.float_equipment.append({"casingId": None, "md": ft(mv), "type": "float shoe", "name": name})
        w.tr(f"floatEquipment[{len(w.float_equipment) - 1}].md", ft(mv), w.where("diagram"), orig, f"{mv} m x {FT_PER_M} ({datum})")
        w.issue("Unclear", w.where("diagram"), f"'{orig}' mapped to floatEquipment (no Type_Key for a float shoe); key names from the template sheet Float_Equipment; casingId null (shoe on the screen/inner-string assembly)")
    used_obj = set()
    for it in S["INT"]:
        text, depth, datum, obj = (it + [""] * 4)[:4]
        where = f"slide {w.slide} obj #{obj}" if obj.isdigit() else w.where(obj)
        if obj.isdigit():
            used_obj.add(obj)
        if re.search(r"(?i)\bretriev|recover", text):
            w.unmapped.append((where, text, "intervention note says the item was retrieved — not installed equipment"))
            continue
        typ, closest, unres = classify(text)
        if typ is None:
            w.unmapped.append((where, text, "intervention note without an equipment type"))
            continue
        dv, _ = num(depth)
        top, extra = None, []
        if dv is not None and family(datum) == w.prim:
            top = (dv, text, f"m {datum} as printed in the note")
        elif dv is not None:
            extra.append(f"depth {dv} m {datum or '(no datum)'} in the note is NOT mapped to topMd (datum differs from {S['datum']})")
            w.issue("Unclear", where, f"'{text[:90]}': depth {dv} m {datum or 'without datum'} not mapped (datum differs)")
        add_comp(w, cid=w.cid("int-" + typ.replace(" (", "-").replace(")", "")), typ=typ, name=text, where=where, top=top,
                 closest=closest, unresolved=unres, notes=["intervention recorded by a note; status as of the note date; string not stated"], extra=extra)
    for n in S["NOTE"]:
        w.unmapped.append((n[0], n[1], "note (status / history / survey / other revision)"))
    for i in S["ISSUE"]:
        w.issue(i[0], i[1], i[2])
    for s in w.S["slides"]:
        d = dump_by_slide[s]
        if d["notes"].strip() and not re.fullmatch(r"(?s)\s*(Open\s*)?\d*\s*", d["notes"]):
            w.unmapped.append((f"slide {s} speaker notes", re.sub(r"\s+", " ", d["notes"])[:400], "slide notes"))
        for o in d["objects"]:
            t = (o.get("text") or "").strip()
            if t and t != "Open" and o["id"] not in used_obj:
                w.unmapped.append((f"slide {s} obj #{o['id']}", re.sub(r"\s+", " ", t)[:400], "slide text box over/beside the picture (status / history / note)"))
            if o.get("table"):
                w.unmapped.append((f"slide {s} obj #{o['id']}", " | ".join(" / ".join(c.replace("\n", " ") for c in r) for r in o["table"])[:400],
                                   "slide table (SPM / survey / intervention history)"))
            if o.get("graphic", "") and "chart" in o.get("graphic", ""):
                w.unmapped.append((f"slide {s} obj #{o['id']}", "chart", "embedded chart — not schematic data"))
            img = (o.get("image") or "").split("/")[-1]
            if (s, img) in OFFSLIDE and not any(img in u[0] for u in w.unmapped):
                w.unmapped.append((f"slide {s} {img}", OFFSLIDE[(s, img)], "picture outside the visible slide area (hidden content)"))


def well_info(w):
    S = w.S
    for k, v in S["INFO"]:
        where = w.where("header")
        if k == "completionDate":
            iso, note = iso_date(v)
            w.info[k] = (iso or v, where, v, f"ISO date ({note})" if iso else "kept as printed (not parseable)")
            if "0CTOBER" in v.upper():
                w.issue("Unclear", where, f"completion date '{v}': '0CTOBER' read as OCTOBER")
        elif k == "xmas":
            add_comp(w, cid=w.cid("christmas-tree"), typ="christmas-tree", name=v, where=where)
        elif k == "rte":
            w.unmapped.append((where, f"RTE / drill floor to tubing hanger : {v}", "offset, not an elevation above datum -> kbElevation left empty (no datum shift)"))
        elif k == "totalDepth":
            w.unmapped.append((where, f"header TOTAL DEPTH : {v}", "header value; TD taken from the TD record (diagram/footer preferred)"))
        else:
            w.info[k] = (v, where, v, "as printed")
    surf = [c for c in S["CASING"] if re.search(r"(?i)SURF|INT\. CASING", c[0])]
    if surf:
        w.info["surfaceCasing"] = (surf[0][0], w.where("header"), surf[0][0], "SURF / INT. CASING header line as printed")


def finish(w, tpl):
    wi = {}
    for k, tv in tpl["wellInfo"].items():
        if k == "name":
            wi[k] = w.name
            w.tr("wellInfo.name", w.name, f"slide {w.slide}", w.label, "identifier rule DULANG-A<no><suffix>; original label kept in the summary")
        elif k in w.info:
            val, where, orig, method = w.info[k]
            wi[k] = val
            w.tr(f"wellInfo.{k}", val, where, orig, method)
        else:
            wi[k] = "-" if isinstance(tv, str) else None
            w.tr(f"wellInfo.{k}", wi[k], "", "", "not supported by the source -> template empty value ('-' text / null number)")
    doc = {k: (wi if k == "wellInfo" else w.comps if k == "components" else []) for k in tpl}
    doc["floatEquipment"] = w.float_equipment
    doc["plugBackCement"] = w.plug_cement
    return doc


def validate(w, doc, tpl, keys):
    back = json.loads(json.dumps(doc))
    cs = back["components"]
    sample = {c.get("name") for c in tpl["components"]} | {t.get("name") for t in tpl["formationTops"]}
    return {
        "sections and order as template": list(back) == list(tpl),
        "wellInfo keys and order as template": list(back["wellInfo"]) == list(tpl["wellInfo"]),
        "wellInfo types (number or null; string)": all((isinstance(v, (int, float)) or v is None) if isinstance(tpl["wellInfo"][k], (int, float)) else isinstance(v, str)
                                                      for k, v in back["wellInfo"].items()),
        "unsupported sections empty lists": all(back[k] == [] for k in ("surveys", "cement", "centralizers", "formations", "formationTops", "fluidContacts")),
        "approved Type_Keys only": all(c["type"] in keys for c in cs),
        "unique component ids": len({c["id"] for c in cs}) == len(cs),
        "numeric fields numeric or null": all(isinstance(c.get(k), (int, float)) or c.get(k) is None for c in cs
                                              for k in ("topMd", "bottomMd", "outerDiameter", "innerDiameter", "topMdM", "bottomMdM")),
        "ft = m x 3.28084 (3 dp) everywhere": all((c[f] == ft(c[f + "M"])) if c[f + "M"] is not None else c[f] is None for c in cs for f in ("topMd", "bottomMd")),
        "nearest-type mappings flagged": all(("type" in c.get("estimatedFields", [])) == ("actualType" in c) for c in cs),
        "flags complete (estimated -> fields + note)": all(c["estimated"] is True and c["estimatedFields"] and c["estimationNote"] for c in cs if "estimated" in c),
        "no template sample values": not any(c["name"] in sample for c in cs) and back["wellInfo"]["name"] != tpl["wellInfo"]["name"],
        "well identifier = file name": back["wellInfo"]["name"] == f"DULANG-{w.key}",
        "every component traced to the source": all(c.get("source", "").startswith(SRC_NAME) for c in cs),
        "every transcribed row mapped or listed": w.row_count + sum(1 for u in w.unmapped if u[2] == "no Type_Key could be assigned") == len(w.S["ROW"]),
    }


def main():
    local = W / "src" / NAS.name
    inputs = [local, TPL, XTPL]
    before = {str(p): sha(p) for p in inputs}
    assert before[str(local)] == NAS_SHA_AT_COPY, "local copy differs from the NAS file hashed at the start"
    nas_state = sha(NAS) if NAS.exists() else None
    tpl, keys = load_template(), type_keys()
    dumps = {int(p.stem[5:]): json.loads(p.read_text()) for p in (W / "dump").glob("slide*.json")}
    dry = "--dry" in sys.argv
    dest = (W / "dry") if dry else OUT
    dest.mkdir(exist_ok=True)
    specs = sorted((W / "specs").glob("A*.txt"), key=lambda p: (int(re.match(r"A(\d+)", p.stem).group(1)), p.stem))
    wells = [Well(parse_spec(p)) for p in specs]
    targets = {w.key: dest / f"DULANG-{w.key}.json" for w in wells}
    clash = [str(p) for p in list(targets.values()) if p.exists()]
    if clash and not dry:
        raise SystemExit("outputs exist — not overwriting without confirmation:\n" + "\n".join(clash))
    results, docs = {}, []
    for w in wells:
        well_info(w)
        rows = build_rows(w)
        build_tubulars(w, rows)
        build_rest(w, dumps)
        doc = finish(w, tpl)
        results[w.key] = validate(w, doc, tpl, keys)
        docs.append((w, doc))
    covered = sorted({s for w in wells for s in w.S["slides"]})
    for w, doc in docs:
        targets[w.key].write_text(json.dumps(doc, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    after = {str(p): sha(p) for p in inputs}
    import d_summary
    d_summary.write(W / "_completion_summary_A_pptx.md", docs, results, before, after, covered, len(dumps), NAS, NAS_SHA_AT_COPY, nas_state)
    assert before == after, "input changed!"
    for w, doc in docs:
        cs = doc["components"]
        print(f"{w.name:<16} slides {','.join(map(str, w.S['slides'])):<6} comps {len(cs):>3}  flagged {sum(1 for c in cs if c.get('estimated')):>3}  "
              f"issues {len(w.issues):>3}  unmapped {len(w.unmapped):>3}  valid {all(results[w.key].values())}")


if __name__ == "__main__":
    main()
