#!/usr/bin/env python3
"""Map 'Dulang C_2023 Updating.pptx' (NAS, read-only) into one well-schematic JSON per well (30 Oct 2026).

Inputs (read only, SHA-256 checked before and after): the PPTX, the JSON template (RTF), the component reference workbook.
Working files: output_C/_work (local copy of the PPTX, unzip, dumps, renders, OCR, EMF text). Outputs: output_C/Dulang_C*.json
and output_C/_completion_summary.md (traceability, conflicts, missing data, unmapped content). Existing outputs are never
overwritten: if a target exists, the run stops.

Rules: ft = m x 3.28084 (3 dp), original metres in topMdM/bottomMdM; depth datums kept as printed (MDTHF table depths; MDDF /
WLTHF / RKB values are never shifted — values on another datum are flagged, WLTHF-only depths are not mapped to MD);
user-defined size rule: a size with no ID/OD designation is ID (MIN ID columns are explicit ID; explicit MAX OD -> OD);
Type_Keys only from Component_Types; nearest-type mappings flagged (estimated / estimatedFields ['type'] / estimationNote /
actualType); unresolved types flagged 'UNRESOLVED'; no sample values; tubular tops of 0 only from the THF depth reference
(flagged); no TD/PBTD substitution for tubular ends.
"""
import datetime as dt
import hashlib
import json
import os
import re
import subprocess
import sys
from pathlib import Path

from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parent))
from c_header import header  # noqa: E402
from c_rules import ANNOT, FT_PER_M, SRC_NAME, classify, frac_in, ft, iso_date  # noqa: E402
from c_table_wells import num, pair_slide, zone_records  # noqa: E402
import c_image_wells  # noqa: E402

W = Path(__file__).resolve().parent
OUT = W.parent
ROOT = OUT.parent
NAS = Path("/Volumes/homes/public/DulangOSS/20261022/WellSchematic_Dulang-C/Dulang C_2023 Updating.pptx")
TPL = ROOT / "well-schematic-template-json.rtf"
XTPL = ROOT / "well-schematic-sample-template.xlsx"
SLIDE_WELL = {1: "C1", 2: "C2", 3: "C3ST1", 4: "C4ST1", 5: "C5", 6: "C6", 7: "C7", 8: "C8", 9: "C9", 10: "C10", 11: "C11",
              12: "C12", 13: "C13", 14: "C14", 15: "C15", 16: "C16", 17: "C17", 18: "C18", 19: "C19", 20: "C20", 21: "C21",
              22: "C22ST2", 23: "C23", 24: "C24", 25: "C25", 26: "C26", 27: "C27", 28: "C28", 29: "C29", 31: "C31"}
IMAGE_SLIDES = {3, 16, 22, 25, 29, 31}


def sha(p):
    h = hashlib.sha256()
    fd = os.open(p, os.O_RDONLY)
    while (b := os.read(fd, 1 << 20)):
        h.update(b)
    try:
        os.close(fd)
    except OSError:  # SMB handles can report EBADF on close
        pass
    return h.hexdigest()


def load_template():
    txt = subprocess.run(["textutil", "-convert", "txt", "-stdout", str(TPL)], capture_output=True, text=True, check=True).stdout
    return json.loads(txt)


def type_keys():
    ws = load_workbook(XTPL, read_only=True)["Component_Types"]
    return {r[2] for r in ws.iter_rows(min_row=2, values_only=True) if r[2]}


class Well:
    def __init__(self, key, slide, label):
        self.key, self.slide, self.label = key, slide, label
        self.name = f"Dulang_{key}"
        self.comps, self.trace, self.issues, self.unmapped = [], [], [], []
        self.ids = {}

    def cid(self, base):
        base = re.sub(r"[^a-z0-9]+", "-", base.lower()).strip("-")
        n = self.ids.get(base, 0) + 1
        self.ids[base] = n
        return f"{base}-{n}"

    def tr(self, path, value, where, original, method):
        self.trace.append(dict(path=path, value=value, where=where, original=original, method=method))

    def issue(self, kind, where, text):
        self.issues.append((kind, where, text))


def loc(slide, obj):
    return f"slide {slide} obj #{obj}"


def add_comp(w, *, cid, typ, name, name_src, string=None, top=None, bot=None, idd=None, od=None, weight=None, grade=None,
             zone=None, closest=None, unresolved=False, est=None, notes=(), src_extra=(), status=None):
    """top/bot: (metres, where, original, unit-note) ; idd/od: (inches, where, original, method)"""
    p = f"components[{cid}]"
    d = {"id": cid, "type": typ, "name": name, "topMd": None, "bottomMd": None, "outerDiameter": None, "innerDiameter": None,
         "topMdM": None, "bottomMdM": None}
    w.tr(f"{p}.name", name, name_src[0], name_src[1], name_src[2])
    w.tr(f"{p}.type", typ, name_src[0], name, "NEAREST TYPE (flagged)" if closest else "EXACT TYPE (Component_Types)")
    src = [f"{SRC_NAME} > {name_src[0]} name '{name}' ({name_src[2]})"]
    estf, enote = list(est or []), list(notes)
    if closest:
        estf.insert(0, "type")
        enote.insert(0, ("UNRESOLVED type: " if unresolved else "type: ") + closest)
    for key, field in ((top, "topMd"), (bot, "bottomMd")):
        if key is None:
            continue
        m, where, orig, unit = key
        if m is None:
            continue
        m = float(m)
        d[field], d[field + "M"] = ft(m), m
        w.tr(f"{p}.{field}", d[field], where, orig, f"{m} m x {FT_PER_M} = {d[field]} ft (3 dp); {unit}")
        w.tr(f"{p}.{field}M", m, where, orig, f"original metres; {unit}")
        src.append(f"{field} {where} '{orig}' = {m} m -> {d[field]} ft ({unit})")
    for key, field in ((idd, "innerDiameter"), (od, "outerDiameter")):
        if key is None or key[0] is None:
            continue
        v, where, orig, method = key
        d[field] = v
        w.tr(f"{p}.{field}", v, where, orig, method)
        src.append(f"{field} {where} '{orig}' ({method})")
    if weight is not None:
        d["weight"] = weight[0]
        w.tr(f"{p}.weight", weight[0], weight[1], weight[2], "parsed from the label (ppf = lb/ft)")
    if grade is not None:
        d["grade"] = grade[0]
        w.tr(f"{p}.grade", grade[0], grade[1], grade[2], "parsed from the label")
    if zone:
        d["zone"] = zone
        w.tr(f"{p}.zone", zone, name_src[0], zone, "sand label as printed")
    if string:
        d["string"] = string
        w.tr(f"{p}.string", string, name_src[0], "", "column of the table (SHORT STRING / LONG STRING) or DUAL packer label")
    if status:
        src.append(f"status {status[0]} '{status[1]}'")
    src += list(src_extra)
    if estf:
        d.update(estimated=True, estimatedFields=estf, estimationNote="; ".join(enote))
        if "type" in estf:
            d["actualType"] = name
    elif enote:
        src.append("NOTE: " + "; ".join(enote))
    d["source"] = "; ".join(src)
    w.comps.append(d)
    return d


# ------------------------------------------------------------------ table-style slides
def depth_of(cell, side_unit="m MDTHF (column header 'DEPTH m MDTHF')"):
    v = num(cell)
    return v


def build_table_well(w, dump):
    s = w.slide
    P = pair_slide(s)
    H = header(dump)
    single = s == 13
    seq = {"S": [], "L": []}
    end_items = {"S": None, "L": None}
    for r in sorted(P["rows"], key=lambda r: (r["side"], r["parts"][0]["y"])):
        parts, side = r["parts"], r["side"]
        text = " / ".join(re.sub(r"[ \t]+", " ", p["text"].strip()).replace("\n", " / ") for p in parts)
        ids = "+".join(f"#{p['id']}" for p in parts)
        where = f"slide {s} obj {ids}"
        if r["kind"] == "annotation":
            if not re.search(r"(?i)CONTROL\s*LINE", text):
                w.unmapped.append((where, text, r.get("override") or "annotation in the equipment-name column (not equipment)"))
            else:
                w.unmapped.append((where, text, "control line (no Type_Key; not a tubular component)"))
            continue
        if r["kind"] == "intervention-row":
            continue  # handled with the other interventions below
        typ, closest, unres = classify(text)
        if r["kind"] == "spm-label":
            typ, closest, unres = "spm", r["override"], False
        string = None if single else ("Short" if side == "S" else "Long")
        if typ == "packer":
            if re.search(r"(?i)DUAL", text):
                string = "Short & Long"
            elif re.search(r"(?i)SINGLE|PERM", text) and not single:
                string = "Long"
        dcell = r.get("depth")
        dv = num(dcell)
        notes, est, extra = [], [], []
        top = None
        if dcell is not None and dv is None:
            extra.append(f"depth cell #{dcell['id']} '{dcell['text']}' is not a number (shared packer label / text) -> no depth")
        elif dcell is not None and dv == 0.0:
            w.issue("Missing", where, f"'{text}': depth printed as '0.0' (placeholder) -> topMd left empty")
            extra.append(f"depth cell #{dcell['id']} '0.0' treated as a placeholder -> topMd empty")
        elif dcell is not None:
            top = (dv, f"slide {s} obj #{dcell['id']}", dcell["text"].strip(), "m MDTHF (table column 'DEPTH m MDTHF')")
        bot = None
        if typ == "blast-joint" and top is None:
            ys = parts[0]["y"]
            ups = [c for c in P["unpaired"] if c["col"] == f"DEPTH_{side}" and num(c) and 0 < ys - c["y"] <= 3.5]
            dns = [c for c in P["unpaired"] if c["col"] == f"DEPTH_{side}" and num(c) and 0 <= c["y"] - ys <= 3.5]
            if ups and dns:
                u, dn = max(ups, key=lambda c: c["y"]), min(dns, key=lambda c: c["y"])
                if num(u) and num(dn) and num(u) < num(dn):
                    top = (num(u), f"slide {s} obj #{u['id']}", u["text"].strip(), "m MDTHF; depth printed at the top end of the blast joint drawing")
                    bot = (num(dn), f"slide {s} obj #{dn['id']}", dn["text"].strip(), "m MDTHF; depth printed at the bottom end of the blast joint drawing")
                    P["unpaired"] = [c for c in P["unpaired"] if c["id"] not in (u["id"], dn["id"])]
                    est += ["topMd", "bottomMd"]
                    notes.append("INTERPRETED: unlabelled depths printed just above/below the BLAST JOINT label at the ends of its drawing taken as its top/bottom")
        icell = r.get("minid")
        idd = None
        if icell is not None:
            iv = num(icell)
            if iv is not None and iv > 0:
                idd = (iv, f"slide {s} obj #{icell['id']}", icell["text"].strip(), "column 'MIN ID (in)' = explicit ID")
        status = None
        if r.get("status"):
            status = (f"slide {s} obj #{r['status']['id']}", re.sub(r"\s+", " ", r["status"]["text"]).strip())
        if typ is None:
            w.unmapped.append((where, text, "no Type_Key could be assigned"))
            continue
        cid = w.cid(("" if single or string in (None, "Short & Long") else string[0].lower() + "-") + (typ.replace(" (", "-").replace(")", "")))
        d = add_comp(w, cid=cid, typ=typ, name=text, name_src=(where, text, "native slide text"), string=string, top=top, bot=bot,
                     idd=idd, closest=closest, unresolved=unres, est=est, notes=notes, src_extra=extra, status=status)
        y0 = parts[0]["y"]
        if string in ("Short", "Long") or single:
            seq["S" if string == "Short" else "L"].append((y0, d))
            k = "S" if string == "Short" else "L"
            if typ == "eot" or (typ == "mule-shoe" and (end_items[k] is None or end_items[k]["type"] != "eot")):
                end_items[k] = d  # documented EOT label preferred over the shoe/guide at the string end
        elif string == "Short & Long":
            seq["S"].append((y0, d))
            seq["L"].append((y0, d))
    # sequence check (depth must increase down each string)
    for side, lst in seq.items():
        prev = None
        for _, d in sorted(lst, key=lambda t: t[0]):
            if d["topMdM"] is None:
                continue
            if prev and d["topMdM"] < prev["topMdM"]:
                w.issue("Conflict", d["id"], f"depth {d['topMdM']} m is shallower than the item above ({prev['id']} {prev['topMdM']} m) "
                        f"on the {'short' if side == 'S' else 'long'} string — values kept as printed")
                d["source"] += f"; CONFLICT: out of depth sequence (above item {prev['id']} at {prev['topMdM']} m)"
            prev = d
    zones(w, P)
    interventions(w, P)
    tubulars(w, H, end_items, single)
    well_info_table(w, H, P)
    for c in P["unpaired"]:
        t = re.sub(r"\s+", " ", c["text"]).strip()
        if num(c) is not None:
            w.unmapped.append((f"slide {s} obj #{c['id']}", t, f"depth value in column {c['col']} with no equipment label on its row -> not mapped"))
        elif not re.search(r"(?i)^\(in\)$|^m MDTHF$|^PKR[- ]?\d$", t):
            w.unmapped.append((f"slide {s} obj #{c['id']}", t, f"text in {c['col']} (annotation / status / reservoir property)"))
    return w


def zones(w, P):
    s = w.slide
    recs, other = zone_records(P["zones"])
    labels = [(z["y"], re.search(r"(?i)SAND\s*:\s*([^\n(]+)", z["text"])) for z in P["zones"]]
    by_obj = {z["id"]: z for z in P["zones"]}
    td = None
    for c in P["unpaired"]:
        m = re.search(r"TD\s*@\s*(\d+(?:\.\d+)?)", c["text"])
        if m:
            td = float(m.group(1))
    for r in recs:
        zone = r["zone"]
        if zone is None:
            y = by_obj[r["obj"]]["y"]
            above = [(yy, m.group(1).strip()) for yy, m in labels if m and yy <= y + 0.5 and y - yy <= 6]
            if above:
                zone = re.sub(r"\s+SAND$", "", max(above)[1], flags=re.I).strip()
                r["zone_note"] = "zone taken from the SAND label box directly above the PERF text"
        zone_s = re.sub(r"\s+", "", zone) if zone else None
        where = f"slide {s} obj #{r['obj']}"
        top, bot = r["top"], r["bot"]
        notes, est = [], []
        line = r["line"]
        datum = "m MDTHF" if re.search(r"(?i)MDTHF|MD-THF", line) or re.search(r"(?i)MDTHF", r["box"]) else None
        unit = "m-MDTHF as printed" if datum else "datum not printed in the perforation text; table depths are m MDTHF (assumed same reference)"
        bad = []
        lim = (td or 4000) * 1.2
        if top > lim:
            bad.append(("topMd", top))
            top = None
        if bot > lim or (top is not None and bot < top):
            bad.append(("bottomMd", bot))
            bot = None
        for f, v in bad:
            w.issue("Conflict", where, f"perforation '{line.strip()}': {f} printed as {v} m is impossible (beyond TD / below top) -> left empty, not corrected")
        cid = w.cid(f"perf-{zone_s or 'unlabelled'}")
        add_comp(w, cid=cid, typ="perforation", name=f"{zone} perforation {r['top']} - {r['bot']} m" if zone else f"perforation {r['top']} - {r['bot']} m",
                 name_src=(where, line.strip(), "DERIVED name: sand label + interval from the perforation text"), zone=zone,
                 top=(top, where, line.strip(), unit) if top is not None else None,
                 bot=(bot, where, line.strip(), unit) if bot is not None else None,
                 notes=[r.get("zone_note")] if r.get("zone_note") else [], src_extra=[f"zone box text: {re.sub(chr(10), ' / ', r['box'])[:200]}"])
        if not datum:
            w.issue("Unclear", where, f"perforation '{line.strip()}' has no datum; assumed m MDTHF like the depth columns")
    for obj, o in other:
        w.unmapped.append((f"slide {s} obj #{obj}", o, "zone box line (pressure survey / reservoir note) — no template field"))


INTERV = re.compile(r"(?i)(plug|storm choke|insert valve|\bIV\b|TRSV)")


def interventions(w, P):
    """plugs / insert valves / storm chokes written as notes (status column, margins or inside an equipment row)"""
    s = w.slide
    rows = [r for r in P["rows"] if r["kind"] == "intervention-row"]
    cand = [(r["parts"][0], r) for r in rows]
    cand += [(c, None) for c in P["unpaired"] if INTERV.search(c["text"]) and not re.search(r"(?i)retriev|recover|reset", c["text"])]
    # status cells already attached to equipment rows may also carry plugs
    for r in P["rows"]:
        st = r.get("status")
        if st and INTERV.search(st["text"]) and re.search(r"(?i)plug", st["text"]):
            cand.append((st, r))
    done = set()
    for c, row in cand:
        if c["id"] in done:
            continue
        done.add(c["id"])
        t = re.sub(r"\s+", " ", c["text"]).strip()
        if re.search(r"(?i)suspect", t):
            w.unmapped.append((f"slide {s} obj #{c['id']}", t, "suspicion / diagnosis note (not confirmed installed equipment)"))
            continue
        if re.search(r"(?i)retriev|recover|reset|Retrieve|absorber", t):
            w.unmapped.append((f"slide {s} obj #{c['id']}", t, "plug/valve note that says it was retrieved / reset — not installed equipment"))
            continue
        typ, closest, unres = classify(t)
        if typ is None or typ not in ("mechanical-plug", "bridge-plug", "storm-choke", "wrsv", "trsv", "scssv"):
            w.unmapped.append((f"slide {s} obj #{c['id']}", t, "intervention note without an equipment type"))
            continue
        where = f"slide {s} obj #{c['id']}"
        m = re.search(r"@\s*@?\s*(\d+(?:\.\d+)?)\s*m?\s*[- ]?\s*(MDTHF|WLTHF|THF|MDRKB|MDDF|MD)?", t) or \
            re.search(r"(\d{3,4}(?:\.\d+)?)\s*m\s*-?\s*(MDTHF|WLTHF|THF|MDRKB|MDDF)", t)
        top, notes, est, extra = None, [], [], []
        dep = (row or {}).get("override") or ""
        if row is not None and dep.startswith("depth:") and dep[6:] in P["by"]:
            dc = P["by"][dep[6:]]
            top = (num(dc), f"slide {s} obj #{dc['id']}", dc["text"].strip(), "m MDTHF (table column 'DEPTH m MDTHF'; value printed beside the note)")
            P["unpaired"] = [x for x in P["unpaired"] if x["id"] != dc["id"]]
        elif m:
            v, datum = float(m.group(1)), (m.group(2) or "").upper()
            if datum in ("MDTHF", "THF"):
                top = (v, where, t, "m MDTHF as printed in the note")
            else:
                extra.append(f"depth {v} m {datum or '(no datum)'} in the note is NOT mapped to topMd (wireline/other datum, not MD-THF)")
                w.issue("Unclear", where, f"'{t}': depth {v} m {datum or 'without datum'} not converted to MD (datum not MDTHF)")
        elif row is not None:
            host = row.get("depth")
            hosts = [r for r in P["rows"] if r["side"] == row["side"] and r["kind"] == "equip" and abs(r["parts"][0]["y"] - c["y"]) <= 1.0 and r.get("depth") and num(r["depth"])]
            if hosts:
                h = min(hosts, key=lambda r: abs(r["parts"][0]["y"] - c["y"]))
                top = (num(h["depth"]), f"slide {s} obj #{h['depth']['id']}", h["depth"]["text"], "m MDTHF; depth of the equipment row the note is written in")
                est.append("topMd")
                notes.append(f"DERIVED: depth of host row '{re.sub(chr(10), ' ', h['parts'][0]['text'])}' (note has no own depth)")
        string = None
        if row is not None and w.slide != 13:
            string = "Short" if row["side"] == "S" else "Long"
        elif c.get("col", "").endswith("_S") or c.get("col") == "LEFT_MARGIN":
            string = "Short" if w.slide != 13 else None
        elif c.get("col", "").endswith("_L") or c.get("col") == "RIGHT_MARGIN":
            string = "Long" if w.slide != 13 else None
        if string and row is None:
            est.append("string") if False else None
            notes.append(f"string inferred from the note position ({c.get('col')})")
        cid = w.cid(("" if not string else string[0].lower() + "-") + typ.replace(" (", "-").replace(")", ""))
        add_comp(w, cid=cid, typ=typ, name=t, name_src=(where, t, "native slide text (intervention note)"), string=string, top=top,
                 closest=closest, unresolved=unres, est=est, notes=notes + ["intervention recorded by note; status as of the note date"], src_extra=extra)


def parse_size_line(t):
    size = frac_in(t)
    wt = re.search(r"(\d+(?:\.\d+)?)\s*(?:ppf|PPF|#)", t)
    gr = re.findall(r"\b([KNLJX]-?\s?\d{2,3})\b", t)
    return size, (float(wt.group(1)) if wt else None), gr


def tubulars(w, H, end_items, single):
    s = w.slide
    # casing lines
    for c in H.get("casing", []):
        line = c["value"]
        where = f"slide {s} obj #{c['obj']}"
        size, wt, gr = parse_size_line(line)
        m = re.search(r"@\s*(\d+(?:\.\d+)?)\s*(m)?\s*-?\s*(MDDF|MDTHF|MD)", line)
        bot, est, notes = None, [], []
        if m:
            v, datum = float(m.group(1)), m.group(3)
            unit = f"m {datum} as printed"
            if not m.group(2):
                unit += " (unit 'm' not printed; metres assumed like the other casing lines)"
                est.append("bottomMd")
                notes.append("ASSUMED: unit m (not printed)")
            if datum != "MDTHF":
                est.append("bottomMd") if "bottomMd" not in est else None
                notes.append(f"DATUM: shoe depth printed in {datum} (drill floor), not MDTHF like the equipment table; not shifted")
            bot = (v, where, line, unit)
        typ = "liner" if re.search(r"(?i)LINER", line) else "casing"
        idd = (size, where, line, "USER-DEFINED RULE: header size without ID/OD designation -> innerDiameter") if size else None
        grade = None
        if len(set(g.replace(" ", "") for g in gr)) == 1:
            grade = (gr[0].replace(" ", ""), where, line)
        elif len(gr) > 1:
            w.issue("Conflict", where, f"casing line '{line}' gives more than one grade {gr} -> grade left empty")
        base = "conductor" if re.search(r"(?i)COND|DRIVE", line) else "surface-casing" if re.search(r"(?i)SURF", line) else "production-casing" if re.search(r"(?i)PROD", line) else typ
        add_comp(w, cid=w.cid(base), typ=typ, name=line, name_src=(where, line, "native slide text (casing header line)"), bot=bot,
                 idd=idd, weight=(wt, where, line) if wt else None, grade=grade, est=est + (["innerDiameter"] if size else []),
                 notes=notes + (["innerDiameter: header size has no ID/OD designation -> recorded as ID (user-defined rule)"] if size else []))
        if bot is None:
            w.issue("Missing", where, f"casing '{line}': setting depth not printed")
    # tubing (header)
    tub = H.get("tubing")
    if tub:
        t = tub[0]["value"]
        where = f"slide {s} obj #{tub[0]['obj']}"
        size, wt, gr = parse_size_line(t)
        strings = [None] if single else ["Short", "Long"]
        for st in strings:
            end = end_items["S" if st == "Short" else "L"] if st else end_items["L"]
            bot = None
            est, notes = ["topMd", "innerDiameter"] if size else ["topMd"], ["ASSUMED: top = 0 m MDTHF (depth reference is the tubing hanger flange; tubing hangs from the THF)"]
            if end is not None and end["topMdM"] is not None:
                bot = (end["topMdM"], f"component {end['id']}", end["name"], f"m MDTHF; depth of the end-of-string item '{end['name']}'")
                est.append("bottomMd")
                notes.append(f"DERIVED: bottom = depth of the string-end item '{end['name']}' ({end['id']})")
            else:
                w.issue("Missing", where, f"{st or 'tubing'} string: no end-of-string depth printed -> tubing bottomMd empty")
            if size:
                notes.append("innerDiameter: header size has no ID/OD designation -> recorded as ID (user-defined rule)")
            add_comp(w, cid=w.cid(("" if not st else st[0].lower() + "-") + "tubing"), typ="tubing", name=t,
                     name_src=(where, t, "native slide text (TUBING header value)"), string=st,
                     top=(0.0, where, "depth reference m MDTHF", "ASSUMED 0 m MDTHF"), bot=bot,
                     idd=(size, where, t, "USER-DEFINED RULE: size without ID/OD designation -> innerDiameter") if size else None,
                     weight=(wt, where, t) if wt else None, grade=(gr[0].replace(" ", ""), where, t) if gr else None, est=est, notes=notes)
    # X-mas tree
    for x in H.get("xmas", [])[:1]:
        where = f"slide {s} obj #{x['obj']}"
        add_comp(w, cid=w.cid("christmas-tree"), typ="christmas-tree", name=x["value"], name_src=(where, x["value"], "native slide text (X-MAS TREE header value)"))


def well_info_table(w, H, P):
    s = w.slide
    w.info = {}
    title = H.get("title", [{}])[0]
    if title:
        w.info["completionType"] = (re.sub(r"(?i)^WELL C-\S+\s*:\s*", "", title["value"]).strip(), f"slide {s} obj #{title['obj']}", title["value"], "text after 'WELL C-n :'")
    if H.get("date"):
        d = H["date"][0]
        iso, note = iso_date(d["value"])
        w.info["completionDate"] = (iso or d["value"], f"slide {s} obj #{d['obj']}", d["value"], f"ISO date {note}".strip())
    if H.get("rig"):
        w.info["rig"] = (H["rig"][0]["value"], f"slide {s} obj #{H['rig'][0]['obj']}", H["rig"][0]["value"], "DIRECT")
    if H.get("tubing"):
        w.info["tubing"] = (H["tubing"][0]["value"], f"slide {s} obj #{H['tubing'][0]['obj']}", H["tubing"][0]["value"], "DIRECT")
    if H.get("packer_fluid"):
        v = " ".join(re.sub(r"\s+", " ", p["value"]) for p in H["packer_fluid"])
        w.info["packerFluid"] = (v, f"slide {s} obj " + "+".join(f"#{p['obj']}" for p in H["packer_fluid"]), v, "DIRECT (lines joined)")
    if H.get("maxdev") and re.search(r"\d", H["maxdev"][0]["value"]):
        w.info["maxDeviation"] = (H["maxdev"][0]["value"], f"slide {s} obj #{H['maxdev'][0]['obj']}", H["maxdev"][0]["value"], "DIRECT")
    surf = [c for c in H.get("casing", []) if re.search(r"(?i)SURF", c["value"])]
    if surf:
        w.info["surfaceCasing"] = (surf[0]["value"], f"slide {s} obj #{surf[0]['obj']}", surf[0]["value"], "DIRECT (SURF casing line)")
    # TD
    for c in P["unpaired"] + [z for z in P["zones"]]:
        m = re.search(r"TD\s*@?\s*(\d+(?:\.\d+)?)\s*m\s*-?\s*(MDDF|MDTHF)", c["text"])
        if m:
            v = float(m.group(1))
            w.info["totalDepth"] = (ft(v), f"slide {s} obj #{c['id']}", c["text"].strip(), f"{v} m x {FT_PER_M} = {ft(v)} ft; datum {m.group(2)} as printed (not shifted)")
            w.td = (v, m.group(2))
            if m.group(2) != "MDTHF":
                w.issue("Unclear", f"slide {s} obj #{c['id']}", f"TD printed in {m.group(2)} ({v} m); equipment depths are MDTHF — no datum shift applied")
    if "totalDepth" not in w.info:
        w.issue("Missing", f"slide {s}", "TD value not printed")
    # PBTD: header value and/or diagram label
    hdr = H.get("pbtd", [])
    hv = None
    for p in hdr:
        m = re.search(r"(\d+(?:\.\d+)?)\s*m\s*(MDDF|MDTHF)", p["value"])
        if m:
            hv = (float(m.group(1)), m.group(2), f"slide {s} obj #{p['obj']}", p["value"])
    dg = None
    for c in P["unpaired"]:
        m = re.search(r"(?i)PBTD\s*:?\s*\n?\s*(\d+(?:\.\d+)?)\s*(m)?", c["text"])
        if m:
            dg = (float(m.group(1)), "m (datum not printed)" if m.group(2) else "unit/datum not printed", f"slide {s} obj #{c['id']}", re.sub(r"\s+", " ", c["text"]).strip())
            P["unpaired"] = [x for x in P["unpaired"] if x["id"] != c["id"]]
    use = dg or hv
    if use:
        notes = []
        if dg and hv and abs(dg[0] - hv[0]) > 0.05:
            w.issue("Conflict", dg[2], f"PBTD on the diagram {dg[0]} vs header {hv[0]} m {hv[1]} ({hv[2]}) -> diagram value used (explicit diagram value has priority), both recorded")
            notes.append(f"CONFLICT: header PBTD {hv[0]} m {hv[1]} ({hv[2]})")
        elif dg and hv:
            notes.append(f"header PBTD {hv[0]} m {hv[1]} ({hv[2]}) agrees")
        unit = ("m " + hv[1] + " (header)") if use is hv else (dg[1] + (f"; header gives {hv[1]}" if hv else ""))
        if use is hv and hv[1] != "MDTHF":
            notes.append(f"DATUM: PBTD printed in {hv[1]}, not MDTHF; not shifted")
        add_comp(w, cid=w.cid("pbtd"), typ="pbtd", name=use[3], name_src=(use[2], use[3], "native slide text"), top=(use[0], use[2], use[3], unit),
                 est=["topMd"] if (use is hv and hv[1] != "MDTHF") or use is dg and not hv else [], notes=notes + (["DATUM: diagram PBTD label has no datum"] if use is dg and not hv else []))
    else:
        w.issue("Missing", f"slide {s}", "PBTD value not printed")
    rte = H.get("rte")
    if rte:
        w.unmapped.append((f"slide {s} obj #{rte[0]['obj']}", f"RTE to tubing hanger {rte[0]['value']}", "rotary-table-to-THF offset, not an elevation above datum -> kbElevation left empty (no datum shift applied)"))
    for k in ("note", "other"):
        for o in H.get(k, []):
            w.unmapped.append((f"slide {s} obj #{o['obj']}", o["value"], "header note"))


def finish(w, tpl):
    wi = {}
    for k, tv in tpl["wellInfo"].items():
        if k == "name":
            wi[k] = w.name
            w.tr("wellInfo.name", w.name, f"slide {w.slide}", w.label, "file/well identifier rule Dulang_C<no><suffix>; original label kept in summary")
        elif k in getattr(w, "info", {}):
            val, where, orig, method = w.info[k]
            wi[k] = val
            w.tr(f"wellInfo.{k}", val, where, orig, method)
        else:
            wi[k] = "-" if isinstance(tv, str) else None
            w.tr(f"wellInfo.{k}", wi[k], "", "", "not supported by the source -> template empty value ('-' text / null number)")
    doc = {}
    for k, v in tpl.items():
        doc[k] = wi if k == "wellInfo" else (w.comps if k == "components" else [])
    if getattr(w, "float_equipment", None):
        doc["floatEquipment"] = w.float_equipment
    return doc


def main():
    inputs = [NAS, TPL, XTPL]
    before = {str(p): sha(p) for p in inputs}
    tpl = load_template()
    keys = type_keys()
    dumps = {int(p.stem[5:]): json.loads(p.read_text()) for p in (W / "dump").glob("slide*.json")}
    dry = "--dry" in sys.argv  # development run: writes to _work/dry (overwritable), never to output_C
    dest = (W / "dry") if dry else OUT
    dest.mkdir(exist_ok=True)
    targets = {k: dest / f"Dulang_{k}.json" for k in SLIDE_WELL.values()}
    clash = [str(p) for p in targets.values() if p.exists()] + ([str(dest / "_completion_summary.md")] if (dest / "_completion_summary.md").exists() else [])
    if clash and not dry:
        raise SystemExit("outputs exist — not overwriting without confirmation:\n" + "\n".join(clash))
    wells = []
    for slide, key in sorted(SLIDE_WELL.items()):
        label = c_image_wells.LABELS.get(slide) or next((re.sub(r"\s+", " ", o["text"]).strip() for o in dumps[slide]["objects"]
                                                          if o.get("text") and re.match(r"WELL C-", o["text"])), key)
        w = Well(key, slide, label)
        if slide in IMAGE_SLIDES:
            c_image_wells.build(w, dumps[slide], add_comp)
        elif slide == 17:
            w.info = {"completionType": ("P & A", "slide 17 obj #38914", "WELL C-17 :  P & A", "text after 'WELL C-17 :'")}
            w.issue("Missing", "slide 17", "slide holds only the title 'WELL C-17 : P & A' — no schematic, header, depths or equipment")
        else:
            build_table_well(w, dumps[slide])
        wells.append((w, finish(w, tpl)))
    c_image_wells.post_checks(wells)
    results = validate(wells, tpl, keys)
    for (w, doc) in wells:
        targets[w.key].write_text(json.dumps(doc, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    after = {str(p): sha(p) for p in inputs}
    import c_summary
    c_summary.write(dest / "_completion_summary.md", wells, results, before, after, dumps)
    assert before == after, "input changed!"
    for w, doc in wells:
        print(f"{w.name:<16} slide {w.slide:>2}  comps {len(doc['components']):>3}  flagged {sum(1 for c in doc['components'] if c.get('estimated')):>3}  "
              f"issues {len(w.issues):>3}  unmapped {len(w.unmapped):>3}  valid {all(results[w.key].values())}")


def validate(wells, tpl, keys):
    res = {}
    for w, doc in wells:
        back = json.loads(json.dumps(doc))
        cs = back["components"]
        chk = {}
        chk["sections and order as template"] = list(back) == list(tpl)
        chk["wellInfo keys and order as template"] = list(back["wellInfo"]) == list(tpl["wellInfo"])
        chk["wellInfo types (number or null; string)"] = all(
            (isinstance(v, (int, float)) or v is None) if isinstance(tpl["wellInfo"][k], (int, float)) else isinstance(v, str)
            for k, v in back["wellInfo"].items())
        chk["unsupported sections empty lists"] = all(isinstance(back[k], list) for k in tpl if k not in ("wellInfo",)) and back["surveys"] == [] \
            and back["cement"] == [] and back["formationTops"] == [] and back["fluidContacts"] == [] and back["formations"] == []
        chk["approved Type_Keys only"] = all(c["type"] in keys for c in cs)
        chk["unique component ids"] = len({c["id"] for c in cs}) == len(cs)
        chk["numeric fields numeric or null"] = all(isinstance(c.get(k), (int, float)) or c.get(k) is None
                                                    for c in cs for k in ("topMd", "bottomMd", "outerDiameter", "innerDiameter", "topMdM", "bottomMdM"))
        chk["ft = m x 3.28084 (3 dp) everywhere"] = all(c[f] == ft(c[f + "M"]) for c in cs for f in ("topMd", "bottomMd") if c[f + "M"] is not None) \
            and all(c[f] is None for c in cs for f in ("topMd", "bottomMd") if c[f + "M"] is None)
        chk["nearest-type mappings flagged"] = all(("type" in c.get("estimatedFields", [])) == ("actualType" in c) for c in cs)
        chk["flags complete (estimated -> fields + note)"] = all(c.get("estimated") is True and c["estimatedFields"] and c["estimationNote"]
                                                                 for c in cs if "estimated" in c)
        sample_names = {c.get("name") for c in tpl["components"]} | {t.get("name") for t in tpl["formationTops"]}
        chk["no template sample values"] = not any(c["name"] in sample_names for c in cs) and back["wellInfo"]["name"] != tpl["wellInfo"]["name"] \
            and back["surveys"] == []
        chk["well identifier = file name"] = back["wellInfo"]["name"] == f"Dulang_{w.key}"
        chk["every component traced to the source"] = all(c.get("source", "").startswith(SRC_NAME) for c in cs)
        res[w.key] = chk
    return res


if __name__ == "__main__":
    main()
