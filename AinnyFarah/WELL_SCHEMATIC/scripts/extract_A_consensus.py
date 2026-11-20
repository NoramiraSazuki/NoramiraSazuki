#!/usr/bin/env python3
"""Exact-match, multi-pass-consensus extraction of the Dulang A well-schematic PDF -> one JSON per well.

Inputs: output_A/ocr_passes (three local Tesseract passes per page + exact PDF text layer), made by ocr_passes_A.py.
A value is written only when the identical text is read in >= 2 independent sources (OCR passes a/b/c, text layer).
Nothing is inferred, converted or corrected; unreadable values stay empty (template convention: "" / null / []).
Component 'type' uses the nearest template type key (user decision) and is flagged estimated when not exact.
"""
import csv
import io
import json
import re
import subprocess
from collections import Counter, defaultdict
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "output_A"
PASS = OUT / "ocr_passes"
TPL = ROOT / "well-schematic-template-json.rtf"
XTPL = ROOT / "well-schematic-sample-template.xlsx"
PASSES = ("a", "b", "c")
DEPTH_RX = re.compile(r"(?<![\d./])(\d{2,4}(?:\.\d{1,2})?)(?![\d./])")
DIM_RX = re.compile(r"(?<![\d.])(\d\.\d{3})(?![\d])")
# equipment keyword -> (Component_Types name, exact?)  checked in order
EQUIP = [
    (r"FLOW\s+COUPLING", "Flow Coupling", True), (r"TRSV|TR-SC-SSSV|TRSCSSV|TRSCSSSV|SCSSV|SCSSI", "SCSSV", False),
    (r"NO\s*-?\s*GO\s+NIPPLE|NOGO\s+NIPPLE", "No-Go Nipple", True), (r"X-?\s*NIPPLE", "Nipple (X)", True),
    (r"NIPPLE", "Landing Nipple", False), (r"\bSPM\b", "SPM (Gas Lift)", True), (r"\bSSD\b|SLIDING\s+SLEEVE", "Sliding Sleeve", False),
    (r"\bPKR\b|PACKER", "Packer", True), (r"BLAST\s+JOINT", "Blast Joint", True), (r"ENTRY\s+GUIDE|\bWEG\b", "Mule Shoe", False),
    (r"MULE\s*SHOE", "Mule Shoe", True), (r"SEAL\s+ASS(?:Y|EMBLY)", "Seal Assembly", True), (r"SEAL\s+BORE", "PBR", False),
    (r"\bCONICAL\b", "Mule Shoe", False), (r"SCREEN", "Screen", True), (r"PERFORATED\s+PUP", "Pup Joint", False),
    (r"CENTRALI[ZS]ER", "Pup Joint", False), (r"HYDRO\s+TRIP\s+SUB|\bSUB\b", "Crossover", False), (r"\bPLUG\b", "Mechanical Plug", False),
    (r"CASING\s+SHOE|LINER\s+SHOE", "Casing", False), (r"\bPBTD\b", "PBTD", True), (r"TUBING\s+HANGER", "Tubing Hanger", True),
]


_LAYER = None


def overlay_boxes(p):
    global _LAYER
    if _LAYER is None:
        _LAYER = json.loads((PASS / "text_layer.json").read_text())
    f = 600 / 72
    return [(l["bbox"][0] * f - 6, l["bbox"][1] * f - 6, l["bbox"][2] * f + 6, l["bbox"][3] * f + 6) for l in _LAYER[p]
            if "Internal" not in l["text"]]


def load_pass(p, k, keep_overlay=False):
    boxes = [] if keep_overlay else overlay_boxes(p)
    tsv = (PASS / f"{p}_{k}.tsv").read_text(errors="ignore")
    lines = defaultdict(list)
    for r in csv.DictReader(io.StringIO(tsv), delimiter="\t", quoting=csv.QUOTE_NONE):
        t = (r.get("text") or "").strip()
        if t and r.get("conf") not in (None, "-1"):
            cx, cy = int(r["left"]) + int(r["width"]) / 2, int(r["top"]) + int(r["height"]) / 2
            if any(x0 <= cx <= x1 and y0 <= cy <= y1 for x0, y0, x1, y1 in boxes):
                continue
            lines[(r["block_num"], r["par_num"], r["line_num"])].append(
                dict(t=t, x=int(r["left"]), y=int(r["top"]), w=int(r["width"]), h=int(r["height"])))
    out = []
    for ws in lines.values():
        ws.sort(key=lambda w: w["x"])
        out.append(dict(words=ws, y=min(w["y"] for w in ws), text=" ".join(w["t"] for w in ws)))
    return out


def text_of(p, k):
    return (PASS / f"{p}_{k}.txt").read_text(errors="ignore")


# ---------------------------------------------------------------- well name
NAME_RX = [re.compile(r"COMPLETION\s+DIAGRAM\s*:\s*(DLA-\d{1,2})"), re.compile(r"\bWELL\s*:\s*(DLA-\d{1,2})"),
           re.compile(r"(?:W|VV|VW|WV|VVV|VVW|VWW|WW|VVM)ELL\s*(?:DULANG\s*)?(A\s?-\s?\d{1,2}(?:\s?ST\s?\d?)?)\s*[:>;}]"),
           re.compile(r"Well\s+No[.,]?\s*[:;,]?\s*(A-\d{1,2}(?:ST\d?)?)\b", re.I),
           re.compile(r"\b(A-\d{1,2}ST\d)\s+WELL\s+COMPLETION\s+DIAGRAM"), re.compile(r"WELL\s+NAME\s*:\s*\w+\s+(A-\d{1,2}(?:ST\s?\d)?)\b"),
           re.compile(r"WELL\s*:\s*DULANG\s+(A\d{1,2}(?:ST\d)?)\b"), re.compile(r"Well\s+No\s*:\s*(DLA\s*#\s*\d{1,2})"),
           re.compile(r"\b(A-\d{1,2})\s*:\s*(?:\(|SINGLE|DUAL)"), re.compile(r"WELL\s+DULANG\s+(A-\d{1,2})\b")]
LAYER_RX = [re.compile(r"^:\s*(A-\d{1,2})[LS]?\s*:"), re.compile(r"^(DLG\s+A\s+\d{1,2})$")]
def std_name(label):
    s = re.sub(r"^(DLG A|DLA\s*#|DLA-|A-|A)", "", label.replace(" ", "").replace("DLGA", "DLG A").upper())
    m = re.match(r"0*(\d+)(ST\d?)?$", s)
    return f"Dulang_A{int(m.group(1))}{m.group(2) or ''}" if m else None


def well_name(p, layer):
    votes = defaultdict(set)
    srcs = {k: text_of(p, k) for k in PASSES + ("d", "t")}
    for f in sorted(PASS.glob(f"{p}_[vhf][0-9].txt")):          # extra local OCR variants of the title / header
        srcs[f.stem.split("_")[1]] = f.read_text(errors="ignore")
    for k, t in srcs.items():
        for rx in NAME_RX:
            for m in rx.finditer(t):
                sn = std_name(m.group(1))
                if sn:
                    votes[(sn, m.group(1).replace(" ", ""))].add(k)
    for j, l in enumerate(layer):        # typed labels (exact digital text); each label is its own source
        for rx in LAYER_RX:
            m = rx.match(l["text"].strip())
            if m:
                sn = std_name(m.group(1).replace("DLG A", "A-"))
                if sn:
                    votes[(sn, l["text"].strip()[:20])].add(f"layer{j}")
    tally = defaultdict(set)
    for (sn, lab), ks in votes.items():
        tally[sn] |= ks
    agreed = {sn: ks for sn, ks in tally.items() if len(ks) >= 2}
    base = lambda sn: re.match(r"Dulang_A(\d+)", sn).group(1)
    if len(agreed) > 1 and len({base(x) for x in agreed}) == 1:     # same number: keep the suffix shown on the page
        suff = {x: ks for x, ks in agreed.items() if re.search(r"ST\d?$", x)}
        if len(suff) == 1:
            agreed = suff
    others = {sn for sn, ks in tally.items() if agreed and base(sn) not in {base(x) for x in agreed}}
    if others:
        return None, None, [], votes          # a different well number is also read on the page -> conflict
    if len(agreed) == 1:
        sn, ks = next(iter(agreed.items()))
        lab = Counter(lab for (s2, lab), k2 in votes.items() if s2 == sn for _ in k2).most_common(1)[0][0]
        return sn, lab, sorted(ks), votes
    return None, None, [], votes


# ---------------------------------------------------------------- equipment items per pass
VOCAB = set("""FLOW COUPLING TRSV TR-SC-SSSV TRSCSSV TRSCSSSV SCSSV NIPPLE X- X NO GO NOGO NO-GO D-NO R XN WP-1 WP1 VVP-1 SPM SSD SLIMLINE
SLIDING SLEEVE PKR PACKER HYD PERM PERMANENT GP SUMP DUAL BWD BLAST JOINT ENTRY GUIDE WL WEG MULE SHOE SEAL ASSY ASSEMBLY BORE EXT.
EXT LOC. LOCATOR CONICAL SCREEN W.W.SCREEN PERFORATED PUP CENTRALIZER CENTRALISER HYDRO TRIP SUB PLUG PBTD CASING LINER TUBING HANGER
CAMCO OTIS BAKER CAM KBMG-M KBMG MAGNUM INSERT VALVE GLOV GLV DUMMY ORIFICE X-PROFILE PROFILE LANDING COLLET RETRIEVABLE""".split())


def label_ok(lab):
    words = [w.strip("()#.,") for w in lab.split()]
    for w in words:
        if not w or re.fullmatch(r"[\d./\"x×]+|#?\d+|\d+(?:\.\d+)?\"?|[A-Z]\d|OD|ID", w):
            continue
        if w.upper() not in VOCAB and not re.fullmatch(r"[A-Z]{1,3}-\d+[A-Z]?|\d/\d+\"?", w.upper()):
            return False
        if w.upper() in VOCAB and w != w.upper() and w not in ("Packer", "Plug", "Casing", "Liner", "Shoe", "Nipple", "Sub", "Screen"):
            return False          # mixed-case OCR artefact (e.g. 'SCSSi')
    return True


def items_of(line):
    """[(label, depth_token, x_label, x_depth, dims[(tok,x)])] for one OCR line (labels paired with nearest depth)."""
    ws = line["words"]
    toks = [w["t"] for w in ws]
    text = " ".join(toks)
    if re.search(r"\b(FLUID|DEVIATION|COMPLETION TYPE|RTE|HANGER\s*:|PTD|DATE)\b.*:|:\s*=", text):
        return []            # header 'label : value' lines are not equipment rows
    labs = []
    for rx, ctype, exact in EQUIP:
        for m in re.finditer(rx, text, re.I):
            if any(m.start() < e and m.end() > s for s, e, *_ in labs):
                continue
            # extend label to the left/right over word tokens (letters, sizes like 7", 3.88", x)
            s, e = m.start(), m.end()
            left = text[:s].split(" ")
            while left and left[-1] and re.fullmatch(r"[A-Za-z][A-Za-z0-9\-/.#]*|\d+(?:\.\d+)?\"|x|X|\(|\d/\d+\"?|#\d", left[-1]) and len(text[:s].split(" ")) - len(left) < 6:
                s -= len(left[-1]) + 1
                left.pop()
            right = text[e:].split(" ")
            i = 0
            while i < len(right) and (right[i] == "" or re.fullmatch(r"[A-Za-z(][A-Za-z0-9\-/.#()]*\)?|#\d|\(\w+", right[i])) and i < 6:
                e += len(right[i]) + 1
                i += 1
            labs.append((max(0, s), e, ctype, exact))
    if not labs:
        return []
    # char offset -> word x
    pos, xs = 0, []
    for w in ws:
        xs.append((pos, pos + len(w["t"]), w["x"], w["w"]))
        pos += len(w["t"]) + 1
    xat = lambda c: next((x for s, e, x, w in xs if s <= c < e), xs[-1][2])
    dcols = line.get("dcols") or []
    depths = [(m.group(1), xat(m.start())) for m in DEPTH_RX.finditer(text) if 10 <= float(m.group(1)) <= 4000
              and any(abs(xat(m.start()) - c) < 260 for c in dcols)]
    dims = [(m.group(1), xat(m.start())) for m in DIM_RX.finditer(text)]
    out, used = [], set()
    for s, e, ctype, exact in sorted(labs):
        lab = text[s:e].strip(" |[]{}—-_.,;:")
        if lab.count("(") != lab.count(")"):
            lab = re.split(r"\s*\(", lab)[0].strip() if lab.count("(") > lab.count(")") else lab.rstrip(")")
        xl = xat(s)
        cand = sorted((abs(xd - xl), i) for i, (tok, xd) in enumerate(depths) if i not in used)
        dep = None
        if cand:
            used.add(cand[0][1])
            dep = depths[cand[0][1]]
        if not label_ok(lab):
            continue
        out.append(dict(label=lab, ctype=ctype, exact=exact, depth=dep[0] if dep else None, x=xl, xd=dep[1] if dep else None, dims=dims))
    return out


def page_items(p):
    per = {}
    dcols = depth_columns(p)
    for k in PASSES:
        lst = []
        for ln in load_pass(p, k):
            ln["dcols"] = dcols
            for it in items_of(ln):
                it["y"] = ln["y"]
                lst.append(it)
        per[k] = lst
    return per


def consensus_items(per, H):
    """item accepted when the same (label, depth) is read in >=2 passes at about the same position."""
    groups = []
    for k, lst in per.items():
        for it in lst:
            key = (it["label"].upper(), it["depth"])
            g = next((g for g in groups if g["key"] == key and abs(g["y"] - it["y"]) < H * 0.012 and k not in g["passes"]), None)
            if g:
                g["passes"].add(k)
                g["items"].append(it)
            else:
                groups.append(dict(key=key, y=it["y"], passes={k}, items=[it]))
    return [g for g in groups if len(g["passes"]) >= 2]


def depth_columns(p):
    hits = defaultdict(list)
    for k in PASSES:
        for ln in load_pass(p, k):
            for w in ln["words"]:
                if re.fullmatch(r"DEPTH|Depth|DEPTH\(m\)|\(m\)|mddf|MDDF|MDTHF|m-MDDF", w["t"].strip(".,:|[]")) and w["t"].upper().startswith("DEPTH"):
                    hits[k].append(w["x"] + w["w"] / 2)
    cols = []
    for k, xs in hits.items():
        for x in xs:
            if sum(1 for k2, xs2 in hits.items() if k2 != k and any(abs(x - y) < 60 for y in xs2)) >= 1 and not any(abs(x - c) < 60 for c in cols):
                cols.append(x)
    return cols


def dim_columns(p):
    """x positions of 'ID' / 'OD' column headers agreed in >= 2 passes."""
    cols = defaultdict(list)
    for k in PASSES:
        for ln in load_pass(p, k):
            for i, w in enumerate(ln["words"]):
                t = w["t"].upper().strip(".,")
                prev = ln["words"][i - 1]["t"].upper() if i else ""
                if t in ("ID", "I.D", "1.D", "LD") and prev.startswith("MIN"):
                    cols["ID"].append((k, w["x"]))
                if t in ("OD", "O.D", "0.D") and prev.startswith("MAX"):
                    cols["OD"].append((k, w["x"]))
    res = {}
    for c, v in cols.items():
        if len({k for k, _ in v}) >= 2:
            res[c] = sorted(x for _, x in v)
    return res


# ---------------------------------------------------------------- zones / perforations
ZONE_RX = re.compile(r"\b(E\d{1,2}(?:/\d{1,2})?[A-Z]?(?:\s?Lower|\s?Upper|L|U)?)\b[^0-9\n]{0,40}?(\d{3,4}(?:\.\d{1,2})?)\s*[-–—]+\s*(\d{3,4}(?:\.\d{1,2})?)")


def zones(p):
    found = defaultdict(set)
    for k in PASSES:
        t = text_of(p, k)
        t2 = re.sub(r"\n(?=\s*(?:Perf|\(|\d{3,4}\s*[-–—]))", " ", t)          # interval on the next line
        for m in ZONE_RX.finditer(t2):
            a, b = float(m.group(2)), float(m.group(3))
            if 50 < a < b < 5000:
                found[(m.group(1), m.group(2), m.group(3))].add(k)
    return [(z, a, b, ks) for (z, a, b), ks in found.items() if len(ks) >= 2]


# ---------------------------------------------------------------- well info
INFO = [
    ("completionDate", re.compile(r"(?:DATE\s+OF\s+COMPLETION|COMPLETION\s+DATE)\s*[:\-]\s*([0-9]{1,2}[- ][A-Za-z]{3,9}[- ][0-9]{2,4}|[0-9]{1,2}\s+[A-Za-z]{3,9}\s+[0-9]{4})", re.I)),
    ("completionFluid", re.compile(r"COMPLETION\s+FLUID\s*:\s*([^\n|]{4,80}?)(?:\s{2,}|$)", re.I | re.M)),
    ("packerFluid", re.compile(r"PACKER\s+FLUID\s*:\s*([^\n|]{4,80}?)(?:\s{2,}|$)", re.I | re.M)),
    ("rig", re.compile(r"\bRIG\s*:\s*([A-Za-z][^\n|]{1,40}?)(?:\s{2,}|$)", re.M)),
    ("tubing", re.compile(r"\bTUBING\s*:\s*([^\n|]{3,60}?)(?:\s{2,}|$)", re.I | re.M)),
    ("wellhead", re.compile(r"\bWELL\s?HEAD\s*:\s*([^\n|]{3,60}?)(?:\s{2,}|RKB|$)", re.I | re.M)),
    ("maxDeviation", re.compile(r"MA[XS]I?(?:MUM)?\.?\s+DEVIATION\s*:\s*([^\n|]{2,40}?)(?:\s{2,}|$)", re.I | re.M)),
    ("completionType", re.compile(r"COMPLETION\s+TYPE\s*:\s*([A-Z][^\n|]{3,60}?)(?:\s{2,}|$)", re.M)),
    ("surfaceCasing", re.compile(r"([0-9 /\"°)']{2,10}\S*\s+[A-Z0-9-]+\s+SURF\.?,?\s+[^\n]{3,40}?)(?:\s{2,}|$)", re.M)),
]
TD_RX = re.compile(r"(?<![A-Z])TD\s*[:=@]\s*(\d{3,4}(?:\.\d{1,2})?)")


FREE_TEXT = ("rig", "completionFluid", "packerFluid", "tubing", "wellhead", "surfaceCasing", "completionType")


def well_info(p, tpl_wi):
    srcs = {k: text_of(p, k) for k in PASSES + ("d",)}
    wi, prov = {}, {}
    for key, rx in INFO:
        votes = defaultdict(set)
        for k, t in srcs.items():
            for m in rx.finditer(t):
                votes[re.sub(r"\s+", " ", m.group(1)).strip(" .,;")].add(k)
        LABELS = r"(COMPLETION|PACKER|TUBING|CASING|WELL|RIG|DATE|FLUID|TYPE|STRING|DEVIATION|HEAD)\b"
        need = 3 if key in FREE_TEXT else 2
        votes = {v: ks for v, ks in votes.items() if len(ks) >= need and not re.search(r"\b[A-Z][a-z]+[A-Z]|[A-Z]{2,}[a-z][A-Z]", v)}
        ok = [(v, ks) for v, ks in votes.items() if len(ks) >= 2 and v and ":" not in v
              and not re.fullmatch(r"[A-Z ]+", v) or (len(ks) >= 2 and v and ":" not in v and key in ("completionType", "rig") and not re.search(LABELS, v))]
        if len(ok) == 1:
            wi[key], prov[key] = ok[0][0], sorted(ok[0][1])
    votes = defaultdict(set)
    for k, t in srcs.items():
        for m in TD_RX.finditer(t):
            votes[m.group(1)].add(k)
    ok = [(v, ks) for v, ks in votes.items() if len(ks) >= 2]
    if len(ok) == 1:
        v = ok[0][0]
        wi["totalDepth"] = float(v) if "." in v else int(v)
        prov["totalDepth"] = sorted(ok[0][1])
    dd = re.match(r"\s*(\d{1,2})", wi.get("completionDate", "") or "")
    if wi.get("completionDate") and (not dd or not 1 <= int(dd.group(1)) <= 31):
        prov["rejected completionDate"] = wi.pop("completionDate")
    md = re.match(r"\s*(\d+(?:\.\d+)?)", wi.get("maxDeviation", "") or "")
    if wi.get("maxDeviation") and (not md or float(md.group(1)) > 90):
        prov["rejected maxDeviation"] = wi.pop("maxDeviation")
    for key in ("surfaceCasing", "tubing", "completionFluid", "packerFluid", "wellhead", "rig", "completionType"):
        if wi.get(key) and re.search(r"[*_~|{}»«¥]|\"\"|/,|,\"", wi[key]):
            prov[f"rejected {key}"] = wi.pop(key)
    out = {k: wi.get(k, None if isinstance(d, (int, float)) and not isinstance(d, bool) else "") for k, d in tpl_wi.items()}
    return out, prov


def num(tok):
    return float(tok) if "." in tok else int(tok)


def main():
    tpl = json.loads(subprocess.run(["textutil", "-convert", "txt", "-stdout", str(TPL)], capture_output=True, text=True, check=True).stdout)
    keymap = {r[0]: r[2] for r in load_workbook(XTPL)["Component_Types"].iter_rows(min_row=2, values_only=True)}
    layer = json.loads((PASS / "text_layer.json").read_text())
    report, written = [], []
    for i in range(1, 31):
        p = f"p{i:02d}"
        H = max(w["y"] + w["h"] for ln in load_pass(p, "a") for w in ln["words"])
        name, label, nsrc, votes = well_name(p, layer[p])
        if not name:
            seen = sorted({f"{lab}->{sn} [{','.join(sorted(ks))}]" for (sn, lab), ks in votes.items()}) or ["(none)"]
            report.append(dict(page=i, well="", status="SKIPPED — well name unreadable or conflicting (no single name read identically in 2 sources)", detail=f"readings: {seen}"))
            continue
        wi, wprov = well_info(p, tpl["wellInfo"])
        wi["name"] = name
        cols = dim_columns(p)
        comps, unread = [], 0
        for g in sorted(consensus_items(page_items(p), H), key=lambda g: g["y"]):
            it = g["items"][0]
            if it["depth"] is None:
                unread += 1
                continue
            d = {"id": f"{name.lower()}-{len(comps) + 1:02d}", "type": keymap[it["ctype"]], "name": it["label"],
                 "topMd": num(it["depth"]), "bottomMd": None, "outerDiameter": None, "innerDiameter": None}
            for c in ("OD", "ID"):
                if c in cols:
                    xs = cols[c]
                    agree = Counter()
                    for jt in g["items"]:
                        for tok, x in jt["dims"]:
                            if min(abs(x - cx) for cx in xs) < 0.04 * H:
                                agree[tok] += 1
                    best = [t for t, n in agree.items() if n >= 2]
                    if len(best) == 1:
                        d["outerDiameter" if c == "OD" else "innerDiameter"] = float(best[0])
            if not it["exact"]:
                d.update(estimated=True, estimatedFields=["type"], estimationNote=f"label '{it['label']}' is not literally a template type; nearest '{it['ctype']}'")
            d["source"] = f"PDF page {i}: '{it['label']}' @ {it['depth']} (read identically in OCR passes {''.join(sorted(g['passes']))}; depth units as printed on the page)"
            comps.append(d)
        for z, a, b, ks in sorted(zones(p), key=lambda t: float(t[1])):
            comps.append({"id": f"{name.lower()}-perf-{z.lower().replace('/', '-').replace(' ', '')}", "type": "perforation", "name": z,
                          "topMd": num(a), "bottomMd": num(b), "outerDiameter": None, "innerDiameter": None, "zone": z,
                          "source": f"PDF page {i}: zone '{z}' interval {a} - {b} (OCR passes {''.join(sorted(ks))}; units as printed)"})
        for l in layer[p]:          # typed update boxes (exact PDF text) — kept as notes, not depths
            pass
        doc = {k: {"wellInfo": wi, "surveys": [], "components": comps, "cement": [], "plugBackCement": [], "centralizers": [],
                   "floatEquipment": [], "formations": [], "formationTops": [], "fluidContacts": []}[k] for k in tpl}
        f = OUT / f"{name}.json"
        if f.exists():
            report.append(dict(page=i, well=name, status="SKIPPED — duplicate well name (file already written from an earlier page)", detail=label))
            continue
        f.write_text(json.dumps(doc, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
        filled = [k for k, v in wi.items() if v not in ("", None) and k != "name"]
        written.append(f.name)
        report.append(dict(page=i, well=name, status="WRITTEN", rejected={k: v for k, v in wprov.items() if k.startswith("rejected")}, detail=f"PDF label '{label}' (sources {','.join(nsrc)}); {len(comps)} components; wellInfo filled: {filled}; "
                                                                          f"items with no agreed depth: {unread}; OD/ID columns found: {sorted(cols)}"))
    rep = OUT / "_extraction_report.json"
    if rep.exists():
        rep.unlink()
    (PASS / "extraction_report.json").write_text(json.dumps(report, indent=1, ensure_ascii=False))
    for r in report:
        print(f"p{r['page']:02d} {r['well'] or '-':16} {r['status'][:60]:60} {r['detail'][:150]}")
    print("written:", len(written))


if __name__ == "__main__":
    main()
