"""Table-style slides (native PowerPoint text): row pairing -> component records with full traceability.

Every record carries: value, slide, object id(s), original text, method. The build script turns records into JSON.
"""
import json
import re
from pathlib import Path

from c_rules import ANNOT, classify, parse_zone_box

W = Path(__file__).resolve().parent
NUMV = re.compile(r"^\s*(\d{1,4}(?:\.\d+)?)\s*$")
CONT = re.compile(r"(?i)^(DUAL|SINGLE)\s*HYD\.?\s*PKR\.?$|^BAKER\s*PERM\.?\s*PKR$|^PERM\.?\s*PKR\.?$|^DUAL HYD PACKER$")

# per-slide corrections found in the visual review (object ids are PowerPoint shape ids)
OVERRIDES = {
    # slide: {obj id: action}
    2: {"8": ("intervention-row", None), "9": ("intervention-row", None)},
    4: {"46": ("intervention-row", "depth:45"), "34": ("annotation", "label '1st Packer (7th October 2023)' beside the long-string side of the dual packer row (PKR depth 2181.0 m printed in the long depth column) — cement-packer job note, not separate equipment")},
    5: {"5": ("intervention-row", None)},
    6: {"8": ("intervention-row", None), "4": ("annotation", "MA plug + gauge set at SSD#2 for an SGS survey and retrieved after the job (text says 'retrieved') — not installed equipment")},
    7: {"18695": ("spm-label", "row label is valve history 'GLV 3/16 975psi - retrieved' in place of the 'KBMG-M SPM' label; depth 1049.3 m and MIN ID 2.371 in the SPM column pattern -> SPM (Gas Lift) inferred from row position")},
    23: {},
    27: {"17": ("intervention-row", None), "7": ("annotation", "HUD note")},
}


def num(c):
    if c is None:
        return None
    m = NUMV.match(c["text"].replace("`", ""))
    return float(m.group(1)) if m else None


def load(slide):
    rows = json.loads((W / "rows" / f"slide{slide:02d}.json").read_text())
    return rows["cells"], rows["columns"]


def pair_slide(slide):
    """returns dict(rows=[...], zones=[...], unpaired=[...]) with merged multi-cell names"""
    cells, cols = load(slide)
    ov = OVERRIDES.get(slide, {})
    by = {c["id"]: c for c in cells}
    used, zones = set(), []
    for c in cells:
        t = c["text"].strip()
        if re.match(r"(?i)^\s*(Add Perf.*\n)?\s*SAND\s*[:]", t) or re.match(r"(?i)^\s*(PERFT?|Perft)\s*:", t) or re.search(r"(?i)\n\s*SAND\s*:", t):
            zones.append(c)
            used.add(c["id"])
    rows = []
    for side in "SL":
        names = sorted([c for c in cells if c["col"] == f"NAME_{side}" and c["id"] not in used], key=lambda c: c["y"])
        # merge continuation lines (older layout splits 'CAMCO HSD-3-SP' / 'DUAL HYD. PKR')
        merged, skip = [], set()
        for i, n in enumerate(names):
            if n["id"] in skip:
                continue
            parts = [n]
            for m in names[i + 1:]:
                if m["y"] - parts[-1]["y"] <= 1.3 and CONT.match(m["text"].strip()):
                    parts.append(m)
                    skip.add(m["id"])
                else:
                    break
            merged.append(parts)
        # primary equipment first, interventions/annotations afterwards (they must not steal the row's depth)
        def kind(parts):
            t = " / ".join(p["text"].strip() for p in parts)
            a = ov.get(parts[0]["id"])
            if a:
                return a[0]
            if ANNOT.search(t) or not classify(t)[0]:
                return "annotation" if not re.search(r"(?i)plug|insert valve|storm choke|\bIV\b|TRSV", t) else "intervention-row"
            return "equip"
        ordered = sorted(merged, key=lambda p: (kind(p) not in ("equip", "spm-label"), p[0]["y"]))
        for parts in ordered:
            k = kind(parts)
            row = {"side": side, "parts": parts, "kind": k, "override": ov.get(parts[0]["id"], (None, None))[1]}
            for p in parts:
                used.add(p["id"])
            if k in ("equip", "spm-label"):
                y = parts[0]["y"]
                bj = re.search(r"(?i)BLAST\s*JOINT", parts[0]["text"])  # BJ depths sit at the ends of its drawing, not on its row
                tol_d = 1.0 if slide == 13 else 0.9
                for col, tol in ((f"DEPTH_{side}", tol_d), (f"MINID_{side}", 0.75), (f"STATUS_{side}", 1.0)):
                    if bj and col.startswith("DEPTH"):
                        continue
                    cand = [c for c in cells if c["col"] == col and c["id"] not in used and abs(c["y"] - y) <= tol]
                    if col.startswith("STATUS"):
                        cand = [c for c in cand if not ANNOT.search(c["text"]) or re.search(r"(?i)CLOSED|OPEN|DV|GLV|GLOV|DUMMY|UNLOADING|ORIFICE|RIPO|BK", c["text"])]
                    if cand:
                        best = min(cand, key=lambda c: abs(c["y"] - y))
                        row[col.split("_")[0].lower()] = best
                        used.add(best["id"])
            rows.append(row)
    unpaired = [c for c in cells if c["id"] not in used]
    return dict(rows=rows, zones=zones, unpaired=unpaired, by=by, cols=cols)


def zone_records(zones):
    out, other = [], []
    for z in zones:
        perfs, oth = parse_zone_box(z["text"])
        for zone, top, bot, line in perfs:
            out.append(dict(zone=zone, top=top, bot=bot, line=line, obj=z["id"], box=z["text"]))
        other += [(z["id"], o) for o in oth]
    return out, other
