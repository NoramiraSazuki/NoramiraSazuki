"""Header block (completion data / casing / PBTD / RTE / max deviation) of the table-style slides.

Standard layout: label column comes from the slide layout (not slide text) and values are separate text boxes, so values
are identified by content pattern (date, rig, tubing, X-mas tree, packer fluid; casing lines; 'xxxx m MDDF' = PBTD;
'nn.nn m' = RTE to tubing hanger; 'nn DEG' = maximum deviation). Older layout (slides 23/24/26/27) prints 'LABEL : value'.
"""
import re


def header(dump):
    objs = [o for o in dump["objects"] if o.get("text", "").strip() and o.get("pct") and 4 < o["pct"][1] < 25.5]
    h = {}

    def put(key, o, val=None):
        h.setdefault(key, []).append(dict(obj=o["id"], text=o["text"], value=(val if val is not None else re.sub(r"\s+", " ", o["text"]).strip())))

    for o in sorted(objs, key=lambda o: (o["pct"][1], o["pct"][0])):
        t = o["text"].strip()
        flat = re.sub(r"\s+", " ", t)
        x = o["pct"][0]
        if re.search(r"(?i)Latest Update|Flowing|Injecting|DULANG WELL COMPLETION", flat):
            put("note", o)
            continue
        if re.match(r"(?i)WELL C-\d", flat):
            put("title", o)
            continue
        lab = re.match(r"(?i)^(DATE OF COMPLETION|RIG|TUBING|X-MAS TREE|PACKER FLUID|PBTD|RTE TO TUBING HANGER|MAXIMUM DEVIATION|"
                       r"Date of Completion|Rig|Tubing|Txmas Tree|X-MAS Tree|Packer Fluid)\s*:?\s*(.*)$", flat, re.S)
        if lab and lab.group(2).strip():
            key = {"DATE OF COMPLETION": "date", "RIG": "rig", "TUBING": "tubing", "X-MAS TREE": "xmas", "TXMAS TREE": "xmas",
                   "PACKER FLUID": "packer_fluid", "PBTD": "pbtd", "RTE TO TUBING HANGER": "rte", "MAXIMUM DEVIATION": "maxdev"}[lab.group(1).upper()]
            put(key, o, lab.group(2).strip())
            continue
        if re.fullmatch(r"(?i)(Casing|PBTD|RTE to Tubing Hanger|Maximum Deviation|Packer Fluid|CASING)\s*:?", flat):
            continue  # bare labels (slide 7)
        if re.search(r"(?i)COND|SURF|PROD|DRIVE PIPE|CONDUCTOR", flat) and re.search(r'"', flat):
            for line in t.split("\n"):
                if line.strip():
                    h.setdefault("casing", []).append(dict(obj=o["id"], text=line, value=re.sub(r"\s+", " ", line).strip()))
            continue
        if re.search(r"(?i)\bDEG\b", flat):
            put("maxdev", o)
        elif re.fullmatch(r"\d+(\.\d+)?\s*m", flat):
            put("rte", o)
        elif re.fullmatch(r"(\d+(\.\d+)?\s*)?m\s*MD(DF|THF)", flat):
            put("pbtd", o)
        elif x < 52 and re.search(r"(?i)JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC", flat) and re.search(r"\d{4}", flat) and len(flat) < 30:
            put("date", o)
        elif x < 52 and re.search(r"(?i)PPF", flat) and re.match(r'\s*\d', flat):
            put("tubing", o)
        elif x < 52 and re.search(r"(?i)FMC|CACTUS|INGRAM", flat):
            put("xmas", o)
        elif x < 52 and re.search(r"(?i)PPG|KC[Il]|NaC|BRINE|BIOCIDE", flat):
            put("packer_fluid", o)
        elif x < 52 and re.search(r"(?i)TIOMAN|T-3|TENDER|KM-1", flat):
            put("rig", o)
        elif re.search(r"(?i)Flowing|Injecting|Latest Update|DULANG WELL COMPLETION", flat):
            put("note", o)
        else:
            put("other", o)
    return h
