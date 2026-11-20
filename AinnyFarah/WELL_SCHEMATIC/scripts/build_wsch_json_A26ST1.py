"""Convert latest/Dulang A26ST1 Well Schematic.xlsx into the .wsch JSON format of
latest/THF M DULANG-B30ST1-2-7.wsch 1.json (used as the template; read only).

Output: latest/THF M DULANG-A26ST1.wsch.json
        latest/THF M DULANG-A26ST1.wsch_extraction_notes.md  (what has no slot in the format)

Conventions copied from the template:
  * depthUnit m, depthDatum DF, displayDatum THF; topMd / bottomMd in m MDDF,
    topMdM / bottomMdM = the metre value as printed in the source (null when not printed)
  * point equipment: bottomMd = topMd
  * strings (casing / liner / tubing) from the header table, with weight + grade
  * zones as 'perforation' components with 'zone'
  * nearest-type mappings flagged: estimated / estimatedFields ['type'] / estimationNote / actualType
  * component keys sorted alphabetically, 2-space indent, UTF-8, no trailing newline
Depths that the input does not give stay null (never invented).
"""
import hashlib
import json
import os
import re
import warnings
import zipfile

import openpyxl

warnings.filterwarnings("ignore")

BASE = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "latest")
SRC = os.path.join(BASE, "Dulang A26ST1 Well Schematic.xlsx")
TEMPLATE = os.path.join(BASE, "THF M DULANG-B30ST1-2-7.wsch 1.json")
EDAFY = os.path.join(BASE, "EDAFY-Data_Dulang_Well.xlsx")
EDAFY_NAME = "DULANG-A26S1"   # EDAFY well name for the A26ST1 sidetrack (EDAFY base depth 1768 m = input TD)
OUT = os.path.join(BASE, f"{EDAFY_NAME}.wsch.json")
NOTES = os.path.join(BASE, f"{EDAFY_NAME}.wsch_extraction_notes.md")
M_TO_FT = 3.28084
SRCNAME = "Dulang A26ST1 Well Schematic.xlsx > sheet 'A26ST1'"

# equipment label (upper-case prefix match) -> (Type_Key, nearest-type note or None)
TYPE_RULES = [
    ("TUBING HANGER", "tubing-hanger", None),
    ("FLOW COUPLING", "flow-coupling", None),
    ("SCSSV", "scssv", None),
    ("SPM", "spm", None),
    ("SSD", "sliding-sleeve", None),
    ("SLIMLINE SSD", "sliding-sleeve", None),
    ("7\" X 3-1/2\" AHC PACKER", "packer", None),
    ("4-1/2\" X 7\" VTA PACKER", "packer", None),
    ("OIL SWELL PACKER", "packer", None),
    ("2.690\" XN NIPPLE", "nipple (no-go)", None),
    ("XN NIPPLE", "nipple (no-go)", None),
    ("NO-GO LOCATOR", "st-slot-locator", "no-go locator -> ST Slot Locator is the nearest type"),
    ("SELF ALIGNING MULE SHOE", "mule-shoe", None),
    ("SDG SELF ALIGN MULESHOE", "mule-shoe", None),
    ("3.88\" SEAL BORE EXTENSION", "pbr", "seal bore extension -> PBR is the nearest type"),
    ("3 WAY SUB", "crossover", "3-way sub (4-1/2 x 4-1/2 x 2-7/8) -> Crossover is the nearest type"),
    ("SEAL ASSY", "seal-assembly", None),
    ("3-1/2\" 9.2# 13CR TUBING", "tubing", None),
]


def sha(p):
    with open(p, "rb") as fh:
        return hashlib.sha256(fh.read()).hexdigest()


def r3(v):
    if v is None:
        return None
    x = round(float(v), 3)
    return int(x) if x == int(x) else x


def classify(name):
    up = name.upper()
    for prefix, key, note in sorted(TYPE_RULES, key=lambda t: -len(t[0])):
        if up.startswith(prefix):
            return key, note
    raise ValueError(f"no type rule for {name!r}")


def drawing_texts(path):
    with zipfile.ZipFile(path) as z:
        xml = z.read("xl/drawings/drawing1.xml").decode("utf-8")
    out = []
    for a in re.findall(r"<xdr:(?:twoCellAnchor|oneCellAnchor)[^>]*>(.*?)</xdr:(?:twoCellAnchor|oneCellAnchor)>",
                        xml, re.S):
        row = int(re.search(r"<xdr:from>.*?<xdr:row>(\d+)</xdr:row>", a, re.S).group(1)) + 1
        paras = ["".join(re.findall(r"<a:t>([^<]*)</a:t>", p)) for p in re.findall(r"<a:p>(.*?)</a:p>", a, re.S)]
        txt = re.sub(r"\s+", " ", " ".join(p.strip() for p in paras if p.strip())).replace("&amp;", "&")
        if txt:
            out.append((row, txt))
    return sorted(out)


def main():
    hashes = {p: sha(p) for p in (SRC, TEMPLATE, EDAFY)}
    ed = openpyxl.load_workbook(EDAFY, read_only=True, data_only=True)["Dulang Wells"]
    ed_rows = {r[1]: r for r in ed.iter_rows(min_row=2, values_only=True) if r[1]}
    assert EDAFY_NAME in ed_rows, f"{EDAFY_NAME} not in EDAFY"
    with open(TEMPLATE, encoding="utf-8") as fh:
        tpl = json.load(fh)
    ws = openpyxl.load_workbook(SRC).worksheets[0]
    wv = openpyxl.load_workbook(SRC, data_only=True).worksheets[0]
    raw = lambda c: ws[c].value          # noqa: E731
    val = lambda c: wv[c].value          # noqa: E731
    txt = lambda c: raw(c).strip() if isinstance(raw(c), str) and raw(c).strip() else None  # noqa: E731

    dfe_thf = val("F9")
    counters = {}
    components = []
    notes = []

    def new_id(prefix):
        counters[prefix] = counters.get(prefix, 0) + 1
        return f"{prefix}-{counters[prefix]}"

    def comp(cid, typ, name, top, bottom, top_m, bottom_m, od, id_, source, extra=None,
             est_fields=None, est_note=None, actual=None):
        c = {"bottomMd": r3(bottom), "bottomMdM": r3(bottom_m), "id": cid, "innerDiameter": r3(id_),
             "name": name, "outerDiameter": r3(od), "source": source, "topMd": r3(top),
             "topMdM": r3(top_m), "type": typ}
        if actual:
            c["actualType"] = actual
        if est_fields:
            c.update({"estimated": True, "estimatedFields": est_fields, "estimationNote": est_note})
        c.update(extra or {})
        components.append(dict(sorted(c.items())))

    # ---------- header ----------
    name_raw = raw("A1")                      # 'WELL NAME: A26ST1'
    well_no = name_raw.split(":", 1)[1].strip()
    well = f"DULANG-{well_no}"
    tub = {}
    for r in range(6, 10):
        tub[raw(f"N{r}")] = {k: val(f"{c}{r}") for k, c in
                             (("size", "Q"), ("weight", "R"), ("od", "S"), ("id", "U"), ("drift", "W"),
                              ("grade", "X"), ("thread", "Y"))} | {"row": r, "depthCell": raw(f"AA{r}")}

    # ---------- left diagram (casing / liner markers) ----------
    left = {}
    for r in range(15, 90):
        if txt(f"E{r}"):
            left[txt(f"E{r}")] = (r, val(f"I{r}"))
    liner_top = left.get('TOP OF 7" LINER')
    casing = left.get('9-5/8" 47# L80 VAMTOP CASING')
    liner_shoe = left.get('7" 29# L80 VAMTOP LINER CASING SHOE')

    # christmas tree (header only, surface)
    comp("christmas-tree-1", "christmas-tree", raw("F11"), 0, None, None, None, None, None,
         f"{SRCNAME} header X-MAS TREE F11 '{raw('F11')}' (no depth printed; placed at surface)",
         est_fields=["topMd"], est_note="topMd: X-mas tree has no depth in the input; 0 = surface")

    # ---------- right diagram: completion string ----------
    orphans = []
    labelled_rows = [r for r in range(15, 90) if txt(f"S{r}")]
    for r in range(15, 90):
        s = txt(f"S{r}")
        if not s:
            if val(f"Y{r}") is not None or val(f"AA{r}") is not None:
                orphans.append((r, val(f"Y{r}"), val(f"AA{r}")))
            continue
        if s in ("DOUBLE VALVE FLOAT SHOE",):
            continue  # -> floatEquipment
        typ, nearest = classify(s)
        depth, id_, od = val(f"Q{r}"), val(f"Y{r}"), val(f"AA{r}")
        src = [f"{SRCNAME} row {r} S{r} '{s}'"]
        if depth is not None:
            src.append(f"topMd Q{r} {r3(depth)} m MDDF (column 'DEPTH m-MDDF'; R{r} = "
                       f"{r3(val(f'R{r}'))} m MDTHF)")
        else:
            src.append("no depth printed")
        if id_ is not None:
            src.append(f"innerDiameter Y{r} '{id_}' (column 'MIN ID')")
        if od is not None:
            src.append(f"outerDiameter AA{r} '{od}' (column 'MAX OD')")
        est_fields, est_note, actual = None, None, None
        if nearest:
            est_fields, est_note, actual = ["type"], f"type: '{s}' ({nearest})", s
        extra = {}
        if typ == "tubing":
            t = tub["PRODUCTION TUBING #1"]
            extra = {"grade": t["grade"], "weight": r3(t["weight"])}
            src.append("section of PRODUCTION TUBING #1 (header row 8); OD column value is the coupling OD")
        prefix = "tubing-section" if typ == "tubing" else re.sub(r"[^a-z0-9]+", "-", typ).strip("-")
        comp(new_id(prefix), typ, s, depth, depth, depth, None, od, id_, "; ".join(src), extra,
             est_fields, est_note, actual)

    # SPM rows with no ID/OD on their own row: note the unlabelled ID/OD rows next to them
    for orow, y, aa in orphans:
        above = max(k for k in labelled_rows if k < orow)
        below = min(k for k in labelled_rows if k > orow)
        notes.append(f"Row {orow}: MIN ID {y} / MAX OD {aa} printed on a row with no equipment label "
                     f"(between row {above} '{txt(f'S{above}')}' and row {below} '{txt(f'S{below}')}'); "
                     "not assigned to a component.")

    # ---------- strings from header ----------
    def string_comp(cid, typ, label, top, bottom, bottom_src, top_m=None, bottom_m=None):
        t = tub[label]
        nm = f"{label} {t['size']} {r3(t['weight'])} ppf {t['grade']} {t['thread']}"
        src = (f"{SRCNAME} header row {t['row']} N{t['row']} '{label}': size '{t['size']}', weight {t['weight']}, "
               f"OD {t['od']}, ID {t['id']}, drift {t['drift']}, grade '{t['grade']}', thread '{t['thread']}'; "
               f"header depth cell AA{t['row']} = {t['depthCell']} (refers to an empty cell -> 0); {bottom_src}")
        comp(cid, typ, nm, top, bottom, top_m, bottom_m, t["od"], t["id"], src,
             {"grade": t["grade"], "weight": r3(t["weight"])})

    string_comp("production-casing-1", "casing", "PRODUCTION CASING", 0, casing[1],
                f"bottom from left diagram E{casing[0]} '9-5/8\" 47# L80 VAMTOP CASING' I{casing[0]} "
                f"{casing[1]} m MDDF", bottom_m=casing[1])
    string_comp("liner-1", "liner", "PRODUCTION LINER #1", liner_top[1], liner_shoe[1],
                f"top from E{liner_top[0]} 'TOP OF 7\" LINER' I{liner_top[0]} {liner_top[1]} m MDDF; bottom from "
                f"E{liner_shoe[0]} '7\" 29# L80 VAMTOP LINER CASING SHOE' I{liner_shoe[0]} {liner_shoe[1]} m MDDF",
                top_m=liner_top[1], bottom_m=liner_shoe[1])
    string_comp("tubing-1", "tubing", "PRODUCTION TUBING #1", None, None, "no tubing depths printed")
    string_comp("tubing-2", "tubing", "PRODUCTION TUBING #2", None, None, "no tubing depths printed")

    # ---------- zones ----------
    zone = None
    for r in range(15, 90):
        a = txt(f"A{r}")
        if not a:
            continue
        if not a.lower().startswith("perf interval"):
            zone = a
            continue
        m = re.search(r"(\d+(?:\.\d+)?)\s*m\s*-\s*(\d+(?:\.\d+)?)\s*m\s*(\S+)", a)
        top, bot = float(m.group(1)), float(m.group(2))
        comp(f"zone-{zone.lower()}-1", "perforation", f"{zone} {a}", top, bot, top, bot, None, None,
             f"{SRCNAME} row {r} A{r} '{a}' (zone label '{zone}')", {"zone": zone})

    # ---------- intervention text boxes ----------
    for row, t in drawing_texts(SRC):
        base = f"{SRCNAME} drawing text box at row {row}: '{t}'"
        if t.startswith("Slickplug @ 1608m MDDF"):
            comp(new_id("mechanical-plug"), "mechanical-plug", t, 1608, 1608, 1608, None, None, None, base,
                 est_fields=["type"], est_note="type: slickplug -> Mechanical Plug is the nearest type",
                 actual="Slickplug")
        elif t.startswith("Junk Catcher & Slickplug @ 1562m MDTHF"):
            md = 1562 + dfe_thf
            comp(new_id("mechanical-plug"), "mechanical-plug", t, md, md, None, None, None, None,
                 base + f"; 1562 m MDTHF + DFE-THF {dfe_thf} m = {r3(md)} m MDDF",
                 est_fields=["type", "topMd"],
                 est_note=f"type: junk catcher & slickplug -> Mechanical Plug is the nearest type; topMd: "
                          f"printed in m MDTHF, converted to MDDF with DFE-THF {dfe_thf} m",
                 actual="Junk Catcher & Slickplug")
        elif t.startswith("13Cr Widepak"):
            md = 1551 + dfe_thf
            comp(new_id("sand-screen"), "sand-screen", t, md, md, None, None, None, None,
                 base + f"; 1551 m MDTHF + DFE-THF {dfe_thf} m = {r3(md)} m MDDF",
                 est_fields=["topMd"],
                 est_note=f"topMd: printed in m MDTHF (top), converted to MDDF with DFE-THF {dfe_thf} m; "
                          "length of the 3 joints not given",
                 actual="13Cr Widepak + 3jts Boron ERTTSS")
        elif t.startswith("Tubing punch"):
            comp(new_id("perforation"), "perforation", t, 1588, 1598, 1588, 1598, None, None, base,
                 est_fields=["type"], est_note="type: tubing punch -> Perforation is the nearest type",
                 actual="Tubing punch", extra={"zone": "E34B"})
        else:
            notes.append(f"Drawing text box at row {row} not mapped to a component: '{t}' "
                         "(depth is wireline WLTHF, not MD).")

    # order like the template: equipment (input order), then strings, then zones
    eq = [c for c in components if c["type"] not in ("casing", "liner", "perforation")
          and c["id"] not in ("tubing-1", "tubing-2")]
    rest = [c for c in components if c not in eq]
    ids = [c["id"] for c in components]
    assert len(ids) == len(set(ids)), "duplicate component ids"

    out = {k: v for k, v in tpl.items()}
    out["id"] = EDAFY_NAME
    wi = dict(tpl["wellInfo"])
    wi.update({
        "completionDate": val("F4").date().isoformat(),
        "completionFluid": raw("Q10"),
        "completionType": "-",
        "depthDatum": "DF",
        "depthUnit": "m",
        "dfToThf": r3(dfe_thf),
        "dfToThfSource": "Input workbook header 'DFE-THF' (F9)",
        "displayDatum": tpl["wellInfo"]["displayDatum"],
        "groundElevation": None,
        "kbElevation": None,
        "maxDeviation": f"{r3(val('F7'))} {raw('G7')}",
        "name": EDAFY_NAME,
        "packerFluid": raw("Q11"),
        "rig": raw("F5"),
        "surfaceCasing": "-",
        "totalDepth": r3(val("F6")),
        "trajectoryType": "-",
        "tubing": "; ".join(f"{k} {t['size']} {r3(t['weight'])} {t['grade']} {t['thread']}"
                            for k, t in tub.items() if k.startswith("PRODUCTION TUBING")),
        "wellhead": raw("F10"),
    })
    out["wellInfo"] = wi
    out["components"] = eq + rest
    for k in ("surveys", "cement", "plugBackCement", "centralizers", "formations", "formationTops",
              "fluidContacts"):
        out[k] = []
    out["floatEquipment"] = [{"casingId": None, "md": r3(val("Q84")), "type": "float shoe",
                              "name": raw("S84")}]
    assert list(out) == list(tpl) and list(out["wellInfo"]) == list(tpl["wellInfo"])

    # ---------- metres -> feet (topMdM / bottomMdM keep the printed metres) ----------
    ft = lambda v: None if v is None else round(v * M_TO_FT, 3)  # noqa: E731
    for c in out["components"]:
        c["topMd"], c["bottomMd"] = ft(c["topMd"]), ft(c["bottomMd"])
    for fe in out["floatEquipment"]:
        fe["md"] = ft(fe["md"])
    wi["depthUnit"] = "ft"
    wi["totalDepth"] = ft(wi["totalDepth"])
    wi["dfToThf"] = ft(wi["dfToThf"])
    wi["dfToThfSource"] = f"Input workbook header 'DFE-THF' (F9) = {r3(dfe_thf)} m -> ft"

    text = json.dumps(out, indent=2, ensure_ascii=False)
    text = re.sub(r"\[\s*\]", "[]", text)
    with open(OUT, "w", encoding="utf-8") as fh:
        fh.write(text)

    unmapped = [
        f"Well name: input label '{name_raw}' -> EDAFY name '{EDAFY_NAME}' (UWI {ed_rows[EDAFY_NAME][0]}; "
        f"EDAFY base depth {ed_rows[EDAFY_NAME][14]} {ed_rows[EDAFY_NAME][15]} = input TD {r3(val('F6'))} m).",
        f"Units: all depths converted m -> ft (x {M_TO_FT}, 3 dp): topMd, bottomMd, totalDepth, dfToThf, "
        "floatEquipment md. topMdM / bottomMdM keep the printed metres; diameters stay in inches; "
        "source texts quote the original metre values.",
        f"Water depth: {val('F8')} {raw('G8').strip()} = {round(val('F8') * M_TO_FT, 3)} ft (F8) — "
        "no wellInfo field in the template.",
        "PBTD: label only (E87), no depth printed — no pbtd component created.",
        "Tubular OD / ID / drift from the header table are in the string components' source text; "
        "drift has no field.",
        "m-MDTHF depth column (R) is MDDF - DFE-THF; kept in each component's source text.",
        f"Updated by: {raw('U91')} (U91); date cell U92 is =TODAY() (no fixed date).",
        "Completion type / schematic title: not stated in the input -> '-'.",
        "Pictures: equipment/perforation symbol images and a radioactive-hazard symbol (no data).",
        "Items with no depth printed are kept with topMd/bottomMd null (flow couplings, tubing sections, "
        "NO-GO LOCATOR, 4-1/2\" x 7\" VTA packer, 3.88\" seal bore extension, 3-way sub, SDG mule shoe, "
        "tubing strings).",
    ] + notes
    with open(NOTES, "w", encoding="utf-8") as fh:
        fh.write(f"# {os.path.basename(OUT)} — extraction notes\n\n"
                 f"Input: `{os.path.basename(SRC)}` (read only, SHA-256 {hashes[SRC]})  \n"
                 f"Template: `{os.path.basename(TEMPLATE)}` (read only, SHA-256 {hashes[TEMPLATE]})  \n"
                 f"Well names: `{os.path.basename(EDAFY)}` (read only, SHA-256 {hashes[EDAFY]})\n\n"
                 "## Information with no slot in the template format\n\n"
                 + "\n".join(f"- {u}" for u in unmapped) + "\n")

    for p, h in hashes.items():
        assert sha(p) == h, f"{p} changed"
    print("Inputs unchanged.")
    print(f"Wrote {OUT}: {len(out['components'])} components, floatEquipment {len(out['floatEquipment'])}")
    print(f"Wrote {NOTES}")


if __name__ == "__main__":
    main()
