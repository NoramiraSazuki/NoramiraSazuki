"""Fill a copy of the Dulang A04ST1 Excel well-schematic template with the data in
latest/THF M DULANG-B30ST1-2-7.wsch 1.json.

The template file is opened read-only (never saved); the result goes to
latest/DULANG-B30ST1 Well Schematic.xlsx.

Layout used in the copy (template is a dual-string sheet; B30ST1 is single string):
  * header block rows 2-18   -> wellInfo + casing / tubing strings
  * right table (S depth, U equipment, Z MIN ID, AB MAX OD) -> completion string,
    upper and lower completion together in depth order ("[LC]" = lower completion)
  * left table  (F equipment, K depth)                       -> reference markers:
    casing shoes, perforated / screened zones, PBTD, TD
  * new sheet "Extracted Data" -> every wellInfo field and every component with all keys
Depths are written as stored in the JSON (wellInfo.depthUnit = m, depthDatum = DF).
"""
import hashlib
import json
import os
import warnings

import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

warnings.filterwarnings("ignore")

BASE = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "latest")
SRC = os.path.join(BASE, "THF M DULANG-B30ST1-2-7.wsch 1.json")
TEMPLATE = os.path.join(BASE, "Dulang A04ST1 Well Schematic.xlsx")
OUT = os.path.join(BASE, "DULANG-B30ST1 Well Schematic.xlsx")

WELL = "DULANG-B30ST1"
FIRST_ROW, LAST_ROW, STEP = 25, 188, 3
# components whose topMd 0 / bottomMd 30.48 is a placeholder (no depth printed for them)
PLACEHOLDER_IDS = {"scssv-1", "flow-coupling-2"}
CASING_IDS = {"surface-casing-1", "production-casing-1", "production-casing-2"}


def sha(path):
    with open(path, "rb") as fh:
        return hashlib.sha256(fh.read()).hexdigest()


def put(ws, coord, value):
    """Write to a cell, redirecting to the top-left cell when it sits inside a merge."""
    cell = ws[coord]
    if isinstance(cell, MergedCell):
        for rng in ws.merged_cells.ranges:
            if coord in rng:
                cell = ws.cell(row=rng.min_row, column=rng.min_col)
                break
    cell.value = value


def num(v):
    if v is None:
        return None
    r = round(v, 3)
    return int(r) if r == int(r) else r


def mtxt(v):
    return f"{num(v)} m MDDF" if v is not None else "-"


def main():
    tpl_hash = sha(TEMPLATE)
    with open(SRC, encoding="utf-8") as fh:
        data = json.load(fh)
    wi = data["wellInfo"]
    comps = data["components"]
    by_id = {c["id"]: c for c in comps}

    wb = openpyxl.load_workbook(TEMPLATE)  # in memory only; template is never saved
    ws = wb.worksheets[0]
    ws.title = WELL

    # ---- remove the A04ST1-specific wellbore drawing and clear every A04ST1 value ----
    ws._images = []
    ws._charts = []
    for row in ws.iter_rows(min_row=2, max_row=206, min_col=1, max_col=42):
        for c in row:
            if not isinstance(c, MergedCell) and c.value is not None:
                c.value = None

    # ---- header (labels re-instated as in the template) ----
    labels = {
        "B5": None, "B8": "DATE OF COMPLETION", "B9": "RIG", "B10": "TD", "B11": "PBTD",
        "B12": "MAXIMUM DEVIATION", "B13": "WATER DEPTH", "B14": "DRILL FLOOR HEIGHT",
        "B15": "WELL HEAD", "V8": "SIZE", "X8": "WEIGHT", "Y8": "GRADE", "Z8": "THREAD",
        "AB8": "DEPTH", "P16": "COMPLETION FLUID", "P17": "PACKER FLUID",
    }
    for k, v in labels.items():
        if v:
            put(ws, k, v)
    for r in range(8, 16):
        put(ws, f"F{r}", ":")
    for r in (9, 10, 11, 12, 13, 14, 16, 17):
        put(ws, f"U{r}", ":")
    put(ws, "U14", None)

    pbtd = next((c for c in comps if c["type"] == "pbtd"), None)
    put(ws, "B2", f"WELL NAME: {WELL}")
    put(ws, "B3", "LOCATION / SLOT: -")
    put(ws, "B5", wi.get("completionType") or "-")
    put(ws, "G8", wi.get("completionDate") or "-")
    put(ws, "G9", wi.get("rig") or "-")
    put(ws, "G10", mtxt(wi.get("totalDepth")))
    put(ws, "G11", mtxt(pbtd["topMd"]) if pbtd else "-")
    put(ws, "G12", wi.get("maxDeviation") or "-")
    put(ws, "G13", "-")
    put(ws, "G14", "-")
    put(ws, "G15", wi.get("wellhead") or "-")
    put(ws, "V16", wi.get("completionFluid") or "-")
    put(ws, "V17", wi.get("packerFluid") or "-")

    # casing / tubing strings: label, component id, size, thread
    strings = [
        (9, "INT. CASING", "surface-casing-1", '13-3/8"', "BTC"),
        (10, "PRODUCTION CASING", "production-casing-1", '9-5/8"', "JFE BEAR"),
        (11, "PRODUCTION CASING", "production-casing-2", '9-5/8"', "JFE BEAR"),
        (12, "LOWER COMP STRING", "lower-completion-string-1", '5-1/2"', "JFE BEAR"),
        (13, "TUBING #1", "tubing-1", '3-1/2"', "JFE BEAR"),
        (14, "TUBING #2", "tubing-2", '3-1/2"', "JFE BEAR"),
    ]
    for r, label, cid, size, thread in strings:
        c = by_id[cid]
        weight = f"{num(c['weight'])} PPF"
        if "?" + str(num(c["weight"])) in c["name"]:
            weight = "?" + weight  # weight digits marked uncertain in the source
        put(ws, f"P{r}", label)
        put(ws, f"V{r}", size)
        put(ws, f"X{r}", weight)
        put(ws, f"Y{r}", c.get("grade"))
        put(ws, f"Z{r}", thread)
        put(ws, f"AB{r}", mtxt(c["bottomMd"]))

    put(ws, "B18", "Note: All depths are in meter from RKB to top of equipment. "
                   f"Datum {wi.get('depthDatum')}; DF to THF = {wi.get('dfToThf')} m.")

    # ---- table headings ----
    put(ws, "B19", "REFERENCE (single string well - no short string)")
    put(ws, "B20", "CASING / ZONES / PBTD / TD")
    put(ws, "U19", "COMPLETION STRING")
    put(ws, "U20", "(OIL PRODUCER)  [LC] = lower completion")
    put(ws, "J19", "Deviation")
    put(ws, "J20", "Degree")
    for col in ("K", "S"):
        put(ws, f"{col}19", "DEPTH")
        put(ws, f"{col}20", "m-MDDF")
    for col, txt in (("B", "MAX OD"), ("D", "MIN ID"), ("F", "EQUIPMENT"),
                     ("U", "EQUIPMENT"), ("Z", "MIN ID"), ("AB", "MAX OD")):
        put(ws, f"{col}22", txt)
    for col in ("B", "D", "Z", "AB"):
        put(ws, f"{col}23", "IN")

    # ---- build the depth-ordered item list ----
    items = []  # (sort depth, json order, side, equipment, depth, id, od)
    prev_key = 0.0
    for order, c in enumerate(comps):
        cid, typ = c["id"], c["type"]
        top = c.get("topMd")
        if cid in PLACEHOLDER_IDS:
            key = prev_key  # keep the source order inside the FC / SCSSV / FC group
            depth = None
        else:
            key, depth = top or 0.0, top
        if cid in CASING_IDS:
            shoe = c["bottomMd"]
            name = c["name"].split("(")[-1].rstrip(")") if "SHOE" in c["name"] else f"{c['name']} (bottom)"
            items.append((shoe, order, "L", name, shoe, None, None))
            continue
        if typ == "perforation":
            items.append((top, order, "L", f"ZONE {c['zone']}: {num(top)} - {num(c['bottomMd'])} m MDDF",
                          top, None, None))
            continue
        if typ == "pbtd":
            items.append((top, order, "L", "PBTD", top, None, None))
            continue
        if typ == "tubing" and cid in ("tubing-1", "tubing-2", "lower-completion-string-1"):
            # string description row, placed where the string starts in the diagram
            key = {"tubing-1": 21.0001, "tubing-2": 250.0, "lower-completion-string-1": 1573.9}[cid]
            items.append((key, order, "R", f"{c['name']} (to {num(c['bottomMd'])} m MDDF)", None,
                          c.get("innerDiameter"), c.get("outerDiameter")))
            continue
        name = ("[LC] " if cid.startswith("lc-") else "") + c["name"]
        items.append((key, order, "R", name, depth, c.get("innerDiameter"), c.get("outerDiameter")))
        prev_key = key
    items.append((wi["totalDepth"], 999, "L", "TD", wi["totalDepth"], None, None))
    items.sort(key=lambda t: (t[0], t[1]))

    # rows where every target cell of both tables is a standalone cell or the top-left of
    # its merge (template rows with wide / multi-row merges are skipped)
    def free(r, cols):
        for col in cols:
            coord = f"{col}{r}"
            for rng in ws.merged_cells.ranges:
                if coord in rng and (rng.min_row != r or rng.coord.split(":")[0] != coord
                                     or rng.max_row != r):
                    return False
        return True
    good = [r for r in range(FIRST_ROW, LAST_ROW + 1)
            if free(r, ("B", "D", "F", "K", "S", "U", "Z", "AB"))]
    step = max(1, len(good) // len(items))
    assert len(good) >= len(items), "not enough free rows"
    for i, (_, _, side, name, depth, cid_, od) in enumerate(items):
        r = good[i * step]
        if side == "R":
            put(ws, f"S{r}", num(depth))
            put(ws, f"U{r}", name)
            put(ws, f"Z{r}", num(cid_))
            put(ws, f"AB{r}", num(od))
        else:
            put(ws, f"F{r}", name)
            put(ws, f"K{r}", num(depth))

    put(ws, "U191", "Source: Dulang B Well Schematic updated .pptx > slide 30 (image30.png), "
                    "via THF M DULANG-B30ST1-2-7.wsch 1.json")
    put(ws, "U193", "Date: 20-May-22")
    put(ws, "Z193", "Rev: 2 (PROPOSED)")

    # ---- full extracted data sheet ----
    ds = wb.create_sheet("Extracted Data")
    hdr_fill = PatternFill("solid", fgColor="1F4E78")
    hdr_font = Font(bold=True, color="FFFFFF")
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def header(row, values):
        for col, v in enumerate(values, 1):
            cell = ds.cell(row=row, column=col, value=v)
            cell.fill, cell.font, cell.border = hdr_fill, hdr_font, border

    ds["A1"] = f"{WELL} — all data from {os.path.basename(SRC)}"
    ds["A1"].font = Font(bold=True, size=13)
    ds["A2"] = (f"Depths as stored: unit {wi.get('depthUnit')}, datum {wi.get('depthDatum')} "
                f"(display datum {wi.get('displayDatum')}, DF to THF {wi.get('dfToThf')} m). "
                "'?' in a value = digit unclear in the source picture.")
    header(4, ["wellInfo field", "Value"])
    r = 5
    ds.cell(row=r, column=1, value="id").border = border
    ds.cell(row=r, column=2, value=data.get("id")).border = border
    for k, v in wi.items():
        r += 1
        ds.cell(row=r, column=1, value=k).border = border
        ds.cell(row=r, column=2, value=v).border = border

    r += 2
    ds.cell(row=r, column=1, value="Empty sections in the input").font = Font(bold=True)
    empties = [k for k, v in data.items() if isinstance(v, list) and not v]
    ds.cell(row=r, column=2, value=", ".join(empties))

    r += 2
    keys = ["id", "type", "actualType", "name", "zone", "topMd", "bottomMd", "topMdM", "bottomMdM",
            "outerDiameter", "innerDiameter", "weight", "grade", "estimated", "estimatedFields",
            "estimationNote", "Shown on schematic", "Note", "source"]
    header(r, keys)
    comp_start = r
    for c in comps:
        r += 1
        note = []
        if c["id"] in PLACEHOLDER_IDS:
            note.append("topMd 0 / bottomMd 30.48 look like placeholders (one printed depth for the "
                        "FC / SCSSV / FC group); depth left blank on the schematic")
        if c.get("topMdM") is not None and abs(c["topMdM"] - (c.get("topMd") or 0)) > 0.01:
            note.append(f"topMd {num(c['topMd'])} differs from printed topMdM {num(c['topMdM'])}; "
                        "schematic uses topMd (current value in the input)")
        if c["id"] in CASING_IDS or c["type"] == "perforation" or c["type"] == "pbtd":
            shown = "Left table / header"
        elif c["id"] in ("tubing-1", "tubing-2", "lower-completion-string-1"):
            shown = "Header + right table"
        else:
            shown = "Right table"
        vals = [c.get(k) for k in keys[:16]]
        vals[14] = ", ".join(vals[14] or []) or None
        vals[13] = "Yes" if vals[13] else None
        for col, v in enumerate(vals + [shown, "; ".join(note) or None, c.get("source")], 1):
            ds.cell(row=r, column=col, value=v).border = border
    for col, w in zip("ABCDEFGHIJKLMNOPQRS", [26, 18, 30, 50, 10, 11, 11, 10, 10, 10, 10, 8, 12, 9, 28,
                                                60, 20, 60, 90]):
        ds.column_dimensions[col].width = w
    for row in ds.iter_rows(min_row=comp_start + 1, max_row=r):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=cell.column_letter in "DPR")
    ds.freeze_panes = ds.cell(row=comp_start + 1, column=2)

    wb.active = 0
    wb.save(OUT)
    assert sha(TEMPLATE) == tpl_hash, "template changed!"
    print(f"Template unchanged: {tpl_hash}")
    print(f"Wrote {OUT}")
    print(f"Schematic rows: {len(items)} (step {step}); components in data sheet: {len(comps)}")


if __name__ == "__main__":
    main()
