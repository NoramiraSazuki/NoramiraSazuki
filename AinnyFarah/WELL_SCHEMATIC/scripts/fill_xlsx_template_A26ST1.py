"""Fill a copy of the Dulang A04ST1 Excel well-schematic template with the data in
latest/Dulang A26ST1 Well Schematic.xlsx.

Both workbooks are opened read-only (never saved); the result goes to
latest/DULANG-A26ST1 Well Schematic (from A04ST1 template).xlsx.

Same layout as fill_xlsx_template_B30ST1.py:
  * header block rows 2-18 -> header fields + tubulars
  * right table (S depth, U equipment, Z MIN ID, AB MAX OD) -> completion string
  * left table  (F equipment, K depth) -> liner top, casing shoe, zones, PBTD, intervention notes
  * new sheet "Extracted Data" -> every value from the input, with its source cell
Items keep the input's top-to-bottom row order (the input sheet is drawn in depth order).
"""
import os
import re
import sys
import warnings
import zipfile

import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from fill_xlsx_template_B30ST1 import num, put, sha  # noqa: E402

warnings.filterwarnings("ignore")

BASE = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "latest")
SRC = os.path.join(BASE, "Dulang A26ST1 Well Schematic.xlsx")
TEMPLATE = os.path.join(BASE, "Dulang A04ST1 Well Schematic.xlsx")
OUT = os.path.join(BASE, "DULANG-A26ST1 Well Schematic (from A04ST1 template).xlsx")
FIRST_ROW, LAST_ROW = 25, 188


def drawing_texts(path):
    """Text boxes of the input drawing: (anchor row, text)."""
    with zipfile.ZipFile(path) as z:
        xml = z.read("xl/drawings/drawing1.xml").decode("utf-8")
    out = []
    for a in re.findall(r"<xdr:(?:twoCellAnchor|oneCellAnchor)[^>]*>(.*?)</xdr:(?:twoCellAnchor|oneCellAnchor)>",
                        xml, re.S):
        row = int(re.search(r"<xdr:from>.*?<xdr:row>(\d+)</xdr:row>", a, re.S).group(1)) + 1
        paras = ["".join(re.findall(r"<a:t>([^<]*)</a:t>", p)) for p in re.findall(r"<a:p>(.*?)</a:p>", a, re.S)]
        txt = " ".join(p.strip() for p in paras if p.strip())
        txt = txt.replace("&amp;", "&")
        if txt:
            out.append((row, re.sub(r"\s+", " ", txt)))
    return sorted(out)


def text(v):
    return v.strip() if isinstance(v, str) and v.strip() else None


def main():
    hashes = {p: sha(p) for p in (TEMPLATE, SRC)}
    src = openpyxl.load_workbook(SRC).worksheets[0]
    srcv = openpyxl.load_workbook(SRC, data_only=True).worksheets[0]

    def raw(c):
        return src[c].value

    def val(c):
        return srcv[c].value

    # ---------------- parse input ----------------
    header = {lbl: (f"F{r}", val(f"F{r}"), val(f"G{r}")) for r in range(4, 12)
              for lbl in [text(raw(f"A{r}"))] if lbl}
    tubulars = []
    for r in range(5, 10):
        if text(raw(f"N{r}")):
            tubulars.append({
                "row": r, "label": raw(f"N{r}"), "size": raw(f"Q{r}"), "weight": raw(f"R{r}"),
                "od": raw(f"S{r}"), "id": raw(f"U{r}"), "drift": raw(f"W{r}"), "grade": raw(f"X{r}"),
                "thread": raw(f"Y{r}"), "depth_formula": raw(f"AA{r}"), "depth_value": val(f"AA{r}")})
    fluids = {raw(f"N{r}"): raw(f"Q{r}") for r in (10, 11)}

    events = []        # (row, side, name, depth_mddf, depth_mdthf, id, od, extra)
    orphans = []       # ID / OD values on rows without an equipment label
    zone_rows = {}
    for r in range(15, 90):
        a = text(raw(f"A{r}"))
        if a and not a.lower().startswith("perf interval"):
            zone_rows[r] = a
        if a and a.lower().startswith("perf interval"):
            zname = next((zone_rows[k] for k in sorted(zone_rows, reverse=True) if k < r), "?")
            m = re.search(r"(\d+(?:\.\d+)?)\s*m\s*-\s*(\d+(?:\.\d+)?)\s*m", a)
            top = float(m.group(1)) if m else None
            events.append((r, "L", f"ZONE {zname}: {a}", top, None, None, None, f"A{r}"))
        s, e = text(raw(f"S{r}")), text(raw(f"E{r}"))
        if s and r < 90:
            events.append((r, "R", s, val(f"Q{r}"), val(f"R{r}"), val(f"Y{r}"), val(f"AA{r}"), f"S{r}"))
        elif val(f"Y{r}") is not None or val(f"AA{r}") is not None:
            orphans.append((r, val(f"Y{r}"), val(f"AA{r}")))
        if e:
            events.append((r, "L", e, val(f"I{r}"), val(f"J{r}"), None, None, f"E{r}"))
    notes_drawn = drawing_texts(SRC)
    for r, t in notes_drawn:
        events.append((r + 0.5, "L", f"NOTE: {t}", None, None, None, None, f"text box @ row {r}"))
    events.append((90, "L", "TD", header["TOTAL DEPTH"][1], None, None, None, "F6"))
    events.sort(key=lambda t: (t[0], t[1] == "R"))

    # left-diagram depths used for header tubular depths whose formula points at an empty cell
    left_depth = {n: d for _, sd, n, d, *_ in events if sd == "L"}
    casing_depth = left_depth.get('9-5/8" 47# L80 VAMTOP CASING')
    liner_top = left_depth.get('TOP OF 7" LINER')
    liner_shoe = left_depth.get('7" 29# L80 VAMTOP LINER CASING SHOE')

    # ---------------- fill template copy ----------------
    wb = openpyxl.load_workbook(TEMPLATE)
    ws = wb.worksheets[0]
    ws.title = "DULANG-A26ST1"
    ws._images, ws._charts = [], []
    for row in ws.iter_rows(min_row=2, max_row=206, min_col=1, max_col=42):
        for c in row:
            if not isinstance(c, MergedCell) and c.value is not None:
                c.value = None

    labels = {"B8": "DATE OF COMPLETION", "B9": "RIG", "B10": "TD", "B11": "PBTD",
              "B12": "MAXIMUM DEVIATION", "B13": "WATER DEPTH", "B14": "DFE-THF",
              "B15": "WELL HEAD", "B16": "X-MAS TREE", "V8": "SIZE", "X8": "WEIGHT", "Y8": "GRADE",
              "Z8": "THREAD", "AB8": "DEPTH", "P16": "COMPLETION FLUID", "P17": "PACKER FLUID"}
    for k, v in labels.items():
        put(ws, k, v)
    for r in range(8, 17):
        put(ws, f"F{r}", ":")

    def hv(label, unit=None):
        _, v, u = header.get(label, (None, None, None))
        if v is None:
            return "-"
        if hasattr(v, "date"):
            return v.date().isoformat()
        u = unit or (u.strip() if isinstance(u, str) else "")
        return f"{num(v) if isinstance(v, (int, float)) else v} {u}".strip()

    put(ws, "B2", f"WELL NAME: {raw('A1').split(':', 1)[1].strip()}")
    put(ws, "B3", raw("A2"))
    put(ws, "B5", "WELL COMPLETION SCHEMATIC")
    put(ws, "G8", hv("DATE OF COMPLETION"))
    put(ws, "G9", hv("RIG"))
    put(ws, "G10", hv("TOTAL DEPTH"))
    put(ws, "G11", "-")
    put(ws, "G12", hv("MAXIMUM DEVIATION"))
    put(ws, "G13", hv("WATER DEPTH", "m"))
    put(ws, "G14", hv("DFE-THF", "m"))
    put(ws, "G15", hv("WELLHEAD"))
    put(ws, "G16", hv("X-MAS TREE"))
    put(ws, "V16", fluids.get("COMPLETION FLUID") or "-")
    put(ws, "V17", fluids.get("PACKER FLUID") or "-")

    tub_depth = {
        "PRODUCTION CASING": f"{num(casing_depth)} m MDDF" if casing_depth else "-",
        "PRODUCTION LINER #1": (f"{num(liner_top)} - {num(liner_shoe)} m MDDF"
                                if liner_top and liner_shoe else "-"),
    }
    for i, t in enumerate(tubulars):
        r = 9 + i
        put(ws, f"P{r}", t["label"])
        put(ws, f"U{r}", ":")
        put(ws, f"V{r}", f'{t["size"]}"')
        put(ws, f"X{r}", f'{num(t["weight"])} PPF')
        put(ws, f"Y{r}", t["grade"])
        put(ws, f"Z{r}", t["thread"])
        put(ws, f"AB{r}", tub_depth.get(t["label"], "-"))

    put(ws, "B18", f"{raw('A12')}. Depths m-MDDF; m-MDTHF = MDDF - {hv('DFE-THF', 'm')}.")
    put(ws, "B19", "REFERENCE (single string well - no short string)")
    put(ws, "B20", "LINER / CASING / ZONES / NOTES")
    put(ws, "U19", "COMPLETION STRING")
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

    def free(r):
        for col in ("B", "D", "F", "K", "S", "U", "Z", "AB"):
            coord = f"{col}{r}"
            for rng in ws.merged_cells.ranges:
                if coord in rng and (rng.min_row != r or rng.max_row != r
                                     or rng.coord.split(":")[0] != coord):
                    return False
        return True

    good = [r for r in range(FIRST_ROW, LAST_ROW + 1) if free(r)]
    assert len(good) >= len(events), "not enough free rows"
    step = max(1, len(good) // len(events))
    for i, (_, side, name, d, _, id_, od, _) in enumerate(events):
        r = good[i * step]
        if side == "R":
            put(ws, f"S{r}", num(d) if isinstance(d, (int, float)) else None)
            put(ws, f"U{r}", name)
            put(ws, f"Z{r}", num(id_) if isinstance(id_, (int, float)) else None)
            put(ws, f"AB{r}", num(od) if isinstance(od, (int, float)) else None)
        else:
            put(ws, f"F{r}", name)
            put(ws, f"K{r}", num(d) if isinstance(d, (int, float)) else None)

    put(ws, "U191", f"{raw('S91')} {raw('U91')}")
    put(ws, "U192", "Source: Dulang A26ST1 Well Schematic.xlsx (sheet A26ST1)")
    put(ws, "U193", "Date: not stated (source date cell is =TODAY())")
    put(ws, "Z193", "Rev: -")

    # ---------------- Extracted Data sheet ----------------
    ds = wb.create_sheet("Extracted Data")
    fill = PatternFill("solid", fgColor="1F4E78")
    hfont = Font(bold=True, color="FFFFFF")
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    r = [1]

    def line(values, head=False, bold=False):
        for col, v in enumerate(values, 1):
            cell = ds.cell(row=r[0], column=col, value=v)
            if head:
                cell.fill, cell.font = fill, hfont
            elif bold:
                cell.font = Font(bold=True, size=12)
            if head or (not bold and v is not None):
                cell.border = border
            cell.alignment = Alignment(vertical="top", wrap_text=False)
        r[0] += 1

    def gap():
        r[0] += 1

    line([f"DULANG-A26ST1 — all data from {os.path.basename(SRC)} (sheet A26ST1)"], bold=True)
    gap()
    line(["Header"], bold=True)
    line(["Source cell", "Field", "Value", "Unit"], head=True)
    line(["A1", "Well name", raw("A1"), None])
    line(["A2", "Location / slot", raw("A2"), None])
    for lbl, (cell, v, u) in header.items():
        line([cell, lbl, v.date().isoformat() if hasattr(v, "date") else v, text(u)])
    for k, v in fluids.items():
        line([None, k, v, None])
    line(["A12", "Note", raw("A12"), None])
    line(["S91/U91", "Updated by", raw("U91"), None])
    line(["U92", "Date", "=TODAY() in source (no fixed date)", None])
    gap()
    line(["Tubulars"], bold=True)
    line(["Source row", "Tubular", "Size (in)", "Weight (ppf)", "OD (in)", "ID (in)", "Drift (in)",
          "Material & grade", "Thread", "Depth cell (source)", "Depth cell value", "Depth used in template"],
         head=True)
    for t in tubulars:
        line([t["row"], t["label"], t["size"], t["weight"], t["od"], t["id"], t["drift"], t["grade"],
              t["thread"], t["depth_formula"], t["depth_value"], tub_depth.get(t["label"], "-")])
    gap()
    line(["Equipment, zones, markers and notes (input order, top to bottom)"], bold=True)
    line(["Source", "Side in template", "Name / text", "Depth m-MDDF", "Depth m-MDTHF", "MIN ID (in)",
          "MAX OD (in)"], head=True)
    for _, side, name, d, dt, id_, od, cell in events:
        line([cell, "Right (completion string)" if side == "R" else "Left (reference)", name, d, dt, id_, od])
    gap()
    line(["ID / OD values on rows without an equipment label (not placed on the schematic)"], bold=True)
    line(["Source row", "MIN ID (Y)", "MAX OD (AA)", "Rows above / below with equipment"], head=True)
    labelled = {int(e[0]): e[2] for e in events if e[1] == "R"}
    for orow, y, aa in orphans:
        above = max((k for k in labelled if k < orow), default=None)
        below = min((k for k in labelled if k > orow), default=None)
        line([orow, y, aa, f"above: row {above} {labelled.get(above)} | below: row {below} {labelled.get(below)}"])
    gap()
    line(["Pictures in the input (no data)"], bold=True)
    line(["Equipment / perforation symbol images and a radioactive-hazard symbol near row 98; "
          "not copied (the template's A04ST1 drawing was removed as well)."])

    for col, w in zip("ABCDEFGHIJKL", [16, 26, 70, 14, 14, 12, 12, 16, 12, 16, 14, 22]):
        ds.column_dimensions[col].width = w

    wb.active = 0
    wb.save(OUT)
    for p, h in hashes.items():
        assert sha(p) == h, f"{p} changed!"
    print("Inputs unchanged:", *(f"{os.path.basename(p)} {h[:12]}" for p, h in hashes.items()))
    print(f"Wrote {OUT}")
    print(f"Items on schematic: {len(events)} (step {step}); orphan ID/OD rows: {orphans}")


if __name__ == "__main__":
    main()
