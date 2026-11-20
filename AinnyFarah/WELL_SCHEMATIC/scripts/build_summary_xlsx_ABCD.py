"""Build a summary workbook from the well-schematic JSONs in
output_A_v2, output_B, output_C and output_D.

Usage:  python3 scripts/build_summary_xlsx_ABCD.py
Output: WELL_SCHEMATIC/Dulang_Well_Schematic_Summary_ABCD.xlsx
"""
import collections
import datetime
import glob
import json
import os
import re

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FOLDERS = [("A", "output_A_v2"), ("B", "output_B"), ("C", "output_C"), ("D", "output_D")]
OUT = os.path.join(BASE, "Dulang_Well_Schematic_Summary_ABCD.xlsx")
EDAFY = os.path.join(BASE, "EDAFY-Data_Dulang_Well.xlsx")
# Schematic wells with no exact EDAFY name: shown with a candidate, never linked automatically.
EDAFY_CANDIDATES = {"DULANG-C22ST2": "DULANG-C22R1ST2"}

HDR_FILL = PatternFill("solid", fgColor="1F4E78")
HDR_FONT = Font(bold=True, color="FFFFFF")
TOTAL_FONT = Font(bold=True)
WARN_FILL = PatternFill("solid", fgColor="FCE4D6")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def well_sort_key(name):
    m = re.match(r"DULANG-([A-Z])(\d+)(.*)", name)
    return (m.group(1), int(m.group(2)), m.group(3)) if m else (name, 0, "")


def blank(v):
    return None if v in (None, "", "-") else v


def load():
    wells = []
    for platform, folder in FOLDERS:
        for path in glob.glob(os.path.join(BASE, folder, "*.json")):
            with open(path, encoding="utf-8") as fh:
                data = json.load(fh)
            wells.append((platform, folder, os.path.basename(path), data))
    wells.sort(key=lambda w: well_sort_key(w[3]["wellInfo"]["name"]))
    return wells


def load_edafy():
    from openpyxl import load_workbook
    ws = load_workbook(EDAFY, read_only=True, data_only=True)["Dulang Wells"]
    rows = list(ws.iter_rows(values_only=True))
    header = list(rows[0])
    data = [list(r) for r in rows[1:] if r[1]]
    return header, data


def edafy_platform(name):
    m = re.match(r"DULANG-([A-D])\d", name)
    return m.group(1) if m else "Exploration / appraisal"


def write_table(ws, headers, rows, widths=None, start_row=1):
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=start_row, column=c, value=h)
        cell.fill, cell.font, cell.border = HDR_FILL, HDR_FONT, BORDER
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    for r, row in enumerate(rows, start_row + 1):
        for c, v in enumerate(row, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top")
    ws.freeze_panes = ws.cell(row=start_row + 1, column=1)
    last = start_row + max(len(rows), 1)
    ws.auto_filter.ref = f"A{start_row}:{get_column_letter(len(headers))}{last}"
    for c, h in enumerate(headers, 1):
        w = (widths or {}).get(h)
        if w is None:
            longest = max([len(str(h))] + [len(str(r[c - 1])) for r in rows if r[c - 1] is not None])
            w = min(max(longest + 2, 8), 45)
        ws.column_dimensions[get_column_letter(c)].width = w


def main():
    wells = load()
    ed_header, ed_data = load_edafy()
    ed_by_name = {r[1]: dict(zip(ed_header, r)) for r in ed_data}
    schem_by_name = {d["wellInfo"]["name"]: (folder, fname) for _, folder, fname, d in wells}
    all_types = sorted({c["type"] for *_, d in wells for c in d["components"]})

    # ---------- Wells sheet ----------
    well_hdr = [
        "No.", "Platform", "Well", "In EDAFY", "EDAFY UWI", "EDAFY well status", "EDAFY well type",
        "Source folder", "File",
        "Completion date", "Completion type", "Tubing", "Surface casing",
        "Packer fluid", "Completion fluid", "Max deviation", "Total depth (ft)",
        "Components", "Flagged (estimated)", "Strings",
        "Casing", "Tubing joints", "Packers", "SPM", "Sliding sleeves", "Perforations",
        "Shallowest top MD (ft)", "Deepest MD (ft)", "Deepest MD (m)",
        "Float equipment", "Plug-back cement", "Note",
    ]
    well_rows = []
    per_platform = collections.defaultdict(collections.Counter)
    type_matrix = []
    comp_rows = []
    other_rows = []

    for i, (platform, folder, fname, d) in enumerate(wells, 1):
        wi = d["wellInfo"]
        comps = d["components"]
        tc = collections.Counter(c["type"] for c in comps)
        flagged = sum(1 for c in comps if c.get("estimated"))
        strings = sorted({c["string"] for c in comps if c.get("string")})
        tops = [c["topMd"] for c in comps if isinstance(c.get("topMd"), (int, float))]
        depth_pairs = []
        for c in comps:
            for ft_key, m_key in (("topMd", "topMdM"), ("bottomMd", "bottomMdM")):
                if isinstance(c.get(ft_key), (int, float)):
                    depth_pairs.append((c[ft_key], c.get(m_key)))
        deepest = max(depth_pairs, key=lambda p: p[0]) if depth_pairs else (None, None)
        ed = ed_by_name.get(wi["name"])
        notes = []
        if ed is None:
            cand = EDAFY_CANDIDATES.get(wi["name"])
            notes.append("No exact EDAFY well name" + (f" (possible match: {cand})" if cand else ""))
        if not comps:
            notes.append("No components extracted")
        if wi.get("totalDepth") is None:
            notes.append("Total depth not stated")
        if not blank(wi.get("completionDate")):
            notes.append("Completion date not stated")

        well_rows.append([
            i, platform, wi["name"], "Yes" if ed else "No",
            ed and ed["UWI"], ed and ed["WELL STATUS"], ed and ed["WELL TYPE"], folder, fname,
            blank(wi.get("completionDate")), blank(wi.get("completionType")),
            blank(wi.get("tubing")), blank(wi.get("surfaceCasing")),
            blank(wi.get("packerFluid")), blank(wi.get("completionFluid")),
            blank(wi.get("maxDeviation")), wi.get("totalDepth"),
            len(comps), flagged, ", ".join(strings) or None,
            tc["casing"], tc["tubing"], tc["packer"], tc["spm"], tc["sliding-sleeve"], tc["perforation"],
            min(tops) if tops else None, deepest[0], deepest[1],
            len(d["floatEquipment"]), len(d["plugBackCement"]),
            "; ".join(notes) or None,
        ])

        p = per_platform[platform]
        p["wells"] += 1
        p["components"] += len(comps)
        p["flagged"] += flagged
        p["float"] += len(d["floatEquipment"])
        p["pbc"] += len(d["plugBackCement"])
        p["empty"] += 0 if comps else 1
        p["no_td"] += 1 if wi.get("totalDepth") is None else 0
        p["no_date"] += 0 if blank(wi.get("completionDate")) else 1

        type_matrix.append([platform, wi["name"]] + [tc.get(t, 0) for t in all_types] + [len(comps)])

        for c in comps:
            comp_rows.append([
                platform, wi["name"], c.get("id"), c.get("type"), c.get("actualType"), c.get("name"),
                c.get("string"), c.get("zone"),
                c.get("topMd"), c.get("bottomMd"), c.get("topMdM"), c.get("bottomMdM"),
                c.get("outerDiameter"), c.get("innerDiameter"), c.get("weight"), c.get("grade"),
                "Yes" if c.get("estimated") else None,
                ", ".join(c.get("estimatedFields") or []) or None,
                c.get("estimationNote"), c.get("source"),
            ])
        for fe in d["floatEquipment"]:
            other_rows.append([platform, wi["name"], "floatEquipment", fe.get("type"), fe.get("name"),
                               fe.get("md"), None, fe.get("casingId")])
        for pb in d["plugBackCement"]:
            other_rows.append([platform, wi["name"], "plugBackCement", None, pb.get("name"),
                               pb.get("topMd"), pb.get("bottomMd"), None])

    wb = Workbook()

    # ---------- Overview ----------
    ws = wb.active
    ws.title = "Overview"
    ws["A1"] = "Dulang Well Schematic Summary — Platforms A, B, C, D"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = (f"Built {datetime.date.today().isoformat()} from WELL_SCHEMATIC/"
                + ", ".join(f for _, f in FOLDERS)
                + ". Depths: MD in ft (…Md) and m (…MdM) as stored in the JSON.")
    ov_hdr = ["Platform", "Source folder", "Wells (JSON files)", "Components", "Flagged (estimated)",
              "% flagged", "Float equipment entries", "Plug-back cement entries",
              "Wells with no components", "Wells with no total depth", "Wells with no completion date",
              "EDAFY wellbores", "EDAFY wellbores with schematic", "EDAFY wellbores without schematic"]
    ed_count = collections.Counter(edafy_platform(r[1]) for r in ed_data)
    ed_with = collections.Counter(edafy_platform(r[1]) for r in ed_data if r[1] in schem_by_name)
    ov_rows = []
    for platform, folder in FOLDERS:
        p = per_platform[platform]
        ov_rows.append([platform, folder, p["wells"], p["components"], p["flagged"],
                        round(p["flagged"] / p["components"], 3) if p["components"] else None,
                        p["float"], p["pbc"], p["empty"], p["no_td"], p["no_date"],
                        ed_count[platform], ed_with[platform], ed_count[platform] - ed_with[platform]])
    xp = "Exploration / appraisal"
    ov_rows.append([xp, "(EDAFY only)", 0, 0, 0, None, 0, 0, 0, 0, 0,
                    ed_count[xp], ed_with[xp], ed_count[xp] - ed_with[xp]])
    tot = [sum(r[i] or 0 for r in ov_rows) for i in range(2, 14)]
    tot[3] = round(tot[2] / tot[1], 3) if tot[1] else None
    ov_rows.append(["Total", None] + tot)
    write_table(ws, ov_hdr, ov_rows, start_row=4)
    for r in range(5, 5 + len(ov_rows)):
        ws.cell(row=r, column=6).number_format = "0.0%"
    for c in range(1, len(ov_hdr) + 1):
        ws.cell(row=4 + len(ov_rows), column=c).font = TOTAL_FONT
    ws.freeze_panes = None

    # component types by platform
    start = 4 + len(ov_rows) + 3
    ws.cell(row=start - 1, column=1, value="Component types by platform").font = Font(bold=True, size=12)
    tp_hdr = ["Component type"] + [p for p, _ in FOLDERS] + ["Total"]
    tp_rows = []
    for t in all_types:
        counts = [sum(row[2 + all_types.index(t)] for row in type_matrix if row[0] == p) for p, _ in FOLDERS]
        tp_rows.append([t] + counts + [sum(counts)])
    tp_rows.sort(key=lambda r: -r[-1])
    write_table(ws, tp_hdr, tp_rows, start_row=start)
    ws.freeze_panes = None
    ws.auto_filter.ref = None
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 16

    # ---------- Wells ----------
    ws = wb.create_sheet("Wells")
    write_table(ws, well_hdr, well_rows, widths={
        "Completion type": 40, "Tubing": 30, "Surface casing": 35, "Packer fluid": 35,
        "Completion fluid": 30, "Note": 35, "File": 22, "Well": 16})
    note_col = len(well_hdr)
    for r in range(2, len(well_rows) + 2):
        if ws.cell(row=r, column=note_col).value:
            for c in range(1, note_col + 1):
                ws.cell(row=r, column=c).fill = WARN_FILL

    # ---------- Component counts ----------
    ws = wb.create_sheet("Component Counts")
    write_table(ws, ["Platform", "Well"] + all_types + ["Total"], type_matrix)
    for c in range(3, len(all_types) + 4):
        ws.column_dimensions[get_column_letter(c)].width = 11
    ws.freeze_panes = "C2"

    # ---------- Components ----------
    ws = wb.create_sheet("Components")
    comp_hdr = ["Platform", "Well", "Component id", "Type", "Actual type (as printed)", "Name",
                "String", "Zone", "Top MD (ft)", "Bottom MD (ft)", "Top MD (m)", "Bottom MD (m)",
                "OD (in)", "ID (in)", "Weight", "Grade", "Estimated", "Estimated fields",
                "Estimation note", "Source"]
    write_table(ws, comp_hdr, comp_rows, widths={"Name": 32, "Estimation note": 45, "Source": 60,
                                                 "Actual type (as printed)": 28, "Component id": 24})

    # ---------- Other sections ----------
    ws = wb.create_sheet("Float & Plug-back")
    write_table(ws, ["Platform", "Well", "Section", "Type", "Name", "Top / MD (ft)", "Bottom MD (ft)",
                     "Casing ID"], other_rows, widths={"Name": 50})

    # ---------- EDAFY wells ----------
    ws = wb.create_sheet("EDAFY Wells")
    ed_hdr = ["No.", "Platform", "Schematic JSON", "Schematic file", "Note"] + ed_header
    ed_rows = []
    rev_cand = {v: k for k, v in EDAFY_CANDIDATES.items()}
    for i, r in enumerate(sorted(ed_data, key=lambda r: well_sort_key(r[1])), 1):
        sc = schem_by_name.get(r[1])
        note = None
        if r[1] in rev_cand:
            note = f"Possible match for schematic {rev_cand[r[1]]} (name differs; not linked)"
        vals = [v.date() if isinstance(v, datetime.datetime) else v for v in r]
        ed_rows.append([i, edafy_platform(r[1]), "Yes" if sc else "No",
                        f"{sc[0]}/{sc[1]}" if sc else None, note] + vals)
    write_table(ws, ed_hdr, ed_rows, widths={"Note": 40, "OPERATOR": 28, "Schematic file": 30})
    for r in range(2, len(ed_rows) + 2):
        if ws.cell(row=r, column=3).value == "No":
            for c in range(1, 6):
                ws.cell(row=r, column=c).fill = WARN_FILL
    ws.freeze_panes = "G2"

    wb.save(OUT)
    print(f"Wrote {OUT}")
    print(f"Wells: {len(well_rows)}  Components: {len(comp_rows)}  Float/PBC rows: {len(other_rows)}")
    for r in ov_rows:
        print(r)


if __name__ == "__main__":
    main()
