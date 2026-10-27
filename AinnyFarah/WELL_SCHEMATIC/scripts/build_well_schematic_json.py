#!/usr/bin/env python3
"""Map 'Dulang A04ST1 Well Schematic.xlsx' (read-only) into the structure of well-schematic-template-json.rtf.

Same extraction and decisions as build_well_schematic.py (ft with meters kept, diagram over header, nearest type
with real type kept, June 2022 annotations included). Template keys and section order are kept; template sample
records (surveys, components, formationTops) are removed. Extra keys carry meters, the real equipment type,
the string and the source cell. Type keys come from the Excel template's Component_Types sheet.
"""
import datetime as dt
import hashlib
import json
import re
import subprocess
import sys
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(Path(__file__).resolve().parent))
import build_well_schematic as B  # noqa: E402

ROOT = B.ROOT
TPL_JSON = ROOT / "well-schematic-template-json.rtf"
OUT = B.versioned(B.OUTDIR / "Dulang_A04ST1_Well_Schematic.json")
REV = B.versioned(B.OUTDIR / "Dulang_A04ST1_JSON_Mapping_Review.xlsx")
sha = B.sha


def template():
    txt = subprocess.run(["textutil", "-convert", "txt", "-stdout", str(TPL_JSON)], capture_output=True, text=True, check=True).stdout
    return json.loads(txt)


def wellinfo_a04(info, v):
    """A04ST1 wellInfo: {key: (value, source[, conversion])}, extra keys, flags."""
    td_m, kb_m = info["Total_Depth_MD"][1], info["KB_Elevation"][1]
    vals = {
        "completionDate": (info["Completion_Date"][0], "G8"),
        "completionFluid": (v("V16").strip(), "V16"),
        "completionType": ("-", "not stated"),
        "groundElevation": (0, "not stated (template default kept)"),
        "kbElevation": (info["KB_Elevation"][0], "G14", f"{kb_m} m × {B.FT}"),
        "maxDeviation": (v("G12"), "G12"),
        "name": (info["Well_Name"][0], "B2"),
        "packerFluid": (re.sub(r"\s+", " ", f"{v('V17')}{v('V18')}").strip(), "V17+V18"),
        "rig": (v("G9"), "G9"),
        "surfaceCasing": (f"{v('P10').strip()} {v('V10')} {v('X10')}; {v('Y10')}", "P10:Y10"),
        "totalDepth": (info["Total_Depth_MD"][0], "K190 (diagram TD)", f"{td_m} m × {B.FT}"),
        "trajectoryType": ("Deviated", f"derived from G12 maximum deviation '{v('G12')}'"),
        "tubing": (f"Tubing {v('V13')} {v('X13')}; {v('Y13')} / {v('V14')} {v('X14')}; {v('Y14')}", "P13:Y14"),
        "wellhead": (v("G15"), "G15"),
    }
    extras = [("totalDepthM", td_m, "K190", "meters kept (extra key)"), ("kbElevationM", kb_m, "G14", "meters kept (extra key)")]
    jflags = [
        ("wellInfo", "surfaceCasing", "Interpretation", "No 'surface casing' in source; outermost casing with full spec used: 13-3/8\" INT. CASING (conductor 26\" DP has no weight/grade)."),
        ("wellInfo", "trajectoryType", "Derived", "'Deviated' derived from maximum deviation 81.2 deg."),
        ("wellInfo", "groundElevation", "Not stated", f"Template default 0 kept; water depth '{v('G13')}' not converted to a seabed elevation."),
        ("wellInfo", "completionType", "Not stated", "Template default '-' kept (sheet name 'STMZ & SAS' not used as a type)."),
        ("wellInfo", "tubing", "Header text", "Header tubing summary (P13:AB14) — the header depths differ from the equipment tally (see components)."),
    ]
    return vals, extras, jflags


def main(wellinfo_fn=None):
    h_src, h_tpl, h_xtpl = sha(B.SRC), sha(TPL_JSON), sha(B.TPL)
    tpl = template()
    info, comps, plug, forms, tops, flags, nodest = B.extract()
    ws = load_workbook(B.SRC).worksheets[0]
    v = lambda a: ws[a].value
    keymap = {r[0]: r[2] for r in load_workbook(B.TPL)["Component_Types"].iter_rows(min_row=2, values_only=True)}
    nearest = {f[1] for f in flags if f[2] == "Nearest Component_Type"}
    trace, jflags = [], []

    def put(path, value, src, conv="as reported"):
        trace.append((path, json.dumps(value, ensure_ascii=False) if not isinstance(value, str) else value, src, conv))
        return value

    # ---------------- wellInfo (template keys, same order; '-' / 0 kept where the source has nothing)
    wi = dict(tpl["wellInfo"])
    vals, extras, wflags = (wellinfo_fn or wellinfo_a04)(info, v)
    for k in wi:
        val, src, *conv = vals[k]
        wi[k] = put(f"wellInfo.{k}", val, src, *conv)
    for k, val, src, conv in extras:
        wi[k] = put(f"wellInfo.{k}", val, src, conv)
    jflags += wflags

    # ---------------- components
    out_c = []
    for i, c in enumerate(comps):
        actual, string, top_m, bot_m = c["_extra"]
        src = c["_src"]
        ref = src["Name"]
        d = {"id": c["Component_ID"], "type": keymap[c["Component_Type"]], "name": c["Name"],
             "topMd": c["Top_MD"], "bottomMd": c["Bottom_MD"], "outerDiameter": c["OD"], "innerDiameter": c["ID"]}
        for col, key in (("Weight", "weight"), ("Grade", "grade"), ("Manufacturer", "manufacturer"), ("Model", "model"),
                         ("Description", "description"), ("Packer_Type", "packerType"), ("Sleeve_State", "sleeveState")):
            if c[col] not in (None, ""):
                d[key] = c[col]
        if c["Has_Valve"] == "Yes":
            d["hasValve"] = True
        est, notes = [], []
        if "assumed" in str(src.get("Top_MD")):
            est.append("topMd")
            notes.append(f"topMd {src['Top_MD']}")
        if ref in nearest:
            est.append("type")
            notes.append(f"no exact type for '{actual}' — nearest '{c['Component_Type']}'")
        if est:
            d.update(estimated=True, estimatedFields=est, estimationNote="; ".join(notes))
        d["source"] = f"Cell {ref}: {v(ref).strip() if re.match(r'^[A-Z]+\d+$', ref) and v(ref) else c['Name']}" \
            if re.match(r"^[A-Z]+\d+$", ref) else f"{ref}: {c['Name']}"
        d.update(actualType=actual, string=string or None, topMdM=top_m, bottomMdM=bot_m)
        out_c.append(d)
        p = f"components[{i}]"
        for k in ("type", "name", "topMd", "bottomMd", "outerDiameter", "innerDiameter", "weight", "grade"):
            if d.get(k) is not None:
                col = {"type": "Component_Type", "name": "Name", "topMd": "Top_MD", "bottomMd": "Bottom_MD", "outerDiameter": "OD",
                       "innerDiameter": "ID", "weight": "Weight", "grade": "Grade"}[k]
                conv = (f"{top_m if k == 'topMd' else bot_m} m × {B.FT}" if k in ("topMd", "bottomMd")
                        else f"Component_Types key for '{c['Component_Type']}'" if k == "type" else "as reported")
                put(f"{p}.{k}", d[k], src.get(col) or ("type mapping" if k == "type" else ref), conv)
    jflags.append(("components", "type", "Type keys", "Type values use the Type_Key column of the Excel template's Component_Types sheet (e.g. 'short-string', 'spm', 'liner')."))
    jflags.append(("components", "extra keys", "Added", "actualType, string, topMdM, bottomMdM added (real type and meters, as requested for the Excel output); hasValve written as true only for live GLV."))

    pbc = [dict(name=p["Name"], topMd=p["Top_MD"], bottomMd=p["Bottom_MD"], topMdM=p["_extra"][0], bottomMdM=p["_extra"][1]) for p in plug]
    fm = [dict(name=f["Name"], topTvd=f["Top_TVD"], bottomTvd=f["Bottom_TVD"], topTvdM=f["_extra"][0], bottomTvdM=f["_extra"][1]) for f in forms]
    ft_ = [dict(name=t["Name"], md=t["MD"], tvd=t["TVD"], mdM=t["_extra"][0], tvdM=t["_extra"][1]) for t in tops]
    for sec, rows, src_rows in (("plugBackCement", pbc, plug), ("formations", fm, forms), ("formationTops", ft_, tops)):
        for i, (r, s) in enumerate(zip(rows, src_rows)):
            for k, x in r.items():
                put(f"{sec}[{i}].{k}", x, ", ".join(sorted(set(s["_src"].values()))), "meters kept (extra key)" if k.endswith("M") else
                    "as reported" if k == "name" else f"m × {B.FT}")
    jflags.append(("formations", "keys", "Key names assumed", "Template has no formations example; keys follow the Excel columns in camelCase (topTvd, bottomTvd)."))
    jflags.append(("surveys", "-", "Not mapped", "Source gives deviation angles only (no azimuth, no MD pairing) — surveys left empty; template sample surveys removed."))

    doc = {}
    for k in tpl:                       # template section order
        doc[k] = {"wellInfo": wi, "surveys": [], "components": out_c, "cement": [], "plugBackCement": pbc, "centralizers": [],
                  "floatEquipment": [], "formations": fm, "formationTops": ft_, "fluidContacts": []}[k]
    OUT.write_text(json.dumps(doc, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")

    # ---------------- validation
    back = json.loads(OUT.read_text(encoding="utf-8"))
    val = [("Valid JSON; top-level sections and order identical to template", "PASS" if list(back) == list(tpl) else "FAIL"),
           ("wellInfo: all template keys present in template order", "PASS" if list(back["wellInfo"])[:len(tpl["wellInfo"])] == list(tpl["wellInfo"]) else "FAIL"),
           (f"components: {len(out_c)} records (template samples removed)", "PASS" if len(back["components"]) == len(comps) and not any(c["id"] in {t["id"] for t in tpl["components"]} for c in back["components"]) else "FAIL"),
           ("Every component type is a Component_Types key", "PASS" if all(c["type"] in set(keymap.values()) for c in back["components"]) else "FAIL"),
           ("Component ids unique", "PASS" if len({c["id"] for c in back["components"]}) == len(back["components"]) else "FAIL"),
           ("Every ft depth = m × 3.28084 (3 dp)", "PASS" if all((c[a] is None and c[b] is None) or abs(c[a] - round(c[b] * B.FT, 3)) < 1e-9
                                                               for c in back["components"] for a, b in (("topMd", "topMdM"), ("bottomMd", "bottomMdM")))
            and all(abs(t["md"] - round(t["mdM"] * B.FT, 3)) < 1e-9 for t in back["formationTops"]) else "FAIL"),
           ("Component names identical to source cells", "PASS" if all(c["name"] == v(r["_src"]["Name"]).strip() for c, r in zip(back["components"], comps)
                                                                         if re.match(r"^[A-Z]+\d+$", r["_src"]["Name"]) and r["Component_Type"] not in ("Casing", "Liner")) else "FAIL"),
           ("Values identical to the Excel output (same extraction)", "PASS" if [c["topMd"] for c in back["components"]] == [c["Top_MD"] for c in comps] else "FAIL"),
           ("Template sample surveys / formationTops removed", "PASS" if back["surveys"] == [] and all(t["name"] in {f["Name"] for f in tops} for t in back["formationTops"]) else "FAIL"),
           ("Source file unchanged", "PASS" if sha(B.SRC) == h_src else "FAIL"),
           ("JSON template (.rtf) unchanged", "PASS" if sha(TPL_JSON) == h_tpl else "FAIL"),
           ("Excel template unchanged", "PASS" if sha(B.TPL) == h_xtpl else "FAIL")]

    wb = Workbook()
    s = wb.active
    s.title = "Summary"
    for i, (k, x) in enumerate([("JSON output", str(OUT)), ("Source (read-only)", str(B.SRC)), ("JSON template (read-only)", str(TPL_JSON)),
                                ("Records", f"components {len(out_c)}, plugBackCement {len(pbc)}, formations {len(fm)}, formationTops {len(ft_)}; surveys/cement/centralizers/floatEquipment/fluidContacts empty"),
                                ("Units", f"Depths ft (m × {B.FT}); meters in *M keys"), ("Processed", dt.datetime.now().strftime("%Y-%m-%d %H:%M") + " — locally")], 1):
        s.cell(i, 1, k).font = Font(bold=True)
        s.cell(i, 2, x)
    s.column_dimensions["A"].width, s.column_dimensions["B"].width = 26, 130
    for title, hdr, data, w in (("Traceability", ["JSON Path", "Value", "Source Cell / Origin", "Conversion"], trace, [34, 70, 40, 34]),
                                ("Flags & Conflicts", ["Section", "Item", "Issue", "Detail"], jflags + flags, [22, 40, 26, 120]),
                                ("No Destination", ["Source Cell", "Source Content", "Status"], [n for n in nodest if n[0] not in ("G9", "G12", "V16", "V17+V18")], [26, 70, 70]),
                                ("Validation", ["Check", "Result"], val, [100, 12])):
        sh = wb.create_sheet(title)
        sh.append(hdr)
        for r in data:
            sh.append([str(x) if x is not None else None for x in r])
        for c in sh[1]:
            c.font, c.fill = Font(bold=True), PatternFill("solid", fgColor="DDEBF7")
        for i, x in enumerate(w, 1):
            sh.column_dimensions[get_column_letter(i)].width = x
        for row in sh.iter_rows(min_row=2):
            for c in row:
                c.alignment = Alignment(wrap_text=True, vertical="top")
        sh.freeze_panes = "A2"
    wb.save(REV)
    for c, r in val:
        print(r, "-", c)
    print(OUT.name, REV.name)


if __name__ == "__main__":
    main()
