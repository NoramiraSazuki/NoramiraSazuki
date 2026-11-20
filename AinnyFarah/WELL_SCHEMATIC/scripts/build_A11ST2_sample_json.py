#!/usr/bin/env python3
"""Map 'Dulang A11ST2 Well Schematic.xlsx' into the structure of output_A11ST2/well-schematic-sample.rtf (JSON).

Rule for this output: data not stated in the source is left empty (null for numbers, "" for text) — no template
defaults, no assumed values (e.g. casing tops not given -> null; trajectoryType not stated -> "").
Extraction reuses build_A11ST2.extract (cells + drawing text; images OCR'd earlier, no text).
Writes one new file; every existing file in output_A11ST2 is left untouched (verified by sha256).
"""
import hashlib
import json
import re
import subprocess
import sys
from pathlib import Path

from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parent))
import build_A11ST2 as A  # noqa: E402
import build_well_schematic as B  # noqa: E402

FOLDER = A.OUTDIR
TPL = FOLDER / "well-schematic-sample.rtf"
OUT = B.versioned(FOLDER / "Dulang_A11ST2_well-schematic-sample.json")
sha = lambda p: hashlib.sha256(p.read_bytes()).hexdigest()


def main():
    before = {p.relative_to(FOLDER): sha(p) for p in FOLDER.rglob("*") if p.is_file()}
    h_src = sha(A.SRC)
    tpl = json.loads(subprocess.run(["textutil", "-convert", "txt", "-stdout", str(TPL)], capture_output=True, text=True, check=True).stdout)
    info, comps, plug, forms, tops, flags, nodest = A.extract()
    ws = load_workbook(A.SRC).worksheets[0]
    v = lambda a: ws[a].value
    keymap = {r[0]: r[2] for r in load_workbook(B.TPL)["Component_Types"].iter_rows(min_row=2, values_only=True)}
    nearest = {f[1] for f in flags if f[2] == "Nearest Component_Type"}

    # wellInfo: template keys in template order; only what the source states
    stated = {
        "completionDate": info["Completion_Date"][0],          # G11
        "completionFluid": v("T12").strip(),                    # T12
        "completionType": v("G12"),                             # G12 WELL TYPE
        "maxDeviation": v("G15"),                               # G15
        "name": info["Well_Name"][0],                           # B2
        "packerFluid": re.sub(r"\s+", " ", v("T13")).strip(),   # T13
        "surfaceCasing": v("T8").strip(),                       # T8 (outermost casing, 18-5/8" @ 563 m)
        "totalDepth": info["Total_Depth_MD"][0],                # L82 TD 2185 m -> ft
        "tubing": v("T11"),                                     # T11
        "wellhead": v("G14"),                                   # G14
    }
    wi = {k: stated.get(k, None if isinstance(x, (int, float)) and not isinstance(x, bool) else "") for k, x in tpl["wellInfo"].items()}
    wi["totalDepthM"] = info["Total_Depth_MD"][1]

    out_c = []
    for c in comps:
        actual, string, top_m, bot_m = c["_extra"]
        assumed = "assumed" in str(c["_src"].get("Top_MD"))
        if assumed:                                     # top not stated in source -> empty
            top_m = None
        d = {"id": c["Component_ID"], "type": keymap[c["Component_Type"]], "name": c["Name"],
             "topMd": B.ft(top_m), "bottomMd": c["Bottom_MD"], "outerDiameter": c["OD"], "innerDiameter": c["ID"]}
        for col, key in (("Weight", "weight"), ("Grade", "grade"), ("Manufacturer", "manufacturer"), ("Model", "model"),
                         ("Description", "description"), ("Packer_Type", "packerType"), ("Sleeve_State", "sleeveState")):
            if c[col] not in (None, ""):
                d[key] = c[col]
        if c["Has_Valve"] == "Yes":
            d["hasValve"] = True
        ref = c["_src"]["Name"]
        if ref in nearest:
            d.update(estimated=True, estimatedFields=["type"],
                     estimationNote=f"no exact type for '{actual}' — nearest '{c['Component_Type']}'")
        d["source"] = f"Cell {ref}: {c['Name']}"
        d.update(actualType=actual, string=string or "", topMdM=top_m, bottomMdM=bot_m)
        out_c.append(d)

    sections = {"wellInfo": wi, "surveys": [], "components": out_c, "cement": [], "plugBackCement": [], "centralizers": [],
                "floatEquipment": [], "fluidContacts": [],
                "formations": [dict(name=f["Name"], topTvd=f["Top_TVD"], bottomTvd=f["Bottom_TVD"], topTvdM=f["_extra"][0],
                                    bottomTvdM=f["_extra"][1]) for f in forms],
                "formationTops": [dict(name=t["Name"], md=t["MD"], tvd=t["TVD"], mdM=t["_extra"][0], tvdM=t["_extra"][1]) for t in tops]}
    doc = {k: sections[k] for k in tpl}
    OUT.write_text(json.dumps(doc, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")

    # validation
    back = json.loads(OUT.read_text(encoding="utf-8"))
    after = {p.relative_to(FOLDER): sha(p) for p in FOLDER.rglob("*") if p.is_file() and p != OUT}
    chk = [
        ("Valid JSON; sections and order identical to sample", list(back) == list(tpl)),
        ("wellInfo keys identical to sample, in order (+ totalDepthM)", list(back["wellInfo"])[:len(tpl["wellInfo"])] == list(tpl["wellInfo"])),
        ("No template sample values carried over", not any(c["id"] in {t["id"] for t in tpl["components"]} for c in back["components"])
         and back["surveys"] == [] and all(t["name"] not in {x["name"] for x in tpl["formationTops"]} for t in back["formationTops"])),
        ("Unstated wellInfo values empty (null / \"\")", all(back["wellInfo"][k] in (None, "") for k in tpl["wellInfo"] if k not in stated)),
        ("No assumed depths (casing tops not stated -> null)", all(c["topMd"] is None for c in back["components"] if c["type"] == "casing")),
        ("Every ft depth = m × 3.28084", all((c[a] is None and c[b] is None) or abs(c[a] - round(c[b] * B.FT, 3)) < 1e-9
                                            for c in back["components"] for a, b in (("topMd", "topMdM"), ("bottomMd", "bottomMdM")))),
        ("Every component type is a template type key", all(c["type"] in set(keymap.values()) for c in back["components"])),
        ("Existing files in output_A11ST2 unchanged", before == after),
        ("Source file unchanged", sha(A.SRC) == h_src),
    ]
    for name, ok in chk:
        print("PASS" if ok else "FAIL", "-", name)
    print(OUT.name, "| components", len(out_c), "| formations", len(sections["formations"]), "| formationTops", len(sections["formationTops"]))


if __name__ == "__main__":
    main()
