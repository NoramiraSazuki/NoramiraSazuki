#!/usr/bin/env python3
"""Map 'Dulang A11ST2 Well Schematic.xlsx' (read-only) into the Excel template and the JSON template.

Text comes from the sheet cells and drawing text boxes; the four embedded images were OCR'd locally with
Tesseract (output_A11ST2/ocr) and contain no text (equipment symbols / vector logo).
Same decisions as A04ST1: depths m -> ft with meters kept; nearest Component_Type with the real type kept;
later annotations (Oct 2023 gauge installation) included. Writers are reused from build_well_schematic(.json).
"""
import re
import sys
import zipfile
from pathlib import Path

from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parent))
import build_well_schematic as B  # noqa: E402
import build_well_schematic_json as J  # noqa: E402

SRC = B.ROOT / "Dulang A11ST2 Well Schematic.xlsx"
OUTDIR = B.ROOT / "output_A11ST2"

# (row, Component_Type, actual type as reported, nearest?, extra fields)  name T, top P, min ID X, max OD Z
EQUIP = [
    (21, "Tubing Hanger", "TUBING HANGER", False, {}),
    (26, "Flow Coupling", "FLOW COUPLING", False, {}),
    (27, "SCSSV", "TRSCSSSV (tubing-retrievable SCSSV)", False, {}),
    (28, "Flow Coupling", "FLOW COUPLING", False, {}),
    (32, "SPM (Gas Lift)", "IPO gas-lift valve 10/64", True, {"Has_Valve": "Yes", "Description": "T33"}),
    (35, "SPM (Gas Lift)", "IPO gas-lift valve 10/64", True, {"Has_Valve": "Yes", "Description": "T36"}),
    (38, "SPM (Gas Lift)", "IPO gas-lift valve 10/64", True, {"Has_Valve": "Yes", "Description": "T39"}),
    (41, "SPM (Gas Lift)", "ORIFICE gas-lift valve 12/64", True, {"Has_Valve": "Yes"}),
    (45, "Sliding Sleeve", "SSD", False, {}),
    (50, "Packer", "AHR PACKER (Cut-to-release)", False, {"Model": "AHR"}),
    (52, "Landing Nipple", "X LANDING NIPPLE", False, {}),
    (55, "No-Go Nipple", "NO-GO NIPPLE", False, {}),
    (57, "Mule Shoe", "WEG (wireline entry guide)", True, {}),
    (65, "Packer", "GP VGH PACKER ASSY", False, {"Model": "VGH"}),
    (66, "PBR", "UPPER SEAL BORE EXTENSION", True, {}),
    (69, "Screen", "WWS (wire-wrap screen)", False, {}),
    (72, "Seal Assembly", "COLLET SEAL ASSEMBLY", False, {}),
    (73, "Packer", "WIRELINE SET PACKER WITH TAIL PIPE (AWG)", False, {}),
]
# casing rows: (cell, label, grade, connection) — grade / connection checked against the cell text
CASINGS = [("T8", "CASING", "X-56", "VETCO RL-4C"), ("T9", "CASING", "K-55", "BTC"), ("T10", "CASING", "L-80", "JFE Bear")]


def extract():
    ws = load_workbook(SRC).worksheets[0]
    v = lambda a: ws[a].value
    drawing = [t for t in re.findall(r"<a:t>(.*?)</a:t>", zipfile.ZipFile(SRC).read("xl/drawings/drawing1.xml").decode()) if t.strip()]
    flags, comps, used = [], [], set()

    def comp(ctype, name, name_ref, top_ref=None, bot_ref=None, string="", actual="", top_m=None, bot_m=None, srcs=None, **kw):
        top = v(top_ref) if top_ref else top_m
        bot = v(bot_ref) if bot_ref else bot_m
        row = {c: None for c in B.COMP_COLS}
        row.update(Component_ID=B.slug(name, used), Component_Type=ctype, Name=name, Top_MD=B.ft(top), Bottom_MD=B.ft(bot), **kw)
        row["_extra"] = [actual, string, top, bot]
        row["_src"] = dict(Name=name_ref, Top_MD=top_ref or kw.pop("_x", None), Bottom_MD=bot_ref, **(srcs or {}))
        comps.append(row)
        return row

    comp("Xmas Tree", v("G13"), "G13", actual="XMASS TREE", Manufacturer=None)
    comp("Wellhead", v("G14"), "G14", actual="WELLHEAD")
    flags.append(("Components", "G13 / G14", "Depth not given", "Xmas tree and wellhead have no depths — Top_MD/Bottom_MD blank."))
    for ref, label, grade, conn in CASINGS:
        t = v(ref)
        assert grade in t and conn in t, (ref, t)
        depth = float(re.search(r"@\s*([\d.]+)\s*m", t).group(1))
        r = comp("Casing", t.strip(), ref, top_m=0, bot_m=depth, actual=label, OD=B.frac(t), Weight=float(re.search(r"([\d.]+)\s*ppf", t).group(1)),
                 Grade=grade, Description=f"Connection: {conn}", srcs=dict(OD=ref, Weight=ref, Grade=ref, Description=ref))
        r["_src"]["Top_MD"], r["_src"]["Bottom_MD"] = "assumed 0 (hung at wellhead)", f"{ref} '@ {depth}m MDDF'"
    flags.append(("Components", "T8:T10", "Assumed top", "Casing Top_MD = 0 (hung at wellhead) — source gives only the '@' depth."))
    flags.append(("Components", "T9", "Window, not shoe", "13-3/8\" depth 1105.38 m is the casing window (sidetrack exit) — used as Bottom_MD; the original shoe depth is not given."))
    flags.append(("Components", "T8", "Grade outside dropdown", "X-56 is written as reported; not in the Grade dropdown (H-40 … Q-125)."))
    # tubing string: hanger (P21 = 0) to WEG top (P57)
    t = comp("Tubing", v("T23").strip(), "T23", top_ref="P21", bot_ref="P57", string="Single", actual="TUBING", OD=v("Z23"), ID=v("X23"),
             Weight=9.2, Grade="L80-13Cr", srcs=dict(OD="Z23", ID="X23", Weight="T11", Grade="T11"))
    assert "9.2ppf" in v("T11") and "L80-13Cr" in v("T11")
    flags.append(("Components", "T23", "Bottom from WEG", f"Tubing Bottom_MD = WEG top P57 {v('P57')} m (only top-of-equipment depths are given)."))
    flags.append(("Components", "T23", "Grade outside dropdown", "Tubing grade 'L80-13Cr' written as reported (T11); not in the Grade dropdown."))
    for r, ctype, actual, near, ex in EQUIP:
        ex = dict(ex)
        srcs = {"OD": f"Z{r}", "ID": f"X{r}"}
        if "Description" in ex:
            srcs["Description"] = ex["Description"]
            ex["Description"] = v(ex["Description"]).strip()
        top_ref = f"P{r}" if v(f"P{r}") is not None else None
        c = comp(ctype, v(f"T{r}").strip(), f"T{r}", top_ref=top_ref, string="Single", actual=actual,
                 OD=v(f"Z{r}"), ID=v(f"X{r}"), srcs=srcs, **ex)
        if near:
            flags.append(("Components", f"T{r}", "Nearest Component_Type", f"'{actual}' has no exact type → '{ctype}'; real type kept in Actual_Equipment_Type."))
        if top_ref is None:
            flags.append(("Components", f"T{r}", "Depth not given", f"'{c['Name']}' has no depth in column P — Top_MD blank (kept in sequence)."))
        if r == 55:
            c["Description"] = re.sub(r"\s+", " ", " — ".join(drawing)).strip()
            c["_src"]["Description"] = "drawing text boxes (+ T77:T79)"
            flags.append(("Components", "T55", "Annotation placement", f"'{c['Description']}' attached to the No-Go Nipple at {v('P55')} m: the note says 'set at XN Nipple' and this is the only no-go (XN-type) nipple. Update note T77: '{v('T77')}'."))
        if r == 69:
            flags.append(("Components", "T69", "Bottom not given", "Screen Bottom_MD blank — only the top (P69) is given; Z69 8.25 is the centralizer OD."))
    flags.append(("Components", "SPM rows", "Has_Valve", "IPO and orifice gas-lift valves are live valves → Has_Valve = Yes; PTRO values in Description."))
    flags.append(("Components", "T27", "Type by category", "TRSCSSSV (tubing-retrievable surface-controlled subsurface safety valve) typed SCSSV."))
    pb = float(re.match(r"\s*([\d.]+)", v("G16")).group(1))
    p = comp("PBTD", "PBTD", "B16", top_m=pb, actual="PBTD")
    p["_src"]["Top_MD"] = "G16"

    # zone E14 -> Formations / Formation_Tops
    tv = re.match(r"\s*([\d.]+)\s*-\s*([\d.]+)\s*mTVDS", v("F69"))
    md = re.match(r"\s*([\d.]+)\s*-\s*([\d.]+)\s*mMDDF", v("F70"))
    t1, t2, m1 = float(tv.group(1)), float(tv.group(2)), float(md.group(1))
    name = v("F68").strip()
    forms = [{"Name": name, "Top_TVD": B.ft(t1), "Bottom_TVD": B.ft(t2), "_extra": [t1, t2], "_src": dict(Name="F68", Top_TVD="F69", Bottom_TVD="F69")}]
    tops = [{"Name": name, "MD": B.ft(m1), "TVD": B.ft(t1), "_extra": [m1, t1], "_src": dict(Name="F68", MD="F70", TVD="F69")}]
    flags.append(("Formations / Formation_Tops", "F69", "TVD reference", "TVD given as 'mTVDS' (likely TVD subsea) while MD is from the drill floor — written as reported; please confirm the datum."))
    flags.append(("Formations / Formation_Tops", "E14", "Not stated", "Lithology, Color, Symbol, Is_Reservoir not given — blank. Zone bottom MD 2112.2 m has no Formation_Tops column."))

    td = v("L82")
    info = {"Well_Name": (v("B2").split(":", 1)[1].strip(), None, "B2"), "Field": (v("G9"), None, "G9"),
            "Total_Depth_MD": (B.ft(td), td, "L82 (TD)"), "Completion_Date": (v("G11").strftime("%Y-%m-%d"), None, "G11")}
    flags.append(("Well_Info", "Well_Name", "Two forms", f"B2 '{info['Well_Name'][0]}' used; G8 WELL gives '{v('G8')}'."))
    flags.append(("Well_Info", "KB_Elevation and others", "Not stated", "UWI, Operator, County, State, Status, KB/Ground elevation, Lat/Long, Spud_Date, Trajectory_Type not given — blank."))
    flags.append(("Components", "R27:R57", "Angles", "Deviation angles per depth (column R) given, but no azimuth — Surveys left empty."))
    angles = ", ".join(f"{v(f'P{r}')} m: {v(f'R{r}')}°" for r in range(21, 60) if v(f"R{r}") is not None)
    nodest = [
        ("G8", v("G8"), "Well identifier (alternate form) — Well_Name taken from B2"),
        ("G10", v("G10"), "Location — no Well_Info row"), ("G12", v("G12"), "Well type — no Well_Info row (JSON: completionType)"),
        ("G15", v("G15"), "Max deviation — no Well_Info row (JSON: maxDeviation)"),
        ("T11", v("T11"), "Tubing summary — used for tubing grade/weight (JSON: tubing)"),
        ("T12", v("T12"), "Completion fluid — no Excel destination (JSON: completionFluid)"),
        ("T13", v("T13"), "Packer fluid — no Excel destination (JSON: packerFluid)"),
        ("R column", angles, "Deviation angles — Surveys needs azimuth"),
        ("T30", v("T30"), "Repeat tubing label — covered by tubing row"),
        ("F61:F62", f"{v('F61')} / {v('F62')}", "Behind-casing zones without depth — not mapped"),
        ("F70", v("F70"), "Zone bottom MD / thickness — no Formation_Tops column"),
        ("F71", v("F71"), "Zone pressure and temperature — no destination"),
        ("T77", v("T77"), "Update note (author/date) — no destination"),
        ("T80:X82", f"{v('T80')} {v('X80')} / {v('T81').strip()} {v('X81')} / {v('T82')}", "Sign-off block — no destination"),
        ("images", "4 embedded images (EMF logo, 3 symbol pictures)", "OCR'd locally: no text"),
    ]
    return info, comps, [], forms, tops, flags, nodest


def wellinfo_a11(info, v):
    td_m = info["Total_Depth_MD"][1]
    vals = {
        "completionDate": (info["Completion_Date"][0], "G11"), "completionFluid": (v("T12").strip(), "T12"),
        "completionType": (v("G12"), "G12 WELL TYPE"), "groundElevation": (0, "not stated (template default kept)"),
        "kbElevation": (0, "not stated (template default kept)"), "maxDeviation": (v("G15"), "G15"), "name": (info["Well_Name"][0], "B2"),
        "packerFluid": (re.sub(r"\s+", " ", v("T13")).strip(), "T13"), "rig": ("-", "not stated (template default)"),
        "surfaceCasing": (v("T8").strip(), "T8"), "totalDepth": (info["Total_Depth_MD"][0], "L82", f"{td_m} m × {B.FT}"),
        "trajectoryType": ("Deviated", f"derived from G15 '{v('G15')}'"), "tubing": (v("T11"), "T11"), "wellhead": (v("G14"), "G14"),
    }
    flags = [("wellInfo", "completionType", "Interpretation", "WELL TYPE 'Single String 3-1/2\"' used as completionType."),
             ("wellInfo", "surfaceCasing", "Interpretation", "Outermost casing (18-5/8\" @ 563 m) used as surface casing."),
             ("wellInfo", "kbElevation / groundElevation / rig", "Not stated", "Template defaults (0 / '-') kept."),
             ("wellInfo", "trajectoryType", "Derived", "'Deviated' derived from max deviation 64.8 degree."),
             ("wellInfo", "packerFluid", "Truncated in source", "T13 ends '+ 2% CI ' — written as reported.")]
    return vals, [("totalDepthM", td_m, "L82", "meters kept (extra key)")], flags


def main():
    OUTDIR.mkdir(exist_ok=True)
    B.SRC, B.OUTDIR, B.extract = SRC, OUTDIR, extract
    B.OUT = B.versioned(OUTDIR / "Dulang_A11ST2_Well_Schematic_Populated.xlsx")
    B.REV = B.versioned(OUTDIR / "Dulang_A11ST2_Mapping_Review.xlsx")
    J.OUT = B.versioned(OUTDIR / "Dulang_A11ST2_Well_Schematic.json")
    J.REV = B.versioned(OUTDIR / "Dulang_A11ST2_JSON_Mapping_Review.xlsx")
    print("== Excel template")
    B.main()
    print("== JSON template")
    J.main(wellinfo_a11)


if __name__ == "__main__":
    main()
