#!/usr/bin/env python3
"""Map 'Dulang A26ST1 Well Schematic.xlsx' (read-only) into the Excel template and the JSON template.

Text comes from sheet cells and drawing text boxes; the 9 embedded images were OCR'd locally with Tesseract
(output_A26ST1/ocr): only stock clip-art labels and the PETRONAS logo — no schematic data.
Same decisions as A04ST1/A11ST2: depths m -> ft with meters kept; nearest Component_Type with real type kept;
later intervention annotations included. Annotation depths in m-MDTHF are converted to m-MDDF with the sheet's
own rule (MDTHF = MDDF - DFE-THF, F9 = 18.8 m). Header depth formulas AA6:AA9 point at empty cells (cached 0)
and are not used.
"""
import html
import re
import sys
import zipfile
from pathlib import Path

from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parent))
import build_well_schematic as B  # noqa: E402
import build_well_schematic_json as J  # noqa: E402

SRC = B.ROOT / "Dulang A26ST1 Well Schematic.xlsx"
OUTDIR = B.ROOT / "output_A26ST1"

# (row, Component_Type, actual type as reported, nearest?, extra fields)  name S, top Q, min ID Y, max OD AA
EQUIP = [
    (15, "Tubing Hanger", "TUBING HANGER", False, {}),
    (20, "Flow Coupling", "FLOW COUPLING", False, {}),
    (21, "SCSSV", "SCSSV", False, {}),
    (22, "Flow Coupling", "FLOW COUPLING", False, {}),
    (26, "SPM (Gas Lift)", "SPM (GLV PTRO 747.10 psi, 10/64\")", False, {"Has_Valve": "Yes"}),
    (29, "SPM (Gas Lift)", "SPM (GLV PTRO 767.99 psi, 10/64\")", False, {"Has_Valve": "Yes"}),
    (32, "SPM (Gas Lift)", "SPM (GLV PTRO 782.04 psi, 10/64\")", False, {"Has_Valve": "Yes"}),
    (36, "SPM (Gas Lift)", "SPM (GLV 792.80 psi, 10/64\")", False, {"Has_Valve": "Yes"}),
    (39, "SPM (Gas Lift)", "SPM (GLOV 12/64\")", False, {"Has_Valve": "Yes"}),
    (43, "Sliding Sleeve", "SSD (CLOSED)", False, {"Sleeve_State": "closed"}),
    (47, "Packer", "AHC PACKER", False, {"Model": "AHC 284"}),
    (51, "No-Go Nipple", "XN NIPPLE (2.690\")", False, {}),
    (52, "ST Slot Locator", "NO-GO LOCATOR", True, {}),
    (54, "Mule Shoe", "SELF ALIGNING MULE SHOE", False, {}),
    (55, "Packer", "VTA PACKER", False, {"Model": "VTA"}),
    (56, "PBR", "SEAL BORE EXTENSION", True, {}),
    (57, "Crossover", "3 WAY SUB (4-1/2\" x 4-1/2\" x 2-7/8\")", True, {}),
    (59, "Packer", "OIL SWELL PACKER", False, {}),
    (61, "Sliding Sleeve", "SLIMLINE SSD (OPEN NOV 2023)", False, {"Sleeve_State": "open"}),
    (63, "Seal Assembly", "SEAL ASSY", False, {}),
    (66, "Sliding Sleeve", "SLIMLINE SSD (CLOSED NOV 2023)", False, {"Sleeve_State": "closed"}),
    (69, "Seal Assembly", "SEAL ASSY", False, {}),
    (72, "Sliding Sleeve", "SLIMLINE SSD (CLOSED JUL 2017)", False, {"Sleeve_State": "closed"}),
    (80, "No-Go Nipple", "XN NIPPLE", False, {}),
    (81, "Mule Shoe", "SELF ALIGN MULESHOE (inner string)", False, {}),
    (84, "EOT", "DOUBLE VALVE FLOAT SHOE", True, {}),
]
UPPER = {15, 20, 21, 22, 26, 29, 32, 36, 39, 43, 47, 51, 52, 54}


def extract():
    wb = load_workbook(SRC)
    ws = wb.worksheets[0]
    cached = load_workbook(SRC, data_only=True).worksheets[0]
    v = lambda a: ws[a].value
    dfe_thf = v("F9")
    xml = zipfile.ZipFile(SRC).read("xl/drawings/drawing1.xml").decode()
    notes = [html.unescape("".join(re.findall(r"<a:t>(.*?)</a:t>", sp))).strip() for sp in re.findall(r"<xdr:sp\b.*?</xdr:sp>", xml, re.S)]
    notes = [n for n in notes if n]
    flags, comps, used = [], [], set()

    def comp(ctype, name, name_ref, top_ref=None, bot_ref=None, string="", actual="", top_m=None, bot_m=None, srcs=None, top_src=None, **kw):
        top = v(top_ref) if top_ref else top_m
        bot = v(bot_ref) if bot_ref else bot_m
        row = {c: None for c in B.COMP_COLS}
        row.update(Component_ID=B.slug(name, used), Component_Type=ctype, Name=name, Top_MD=B.ft(top), Bottom_MD=B.ft(bot), **kw)
        row["_extra"] = [actual, string, top, bot]
        row["_src"] = dict(Name=name_ref, Top_MD=top_ref or top_src, Bottom_MD=bot_ref, **(srcs or {}))
        comps.append(row)
        return row

    comp("Xmas Tree", v("F11"), "F11", actual="X-MAS TREE")
    comp("Wellhead", v("F10"), "F10", actual="WELLHEAD")
    flags.append(("Components", "F10 / F11", "Depth not given", "Xmas tree and wellhead have no depths — Top_MD/Bottom_MD blank."))
    hdr = lambda r: dict(OD=v(f"S{r}"), ID=v(f"U{r}"), Weight=v(f"R{r}"), Grade=v(f"X{r}"),
                         Description=f"Thread: {v(f'Y{r}')}; Drift: {v(f'W{r}')} in",
                         srcs=dict(OD=f"S{r}", ID=f"U{r}", Weight=f"R{r}", Grade=f"X{r}", Description=f"Y{r}, W{r}"))
    comp("Casing", v("E37").strip(), "E37", bot_ref="I37", top_m=0, top_src="assumed 0 (hung at wellhead)", actual=v("N6"), **hdr(6))
    comp("Liner", v("E89").strip(), "E89", top_ref="I32", bot_ref="I89", actual=v("N7"), **hdr(7))
    flags.append(("Components", "E37", "Assumed top", f"9-5/8\" casing Top_MD = 0 (hung at wellhead); bottom = diagram label I37 {v('I37')} m."))
    flags.append(("Components", "E89", "Liner interval", f"7\" liner from 'TOP OF 7\" LINER' I32 {v('I32')} m to shoe I89 {v('I89')} m."))
    flags.append(("Components", "AA6:AA9", "Header depths unusable", "Header DEPTH formulas (=I49, =I50, =R20, =R60) reference empty cells and evaluate to 0 — not used; diagram depths used."))
    t1 = dict(hdr(8)); t1.pop("ID"); t1["srcs"].pop("ID")
    comp("Tubing", v("S18").strip(), "S18", top_ref="Q15", bot_ref="Q21", string="Upper completion", actual=v("N8"), ID=v("Y18"),
         **{k: x for k, x in t1.items() if k != "srcs"}, srcs=dict(t1["srcs"], ID="Y18"))
    comp("Tubing", v("S24").strip(), "S24", top_ref="Q21", bot_ref="Q54", string="Upper completion", actual=v("N8"), ID=v("Y24"),
         **{k: x for k, x in t1.items() if k != "srcs"}, srcs=dict(t1["srcs"], ID="Y24"))
    n9 = f"{v('N9')} - {v('Q9')}\" {v('R9')} ppf {v('X9')} {v('Y9')}"
    comp("Tubing", n9, "N9:Y9", top_ref="Q59", bot_ref="Q84", string="Lower completion", actual=v("N9"), **hdr(9))
    flags.append(("Components", "S18 / S24", "Interval from equipment", f"3-1/2\" tubing: hanger Q15 {v('Q15')} → SCSSV Q21 {v('Q21')} → mule shoe Q54 {v('Q54')} m (labels 'ABOVE/BELOW SCSSV')."))
    flags.append(("Components", "N9", "Interval from equipment", f"2-7/8\" tubing Top_MD = oil swell packer Q59 {v('Q59')} m (VTA packer, seal bore extension and 3-way sub above it have no depths); Bottom_MD = float shoe Q84 {v('Q84')} m."))
    for r, ctype, actual, near, ex in EQUIP:
        top_ref = f"Q{r}" if v(f"Q{r}") is not None else None
        c = comp(ctype, v(f"S{r}").strip(), f"S{r}", top_ref=top_ref, string="Upper completion" if r in UPPER else "Lower completion",
                 actual=actual, OD=v(f"AA{r}"), ID=v(f"Y{r}"), srcs={"OD": f"AA{r}", "ID": f"Y{r}"}, **ex)
        if near:
            flags.append(("Components", f"S{r}", "Nearest Component_Type", f"'{actual}' has no exact type → '{ctype}'; real type kept in Actual_Equipment_Type."))
        if top_ref is None:
            flags.append(("Components", f"S{r}", "Depth not given", f"'{c['Name']}' has no depth in column Q — Top_MD blank (kept in sequence)."))
    flags.append(("Components", "Y31:AA31, Y35:AA35, Y58:AA58", "Sizes not assigned",
                  "ID/OD values on rows without an equipment label (2.88/5.313, 2.88/5.313, 2.813/4.04) — not attached to a component; the SPMs at rows 26–36 therefore have no ID/OD."))
    flags.append(("Components", "S51 / S80", "XN = no-go", "XN nipples typed No-Go Nipple (XN is the no-go profile)."))
    # perforations (zone labels A61/A66/A73 + 'Perf Interval' text)
    tops = []
    for name_r, perf_r in ((61, 62), (66, 67), (73, 74)):
        zone = v(f"A{name_r}").strip()
        m = re.match(r"\s*Perf Interval:\s*([\d.]+)m\s*-\s*([\d.]+)m MDDF", v(f"A{perf_r}"))
        a, b = float(m.group(1)), float(m.group(2))
        comp("Perforation", f"{zone} - {v(f'A{perf_r}').strip()}", f"A{name_r}+A{perf_r}", top_m=a, bot_m=b, top_src=f"A{perf_r}",
             string="", actual="PERF INTERVAL")["_src"]["Bottom_MD"] = f"A{perf_r}"
        tops.append({"Name": zone, "MD": B.ft(a), "TVD": None, "_extra": [a, None], "_src": dict(Name=f"A{name_r}", MD=f"A{perf_r}", TVD="not given")})
    flags.append(("Formation_Tops", "E34A / E34B / E40", "TVD not given", "Zone tops written with MD (top of perforation interval); TVD blank. No Formations rows (needs TVD)."))
    # intervention annotations (drawing)
    def at(pattern):
        return next(n for n in notes if re.search(pattern, n))
    tp = at(r"^Tubing punch")
    m = re.search(r"@\s*([\d.]+)\s*-\s*([\d.]+)m MDDF", tp)
    comp("Perforation", tp, "drawing: tubing punch", top_m=float(m.group(1)), bot_m=float(m.group(2)), top_src="drawing text",
         actual="TUBING PUNCH", Description="Post-completion annotation (Jul 2018)")["_src"]["Bottom_MD"] = "drawing text"
    sp = at(r"^Slickplug")
    comp("Mechanical Plug", sp, "drawing: slickplug 2018", top_m=float(re.search(r"@\s*([\d.]+)m MDDF", sp).group(1)), top_src="drawing text",
         actual="SLICKPLUG", Description="Post-completion annotation (Jul 2018)")
    for pat, ctype, actual, ref in ((r"^Junk Catcher", "Mechanical Plug", "JUNK CATCHER & SLICKPLUG", "drawing: junk catcher"),
                                    (r"Widepak", "Screen", "WIDEPAK + 3 JTS BORON ERTTSS (through-tubing sand control)", "drawing: widepak")):
        n = at(pat)
        thf = float(re.search(r"@\s*([\d.]+)m MDTHF", n).group(1))
        comp(ctype, n, ref, top_m=round(thf + dfe_thf, 3), top_src=f"drawing text {thf} m MDTHF + DFE-THF F9 {dfe_thf} m",
             actual=actual, Description=f"Post-completion annotation (Nov 2023); depth reported {thf} m MDTHF")
        flags.append(("Components", ref, "MDTHF → MDDF", f"'{n}': {thf} m MDTHF + {dfe_thf} m (F9 DFE-THF, the sheet's own R = Q - F9 rule) = {round(thf + dfe_thf, 3)} m MDDF."))
    flags.append(("Components", "drawing: slickplug 2018 / junk catcher", "Nearest Component_Type", "Slickplug / junk catcher typed Mechanical Plug; Widepak ERTTSS typed Screen; tubing punch typed Perforation."))
    flags.append(("Components", "drawing annotations", "Date precision", "Annotation dates are month-only (Jul 2018, Nov 2023) — kept in Description; Set_Date/Perforation_Date blank."))
    hud = at(r"HUD")

    td = v("F6")
    info = {"Well_Name": (v("A1").split(":", 1)[1].strip(), None, "A1"), "Total_Depth_MD": (B.ft(td), td, "F6"),
            "Completion_Date": (v("F4").strftime("%Y-%m-%d"), None, "F4")}
    flags.append(("Well_Info", "Total_Depth_MD", "Header only", f"TD from header F6 {td} m-MDDF (liner shoe I89 {v('I89')} m)."))
    flags.append(("Well_Info", "Field and others", "Not stated", f"Field not stated ('{v('A2')}' is the location); UWI, Operator, County, State, Status, KB/Ground elevation, Lat/Long, Spud_Date, Trajectory_Type blank."))
    flags.append(("Components", "E87", "PBTD without depth", "'PBTD' label has no depth — not mapped."))
    nodest = [
        ("A2", v("A2"), "Location / slot — no Well_Info row"), ("F5", v("F5"), "Rig — no Excel destination (JSON: rig)"),
        ("F7:G7", f"{v('F7')} {v('G7')}", "Max deviation — no Excel destination (JSON: maxDeviation)"),
        ("F8:G8", f"{v('F8')} {v('G8')}".strip(), "Water depth — no destination"),
        ("F9:G9", f"{v('F9')} {v('G9')}".strip(), "DFE-THF — used to convert MDTHF annotations; no destination"),
        ("Q10 / Q11", f"{v('Q10')} / {v('Q11')}", "Completion / packer fluid — JSON only"),
        ("AA6:AA9", "formulas =I49, =I50, =R20, =R60 (cached 0)", "Broken header depth references — not used"),
        ("R / J columns", "depths m-MDTHF (= MDDF - F9)", "Derived values — MDDF used"),
        ("E87", v("E87"), "PBTD label without depth"),
        ("Y31:AA31, Y35:AA35, Y58:AA58", "2.88/5.313, 2.88/5.313, 2.813/4.04", "Sizes on unlabeled rows — not assigned"),
        ("drawing", hud, "Hold-up depth / bailer observation — no destination"),
        ("S91:U92", f"{v('S91')} {v('U91')} / {v('S92')} =TODAY() (cached {cached['U92'].value:%Y-%m-%d})", "Update block — no destination"),
        ("images", "9 embedded images", "OCR'd locally: PETRONAS logo, stock clip-art labels ('Well drilling…Hydraulic fracturing'), symbols — no data"),
    ]
    return info, comps, [], [], tops, flags, nodest


def wellinfo_a26(info, v):
    td_m = info["Total_Depth_MD"][1]
    vals = {
        "completionDate": (info["Completion_Date"][0], "F4"), "completionFluid": (v("Q10"), "Q10"),
        "completionType": ("-", "not stated"), "groundElevation": (0, "not stated (template default kept)"),
        "kbElevation": (0, "not stated (template default kept)"), "maxDeviation": (f"{v('F7')} {v('G7')}", "F7:G7"),
        "name": (info["Well_Name"][0], "A1"), "packerFluid": (v("Q11"), "Q11"), "rig": (v("F5"), "F5"),
        "surfaceCasing": (f"{v('N6')} {v('Q6')}\" {v('R6')} ppf; {v('X6')}", "N6:X6"),
        "totalDepth": (info["Total_Depth_MD"][0], "F6", f"{td_m} m × {B.FT}"),
        "trajectoryType": ("Deviated", f"derived from F7 max deviation {v('F7')} deg"),
        "tubing": (f"Tubing {v('Q8')}\" {v('R8')} ppf; {v('X8')} / {v('Q9')}\" {v('R9')} ppf; {v('X9')}", "N8:X9"),
        "wellhead": (v("F10"), "F10"),
    }
    flags = [("wellInfo", "surfaceCasing", "Interpretation", "Only casing listed is the 9-5/8\" PRODUCTION CASING — used (outermost casing rule)."),
             ("wellInfo", "kbElevation / groundElevation / completionType", "Not stated", "Template defaults kept. DFE-THF 18.8 m is not a KB elevation (AMSL) — not used."),
             ("wellInfo", "trajectoryType", "Derived", "'Deviated' derived from max deviation 52.76 deg.")]
    return vals, [("totalDepthM", td_m, "F6", "meters kept (extra key)")], flags


def main():
    OUTDIR.mkdir(exist_ok=True)
    B.SRC, B.OUTDIR, B.extract = SRC, OUTDIR, extract
    B.OUT = B.versioned(OUTDIR / "Dulang_A26ST1_Well_Schematic_Populated.xlsx")
    B.REV = B.versioned(OUTDIR / "Dulang_A26ST1_Mapping_Review.xlsx")
    J.OUT = B.versioned(OUTDIR / "Dulang_A26ST1_Well_Schematic.json")
    J.REV = B.versioned(OUTDIR / "Dulang_A26ST1_JSON_Mapping_Review.xlsx")
    print("== Excel template")
    B.main()
    print("== JSON template")
    J.main(wellinfo_a26)


if __name__ == "__main__":
    main()
