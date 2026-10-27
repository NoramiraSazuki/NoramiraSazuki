#!/usr/bin/env python3
"""Map 'Dulang A04ST1 Well Schematic.xlsx' (read-only) into a copy of well-schematic-sample-template.xlsx.

User decisions:
  * depths converted m -> ft (x 3.28084, 3 dp) into the template columns; original meters kept in extra columns
  * header-vs-diagram conflicts: diagram wins (TD 3591, PBTD 3550, liner top 2513; liner bottom blank + flagged)
  * equipment with no exact Component_Type: nearest type, real equipment type kept in an extra column
  * June 2022 drawing annotations (bridge plug, TOC, seal-assembly leak) included
Template columns, header notes, styles and (x14) validations are untouched: only the five target sheet XML parts
are rewritten; extra columns are appended after the template's last column. Values are read from source cells.
"""
import datetime as dt
import hashlib
import re
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "Dulang A04ST1 Well Schematic.xlsx"
TPL = ROOT / "well-schematic-sample-template.xlsx"
OUTDIR = ROOT / "output"
FT = 3.28084


def versioned(p):
    k = 2
    q = p
    while q.exists():
        q = p.with_name(f"{p.stem}_v{k}{p.suffix}")
        k += 1
    return q


OUT = versioned(OUTDIR / "Dulang_A04ST1_Well_Schematic_Populated.xlsx")
REV = versioned(OUTDIR / "Dulang_A04ST1_Mapping_Review.xlsx")
sha = lambda p: hashlib.sha256(p.read_bytes()).hexdigest()
ft = lambda m: None if m is None else round(m * FT, 3)

COMP_COLS = ["Component_ID", "Component_Type", "Name", "Top_MD", "Bottom_MD", "OD", "ID", "Weight", "Grade", "Manufacturer",
             "Model", "Description", "Pressure_Rating", "Setting", "Packer_Type", "Set_Date", "Shots_Per_Foot", "Phasing",
             "Gun_Type", "Charge_Size", "Perforation_Date", "Plug_Type", "Sleeve_State", "Ports", "Has_Valve",
             "Wall_Thickness", "Stroke_Length", "Seal_Count", "Color"]
COMP_EXTRA = ["Actual_Equipment_Type (as reported)", "String (as reported)", "Top_MD (m MDDF)", "Bottom_MD (m MDDF)"]

# ---------------------------------------------------------------- source equipment tally
# (row, side, Component_Type, actual type as reported, extra fields)   side S = short string (F/K/B/D), L = long (U/S/AB/Z)
SHORT = [
    (28, "Flow Coupling", "FLOW COUPLING", {}),
    (29, "TRSV", "TRMAXX-5E (TRSV per string label 'From Tubing Hanger to TRSV')", {"Model": "TRMAXX-5E"}),
    (30, "Flow Coupling", "FLOW COUPLING", {}),
    (33, "SPM (Gas Lift)", "SPM (GLV: PPO 10/64\" TRO: 675)", {"Has_Valve": "Yes"}),
    (40, "SPM (Gas Lift)", "SPM (GLV: PPO 10/64\" TRO: 760)", {"Has_Valve": "Yes"}),
    (46, "SPM (Gas Lift)", "AUTO SPM DIAL", {}),
    (48, "SPM (Gas Lift)", "SPM (GLV: Dummy)", {}),
    (53, "SPM (Gas Lift)", "AUTO SPM DIAL", {}),
    (55, "SPM (Gas Lift)", "SPM (GLV: Dummy)", {}),
    (65, "Landing Nipple", "LANDING NIPPLE (2.750 OSN)", {}),
    (67, "Mechanical Plug", "PUMP OUT PLUG", {"Model": "PE-635"}),
    (68, "Pup Joint", "PERFORATED PUP", {}),
    (69, "Mule Shoe", "BULL NOSE", {}),
    (71, "EOT", "EOT", {}),
]
LONG = [
    (27, "Flow Coupling", "FLOW COUPLING", {}),
    (28, "TRSV", "TRMAXX-5E (TRSV per string label 'From Tubing Hanger to TRSV')", {"Model": "TRMAXX-5E"}),
    (29, "Flow Coupling", "FLOW COUPLING", {}),
    (32, "SPM (Gas Lift)", "SPM (GLV: PPO 10/64\" TRO: 660)", {"Has_Valve": "Yes"}),
    (35, "SPM (Gas Lift)", "AUTO SPM DIAL", {}),
    (37, "SPM (Gas Lift)", "SPM (GLV: Dummy)", {}),
    (42, "SPM (Gas Lift)", "AUTO SPM DIAL", {}),
    (44, "SPM (Gas Lift)", "SPM (GLV: Dummy)", {}),
    (50, "SPM (Gas Lift)", "AUTO SPM DIAL", {}),
    (52, "SPM (Gas Lift)", "SPM (GLV: Dummy)", {}),
    (59, "Sliding Sleeve", "SLIDING SLEEVE DOOR", {}),
    (61, "Packer", "DUAL STRING FEED THROUGH PACKER", {"Model": "OSD-1", "_string": "Short & Long (dual packer)"}),
    (73, "Pup Joint", "DIAL SENSOR", {}),
    (80, "Sliding Sleeve", "SLIDING SLEEVE DOOR", {}),
    (82, "Landing Nipple", "LANDING NIPPLE (2.813 OSX)", {}),
    (88, "Crossover", "CROSSOVER", {}),
    (93, "ST Slot Locator", "NO-GO LOCATOR", {}),
    (96, "Packer", "GP PACKER (VCA Versa Trieve)", {"Model": "Versa Trieve"}),
    (101, "Screen", "WIRE-WRAP SCREENS", {"_top": "K101", "_bottom": "K105", "Description": "K102"}),
    (106, "Crossover", "CROSSOVER", {}),
    (107, "PBR", "SEALBORE EXT", {}),
    (108, "ST Slot Locator", "NO-GO LOCATOR", {}),
    (109, "Seal Assembly", "SEAL ASSEMBLY", {}),
    (110, "Mule Shoe", "SELF ALIGN MULESHOE GUIDE", {}),
    (113, "Packer", "SUMP PACKER (VCA Versa Trieve)", {"Model": "Versa Trieve"}),
    (118, "Crossover", "not stated - thread change 2 7/8 VAM FJL Box x VAM TOP Pin", {}),
    (119, "Seal Assembly", "MSA SEAL ASSEMBLY", {}),
    (120, "Mule Shoe", "SELF ALIGN MULESHOE GUIDE", {}),
    (129, "Sliding Sleeve", "SLIMLINE SSD (OPEN)", {"Sleeve_State": "open"}),
    (135, "Blast Joint", "BLAST JOINT", {"_top": "K133", "_bottom": "K137", "Description": "K134"}),
    (140, "Crossover", "CROSSOVER", {}),
    (142, "Packer", "BLUEPACK MAX RH PACKER", {"Model": "BLUEPACK MAX RH"}),
    (144, "Crossover", "CROSSOVER", {}),
    (150, "Sliding Sleeve", "SLIMLINE SSD (OPEN)", {"Sleeve_State": "open"}),
    (157, "Blast Joint", "BLAST JOINT", {"_top": "K155", "_bottom": "K159", "Description": "K156"}),
    (160, "Crossover", "CROSSOVER", {}),
    (162, "Crossover", "CROSSOVER", {}),
    (163, "Packer", "BLUEPACK RH PACKER", {"Model": "BLUEPACK RH"}),
    (164, "Crossover", "CROSSOVER", {}),
    (166, "Crossover", "CROSSOVER", {}),
    (171, "Landing Nipple", "LANDING NIPPLE (2.313 XN)", {}),
    (173, "Crossover", "CROSSOVER", {}),
    (175, "Mechanical Plug", "PUMP OUT PLUG + WIRELINE ENTRY GUIDE", {}),
    (177, "EOT", "EOT", {}),
]
EXACT = {"FLOW COUPLING", "EOT", "CROSSOVER", "SLIDING SLEEVE DOOR", "SLIMLINE SSD (OPEN)", "SEAL ASSEMBLY",
         "MSA SEAL ASSEMBLY", "BLAST JOINT", "WIRE-WRAP SCREENS", "SELF ALIGN MULESHOE GUIDE"}


def frac(t):
    """'9-5/8\"' -> 9.625, '26\" DP' -> 26, 7 -> 7."""
    if isinstance(t, (int, float)):
        return float(t)
    m = re.match(r'\s*(\d+)(?:\s*-\s*(\d+)/(\d+))?\s*"', str(t))
    return int(m.group(1)) + (int(m.group(2)) / int(m.group(3)) if m.group(2) else 0)


def num(t):
    return float(re.match(r"\s*([\d.]+)", str(t)).group(1))


def slug(s, used):
    b = re.sub(r"[^a-z0-9]+", "-", s.lower()).strip("-")[:30].strip("-")
    k, out = 2, b
    while out in used:
        out, k = f"{b}-{k}", k + 1
    used.add(out)
    return out


# ---------------------------------------------------------------- extraction
def extract():
    wb = load_workbook(SRC, data_only=False)
    ws = wb["OP4 dual A04ST1 STMZ & SAS"]
    v = lambda a: ws[a].value
    S = ws.title
    trace, flags, comps, used = [], [], [], set()
    drawing = " ".join(t for t in re.findall(r"<a:t>(.*?)</a:t>", zipfile.ZipFile(SRC).read("xl/drawings/drawing1.xml").decode()))

    def comp(ctype, name, name_ref, top_ref=None, bot_ref=None, string="", actual="", top_m=None, bot_m=None, **kw):
        top = v(top_ref) if top_ref else top_m
        bot = v(bot_ref) if bot_ref else bot_m
        row = {c: None for c in COMP_COLS}
        row.update(Component_ID=slug(f"{string[:1].lower()}-{name}" if string else name, used), Component_Type=ctype,
                   Name=name, Top_MD=ft(top), Bottom_MD=ft(bot))
        row.update({k: x for k, x in kw.items() if not k.startswith("_")})
        row["_extra"] = [actual, string, top, bot]
        row["_src"] = dict(Name=name_ref, Top_MD=top_ref or kw.get("_top_src"), Bottom_MD=bot_ref or kw.get("_bot_src"),
                           **kw.get("_srcs", {}))
        comps.append(row)
        return row

    # tubulars (header block P9:AB14 + diagram depths)
    comp("Wellhead", v("G15"), "G15", string="", actual="WELL HEAD")
    for r, ctype, label in ((9, "Casing", "CONDUCTOR"), (10, "Casing", "INT. CASING"), (11, "Casing", "PRODUCTION CASING"),
                            (12, "Liner", "PRODUCTION LINER")):
        name = " ".join(str(x) for x in (v(f"V{r}"), v(f"X{r}"), v(f"Y{r}"), v(f"Z{r}")) if x)
        name = f'{v(f"P{r}").strip()} - {name}'.replace("- 7 ", '- 7" ')
        kw = dict(OD=frac(v(f"V{r}")), Weight=num(v(f"X{r}")) if v(f"X{r}") else None,
                  Grade=v(f"Y{r}"), Description=f"Thread: {v(f'Z{r}')}" if v(f"Z{r}") else None,
                  _srcs=dict(OD=f"V{r}", Weight=f"X{r}", Grade=f"Y{r}", Description=f"Z{r}"))
        if r in (9, 10):
            kw["Description"] = "; ".join(x for x in (kw["Description"], f"Depth stated as '{v(f'AB{r}')}' (not given)") if x)
            comp(ctype, name, f"P{r}:Z{r}", actual=label, **kw)
            flags.append(("Components", name, "Depth not given", f"AB{r} = '{v(f'AB{r}')}' — Top_MD/Bottom_MD left blank (required for casing)."))
        elif r == 11:
            comp(ctype, name, f"P{r}:Z{r}", bot_ref="K86", top_m=0, actual=label, _top_src="assumed 0 (hung at wellhead)", **kw)
            flags.append(("Components", name, "Assumed top", "Top_MD = 0 (casing hung at wellhead) — the source gives only the shoe depth (K86 '9-5/8\" CSG SHOE' 2572; AB11 '2572 m MDDF' agrees)."))
        else:
            comp(ctype, name, f"P{r}:Z{r}", top_ref="K78", actual=label, **kw)
            flags.append(("Components", name, "Header vs diagram conflict", f"AB12 gives '{v('AB12')}', but TD is {v('K190')} m and the liner covers zones below 3098 m; the diagram's '7\" LINER SHOE' (H189) has no MD. Bottom_MD left blank (required for liner) — please confirm the liner shoe depth."))
    # tubing strings
    comp("Short String", v("B25"), "B25", bot_ref="K71", top_m=0, string="Short", actual="TUBING", OD=3.5, Weight=9.2,
         Grade="HP2-13CR-95", _top_src="assumed 0 (from tubing hanger, B25)", _srcs=dict(OD="B25", Weight="B25", Grade="B25"))
    comp("Long String", v("S25"), "S25", bot_ref="S88", top_m=0, string="Long", actual="TUBING", OD=3.5, Weight=9.2,
         Grade="HP2-13CR-95", _top_src="assumed 0 (from tubing hanger, S25)", _srcs=dict(OD="S25", Weight="S25", Grade="S25"))
    comp("Long String", v("U90"), "U90", top_ref="S88", bot_ref="S177", string="Long", actual="TUBING",
         OD=v("AB90"), ID=v("Z90"), Weight=num(v("X14")), Grade="13CR-L80", _srcs=dict(OD="AB90", ID="Z90", Weight="X14", Grade="U90"))
    flags.append(("Components", "tubing strings", "Assumed top", "Short and long string Top_MD = 0 ('From Tubing Hanger'); the tubing hanger depth is not given."))
    flags.append(("Components", "tubing strings", "Header vs diagram conflict",
                  f"Header AB13 '{v('AB13')}' (3-1/2\") and AB14 '{v('AB14')}' (2-7/8\") differ from the equipment tally; diagram used: 3-1/2\" long string 0 to crossover S88 {v('S88')}; 2-7/8\" from S88 to long EOT S177 {v('S177')}; short string 0 to EOT K71 {v('K71')}."))
    flags.append(("Components", "tubing strings", "Grade outside dropdown", "Tubing grades 'HP2-13CR-95' / '13CR-L80' and liner 'L-80 13Cr' are written as reported; they are not in the Grade dropdown (H-40 … Q-125)."))
    # equipment
    for side, items in (("Short", SHORT), ("Long", LONG)):
        cn, ct, co, ci = ("F", "K", "B", "D") if side == "Short" else ("U", "S", "AB", "Z")
        for r, ctype, actual, ex in items:
            ex = dict(ex)
            top_ref, bot_ref = ex.pop("_top", f"{ct}{r}"), ex.pop("_bottom", None)
            string = ex.pop("_string", side)
            srcs = dict(OD=f"{co}{r}", ID=f"{ci}{r}")
            if "Description" in ex:
                srcs["Description"] = ex["Description"]
                ex["Description"] = v(ex["Description"])
            comp(ctype, str(v(f"{cn}{r}")).strip(), f"{cn}{r}", top_ref=top_ref, bot_ref=bot_ref, string=string, actual=actual,
                 OD=v(f"{co}{r}"), ID=v(f"{ci}{r}"), _srcs=srcs, **ex)
            if actual not in EXACT and not actual.startswith(("SPM (GLV", "LANDING NIPPLE", "TRMAXX")) and "PACKER" not in actual:
                flags.append(("Components", f"{cn}{r}", "Nearest Component_Type", f"'{actual}' has no exact type → '{ctype}'; real type kept in Actual_Equipment_Type."))
    flags.append(("Components", "TRSV rows", "Type from context", "TRMAXX-5E typed TRSV: the string label reads 'From Tubing Hanger to TRSV' (B25/S25)."))
    flags.append(("Components", "SPM rows", "Has_Valve", "Has_Valve = Yes only where a live GLV (PPO/TRO) is stated; 'GLV: Dummy' and AUTO SPM DIAL left blank."))
    flags.append(("Components", "F61/U61", "One row for dual packer", "K61 and S61 both 2481.219 for the same dual-string packer — one component row."))
    pbtd = comp("PBTD", "PBTD", "I187", top_ref="K187", actual="PBTD")
    flags.append(("Components", "PBTD", "Header vs diagram conflict", f"Header G11 '{v('G11')}' vs diagram K187 {v('K187')} — diagram used."))
    # June 2022 annotations
    bp = re.search(r"Bridge Plug @\s*(\d+(?:\.\d+)?)m MDDF \((\w+ \d{4})\)", drawing)
    toc = re.search(r"TOC @\s*(\d+(?:\.\d+)?)m MDDF \((\w+ \d{4})\)", drawing)
    leak = re.search(r"(Leak at Seal Assembly\s+@ (\d+)m MDDF \(June\s+2022\) from WLD\. Not isolated)", drawing)
    bpm = float(bp.group(1))
    comp("Bridge Plug", f"Bridge Plug @ {bp.group(1)}m MDDF ({bp.group(2)})", "drawing text", top_m=bpm, actual="BRIDGE PLUG",
         Description=f"Post-completion drawing annotation, {bp.group(2)}", _top_src="drawing annotation")
    msa = next(c for c in comps if c["_src"]["Name"] == "U119")
    msa["Description"] = re.sub(r"\s+", " ", leak.group(1))
    flags.append(("Components", "U119", "Annotation placement", f"'{msa['Description']}' attached to the MSA seal assembly at {v('S119')} m (nearest to 3108 m; the other seal assembly is at {v('S109')} m)."))
    flags.append(("Components", "Bridge Plug", "Date precision", "Set_Date left blank: annotation gives month only (June 2022) — kept in Description."))
    plug = [{"Name": f"Cement on bridge plug - TOC @ {toc.group(1)}m MDDF ({toc.group(2)})", "Top_MD": ft(float(toc.group(1))),
             "Bottom_MD": ft(bpm), "_extra": [float(toc.group(1)), bpm], "_src": dict(Name="drawing text", Top_MD="drawing 'TOC @'", Bottom_MD="drawing 'Bridge Plug @'")}]
    flags.append(("Plug_Back_Cement", plug[0]["Name"], "Bottom from bridge plug", f"Bottom_MD = bridge plug depth {bpm} m (cement sits on the plug)."))

    # zones -> Formations / Formation_Tops
    forms, tops = [], []
    for r in (101, 133, 155, 177):
        name = v(f"B{r}").strip()
        tv = re.match(r"\s*([\d.]+)\s*m\s*-\s*([\d.]+)\s*m TVD", v(f"B{r + 1}"))
        md = re.match(r"\s*([\d.]+)\s*m\s*-\s*([\d.]+)\s*m MDDF", v(f"B{r + 2}"))
        t1, t2, m1 = float(tv.group(1)), float(tv.group(2)), float(md.group(1))
        forms.append({"Name": name, "Top_TVD": ft(t1), "Bottom_TVD": ft(t2), "_extra": [t1, t2],
                      "_src": dict(Name=f"B{r}", Top_TVD=f"B{r + 1}", Bottom_TVD=f"B{r + 1}")})
        tops.append({"Name": name, "MD": ft(m1), "TVD": ft(t1), "_extra": [m1, t1],
                     "_src": dict(Name=f"B{r}", MD=f"B{r + 2}", TVD=f"B{r + 1}")})
    flags.append(("Formations / Formation_Tops", "zones", "Not stated", "Lithology, Color, Symbol and Is_Reservoir not given — blank. Zone bottom MD (e.g. 3105 m) has no Formation_Tops column; kept in review."))

    # Well_Info
    td, kb = v("K190"), num(v("G14"))
    info = {"Well_Name": (v("B2").split(":", 1)[1].strip(), None, "B2"),
            "Total_Depth_MD": (ft(td), td, "K190 (diagram TD)"),
            "KB_Elevation": (ft(kb), kb, "G14 'DRILL FLOOR HEIGHT 33 m-AMSL'"),
            "Completion_Date": (v("G8").strftime("%Y-%m-%d"), None, "G8")}
    flags.append(("Well_Info", "Total_Depth_MD", "Header vs diagram conflict", f"Header G10 '{v('G10')}' vs diagram K190 {td} — diagram used."))
    flags.append(("Well_Info", "KB_Elevation", "Interpretation", "Drill floor height 33 m AMSL used as KB elevation (depths are from RKB)."))
    flags.append(("Well_Info", "UWI, Field, Operator, County, State, Status, Ground_Elevation, Latitude, Longitude, Spud_Date, Trajectory_Type",
                  "Not stated", f"Left blank. Location '{v('B3')}', water depth '{v('G13')}', max deviation '{v('G12')}' have no Well_Info row (see No Destination)."))
    flags.append(("(all)", "Well name", "Name differs from file name", f"Well name as written in B2 ('{info['Well_Name'][0]}'); file name says 'A04ST1'."))

    nodest = [
        ("B3", v("B3"), "Location / slot — no Well_Info row"), ("G9", v("G9"), "Rig — no destination"),
        ("G10", v("G10"), "Header TD — superseded by diagram TD (K190)"), ("G11", v("G11"), "Header PBTD — superseded by K187"),
        ("G12", v("G12"), "Maximum deviation — no destination"), ("G13", v("G13"), "Water depth — no destination (Ground_Elevation not derived)"),
        ("V16", v("V16"), "Completion fluid — no destination"), ("V17+V18", f"{v('V17')}{v('V18')}", "Packer fluid — no destination"),
        ("B18", v("B18"), "Depth reference note (applied: depths are top of equipment from RKB)"),
        ("U62", v("U62"), "Pup below dual packer — no depth given"), ("F31 / U30", v("F31"), "Tubing joints between equipment — covered by string rows"),
        ("J40,J47,J54,J103,J135,J157,J179,J189", ", ".join(str(v(c)) for c in ("J40", "J47", "J54", "J103", "J135", "J157", "J179", "J189")),
         "Deviation (deg) — Surveys needs MD + azimuth; azimuth not given"),
        ("B105,B137,B159,B181", ", ".join(v(c) for c in ("B105", "B137", "B159", "B181")), "Zone pressures (E34A, E40A, E47A, E51) — no destination"),
        ("B104,B136,B158,B180", "pressure-gradient formulas", "Calculations — no destination"),
        ("B103,B135,B157,B179", ", ".join(v(c) for c in ("B103", "B135", "B157", "B179")), "Zone bottom MD — no Formation_Tops column"),
        ("B118", v("B118"), "Zone named without depth — not mapped"),
        ("B123,B144,B167", ", ".join(v(c).strip() for c in ("B123", "B144", "B167")), "Zone spacing notes — no destination"),
        ("AE11,AD13:AE14", "Total jts required (formulas)", "Joint-count calculation — no destination"),
        ("U191:Z193", f"{v('U191')} / {v('V192').strip()} / {v('U193')} / {v('Z193')}", "Sign-off block — no destination"),
        ("drawing", "TOC in Annulus unknown (June 2022)", "Annotation — no destination"),
        ("images", "2 pictures (logo / symbol)", "Images — no destination"),
    ]
    return info, comps, plug, forms, tops, flags, nodest


# ---------------------------------------------------------------- XML writing
def cell(ref, v, style=None):
    if v is None or v == "":
        return ""
    s = f' s="{style}"' if style else ""
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return f'<c r="{ref}"{s}><v>{repr(float(v)) if v != int(v) else int(v)}</v></c>'
    return f'<c r="{ref}"{s} t="inlineStr"><is><t xml:space="preserve">{escape(str(v))}</t></is></c>'


def rewrite(xml, ncols, extra_hdr, rows, widths, keep_rows=False):
    """Append extra header cells to row 1, replace rows >= 2 with `rows` (lists of values) or, if keep_rows,
    append cells to the existing rows. Updates spans, dimension and <cols>."""
    total = ncols + len(extra_hdr)
    m = re.search(r"<sheetData>(.*)</sheetData>", xml, re.S)
    body = m.group(1)
    rowx = re.findall(r'<row r="(\d+)"[^>]*>.*?</row>|<row r="(\d+)"[^>]*/>', body, re.S)
    rows_xml = re.findall(r'(<row r="(\d+)"[^>]*?)(/>|>(.*?)</row>)', body, re.S)
    out = []
    for open_, rn, _, inner in rows_xml:
        rn = int(rn)
        open_ = re.sub(r'spans="\d+:\d+"', f'spans="1:{total}"', open_)
        if rn == 1:
            hdr = "".join(cell(f"{get_column_letter(ncols + i + 1)}1", h, 4) for i, h in enumerate(extra_hdr))
            out.append(f"{open_}>{inner}{hdr}</row>")
        elif keep_rows:
            add = "".join(cell(f"{get_column_letter(k + 1)}{rn}", x) for k, x in enumerate(rows.get(rn, [])) if x is not None)
            out.append(f"{open_}>{inner or ''}{add}</row>")
    if not keep_rows:
        for i, r in enumerate(rows, 2):
            out.append(f'<row r="{i}" spans="1:{total}">' + "".join(cell(f"{get_column_letter(k + 1)}{i}", x) for k, x in enumerate(r)) + "</row>")
    xml = xml[:m.start(1)] + "".join(out) + xml[m.end(1):]
    last = max(len(rows) + 1, 1) if not keep_rows else max(int(r[1]) for r in rows_xml)
    xml = re.sub(r'<dimension ref="[^"]*"/>', f'<dimension ref="A1:{get_column_letter(total)}{last}"/>', xml, count=1)
    cols = "".join(f'<col min="{ncols + i + 1}" max="{ncols + i + 1}" width="{w}" customWidth="1"/>' for i, w in enumerate(widths))
    return xml.replace("</cols>", cols + "</cols>", 1)


def main():
    OUTDIR.mkdir(exist_ok=True)
    h_src, h_tpl = sha(SRC), sha(TPL)
    info, comps, plug, forms, tops, flags, nodest = extract()
    z = zipfile.ZipFile(TPL)
    parts = {"Well_Info": "xl/worksheets/sheet2.xml", "Components": "xl/worksheets/sheet3.xml",
             "Plug_Back_Cement": "xl/worksheets/sheet7.xml", "Formations": "xl/worksheets/sheet9.xml",
             "Formation_Tops": "xl/worksheets/sheet10.xml"}
    sst = ["".join(re.findall(r"<t[^>]*>(.*?)</t>", si, re.S)) for si in re.findall(r"<si>(.*?)</si>", z.read("xl/sharedStrings.xml").decode(), re.S)]
    wi_xml = z.read(parts["Well_Info"]).decode()
    prop_row = {sst[int(v)]: int(r) for r, v in re.findall(r'<c r="A(\d+)"[^>]*t="s"[^>]*><v>(\d+)</v>', wi_xml)}
    new = {
        "Well_Info": rewrite(wi_xml, 2, ["Value (m)"], {prop_row[k]: [None, val, m] for k, (val, m, _) in info.items()}, [12], keep_rows=True),
        "Components": rewrite(z.read(parts["Components"]).decode(), 29, COMP_EXTRA,
                              [[c[k] for k in COMP_COLS] + c["_extra"] for c in comps], [40, 22, 16, 18]),
        "Plug_Back_Cement": rewrite(z.read(parts["Plug_Back_Cement"]).decode(), 3, ["Top_MD (m MDDF)", "Bottom_MD (m MDDF)"],
                                    [[p["Name"], p["Top_MD"], p["Bottom_MD"]] + p["_extra"] for p in plug], [16, 18]),
        "Formations": rewrite(z.read(parts["Formations"]).decode(), 5, ["Top_TVD (m)", "Bottom_TVD (m)"],
                              [[f["Name"], f["Top_TVD"], f["Bottom_TVD"], None, None] + f["_extra"] for f in forms], [13, 14]),
        "Formation_Tops": rewrite(z.read(parts["Formation_Tops"]).decode(), 6, ["MD (m MDDF)", "TVD (m)"],
                                  [[t["Name"], t["MD"], t["TVD"], None, None, None] + t["_extra"] for t in tops], [14, 10]),
    }
    by_part = {p: n for n, p in parts.items()}
    with zipfile.ZipFile(OUT, "w") as zo:
        for it in z.infolist():
            b = z.read(it.filename)
            if it.filename in by_part:
                b = new[by_part[it.filename]].encode("utf-8")
            zo.writestr(it, b)

    # traceability
    trace = []
    for k, (val, m, src) in info.items():
        trace.append(("Well_Info", prop_row[k], "Value", val, src, f"{m} m × {FT} = {val} ft" if m is not None else "as reported"))
    for i, c in enumerate(comps, 2):
        for k in COMP_COLS:
            if c[k] not in (None, "") and k != "Component_ID":
                src = c["_src"].get(k) or ("type mapping (Actual_Equipment_Type)" if k == "Component_Type" else
                                          "source text" if k in ("Model", "Sleeve_State", "Has_Valve", "Grade") else "")
                conv = f"{c['_extra'][2 if k == 'Top_MD' else 3]} m × {FT}" if k in ("Top_MD", "Bottom_MD") else "as reported"
                trace.append(("Components", i, k, c[k], src, conv))
        trace.append(("Components", i, "Component_ID", c["Component_ID"], "generated key (template requires a unique ID)", "generated"))
    for sheet, rows, cols in (("Plug_Back_Cement", plug, ["Name", "Top_MD", "Bottom_MD"]), ("Formations", forms, ["Name", "Top_TVD", "Bottom_TVD"]),
                              ("Formation_Tops", tops, ["Name", "MD", "TVD"])):
        for i, r in enumerate(rows, 2):
            for j, k in enumerate(cols):
                trace.append((sheet, i, k, r[k], r["_src"][k], "as reported" if j == 0 else f"{r['_extra'][j - 1]} m × {FT}"))

    validation = validate(parts, info, comps, plug, forms, tops, h_src, h_tpl, prop_row)
    write_review(trace, flags, nodest, comps, validation, h_src, h_tpl, len(plug), len(forms), len(tops))
    print("Components", len(comps), "Plug_Back_Cement", len(plug), "Formations", len(forms), "Formation_Tops", len(tops))
    for c, r in validation:
        print(r, "-", c)
    print(OUT.name, REV.name)


def validate(parts, info, comps, plug, forms, tops, h_src, h_tpl, prop_row):
    v = []
    zt, zo = zipfile.ZipFile(TPL), zipfile.ZipFile(OUT)
    changed = {n for n in zt.namelist() if zt.read(n) != zo.read(n)}
    v.append(("Only the 5 target sheet XML parts changed (README, Component_Types, styles, header notes, other sheets byte-identical)",
              "PASS" if changed <= set(parts.values()) else f"FAIL {changed}"))
    x14t = len(re.findall(r"<x14:dataValidation\b", zt.read(parts["Components"]).decode()))
    x14o = len(re.findall(r"<x14:dataValidation\b", zo.read(parts["Components"]).decode()))
    v.append((f"Components dropdowns kept ({x14o} x14 + standard lists)", "PASS" if x14t == x14o else "FAIL"))
    wb = load_workbook(OUT)
    ws = wb["Components"]
    hdr = [c.value for c in ws[1]]
    v.append(("Components: template columns A:AC unchanged, extra columns after AC", "PASS" if hdr[:29] == COMP_COLS and hdr[29:] == COMP_EXTRA else "FAIL"))
    v.append((f"Components: {len(comps)} rows, template example rows removed", "PASS" if ws.max_row == len(comps) + 1 else f"FAIL {ws.max_row}"))
    types = {r[0] for r in wb["Component_Types"].iter_rows(min_row=2, values_only=True)}
    v.append(("Every Component_Type is in the Component_Types list", "PASS" if all(r[1] in types for r in ws.iter_rows(min_row=2, values_only=True)) else "FAIL"))
    ids = [r[0] for r in ws.iter_rows(min_row=2, values_only=True)]
    v.append(("Component_IDs unique", "PASS" if len(ids) == len(set(ids)) else "FAIL"))
    ok = True
    for r, c in zip(ws.iter_rows(min_row=2, values_only=True), comps):
        ok &= [x if x != "" else None for x in r[:33]] == [x if x != "" else None for x in [c[k] for k in COMP_COLS] + c["_extra"]]
        for i_ft, i_m in ((3, 31), (4, 32)):
            ok &= (r[i_ft] is None and r[i_m] is None) or abs(r[i_ft] - round(r[i_m] * FT, 3)) < 1e-9
    v.append(("Components: all written values equal extracted values; every ft depth = m × 3.28084", "PASS" if ok else "FAIL"))
    src = load_workbook(SRC).worksheets[0]
    ok = all(src[c["_src"]["Name"]].value is None or str(src[c["_src"]["Name"]].value).strip() == c["Name"]
             for c in comps if re.match(r"^[A-Z]+\d+$", c["_src"]["Name"] or "") and c["Component_Type"] not in ("Casing", "Liner"))
    v.append(("Equipment names identical to source text", "PASS" if ok else "FAIL"))
    wi = wb["Well_Info"]
    v.append(("Well_Info: property labels unchanged; values written", "PASS" if all(wi.cell(prop_row[k], 2).value == val for k, (val, m, s) in info.items()) else "FAIL"))
    for name, rows, n in (("Plug_Back_Cement", plug, 3), ("Formations", forms, 5), ("Formation_Tops", tops, 6)):
        v.append((f"{name}: {len(rows)} rows written", "PASS" if wb[name].max_row == len(rows) + 1 else "FAIL"))
    empty = [n for n in ("Cement", "Centralizers", "Float_Equipment", "Surveys", "Fluid_Contacts") if wb[n].max_row > 1]
    v.append(("Sheets with no source data left empty (Cement, Centralizers, Float_Equipment, Surveys, Fluid_Contacts)", "PASS" if not empty else f"FAIL {empty}"))
    v.append(("Source file unchanged", "PASS" if sha(SRC) == h_src else "FAIL"))
    v.append(("Template unchanged", "PASS" if sha(TPL) == h_tpl else "FAIL"))
    return v


def write_review(trace, flags, nodest, comps, validation, h_src, h_tpl, n_plug, n_forms, n_tops):
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    for i, (k, x) in enumerate([
        ("Populated workbook", str(OUT)), ("Source (read-only)", f"{SRC}  sha256 {h_src[:16]}…"),
        ("Template (read-only)", f"{TPL}  sha256 {h_tpl[:16]}…"),
        ("Rows", f"Components {len(comps)}, Plug_Back_Cement {n_plug}, Formations {n_forms}, Formation_Tops {n_tops}"),
        ("Unit conversion", f"All depths m MDDF / m TVD → ft (× {FT}, 3 dp). Original meters in the extra '(m …)' columns. OD/ID already inches; weight PPF = lb/ft."),
        ("Extra columns", "Components AD:AG (Actual_Equipment_Type, String, Top/Bottom m); Well_Info C; Plug_Back_Cement D:E; Formations F:G; Formation_Tops G:H"),
        ("Decisions", "Diagram wins over header; nearest Component_Type with real type kept; June 2022 annotations included"),
        ("Processed", dt.datetime.now().strftime("%Y-%m-%d %H:%M") + " — locally")], 1):
        ws.cell(i, 1, k).font = Font(bold=True)
        ws.cell(i, 2, x)
    ws.column_dimensions["A"].width, ws.column_dimensions["B"].width = 22, 140
    for title, hdr, data, w in (
        ("Traceability", ["Target Sheet", "Row", "Column", "Value Written", "Source Cell / Origin", "Conversion"], trace, [18, 6, 18, 60, 44, 26]),
        ("Type Mapping", ["Row", "Name (source text)", "Actual_Equipment_Type", "Component_Type used", "String", "Top m", "Top ft"],
         [(i, c["Name"], c["_extra"][0], c["Component_Type"], c["_extra"][1], c["_extra"][2], c["Top_MD"]) for i, c in enumerate(comps, 2)],
         [6, 70, 44, 18, 22, 10, 10]),
        ("Flags & Conflicts", ["Sheet", "Item", "Issue", "Detail"], flags, [22, 40, 26, 120]),
        ("No Destination", ["Source Cell", "Source Content", "Status"], nodest, [26, 70, 70]),
        ("Validation", ["Check", "Result"], validation, [110, 12])):
        s = wb.create_sheet(title)
        s.append(hdr)
        for r in data:
            s.append(list(r))
        for c in s[1]:
            c.font, c.fill = Font(bold=True), PatternFill("solid", fgColor="DDEBF7")
        for i, x in enumerate(w, 1):
            s.column_dimensions[get_column_letter(i)].width = x
        for row in s.iter_rows(min_row=2):
            for c in row:
                c.alignment = Alignment(wrap_text=True, vertical="top")
        s.freeze_panes = "A2"
    wb.save(REV)


if __name__ == "__main__":
    main()
