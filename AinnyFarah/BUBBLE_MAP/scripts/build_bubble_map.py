#!/usr/bin/env python3
"""Map the Dulang well header listing + well marker file (NAS, read-only) into the bubble map template (read-only).

Outputs (BUBBLE_MAP/output_bubble-map/):
  bubble_map_mapped.xlsx           template sheet + one row per wellbore (202 header wellbores, then marker-only wellbores)
  bubble_map_mapping_and_QC.xlsx   Mapping Log / QC Issues / Summary / Marker Name Matching

Decisions confirmed by the user (27 Oct 2026):
  * rows = header wellbores + marker-only standard wellbores
  * WELL = DULG-<platform><3-digit no.><suffix exactly as in the source name> (S1 / ST1 / R1 / s2 not standardised; user, 28 Oct 2026)
  * exploration wells (DULANG-1..8, 6G-...): PLATFORM and WELL blank
  * X UTM / Y UTM only as given in the source (none at well level -> blank; marker X/Y are subsurface pick positions)
  * isDummyCoord, STRING TYPE, RESERVOIR, both DATE columns and production columns blank (not in the sources)
"""
import hashlib
import os
import re
from collections import Counter, defaultdict
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

HERE = Path(__file__).resolve().parent.parent
NAS = Path("/Volumes/homes/public/DulangOSS/20260929")
HDR_FILE = NAS / "Well Header Info & Data Availability Listing.xlsx"
MKR_FILE = NAS / "Well_Marker_15Jan2026.xlsx"
TEMPLATE = HERE / "bubble_map_template.xlsx"
OUT = HERE / "output_bubble-map"
HDR_SHEET, MKR_SHEET = "Availability Listing_USER", "Sheet1"

HDR_RE = re.compile(r"DULANG-([ABCD])(\d+)(?:R(\d+))?(?:(ST|S)(\d+))?")
MKR_RE = re.compile(r"([ABCD])(\d+)(?:s(\d+))?")
ALIAS_RE = re.compile(r"X\d+(?:_\d+)?(?:s\d+)?\s*=\s*DL[- ](.+)")


def sha(p):
    h = hashlib.sha256()
    fd = os.open(p, os.O_RDONLY)
    while (b := os.read(fd, 1 << 20)):
        h.update(b)
    os.close(fd)
    return h.hexdigest()


def std_well(letter, num, suffix):
    """Platform + 3-digit well number; any suffix (S1, ST1, R1S1, s2 ...) is kept exactly as written in the source."""
    return f"DULG-{letter}{int(num):03d}{suffix}"


def key(letter, num, redrill, side):
    """Wellbore identity used to match the two files (S# and ST# are the same sidetrack)."""
    return (letter, int(num), int(redrill or 0), int(side or 0))


def marker_category(name):
    if MKR_RE.fullmatch(name):
        return "standard"
    if ALIAS_RE.fullmatch(name):
        return "exploration alias"
    if re.fullmatch(r"X\d+(?:_\d+)?", name):
        return "exploration code without alias"
    if re.fullmatch(r"Well \d+", name):
        return "generic 'Well N' name"
    return "planned / old / variant trajectory"


def main():
    before = {p: sha(p) for p in (HDR_FILE, MKR_FILE, TEMPLATE)}

    # ---------------- header listing
    hws = load_workbook(HDR_FILE, data_only=True)[HDR_SHEET]
    hcols = [c.value for c in hws[2]]
    assert hcols[1:23] == ["Unique Wellbore Identifier", "Wellbore Name", "Country Name", "Region", "Original Operator", "Wellbore Spud Date",
                           "Spud Year", "Water Depth", "Water Depth Unit", "Total Depth", "Total Depth Unit", "Reference Point", "Elevation",
                           "Elevation Unit", "Reference Datum", "Latitude WGS84", "Longitude WGS84", "Well Header Info", "Deviation",
                           "Checkshot", "Markers", "REMARKS"], hcols
    H = {n: i for i, n in enumerate(hcols) if n}
    hrows = [(r_i, r) for r_i, r in enumerate(hws.iter_rows(min_row=3, values_only=True), start=3) if r[H["Wellbore Name"]]]
    assert len(hrows) == 202 and len({r[H["Wellbore Name"]] for _, r in hrows}) == 202

    # ---------------- marker file
    mws = load_workbook(MKR_FILE, data_only=True)[MKR_SHEET]
    mcols = [c.value for c in mws[1]]
    wi, si = mcols.index("Well"), mcols.index("Surface")
    mcount, msurf = Counter(), defaultdict(set)
    for r in mws.iter_rows(min_row=2, values_only=True):
        mcount[r[wi]] += 1
        msurf[r[wi]].add(r[si])
    n_picks = sum(mcount.values())

    # ---------------- template
    tpl = load_workbook(TEMPLATE)
    ws = tpl.worksheets[0]
    tcols = [c.value for c in ws[1]]
    assert tcols == ["ORIGINAL WELL NAME", "FIELD", "PLATFORM", "WELL", "STRING TYPE", "RESERVOIR", "DATE", "LAT", "LONG", "X UTM", "Y UTM",
                     "isDummyCoord", "DATE", "WATER CUT, WC (%)", "CUMULATIVE OIL PRODUCTION, NP", "GOR", "NPRADIUS",
                     "CUMULATIVE WATER PRODUCTION, WP", "PRODUCTIVITY"], tcols
    assert len(tpl.worksheets) == 1 and ws.max_row <= 2

    qc = []                       # (original name, issue, source reference, reason for review)
    records = []                  # (row values dict, source tag)
    hdr_by_key, hdr_by_name = {}, {}
    for r_i, r in hrows:
        name = r[H["Wellbore Name"]]
        hdr_by_name[name] = r_i
        m = HDR_RE.fullmatch(name)
        ref = f"{HDR_FILE.name} > {HDR_SHEET} > row {r_i}, col C"
        rec = {"ORIGINAL WELL NAME": name, "FIELD": "DULANG"}
        if m:
            letter, num, rd, st, sd = m.groups()
            k = key(letter, num, rd, sd)
            assert k not in hdr_by_key, (name, k)
            hdr_by_key[k] = r_i
            rec["PLATFORM"] = f"DULG-{letter}"
            rec["WELL"] = std_well(letter, num, name[m.end(2):])
        else:
            qc.append((name, "PLATFORM and WELL blank", ref, "exploration/appraisal well name without a platform letter — no platform can be assigned"))
        lat, lon = r[H["Latitude WGS84"]], r[H["Longitude WGS84"]]
        if lat is not None and lon is not None:
            rec["LAT"], rec["LONG"] = lat, lon
        else:
            qc.append((name, "LAT / LONG missing", f"{HDR_FILE.name} > {HDR_SHEET} > row {r_i}, cols Q-R",
                       "no coordinates in the header listing" + (f"; REMARKS: {r[H['REMARKS']]}" if r[H["REMARKS"]] else "")))
        if r[H["REMARKS"]] and lat is not None:
            qc.append((name, "Source remark", f"{HDR_FILE.name} > {HDR_SHEET} > row {r_i}, col W", str(r[H["REMARKS"]])))
        records.append((rec, "header", r_i))

    # platform location check: header coordinates vs the platform's median surface location
    cent = {}
    for p in "ABCD":
        pts = [(rec["LAT"], rec["LONG"]) for rec, _, _ in records if rec.get("PLATFORM") == f"DULG-{p}" and "LAT" in rec]
        pts_lat, pts_lon = sorted(a for a, _ in pts), sorted(b for _, b in pts)
        cent[p] = (pts_lat[len(pts) // 2], pts_lon[len(pts) // 2])
    for rec, _, r_i in records:
        if "LAT" in rec and rec.get("PLATFORM"):
            near = min(cent, key=lambda p: (rec["LAT"] - cent[p][0]) ** 2 + (rec["LONG"] - cent[p][1]) ** 2)
            if f"DULG-{near}" != rec["PLATFORM"]:
                qc.append((rec["ORIGINAL WELL NAME"], "Coordinates conflict with platform",
                           f"{HDR_FILE.name} > {HDR_SHEET} > row {r_i}, cols Q-R",
                           f"LAT/LONG {rec['LAT']}, {rec['LONG']} lie at platform {near} (median {cent[near][0]:.5f}, {cent[near][1]:.5f}), "
                           f"not at {rec['PLATFORM']}; mapped as given — review"))

    # ---------------- marker names -> header wellbores
    match_rows, marker_only = [], []
    matched_hdr = set()
    for name in sorted(mcount):
        cat = marker_category(name)
        m = MKR_RE.fullmatch(name)
        hit = None
        if m:
            k = key(m.group(1), m.group(2), None, m.group(3))
            if k in hdr_by_key:
                hit = hdr_by_key[k]
            else:
                marker_only.append((name, m))
        elif cat == "exploration alias":
            alias = "DULANG-" + ALIAS_RE.fullmatch(name).group(1).strip().replace(" ", "-")
            hit = hdr_by_name.get(alias)
        hdr_name = None
        if hit:
            hdr_name = next(r[H["Wellbore Name"]] for r_i, r in hrows if r_i == hit)
            matched_hdr.add(hit)
        status = "matched" if hit else ("marker-only wellbore (output row added)" if m else "not matched (QC)")
        match_rows.append((name, cat, mcount[name], hdr_name or "", status))
        if not hit and not m:
            reason = {"exploration alias": "alias names a wellbore not in the header listing",
                      "exploration code without alias": "exploration code with no 'DL-…' alias — wellbore cannot be identified",
                      "generic 'Well N' name": "generic name — wellbore cannot be identified",
                      "planned / old / variant trajectory": "planned well path, old version or variant name — not a separate drilled wellbore in the header listing"}[cat]
            qc.append((name, "Marker well name not matched", f"{MKR_FILE.name} > {MKR_SHEET} > col K 'Well' ({mcount[name]} picks)", reason))

    for name, m in marker_only:
        letter, num, side = m.group(1), m.group(2), m.group(3)
        rec = {"ORIGINAL WELL NAME": name, "FIELD": "DULANG", "PLATFORM": f"DULG-{letter}", "WELL": std_well(letter, num, name[m.end(2):])}
        records.append((rec, "marker", None))
        qc.append((name, "Wellbore only in the marker file", f"{MKR_FILE.name} > {MKR_SHEET} > col K 'Well' ({mcount[name]} picks)",
                   "no header record: LAT/LONG missing; ORIGINAL WELL NAME is the marker file's short name"))

    # header 'Markers' availability flag vs the marker file
    for r_i, r in hrows:
        flag = r[H["Markers"]]
        has = r_i in matched_hdr
        if (flag == "YES") != has:
            qc.append((r[H["Wellbore Name"]], "Marker availability flag differs from the marker file",
                       f"{HDR_FILE.name} > {HDR_SHEET} > row {r_i}, col V",
                       f"header 'Markers' = {flag!r}, but the marker file {'has' if has else 'has no'} picks for this wellbore"))

    qc.append(("(all records)", "STRING TYPE blank", "both input files", "no string type (SS/LS/TS) in either source"))
    qc.append(("(all records)", "RESERVOIR blank", f"{MKR_FILE.name} > col J 'Surface'",
               "markers are tops penetrated by each wellbore, not completed/producing reservoirs — not mapped (user confirmed)"))
    qc.append(("(all records)", "X UTM / Y UTM blank", f"{HDR_FILE.name}; {MKR_FILE.name} > cols A-B",
               "no well-level UTM coordinates in the sources; marker X/Y are subsurface pick positions (user: values only as given in the source)"))
    qc.append(("(all records)", "DATE, production columns, isDummyCoord blank", "both input files", "not available in the sources / user instruction"))

    # duplicates
    wells = Counter(rec.get("WELL") for rec, _, _ in records if rec.get("WELL"))
    assert all(v == 1 for v in wells.values()), [k for k, v in wells.items() if v > 1]
    assert len({rec["ORIGINAL WELL NAME"] for rec, _, _ in records}) == len(records)

    # ---------------- write mapped workbook (template copy)
    OUT.mkdir(exist_ok=True)
    out1, out2 = OUT / "bubble_map_mapped.xlsx", OUT / "bubble_map_mapping_and_QC.xlsx"
    for p in (out1, out2):
        if p.exists():
            raise SystemExit(f"{p} exists — not overwriting")
    g_fmt = ws["G2"].number_format
    for i, (rec, _, _) in enumerate(records, start=2):
        for j, col in enumerate(tcols, start=1):
            c = ws.cell(i, j)
            c.value = rec.get(col) if j not in (7, 13) else None
            if j == 7:
                c.number_format = g_fmt
    last = get_column_letter(len(tcols))
    ws.auto_filter.ref = f"A1:{last}{len(records) + 1}"
    ws.freeze_panes = "A2"
    for col, w in (("A", 18), ("D", 16), ("H", 13), ("I", 13)):
        ws.column_dimensions[col].width = max(ws.column_dimensions[col].width or 0, w)
    tpl.save(out1)

    # ---------------- mapping / QC workbook
    hf, mf = HDR_FILE.name, MKR_FILE.name
    mapping = [
        ("ORIGINAL WELL NAME", hf, HDR_SHEET, "C 'Wellbore Name' (rows 3-204)", "exact text, unchanged"),
        ("ORIGINAL WELL NAME", mf, MKR_SHEET, "K 'Well'", "marker-only wellbores only: exact short name as in the marker file"),
        ("FIELD", "-", "-", "-", "constant 'DULANG'"),
        ("PLATFORM", hf + " / " + mf, f"{HDR_SHEET} / {MKR_SHEET}", "C 'Wellbore Name' / K 'Well'",
         "platform letter in the name: DULANG-<A-D>… / <A-D>… -> DULG-<A-D>; exploration names (no letter) blank. Checked against LAT/LONG platform locations"),
        ("WELL", hf + " / " + mf, f"{HDR_SHEET} / {MKR_SHEET}", "C 'Wellbore Name' / K 'Well'",
         "DULG-<platform><well no. 3 digits> + suffix exactly as in the source name (S1, ST1, R1S1, s2 … not standardised — user instruction); exploration wells blank"),
        ("STRING TYPE", "-", "-", "-", "blank — not in sources"),
        ("RESERVOIR", "-", "-", "-", "blank — markers are penetrated tops, not completion reservoirs (user confirmed)"),
        ("DATE (col G)", "-", "-", "-", "blank (user instruction); spud date not used"),
        ("LAT", hf, HDR_SHEET, "Q 'Latitude WGS84'", "as given, decimal degrees WGS84, no conversion"),
        ("LONG", hf, HDR_SHEET, "R 'Longitude WGS84'", "as given, decimal degrees WGS84, no conversion"),
        ("X UTM / Y UTM", "-", "-", "-", "blank — no well-level UTM in the sources; not calculated (user instruction); marker X/Y (cols A-B) are subsurface pick positions"),
        ("isDummyCoord", "-", "-", "-", "blank (user instruction)"),
        ("DATE (col M), WATER CUT, NP, GOR, NPRADIUS, WP, PRODUCTIVITY", "-", "-", "-", "blank — no production data in the sources"),
        ("(match key)", hf + " / " + mf, f"{HDR_SHEET} / {MKR_SHEET}", "C 'Wellbore Name' / K 'Well'",
         "platform + well number + redrill + sidetrack (S#=ST#; 'A01'='DULANG-A1', 'B28s1'='DULANG-B28ST1'); exploration aliases 'Xnn = DL-…' matched by the alias"),
        ("(not mapped)", hf, HDR_SHEET, "B, D-P, S-V", "UWI, country, region, operator, spud date/year, water depth, TD, reference point, elevation, datum, "
         "data-availability flags — no template column (units in source: feet/meters as listed; nothing converted)"),
        ("(not mapped)", mf, MKR_SHEET, "all other columns", "marker picks (X, Y, Z, MD, TVT, TST, Surface, Interpreter, …) used only to match wellbores; no template column"),
    ]
    q = Workbook()
    hdr_fill, bold = PatternFill("solid", fgColor="FF00B0F0"), Font(bold=True)

    def sheet(title, header, rows, widths):
        s = q.create_sheet(title)
        s.append(header)
        for c in s[1]:
            c.fill, c.font, c.alignment = hdr_fill, bold, Alignment(wrap_text=True, vertical="top")
        for r in rows:
            s.append(list(r))
        for i, w in enumerate(widths, start=1):
            s.column_dimensions[get_column_letter(i)].width = w
        for row in s.iter_rows(min_row=2):
            for c in row:
                c.alignment = Alignment(wrap_text=True, vertical="top")
        s.freeze_panes = "A2"
        s.auto_filter.ref = f"A1:{get_column_letter(len(header))}{len(rows) + 1}"
        return s

    q.remove(q.active)
    sheet("Mapping Log", ["Template column", "Source file", "Source worksheet", "Source column / location", "Transformation / unit conversion"],
          mapping, [30, 45, 28, 32, 90])
    sheet("QC Issues", ["Original well name", "Issue", "Source reference", "Reason for review"], qc, [32, 40, 60, 90])

    out_rows = [rec for rec, _, _ in records]
    blank = lambda col: sum(1 for r in out_rows if r.get(col) in (None, ""))
    cats = Counter(c for _, c, _, _, _ in match_rows)
    st = Counter(s for _, _, _, _, s in match_rows)
    summary = [
        ("Output records (bubble_map_mapped.xlsx)", len(records)),
        ("  from header listing", sum(1 for _, s, _ in records if s == "header")),
        ("  marker-only wellbores", sum(1 for _, s, _ in records if s == "marker")),
        ("Header wellbores", len(hrows)),
        ("  matched to marker file", len(matched_hdr)),
        ("  not in marker file", len(hrows) - len(matched_hdr)),
        ("Marker file picks", n_picks),
        ("Marker well names", len(mcount)),
        *[(f"  {k}", v) for k, v in sorted(st.items())],
        *[(f"  category: {k}", v) for k, v in sorted(cats.items())],
        ("Duplicate output wellbores", 0),
        ("Blank PLATFORM", blank("PLATFORM")), ("Blank WELL", blank("WELL")),
        ("WELL with a suffix kept as in source", sum(1 for r in out_rows if re.search(r"\d[A-Za-z]", r.get("WELL") or ""))),
        ("Blank STRING TYPE", blank("STRING TYPE")), ("Blank RESERVOIR", blank("RESERVOIR")),
        ("Blank LAT / LONG", blank("LAT")), ("Blank X UTM / Y UTM", blank("X UTM")),
        ("Coordinates conflicting with platform", sum(1 for x in qc if x[1] == "Coordinates conflict with platform")),
        ("QC issue rows", len(qc)),
    ]
    sheet("Summary", ["Item", "Count"], summary, [55, 12])
    sheet("Marker Name Matching", ["Marker 'Well' name", "Category", "Picks", "Matched header Wellbore Name", "Status"], match_rows, [48, 32, 8, 30, 38])
    q.save(out2)

    # ---------------- verification
    v = load_workbook(out1).worksheets[0]
    vrows = list(v.iter_rows(values_only=True))
    hdr_vals = {r[H["Wellbore Name"]]: r for _, r in hrows}
    checks = [
        ("header row = template", list(vrows[0]) == tcols),
        ("header styles kept", all(v.cell(1, j).fill.fgColor.rgb == ws.cell(1, j).fill.fgColor.rgb for j in range(1, 20))),
        ("record count", len(vrows) - 1 == len(records)),
        ("ORIGINAL WELL NAME exact + LAT/LONG equal source", all(
            (r[0] in hdr_vals and r[7] == hdr_vals[r[0]][H["Latitude WGS84"]] and r[8] == hdr_vals[r[0]][H["Longitude WGS84"]])
            or (r[0] in mcount and r[7] is None) for r in vrows[1:])),
        ("FIELD = DULANG everywhere", all(r[1] == "DULANG" for r in vrows[1:])),
        ("WELL unique", len({r[3] for r in vrows[1:] if r[3]}) == sum(1 for r in vrows[1:] if r[3])),
        ("unsupported columns blank", all(all(r[j] is None for j in (4, 5, 6, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18)) for r in vrows[1:])),
        ("filter + freeze", v.auto_filter.ref == f"A1:S{len(records) + 1}" and v.freeze_panes == "A2"),
        ("inputs and template unchanged (SHA-256)", all(sha(p) == h for p, h in before.items())),
    ]
    for c, ok in checks:
        print("PASS" if ok else "FAIL", "-", c)
    for k, val in summary:
        print(f"{k}: {val}")


if __name__ == "__main__":
    main()
