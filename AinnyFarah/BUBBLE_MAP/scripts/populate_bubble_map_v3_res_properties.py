#!/usr/bin/env python3
"""Create bubble_map_dulang_v3.xlsx as a copy of v2 and fill blank Pay Thickness / Porosity / Water Saturation cells from
RES-flag records of DULANG_RESSUM.xlsx (matched on well name + full reservoir label; exact label first, '.smk' only when
no exact-label RES record exists).

The copy is made at package level: every zip part except the 'Bubble Map Sample' sheet XML is copied byte for byte, and
in that sheet the only edits are (a) new V/W/X cells in matched rows and (b) the cached result of the existing
Hydrocarbon Saturation formula (=IF(X="","",1-X)) in those rows. Formulas, styles, row order and the other sheet are
untouched. Inputs are read only; an existing v3 is never overwritten.
"""
import collections
import csv
import datetime as dt
import hashlib
import math
import re
import sys
import zipfile
from pathlib import Path

import openpyxl

BASE = Path(__file__).resolve().parent.parent
SRC = BASE / "DULANG_RESSUM.xlsx"
TPL = BASE / "bubble_map_dulang_sample" / "bubble_map_dulang_v2.xlsx"
OUT = BASE / "bubble_map_dulang_sample" / "bubble_map_dulang_v3.xlsx"
REP_DIR = BASE / "output_bubble-map"
REP_MD = REP_DIR / "bubble_map_dulang_v3_completion_summary.md"
REP_CSV = REP_DIR / "bubble_map_dulang_v3_match_detail.csv"
FT_PER_M = 3.28084
SENTINEL = -999.25  # Petrel/Techlog null value in RESSUM
HDR = {"V": "Pay Thickness, h (ft)", "W": "Porosity, F (fraction)", "X": "Water Saturation, Sw (fraction)",
       "Y": "Hydrocarbon Saturation, Shc (fraction)"}


def sha(p):
    return hashlib.sha256(Path(p).read_bytes()).hexdigest()


def well_key(name):
    """DULANG-A1 / DULG-A001 -> ('A', 1, ()); suffix tokens kept (ST and S are both 'sidetrack', zero padding ignored).
    Names that do not parse completely return None (never matched)."""
    m = re.fullmatch(r"(?:DULANG|DULG)-([A-D])0*(\d+)((?:(?:ST|S|W|H|P|R)\d+)*)", str(name or "").strip())
    if not m:
        return None
    toks = tuple(("S" if t == "ST" else t, int(d)) for t, d in re.findall(r"(ST|S|W|H|P|R)0*(\d+)", m.group(3)))
    return m.group(1), int(m.group(2)), toks


def num(v):
    """-> (value, problem) ; problem in None / 'missing' / 'null -999.25' / 'unreadable'"""
    if v is None or (isinstance(v, str) and not v.strip()):
        return None, "missing"
    if isinstance(v, bool):
        return None, "unreadable"
    if isinstance(v, str):
        try:
            v = float(v.strip())
        except ValueError:
            return None, "unreadable"
    v = float(v)
    if math.isnan(v) or math.isinf(v):
        return None, "unreadable"
    if v == SENTINEL:
        return None, "null -999.25"
    return v, None


def load_res():
    ws = openpyxl.load_workbook(SRC, read_only=True, data_only=True).worksheets[0]
    it = ws.iter_rows(values_only=True)
    hdr = list(next(it))
    ci = {h: hdr.index(h) for h in ("Well Name", "Zones", "Flag Name", "Gross Interval [m]", "AV_POROSITY [v/v]",
                                    "AV_WATER_SATURATION [v/v]")}
    recs = collections.defaultdict(list)
    names = collections.defaultdict(set)
    n_res = 0
    for i, r in enumerate(it, start=2):
        if r[ci["Flag Name"]] != "RES":
            continue
        n_res += 1
        k = well_key(r[ci["Well Name"]])
        if k is None:
            continue
        names[k].add(r[ci["Well Name"]])
        recs[(k, r[ci["Zones"]])].append(dict(row=i, well=r[ci["Well Name"]], zone=r[ci["Zones"]],
                                              gross=r[ci["Gross Interval [m]"]], por=r[ci["AV_POROSITY [v/v]"]],
                                              sw=r[ci["AV_WATER_SATURATION [v/v]"]]))
    return ws.title, recs, names, n_res


def resolve(recs, prop, lo=None, hi=None):
    """value for one property from the selected record group. -> (value or None, note)"""
    vals, probs = [], []
    for r in recs:
        v, p = num(r[prop])
        if p is None and ((lo is not None and v < lo) or (hi is not None and v > hi)):
            v, p = None, f"invalid ({r[prop]})"
        (vals if p is None else probs).append((r["row"], v if p is None else p))
    distinct = sorted({v for _, v in vals})
    if len(distinct) > 1:
        return None, "CONFLICT " + "; ".join(f"row {row}={v}" for row, v in vals)
    if not distinct:
        return None, "; ".join(f"row {row} {p}" for row, p in probs)
    note = "" if not probs else "other duplicate(s) without value: " + "; ".join(f"row {row} {p}" for row, p in probs)
    return distinct[0], note


def template_rows():
    ws = openpyxl.load_workbook(TPL, read_only=True)["Bubble Map Sample"]
    it = ws.iter_rows(values_only=True)
    hdr = list(next(it))
    assert [hdr[21], hdr[22], hdr[23], hdr[24]] == [HDR[c] for c in "VWXY"], hdr[21:25]
    rows = []
    for n, r in enumerate(it, start=2):
        rows.append(dict(n=n, well=r[2], ru=r[4], V=r[21], W=r[22], X=r[23], Y=r[24]))
    return rows


def blank(v):
    return v is None or v == ""


ROW_RX = re.compile(rb'<row r="(\d+)"[^>]*>(.*?)</row>', re.S)


def edit_sheet(xml, fills):
    """fills: {row: {'V': float, 'W': float, 'X': float}} -> new xml bytes, count of Y caches updated"""
    out, pos, ycount = [], 0, 0
    for m in ROW_RX.finditer(xml):
        n = int(m.group(1))
        if n not in fills:
            continue
        body = m.group(2)
        f = fills[n]
        for c in "VWX":
            assert f'r="{c}{n}"'.encode() not in body, (c, n)
        cells = b"".join(f'<c r="{c}{n}"><v>{repr(f[c])}</v></c>'.encode() for c in "VWX" if c in f)
        ytag = f'<c r="Y{n}"'.encode()
        at = body.find(ytag)
        assert at >= 0, n
        body = body[:at] + cells + body[at:]
        if "X" in f:
            ym = re.compile(rb'<c r="Y%d"( s="\d+")? t="str">(<f[^>]*/>|<f[^>]*>[^<]*</f>)<v/></c>' % n)
            new, k = ym.subn(lambda mm: b'<c r="Y%d"%s>%s<v>%s</v></c>' % (n, mm.group(1) or b"", mm.group(2),
                                                                          repr(1 - f["X"]).encode()), body)
            assert k == 1, n
            body = new
            ycount += 1
        out.append(xml[pos:m.start(2)])
        out.append(body)
        pos = m.end(2)
    out.append(xml[pos:])
    return b"".join(out), ycount


def revert_sheet(xml):
    """undo the edits (for validation): drop V/W/X cells written by this script and restore empty Y caches"""
    xml = re.sub(rb'<c r="[VWX]\d+"><v>[^<]*</v></c>', b"", xml)
    return re.sub(rb'<c r="(Y\d+)"( s="\d+")?>(<f[^>]*/>|<f[^>]*>[^<]*</f>)<v>[^<]*</v></c>',
                  lambda m: b'<c r="%s"%s t="str">%s<v/></c>' % (m.group(1), m.group(2) or b"", m.group(3)), xml)


def main():
    if OUT.exists():
        sys.exit(f"STOP: {OUT} already exists - not overwritten.")
    for p in (REP_MD, REP_CSV):
        if p.exists():
            sys.exit(f"STOP: {p} already exists - not overwritten.")
    h0 = {p: sha(p) for p in (SRC, TPL)}
    src_sheet, res, names, n_res = load_res()
    rows = template_rows()

    # Y must hold the template's own 1-X formula for the cached-value update to be valid
    bad_y = [r["n"] for r in rows if not (isinstance(r["Y"], str) and r["Y"] == f'=IF(X{r["n"]}="","",1-X{r["n"]})')]
    assert not bad_y, bad_y[:10]

    pairs = collections.OrderedDict()
    for r in rows:
        pairs.setdefault((r["well"], r["ru"]), []).append(r)

    detail, fills = [], {}
    stats = collections.Counter()
    for (well, ru), trs in pairs.items():
        k = well_key(well)
        resv = ru.split("-", 1)[1] if ru and "-" in ru else ru
        d = dict(well=well, reservoir_unit=ru, reservoir=resv, template_rows=f"{trs[0]['n']}-{trs[-1]['n']}",
                 n_rows=len(trs), source_well="", match="", source_zone="", source_rows="", pay_ft="", porosity="",
                 sw="", pay_note="", por_note="", sw_note="", status="")
        if k is None or not any(kk == k for kk, _ in res):
            d["match"] = "none"
            d["status"] = "well not in RES records" if k else "well name not parseable"
            stats["unmatched_well"] += 1
            detail.append(d)
            continue
        if (k, resv) in res:
            lvl, zone = "exact", resv
        elif (k, resv + ".smk") in res:
            lvl, zone = ".smk", resv + ".smk"
        else:
            d["match"] = "none"
            d["source_well"] = "/".join(sorted(names[k]))
            d["status"] = "reservoir not in RES records for this well"
            stats["unmatched_res"] += 1
            detail.append(d)
            continue
        recs = res[(k, zone)]
        stats[lvl] += 1
        d.update(match=lvl, source_zone=zone, source_well="/".join(sorted({r["well"] for r in recs})),
                 source_rows=", ".join(str(r["row"]) for r in recs))
        gross, d["pay_note"] = resolve(recs, "gross", lo=0)
        por, d["por_note"] = resolve(recs, "por", 0, 1)
        sw, d["sw_note"] = resolve(recs, "sw", 0, 1)
        f = {}
        if gross is not None:
            f["V"] = gross * FT_PER_M
            d["pay_ft"] = f["V"]
        if por is not None:
            f["W"] = por
            d["porosity"] = por
        if sw is not None:
            f["X"] = sw
            d["sw"] = sw
        for tr in trs:
            ff = {c: v for c, v in f.items() if blank(tr[c])}
            if ff:
                fills[tr["n"]] = ff
        d["status"] = "populated: " + ",".join(HDR[c].split(",")[0] for c in f) if f else "matched, no usable value"
        detail.append(d)

    # write v3: copy every part, edit only the Bubble Map Sample sheet
    zin = zipfile.ZipFile(TPL)
    sheet_part = "xl/worksheets/sheet1.xml"
    assert b'name="Bubble Map Sample" sheetId="1" r:id="rId1"' in zin.read("xl/workbook.xml")
    assert b'Id="rId1"' in zin.read("xl/_rels/workbook.xml.rels") and b'worksheets/sheet1.xml' in zin.read("xl/_rels/workbook.xml.rels")
    orig = zin.read(sheet_part)
    new, ycount = edit_sheet(orig, fills)
    assert revert_sheet(new) == orig, "sheet edit is not limited to V/W/X cells and Y cached values"
    tmp = OUT.with_name(OUT.stem + ".partial.xlsx")
    with open(tmp, "xb") as fh, zipfile.ZipFile(fh, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = new if info.filename == sheet_part else zin.read(info.filename)
            zi = zipfile.ZipInfo(info.filename, info.date_time)
            zi.compress_type, zi.external_attr = zipfile.ZIP_DEFLATED, info.external_attr
            zout.writestr(zi, data)
    if OUT.exists():
        tmp.unlink()
        sys.exit("STOP: v3 appeared during the run - not overwritten.")
    tmp.rename(OUT)

    # ---- validation (re-read the saved file)
    zo = zipfile.ZipFile(OUT)
    assert [i.filename for i in zo.infolist()] == [i.filename for i in zin.infolist()]
    for i in zin.infolist():
        if i.filename != sheet_part:
            assert zo.read(i.filename) == zin.read(i.filename), i.filename
    assert revert_sheet(zo.read(sheet_part)) == orig
    v3 = openpyxl.load_workbook(OUT, read_only=True, data_only=True)
    assert v3.sheetnames == openpyxl.load_workbook(TPL, read_only=True).sheetnames
    ws3 = v3["Bubble Map Sample"]
    nchk = 0
    for n, r in enumerate(ws3.iter_rows(min_row=2, values_only=True), start=2):
        f = fills.get(n, {})
        for c, i in (("V", 21), ("W", 22), ("X", 23)):
            assert (r[i] is None) if c not in f else (r[i] == f[c]), (n, c, r[i])
        if "W" in f:
            assert 0 <= r[22] <= 1
        if "X" in f:
            assert 0 <= r[23] <= 1 and r[24] == 1 - r[23], (n, r[23], r[24])
        else:
            assert r[24] in (None, ""), (n, r[24])
        nchk += 1
    assert nchk == len(rows)
    for d in detail:
        if d["pay_ft"] != "":
            g = resolve(res[(well_key(d["well"]), d["source_zone"])], "gross", lo=0)[0]
            assert abs(d["pay_ft"] - g * FT_PER_M) < 1e-12
    h1 = {p: sha(p) for p in (SRC, TPL)}
    assert h0 == h1, "input changed"

    # ---- reports (separate files; nothing added to the workbook)
    cols = list(detail[0])
    with open(REP_CSV, "x", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, cols)
        w.writeheader()
        w.writerows(detail)
    cells = collections.Counter(c for f in fills.values() for c in f)
    write_md(detail, fills, cells, stats, ycount, h0, h1, src_sheet, n_res, len(rows))
    print(f"v3 written: {OUT}")
    print(f"pairs={len(pairs)} exact={stats['exact']} smk={stats['.smk']} unmatched_well={stats['unmatched_well']} "
          f"unmatched_res={stats['unmatched_res']} rows_filled={len(fills)} cells={dict(cells)} Y_cache={ycount}")


def write_md(detail, fills, cells, stats, ycount, h0, h1, src_sheet, n_res, n_rows):
    L = []
    a = L.append
    a("# bubble_map_dulang_v3.xlsx - RES reservoir-property population: completion summary")
    a("")
    a(f"Run: {dt.datetime.now().isoformat(timespec='seconds')} - local processing only. Script: `BUBBLE_MAP/scripts/{Path(__file__).name}`.")
    a("")
    a("## Files")
    a("")
    a("| role | file | SHA-256 before | SHA-256 after | unchanged |")
    a("|---|---|---|---|---|")
    for role, p in (("input (read only)", SRC), ("template (read only)", TPL)):
        a(f"| {role} | `{p.relative_to(BASE)}` | `{h0[p]}` | `{h1[p]}` | {'YES' if h0[p] == h1[p] else 'NO'} |")
    a(f"| output (new) | `{OUT.relative_to(BASE)}` | - | `{sha(OUT)}` | - |")
    a(f"| match detail | `{REP_CSV.relative_to(BASE)}` | | | |")
    a("")
    a("## Result")
    a("")
    a(f"- Template rows (Bubble Map Sample): {n_rows}; distinct well + reservoir-unit pairs: {len(detail)}.")
    a(f"- Source: worksheet `{src_sheet}`, {n_res} RES-flag records (PAY / ROCK ignored).")
    a(f"- Pairs matched on the exact reservoir label: **{stats['exact']}**; on the `.smk` label (no exact-label RES record): **{stats['.smk']}**.")
    a(f"- Pairs unmatched - well absent from RES records: {stats['unmatched_well']}; well present but reservoir absent: {stats['unmatched_res']}.")
    a(f"- Rows that received at least one value: **{len(fills)}**. Cells written: Pay Thickness {cells['V']}, Porosity {cells['W']}, "
      f"Water Saturation {cells['X']}.")
    a(f"- Hydrocarbon Saturation: every v2 cell already holds the formula `=IF(Xn=\"\",\"\",1-Xn)` (populated, so preserved). It now "
      f"evaluates to 1 - Sw in the {ycount} rows where Sw was written; the formula's stored result was updated in those rows so the value "
      "shows without a full recalculation. No other cell was touched.")
    a("")
    a("## Rules applied")
    a("")
    a("- Well names: temporary key only (displayed names unchanged). `DULANG-` and `DULG-` prefixes are equivalent, the main well number "
      "ignores zero padding (A1 = A001), and every suffix must be present on both sides (A001W01 does not match A001). Sidetrack `ST` and `S` "
      "are treated as the same suffix (S02 = ST2), following the Updated Well Name rule used for this project. `R`, `W`, `H`, `P` suffixes "
      "must match exactly, so e.g. `DULANG-A10R1S2` does not match `DULG-A010S02`, and `DULANG-D19P1` does not match `DULG-D019H01`.")
    a("- Reservoir: the template RESERVOIR UNIT is `<block>-<reservoir>` (e.g. `SE1-E34`); the part after the first hyphen is matched "
      "in full against `Zones`. Exact label first; `<label>.smk` only when no exact-label RES record exists. `.mkr`, `.mkr_base`, `TSND`, "
      "top/base, `' SMK'` (space form), trailing-dot labels and all split/combined reservoirs are not used.")
    a("- Values: Pay Thickness = Gross Interval [m] x 3.28084 (full precision); Porosity = AV_POROSITY; Sw = AV_WATER_SATURATION (already "
      "fractions). `-999.25` is the source null value and is treated as missing. Porosity/Sw outside 0-1 and negative Gross are invalid.")
    a("- Duplicates in the selected record group: identical values used once; differing values = conflict -> cell left blank. A duplicate "
      "without a value does not block the value of the other record (listed below).")
    a("- Output cells are written without a style, like the existing blank cells in these columns (General display format).")
    a("")
    a("## Validation (all passed before the run ended)")
    a("")
    for t in ["Every zip part other than the Bubble Map Sample sheet is byte-identical to v2 (Mapping_Review, styles, shared strings, calcChain).",
              "Removing the written V/W/X cells and restoring the empty Y results reproduces the v2 sheet XML exactly: no other cell, formula, "
              "style, row or row order changed; no rows added, removed or duplicated; no worksheet or column added.",
              "Re-read v3: every written value equals the selected source value; Pay = Gross x 3.28084; porosity and Sw in 0-1; Shc = 1 - Sw.",
              "Unmatched / conflicting / missing cases remain blank.",
              "Input and template SHA-256 unchanged."]:
        a(f"- {t}")
    a("")

    def tbl(title, rows, cols, intro=""):
        a(f"## {title} ({len(rows)})")
        a("")
        if intro:
            a(intro)
            a("")
        if not rows:
            a("None.")
            a("")
            return
        a("| " + " | ".join(c for c, _ in cols) + " |")
        a("|" + "---|" * len(cols))
        for d in rows:
            a("| " + " | ".join(str(fn(d)).replace("|", "/") for _, fn in cols) + " |")
        a("")

    base = [("well", lambda d: d["well"]), ("reservoir unit", lambda d: d["reservoir_unit"]), ("template rows", lambda d: d["template_rows"])]
    src = [("source well", lambda d: d["source_well"]), ("Zones", lambda d: d["source_zone"]),
           ("Sheet1 rows", lambda d: d["source_rows"])]
    pop = [d for d in detail if d["status"].startswith("populated")]
    tbl("Populated pairs", pop, base + [("match", lambda d: d["match"])] + src +
        [("Pay ft", lambda d: "" if d["pay_ft"] == "" else round(d["pay_ft"], 4)), ("Porosity", lambda d: d["porosity"]), ("Sw", lambda d: d["sw"])],
        "Pay ft shown rounded here; the workbook holds full precision. Blank = not populated (see conflicts / missing).")
    conf = [d for d in detail if "CONFLICT" in d["pay_note"] + d["por_note"] + d["sw_note"]]
    tbl("Conflicts - cells left blank", conf, base + src +
        [("Pay", lambda d: d["pay_note"] if "CONFLICT" in d["pay_note"] else ""),
         ("Porosity", lambda d: d["por_note"] if "CONFLICT" in d["por_note"] else ""),
         ("Sw", lambda d: d["sw_note"] if "CONFLICT" in d["sw_note"] else "")],
        "Most duplicates are the same zone reported twice: once from TVDSS tops/bottoms and once from MDDF tops/bottoms. Gross Interval "
        "differs between the two (vertical vs measured thickness), and porosity/Sw sometimes differ in the third decimal.")
    miss = [d for d in detail if d["match"] != "none" and any(n and "CONFLICT" not in n and not n.startswith("other") for n in
                                                              (d["pay_note"], d["por_note"], d["sw_note"]))]
    tbl("Missing / invalid source values - cells left blank", miss, base + src +
        [("Pay", lambda d: "" if "CONFLICT" in d["pay_note"] or d["pay_note"].startswith("other") else d["pay_note"]),
         ("Porosity", lambda d: "" if "CONFLICT" in d["por_note"] or d["por_note"].startswith("other") else d["por_note"]),
         ("Sw", lambda d: "" if "CONFLICT" in d["sw_note"] or d["sw_note"].startswith("other") else d["sw_note"])])
    part = [d for d in detail if any(n.startswith("other") for n in (d["pay_note"], d["por_note"], d["sw_note"]))]
    tbl("Value used although another duplicate had no value", part, base + src +
        [("note", lambda d: " / ".join(n for n in (d["pay_note"], d["por_note"], d["sw_note"]) if n.startswith("other")))])
    um = [d for d in detail if d["match"] == "none"]
    tbl("Unmatched well / reservoir pairs - left blank", um, base + [("reason", lambda d: d["status"]),
                                                                      ("RES well name(s) for this key", lambda d: d["source_well"])])
    REP_MD.write_text("\n".join(L) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
