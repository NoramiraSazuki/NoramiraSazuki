#!/usr/bin/env python3
"""Build bubble_map_dulang_v5.xlsx (all platforms) and bubble_map_dulang_v5_DULANG-{A,B,C,D}.xlsx from v4.

Changes to the 'Bubble Map Sample' sheet (everything else is copied byte for byte from v4, calcChain dropped):
  1. WELL NAME (v4 column C): DULG-/DULANG- prefix -> DULANG-, leading zeros of the base number removed
     (DULG-A032S03 -> DULANG-A32S03); suffixes untouched.
  2. RESERVOIR UNIT (v4 column E) removed.
  3. PLATFORM inserted after FIELD NAME (new column C, DULANG-A..D from the standardised name); WELL NAME -> D,
     STRING TYPE -> E; columns F onward keep their letters. New cells copy the WELL NAME cell style.
  4. One empty row between consecutive ENTITY groups (entities are contiguous in v4, so no rows move).
Formulas are rewritten for the new layout: column refs C->D, D->E; relative row refs follow their cell, the absolute
$2 anchors of the cumulative SUMPRODUCTs stay on row 2. The cumulative-PD test
IF(AND(B=Bprev,C=Cprev,A=Aprev,E=Eprev),...) loses its RESERVOIR UNIT term: ENTITY (A) is always
WELL-STRING-RESERVOIR UNIT in v4, so the remaining terms give the same result. Shared formulas are written out
explicitly. Cached values are kept, except the cumulative Np/Wp (running totals from row 2) in the platform files,
which are recomputed over that file's rows. v4 is read only; existing outputs are never overwritten.
"""
import collections
import hashlib
import html
import re
import sys
import zipfile
from pathlib import Path

import openpyxl
from openpyxl.utils import column_index_from_string as cidx, get_column_letter as clet

BASE = Path(__file__).resolve().parent.parent
DIR = BASE / "bubble_map_dulang_sample"
SRC = DIR / "bubble_map_dulang_v4.xlsx"
OUT_ALL = DIR / "bubble_map_dulang_v5.xlsx"
PLATFORMS = ["DULANG-A", "DULANG-B", "DULANG-C", "DULANG-D"]
OUTS = {None: OUT_ALL, **{p: DIR / f"bubble_map_dulang_v5_{p}.xlsx" for p in PLATFORMS}}
REP_MD = BASE / "output_bubble-map" / "bubble_map_dulang_v5_summary.md"
SHEET = "xl/worksheets/sheet1.xml"
SHEET_NAME = "Bubble Map Sample"
NEW_HDR = "PLATFORM"
REF = re.compile(r"(?<![A-Za-z_\d.$])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\d(A-Za-z_])")
CELL = re.compile(rb'<c r="([A-Z]+)(\d+)"((?:\s+[\w:]+="[^"]*")*)\s*(?:/>|>(.*?)</c>)', re.S)
ROW = re.compile(rb'<row r="(\d+)"([^>]*?)(?:/>|>(.*?)</row>)', re.S)


def sha(p):
    return hashlib.sha256(Path(p).read_bytes()).hexdigest()


def std_name(name):
    m = re.fullmatch(r"(?:DULG|DULANG)-([A-D])0*(\d+)([A-Z0-9]*)", str(name or "").strip())
    return f"DULANG-{m.group(1)}{m.group(2)}{m.group(3)}" if m else None


def platform(std):
    m = re.fullmatch(r"(DULANG-[A-D])\d.*", std or "")
    return m.group(1) if m else None


def new_col(c):
    """v4 column letter -> v5 column letter (None for the removed RESERVOIR UNIT)"""
    return {"C": "D", "D": "E", "E": None}.get(c, c)


def relmove(f, d):
    return REF.sub(lambda m: m.group(0) if m.group(3) else f"{m.group(1)}{m.group(2)}{int(m.group(4)) + d}", f)


def move_formula(f, old_row, new_row):
    # cumulative-PD test: drop the RESERVOIR UNIT (E) comparison, ENTITY already contains it
    f = re.sub(r",\$?E\$?(\d+)=\$?E\$?(\d+)(?=\))", "", f)

    def sub(m):
        c = new_col(m.group(2))
        assert c, f"formula still refers to removed column E: {f}"
        r = int(m.group(4))
        if m.group(3):
            assert r <= 2, f
        else:
            r += new_row - old_row
        return f"{m.group(1)}{c}{m.group(3)}{r}"
    assert "!" not in f, f
    return REF.sub(sub, f)


def read_sst(z):
    xml = z.read("xl/sharedStrings.xml")
    head, rest = xml.split(b"<si>", 1)
    items = re.findall(rb"<si>(.*?)</si>", b"<si>" + rest, re.S)
    texts = [html.unescape("".join(t.decode() for t in re.findall(rb"<t[^>]*>(.*?)</t>", it, re.S))) for it in items]
    return xml, texts


def write_sst(xml, n_old, added, extra_refs):
    new_items = b"".join(b"<si><t>%s</t></si>" % html.escape(t, quote=False).encode() for t in added)
    xml = xml.replace(b"</sst>", new_items + b"</sst>")
    xml = re.sub(rb'uniqueCount="(\d+)"', lambda m: b'uniqueCount="%d"' % (n_old + len(added)), xml, count=1)
    return re.sub(rb' count="(\d+)"', lambda m: b' count="%d"' % (int(m.group(1)) + extra_refs), xml, count=1)


def parse_rows(xml):
    """-> header bytes, list of (row_num, row_attrs, [(col, row, attrs, inner)]), tail bytes"""
    head, rest = xml.split(b"<sheetData>", 1)
    body, tail = rest.split(b"</sheetData>", 1)
    rows = []
    for m in ROW.finditer(body):
        cells = [(c.group(1).decode(), int(c.group(2)), c.group(3).decode(), (c.group(4) or b"").decode())
                 for c in CELL.finditer(m.group(3) or b"")]
        rows.append((int(m.group(1)), m.group(2).decode(), cells))
    return head, rows, tail


def main():
    if any(o.exists() for o in OUTS.values()):
        sys.exit("STOP: output exists - not overwritten: " + ", ".join(o.name for o in OUTS.values() if o.exists()))
    h0 = sha(SRC)
    zin = zipfile.ZipFile(SRC)
    assert b'name="Bubble Map Sample" sheetId="1" r:id="rId1"' in zin.read("xl/workbook.xml")
    sst_xml, sst = read_sst(zin)
    head, rows, tail = parse_rows(zin.read(SHEET))

    # header checks (by name, before any change)
    hdr = {c: sst[int(re.search(r"<v>(\d+)</v>", inner).group(1))] for c, _, a, inner in rows[0][2]}
    assert hdr["A"] == "ENTITY" and hdr["B"] == "FIELD NAME" and hdr["C"] == "WELL NAME" \
        and hdr["E"] == "RESERVOIR UNIT", hdr

    # shared formula masters
    masters = {}
    for _, _, cells in rows:
        for c, r, a, inner in cells:
            m = re.match(r'<f t="shared" ref="[^"]+" si="(\d+)">([^<]*)</f>', inner)
            if m:
                masters[m.group(1)] = (c, r, m.group(2))

    def formula_of(c, r, inner):
        m = re.match(r"<f>([^<]*)</f>", inner)
        if m:
            return m.group(1)
        m = re.match(r'<f t="shared"(?: ref="[^"]+")? si="(\d+)"(?:/>|>[^<]*</f>)', inner)
        if m:
            mc, mr, mf = masters[m.group(1)]
            assert mc == c, (c, r)
            return relmove(mf, r - mr)
        assert "<f" not in inner, inner
        return None

    # records
    added = {}

    def sidx(text):
        if text in sst_index:
            return sst_index[text]
        if text not in added:
            added[text] = len(sst) + len(added)
        return added[text]
    sst_index = {t: i for i, t in enumerate(sst)}  # first occurrence

    recs, changed, unparsed, no_plat, no_entity = [], collections.Counter(), [], [], []
    std_from = collections.defaultdict(set)
    for rn, rattrs, cells in rows[1:]:
        byc = {c: (a, inner) for c, _, a, inner in cells}
        a_txt = None
        if "A" in byc and 't="s"' in byc["A"][0]:
            a_txt = sst[int(re.search(r"<v>(\d+)</v>", byc["A"][1]).group(1))]
        well_attrs, well_inner = byc.get("C", ("", ""))
        assert not well_inner or 't="s"' in well_attrs, rn
        old = sst[int(re.search(r"<v>(\d+)</v>", well_inner).group(1))] if well_inner else None
        new = std_name(old)
        if new is None:
            unparsed.append((rn, old))
            new = old
        elif new != old:
            changed[(old, new)] += 1
        if new:
            std_from[new].add(old)
        plat = platform(new)
        if plat is None:
            no_plat.append((rn, old))
        if not a_txt:
            no_entity.append(rn)
        recs.append(dict(rn=rn, rattrs=rattrs, cells=cells, entity=a_txt, plat=plat, well=new,
                         style=re.search(r'\s+s="\d+"', well_attrs).group(0) if 's="' in well_attrs else ""))

    hdr_style = re.search(r'\s+s="\d+"', dict((c, a) for c, _, a, _ in rows[0][2])["C"]).group(0)

    def build(sel):
        """sel: list of records (in v4 order) -> sheet xml, last row, number of separators"""
        out, n, prev_ent, seps = [], 1, object(), 0
        # header
        hc = []
        for c, r, a, inner in rows[0][2]:
            nc = new_col(c)
            if nc:
                hc.append((nc, f'<c r="{nc}1"{a}>{inner}</c>'))
        hc.append(("C", f'<c r="C1"{hdr_style} t="s"><v>{sidx(NEW_HDR)}</v></c>'))
        out.append(f'<row r="1"{rows[0][1]}>' + "".join(x for _, x in sorted(hc, key=lambda t: cidx(t[0]))) + "</row>")
        sums = [0.0, 0.0]
        for rec in sel:
            if n > 1 and rec["entity"] != prev_ent:
                n += 1
                seps += 1
            prev_ent = rec["entity"]
            n += 1
            vals = {}
            cs = []
            for c, r, a, inner in rec["cells"]:
                nc = new_col(c)
                if nc is None:
                    continue
                if nc == "D":  # standardised well name
                    inner = f"<v>{sidx(rec['well'])}</v>" if rec["well"] is not None else ""
                f = formula_of(c, r, inner)
                if f is not None:
                    v = re.search(r"<v>[^<]*</v>|<v/>", inner)
                    inner = f"<f>{move_formula(f, r, n)}</f>{v.group(0) if v else ''}"
                vm = re.search(r"<v>([^<]*)</v>", inner)
                vals[nc] = vm.group(1) if vm and 't="' not in a else None
                cs.append([nc, a, inner])
            if rec["plat"]:
                cs.append(["C", f'{rec["style"]} t="s"', f"<v>{sidx(rec['plat'])}</v>"])
            if sel is not ALL:  # cumulative Np / Wp over this file's rows
                k = float(vals.get("K") or 0)
                sums[0] += k * float(vals.get("O") or 0)
                sums[1] += k * float(vals.get("P") or 0)
                for x in cs:
                    if x[0] in ("R", "S") and "SUMPRODUCT($K$2:" in x[2]:
                        x[2] = re.sub(r"<v>[^<]*</v>", f"<v>{repr(sums[0] if x[0] == 'R' else sums[1])}</v>", x[2])
            cs.sort(key=lambda t: cidx(t[0]))
            out.append(f'<row r="{n}"{rec["rattrs"]}>' + "".join(
                f'<c r="{c}{n}"{a}/>' if not inner else f'<c r="{c}{n}"{a}>{inner}</c>' for c, a, inner in cs)
                + "</row>")
        h = re.sub(rb'<dimension ref="A1:([A-Z]+)\d+"/>', lambda m: b'<dimension ref="A1:%s%d"/>' % (m.group(1), n), head)
        h = new_cols(h)
        t = re.sub(rb'<autoFilter ref="A1:([A-Z]+)\d+"', lambda m: b'<autoFilter ref="A1:%s%d"' % (m.group(1), n), tail)
        return h + b"<sheetData>" + "".join(out).encode() + b"</sheetData>" + t, n, seps

    def new_cols(h):
        """rebuild <cols>: widths/styles follow their columns; PLATFORM copies WELL NAME's"""
        m = re.search(rb"<cols>(.*?)</cols>", h, re.S)
        per = {}
        for cm in re.finditer(rb'<col min="(\d+)" max="(\d+)"([^>]*)/>', m.group(1)):
            for i in range(int(cm.group(1)), int(cm.group(2)) + 1):
                per[i] = cm.group(3)
        newper = {}
        for i, a in per.items():
            nc = new_col(clet(i))
            if nc:
                newper[cidx(nc)] = a
        if 3 in per:
            newper[3] = per[3]
        parts, keys = [], sorted(newper)
        i = 0
        while i < len(keys):
            j = i
            while j + 1 < len(keys) and keys[j + 1] == keys[j] + 1 and newper[keys[j + 1]] == newper[keys[i]]:
                j += 1
            parts.append(b'<col min="%d" max="%d"%s/>' % (keys[i], keys[j], newper[keys[i]]))
            i = j + 1
        return h[:m.start(1)] + b"".join(parts) + h[m.end(1):]

    ALL = recs
    results = {}
    for p, out in OUTS.items():
        sel = ALL if p is None else [r for r in recs if r["plat"] == p]
        results[p] = build(sel) + (len(sel),)
    sst_new = write_sst(sst_xml, len(sst), list(added), 0)

    for p, out in OUTS.items():
        sheet_xml, last, seps, nrec = results[p]
        tmp = out.with_name(out.stem + ".partial.xlsx")
        with open(tmp, "xb") as fh, zipfile.ZipFile(fh, "w", zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                fn = info.filename
                if fn == "xl/calcChain.xml":
                    continue
                data = zin.read(fn)
                if fn == SHEET:
                    data = sheet_xml
                elif fn == "xl/sharedStrings.xml":
                    data = sst_new
                elif fn == "xl/workbook.xml":
                    data, k = re.subn(rb"('Bubble Map Sample'!\$A\$1:\$[A-Z]+\$)\d+",
                                      lambda m: m.group(1) + str(last).encode(), data)
                    assert k == 1
                elif fn == "xl/_rels/workbook.xml.rels":
                    data, k = re.subn(rb'<Relationship Id="rId\d+" Type="[^"]+/calcChain" Target="calcChain.xml"/>',
                                      b"", data)
                    assert k == 1
                elif fn == "[Content_Types].xml":
                    data, k = re.subn(rb'<Override PartName="/xl/calcChain.xml"[^>]*/>', b"", data)
                    assert k == 1
                zi = zipfile.ZipInfo(fn, info.date_time)
                zi.compress_type, zi.external_attr = zipfile.ZIP_DEFLATED, info.external_attr
                zout.writestr(zi, data)
        tmp.rename(out)
        print(f"{out.name}: records {nrec}, separators {seps}, last row {last}")

    validate(recs, results, std_from)
    assert sha(SRC) == h0, "v4 changed"

    dups = {s: sorted(o) for s, o in std_from.items() if len(o) > 1}
    by_plat = collections.Counter(r["plat"] for r in recs)
    lines = ["# bubble_map_dulang_v5 - summary", "", f"Source: `{SRC.relative_to(BASE)}` (unchanged)", "",
             "| Workbook | Records | Entity groups | Blank separator rows |", "|---|---|---|---|"]
    for p, out in OUTS.items():
        _, last, seps, nrec = results[p]
        lines.append(f"| {out.name} | {nrec} | {seps + 1} | {seps} |")
    lines += ["", f"- Column removed: RESERVOIR UNIT (v4 column E). Column added: {NEW_HDR} (column C). "
              "WELL NAME moved C -> D, STRING TYPE D -> E; F onward unchanged.",
              f"- Well names changed: {len(changed)} distinct names ({sum(changed.values())} cells); "
              f"{len(std_from) - len(changed)} already standard.",
              f"- Duplicate names created by standardisation: {dups or 'none'}",
              f"- Unparsed well names: {unparsed or 'none'}", f"- Unidentified platform: {no_plat or 'none'}",
              f"- Missing ENTITY: {no_entity or 'none'}",
              "- Formulas: cumulative-PD test no longer compares RESERVOIR UNIT (ENTITY = WELL-STRING-RESERVOIR "
              "UNIT in every v4 row, so results are unchanged); cumulative Np/Wp in platform files total that "
              "platform's rows (cached values recomputed).", "", "## Well names changed", "",
              "| v4 | v5 | Rows |", "|---|---|---|"]
    lines += [f"| {o} | {n} | {c} |" for (o, n), c in sorted(changed.items())]
    REP_MD.parent.mkdir(exist_ok=True)
    with open(REP_MD, "x", encoding="utf-8") as fh:
        fh.write("\n".join(lines) + "\n")
    print("by platform:", dict(by_plat), "| names changed:", len(changed), "cells", sum(changed.values()))
    print("duplicates:", dups or "none", "| unparsed:", unparsed or "none", "| no platform:", no_plat or "none",
          "| no entity:", no_entity or "none")


def validate(recs, results, std_from):
    src = list(openpyxl.load_workbook(SRC, read_only=True)[SHEET_NAME].iter_rows(values_only=True))
    srcv = list(openpyxl.load_workbook(SRC, read_only=True, data_only=True)[SHEET_NAME].iter_rows(values_only=True))
    exp_hdr = list(src[0][:2]) + [NEW_HDR] + list(src[0][2:4]) + list(src[0][5:])
    mr_src = list(openpyxl.load_workbook(SRC, read_only=True)["Mapping_Review"].iter_rows(values_only=True))
    v5_rows = {}
    for p, out in OUTS.items():
        wb = openpyxl.load_workbook(out, read_only=True)
        assert wb.sheetnames == [SHEET_NAME, "Mapping_Review"]
        o = list(wb[SHEET_NAME].iter_rows(values_only=True))
        ov = list(openpyxl.load_workbook(out, read_only=True, data_only=True)[SHEET_NAME].iter_rows(values_only=True))
        assert list(o[0][:len(exp_hdr)]) == exp_hdr and all(v is None for v in o[0][len(exp_hdr):]), o[0]
        assert o[-1] and any(v is not None for v in o[-1]), "trailing blank row"
        sel = [r for r in recs if p is None or r["plat"] == p]
        k, prev, sums = 1, None, [0.0, 0.0]
        for rec in sel:
            if k > 1 and rec["entity"] != prev:
                assert all(v is None for v in o[k]) and all(v is None for v in ov[k]), (p, k)
                k += 1
            prev = rec["entity"]
            i = rec["rn"] - 1
            a, b, av, bv = src[i], o[k], srcv[i], ov[k]
            assert b[2] == rec["plat"] and b[3] == std_name(a[2]) and b[3] in std_from, (p, k)
            assert b[3].startswith(b[2]) and b[0] == a[0] == f"{a[2]}-{a[3]}-{a[4]}"
            sa, sb = list(a[:2]) + list(a[3:4]) + list(a[5:]), list(b[:2]) + list(b[4:])
            sav, sbv = list(av[:2]) + list(av[3:4]) + list(av[5:]), list(bv[:2]) + list(bv[4:])
            for c, (x, y) in enumerate(zip(sa, sb)):
                if isinstance(x, str) and x.startswith("="):
                    assert y == "=" + move_formula(x[1:], i + 1, k + 1), (p, k, c, x, y)
                else:
                    assert x == y, (p, k, c, x, y)
            if p is not None:
                sums[0] += (av[10] or 0) * (av[14] or 0)
                sums[1] += (av[10] or 0) * (av[15] or 0)
                for c, s in ((17, sums[0]), (18, sums[1])):
                    if isinstance(a[c], str) and a[c].startswith("=SUMPRODUCT"):
                        assert abs(bv[c] - s) <= 1e-9 * max(1, abs(s)), (p, k, c, bv[c], s)
                        sbv[c - 2] = sav[c - 2]  # R/S sit at index 15/16 of the compared lists
            assert sav == sbv, (p, k)
            if p is None:
                v5_rows[rec["rn"]] = (b, bv)
            else:
                cb, cbv = v5_rows[rec["rn"]]  # reconciles with combined: same values (Np/Wp totals aside)
                assert bv[:17] == cbv[:17] and bv[19:] == cbv[19:] and b[:9] == cb[:9], (p, k)
            k += 1
        assert k == len(o), (p, k, len(o))
        assert list(wb["Mapping_Review"].iter_rows(values_only=True)) == mr_src
    print(f"validated: 5 workbooks, header {len(exp_hdr)} columns, {len(recs)} records in v5, "
          f"platform files reconcile")


if __name__ == "__main__":
    main()
