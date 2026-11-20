#!/usr/bin/env python3
"""Recalculate Cumulative Oil (Np, column R) and Water (Wp, column S) production in one platform workbook:
    python3 fix_v5_dulang_a_cumulative_np_wp.py <A|B|C|D>   -> bubble_map_dulang_v5_DULANG-<P>.xlsx
NPRADIUS (column Z) keeps its formula; its cached values are recomputed from the new Np (same formula as Excel).

Per ENTITY (full perforation id = string + reservoir unit; the FPC allocation is made per perforation), months in
chronological order, oil and water independently:
    monthly volume = 'Alloc. Oil Volume(bbl)' / 'Alloc. Water Volume(bbl)' of that perforation and month in
                     input_FPC_Prod Allocation_11Jun126.xlsx, used directly (equals Producing Days x PD rate wherever
                     both exist - checked; PD rates, not calendar-day rates). Volumes allocated in months with
                     0 producing days are used and flagged; Producing Days = 0 with a blank volume is a valid zero
                     month (user decisions 3 Nov 2026), flagged.
    cumulative     = cumulative of the previous record of the SAME entity that has a value + monthly volume
Nothing is reused from the existing cumulative columns. A month with missing volume gets an empty cell; later totals
of that entity keep accumulating from the last available total and are flagged incomplete. Months before an entity's
first available volume are empty (no opening balances are documented). Formulas reference the previous record with
a value explicitly:  first record =<volume>,  later records =R<q>+<volume>  (q = previous valued row, same entity).

Sibling perforations of one string with identical non-zero volume in a month are flagged (possible unsplit string
volume) but kept, as the allocation file is the documented per-reservoir allocation.
Only R/S data cells change. A backup is made first; the result is validated before it replaces the file.
"""
import collections
import csv
import datetime as dt
import hashlib
import os
import re
import math
import shutil
import sys
import zipfile
from pathlib import Path

import openpyxl

BASE = Path(__file__).resolve().parent.parent
PLAT = sys.argv[1].upper() if len(sys.argv) > 1 else "A"
assert PLAT in "ABCD" and len(PLAT) == 1, PLAT
WB = BASE / "bubble_map_dulang_sample" / f"bubble_map_dulang_v5_DULANG-{PLAT}.xlsx"
SRC = BASE / "input_FPC_Prod Allocation_11Jun126.xlsx"
STAMP = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
BACKUP = BASE / "bubble_map_dulang_sample" / "_backup" / f"bubble_map_dulang_v5_DULANG-{PLAT}_before_NpWp_fix_{STAMP}.xlsx"
REP_CSV = BASE / "output_bubble-map" / f"v5_DULANG-{PLAT}_NpWp_fix_flags_{STAMP}.csv"
REP_MD = BASE / "output_bubble-map" / f"v5_DULANG-{PLAT}_NpWp_fix_summary_{STAMP}.md"
SHEET = "xl/worksheets/sheet1.xml"
PHASES = [
    dict(tag="Np", fluid="oil", col="R", hdr="Cumulative Oil Production, Np (bbl)", wb_rate="OIL PD RATE (bbls/d)",
         src_vol=(12, "Alloc. Oil Volume(bbl)"), src_rate=(19, "Oil PD Rate (bbls/d)")),
    dict(tag="Wp", fluid="water", col="S", hdr="Cumulative Water Production, Wp (bbl)",
         wb_rate="WATER PD RATE (bbls/d)", src_vol=(15, "Alloc. Water Volume(bbl)"),
         src_rate=(22, "Water PD Rate(bbls/d)")),
]


def sha(p):
    return hashlib.sha256(Path(p).read_bytes()).hexdigest()


def isnum(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def same(a, b, tol=1e-9):
    if a in (None, "") or b in (None, ""):
        return a in (None, "") and b in (None, "")
    return abs(a - b) <= tol * max(1.0, abs(a), abs(b))


NPR_FX = re.compile(r'=IF\(COUNT\(R(\d+),U\1,V\1,W\1,Y\1\)<5,"",IF\(OR\(R\1<0,U\1<=0,V\1<=0,V\1>1,W\1<0,'
                    r'W\1>=1,Y\1<=0\),"",SQRT\(\(5\.615\*R\1\*Y\1\)/\(PI\(\)\*U\1\*V\1\*\(1-W\1\)\)\)\)\)')


def npradius(r, u, v, w, y):
    """Excel: IF(COUNT(R,U,V,W,Y)<5,"",IF(OR(R<0,U<=0,V<=0,V>1,W<0,W>=1,Y<=0),"",SQRT((5.615*R*Y)/(PI()*U*V*(1-W)))))"""
    if not all(isnum(x) for x in (r, u, v, w, y)):
        return ""
    if r < 0 or u <= 0 or v <= 0 or v > 1 or w < 0 or w >= 1 or y <= 0:
        return ""
    return math.sqrt((5.615 * r * y) / (math.pi * u * v * (1 - w)))


def fmt(v):
    return "" if v in (None, "") else f"{v:,.0f}"


def sheet_rows(path, data_only=False):
    ws = openpyxl.load_workbook(path, read_only=True, data_only=data_only)["Bubble Map Sample"]
    return [list(r) for r in ws.iter_rows(values_only=True)]


def load_source(entities):
    ws = openpyxl.load_workbook(SRC, read_only=True).worksheets[0]
    hdr = [list(r) for r in ws.iter_rows(min_row=1, max_row=2, values_only=True)]
    assert hdr[1][2] == "String" and hdr[1][3] == "Perforation" and hdr[0][8] == "Alloc. Month" \
        and hdr[0][9] == "Producing Days" and hdr[0][19] == "Producing Day Rate", hdr
    for ph in PHASES:
        for i, h in (ph["src_vol"], ph["src_rate"]):
            assert hdr[1][i] == h, (i, hdr[1][i], h)
    src = collections.defaultdict(list)
    for r in ws.iter_rows(min_row=3, values_only=True):
        if r[3] in entities:
            src[(r[3], r[8])].append(dict(string=r[2], pd=r[9], oil=(r[12], r[19]), water=(r[15], r[22])))
    return src


def main():
    h_src = sha(SRC)
    f = sheet_rows(WB)
    old_v = sheet_rows(WB, data_only=True)
    ix = {h.strip(): i for i, h in enumerate(f[0]) if h}
    E, G, K = ix["ENTITY"], ix["ALLOC. DATE"], ix["PRODUCTION DAYS (PD)"]
    assert E == 0
    for ph in PHASES:
        ph["X"], ph["O"] = ix[ph["hdr"]], ix[ph["wb_rate"]]
        assert ph["col"] == openpyxl.utils.get_column_letter(ph["X"] + 1)

    # entity groups in sheet order; chronological order and duplicates checked before accumulating
    groups, last = collections.OrderedDict(), None
    dup = collections.Counter()
    for n, r in enumerate(f[1:], 2):
        if r[E] is None:
            assert all(v is None for v in r), f"row {n}: data without ENTITY"
            continue
        dup[(r[E], r[G])] += 1
        if r[E] != last:
            assert r[E] not in groups, f"ENTITY {r[E]} not contiguous"
            groups[r[E]] = []
        groups[r[E]].append(n)
        last = r[E]
    dups = {k: c for k, c in dup.items() if c > 1}
    for e in groups:  # chronological processing even if sheet order were not
        groups[e].sort(key=lambda n: f[n - 1][G])
    out_of_order = [e for e, ns in groups.items() if ns != sorted(ns)]
    src = load_source(set(groups))
    src_dups = {k: v for k, v in src.items() if len(v) > 1}

    # sibling perforations of one string with identical non-zero volume in a month
    by_string = collections.defaultdict(list)
    for (e, m), lst in src.items():
        by_string[(lst[0]["string"], m)].append((e, lst[0]))
    identical = {"oil": set(), "water": set()}
    for (s, m), lst in by_string.items():
        if len(lst) > 1:
            for fl in identical:
                vals = [x[fl][0] for _, x in lst]
                if all(v not in (None, 0) for v in vals) and len({round(v, 9) for v in vals}) == 1:
                    identical[fl].update((e, m) for e, _ in lst)

    res = {}  # (phase tag) -> {row: (cum or None, formula or None)}
    flags = collections.defaultdict(list)  # row -> [text]
    per_ent = {ph["tag"]: {} for ph in PHASES}
    issues = {ph["tag"]: collections.Counter() for ph in PHASES}
    pd_dev = {ph["tag"]: 0.0 for ph in PHASES}
    for ph in PHASES:
        tag, fl, col, O, X = ph["tag"], ph["fluid"], ph["col"], ph["O"], ph["X"]
        out = res[tag] = {}
        for e, ns in groups.items():
            cum, q, started, inc_from, miss, seen, vsum = None, None, False, None, collections.Counter(), set(), 0.0
            for n in ns:
                r = f[n - 1]
                s = src.get((e, r[G]))
                assert s, ("no source record", e, r[G])
                vol, rate = s[0][fl]
                pd = s[0]["pd"]
                assert (pd == r[K] or same(pd, r[K])) and (rate == r[O] or same(rate, r[O])), (n, tag)
                fx = None
                if (e, r[G]) in seen:
                    use = None
                    flags[n].append(f"{tag}: DUPLICATE ENTITY/MONTH - not counted")
                    issues[tag]["duplicate entity/month"] += 1
                elif vol is not None:
                    use = float(vol)
                    if isnum(pd) and isnum(rate) and pd > 0:
                        pd_dev[tag] = max(pd_dev[tag], abs(pd * rate - vol) / max(1.0, abs(vol)))
                    if vol != 0 and not (isnum(pd) and pd > 0):
                        flags[n].append(f"{tag}: allocated volume with 0 producing days, no rate - volume used")
                        issues[tag]["allocated volume with 0 producing days"] += 1
                    if (e, r[G]) in identical[fl]:
                        flags[n].append(f"{tag}: same volume as sibling perforation(s) of this string - "
                                        "check split")
                        issues[tag]["identical volume across sibling perforations"] += 1
                elif isnum(pd) and pd == 0:
                    use = 0.0
                    flags[n].append(f"{tag}: producing days 0, volume blank - valid zero")
                    issues[tag]["producing days 0 with blank volume (zero)"] += 1
                else:
                    use = None
                    kind = "producing days blank" if pd is None else f"{fl} rate/volume blank with PD > 0"
                    if started:
                        miss[kind] += 1
                        inc_from = inc_from or r[G]
                        flags[n].append(f"{tag}: MISSING ({kind}) - cell blank")
                    else:
                        miss["before first volume: " + kind] += 1
                        flags[n].append(f"{tag}: missing before first available volume ({kind}) - cell blank")
                seen.add((e, r[G]))
                if use is not None:
                    fx = f"{col}{q}+{repr(use)}" if q else repr(use)
                    cum = (cum or 0.0) + use
                    vsum += use
                    q, started = n, True
                    if inc_from:
                        flags[n].append(f"{tag}: INCOMPLETE - missing month(s) since {inc_from:%Y-%m}")
                out[n] = (cum if use is not None else None, fx)
            per_ent[tag][e] = dict(final=cum, vsum=vsum, inc_from=inc_from, miss=miss,
                                   old_final=old_v[ns[-1] - 1][X],
                                   changed=sum(1 for n in ns if not same(old_v[n - 1][X], out[n][0])))
        assert pd_dev[tag] < 1e-6, (tag, pd_dev[tag])

    # NPRADIUS (Z) from the new Np; the formula itself is checked and left as it is
    Z = ix["NPRADIUS"]
    U, V, W, Y = (ix[h] for h in ("Pay Thickness, h (ft)", "Porosity, F (fraction)",
                                  "Water Saturation, Sw (fraction)", "Bo (rb/stb)"))
    assert (Z, U, V, W, Y) == (25, 20, 21, 22, 24)
    npr = {}
    for n in res["Np"]:
        assert NPR_FX.fullmatch(f[n - 1][Z] or "") and NPR_FX.fullmatch(f[n - 1][Z]).group(1) == str(n), (n, f[n - 1][Z])
        r = old_v[n - 1]
        npr[n] = npradius(res["Np"][n][0], r[U], r[V], r[W], r[Y])

    # backup, then write
    BACKUP.parent.mkdir(exist_ok=True)
    shutil.copy2(WB, BACKUP)
    assert sha(BACKUP) == sha(WB)
    zin = zipfile.ZipFile(BACKUP)
    xml = zin.read(SHEET)
    has_chain = "xl/calcChain.xml" in zin.namelist()

    def cell(m):
        c, n, attrs = m.group(1).decode(), int(m.group(2)), m.group(3).decode()
        tag = {"R": "Np", "S": "Wp"}[c]
        if n not in res[tag]:
            return m.group(0)
        s = re.search(r'\s+s="\d+"', attrs)
        s = s.group(0) if s else ""
        cum, fx = res[tag][n]
        if fx is None:
            return f'<c r="{c}{n}"{s}/>'.encode()
        return f'<c r="{c}{n}"{s}><f>{fx}</f><v>{repr(cum)}</v></c>'.encode()

    new_xml = re.sub(rb'<c r="([RS])(\d+)"((?:\s+[\w:]+="[^"]*")*)\s*(?:/>|>.*?</c>)', cell, xml, flags=re.S)

    def zcell(m):
        n, attrs, fel = int(m.group(1)), m.group(2).decode(), m.group(3).decode()
        if n not in npr:
            return m.group(0)
        attrs = re.sub(r'\s+t="[^"]*"', "", attrs)
        if npr[n] == "":
            return f'<c r="Z{n}"{attrs} t="str">{fel}<v></v></c>'.encode()
        return f'<c r="Z{n}"{attrs}>{fel}<v>{repr(npr[n])}</v></c>'.encode()

    new_xml, kz = re.subn(rb'<c r="Z(\d+)"((?:\s+[\w:]+="[^"]*")*)>(<f[^>]*/>|<f[^>]*>[^<]*</f>)(?:<v>[^<]*</v>|<v/>)?</c>',
                          zcell, new_xml)
    assert kz >= len(npr), (kz, len(npr))
    tmp = WB.with_name(WB.stem + ".partial.xlsx")
    with open(tmp, "xb") as fh, zipfile.ZipFile(fh, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            fn = info.filename
            if fn == "xl/calcChain.xml":
                continue
            data = zin.read(fn)
            if fn == SHEET:
                data = new_xml
            elif fn == "xl/workbook.xml" and b"fullCalcOnLoad" not in data:
                data, c = re.subn(rb"<calcPr ([^>]*?)\s*/>", rb'<calcPr \1 fullCalcOnLoad="1"/>', data)
                assert c == 1
            elif has_chain and fn == "xl/_rels/workbook.xml.rels":
                data, c = re.subn(rb'<Relationship Id="rId\d+" Type="[^"]+/calcChain" Target="calcChain.xml"/>', b"",
                                  data)
                assert c == 1
            elif has_chain and fn == "[Content_Types].xml":
                data, c = re.subn(rb'<Override PartName="/xl/calcChain.xml"[^>]*/>', b"", data)
                assert c == 1
            zi = zipfile.ZipInfo(fn, info.date_time)
            zi.compress_type, zi.external_attr = zipfile.ZIP_DEFLATED, info.external_attr
            zout.writestr(zi, data)

    # ---- validation before replacing the file
    zt = zipfile.ZipFile(tmp)
    for info in zin.infolist():
        if info.filename not in (SHEET, "xl/workbook.xml", "xl/calcChain.xml", "xl/_rels/workbook.xml.rels",
                                 "[Content_Types].xml"):
            assert zt.read(info.filename) == zin.read(info.filename), info.filename
    tf, tv = sheet_rows(tmp), sheet_rows(tmp, data_only=True)
    assert len(tf) == len(f)
    Xs = {ph["X"]: ph["tag"] for ph in PHASES}
    for n, (a, b, av, bv) in enumerate(zip(f, tf, old_v, tv), 1):
        for c in range(max(len(a), len(b))):
            x, y = (a[c] if c < len(a) else None), (b[c] if c < len(b) else None)
            if c in Xs and n in res[Xs[c]]:
                cum, fx = res[Xs[c]][n]
                assert (y is None and bv[c] is None) if fx is None else (y == "=" + fx and same(bv[c], cum, 0)), n
            elif c == Z and n in npr:
                assert x == y and (bv[c] in ("", None) if npr[n] == "" else same(bv[c], npr[n], 0)), (n, bv[c], npr[n])
            else:
                assert x == y, (n, c, x, y)
                if c < len(av):
                    assert av[c] == bv[c], (n, c)
    for ph in PHASES:
        tag, X, col = ph["tag"], ph["X"], ph["col"]
        for e, ns in groups.items():
            prev_valued = None
            for n in ns:  # each formula points at the previous valued record of the same entity
                y = tf[n - 1][X]
                if y is None:
                    continue
                m = re.fullmatch(r"=(?:%s(\d+)\+)?(.+)" % col, y)
                ref = int(m.group(1)) if m.group(1) else None
                assert ref == prev_valued and (ref is None or tf[ref - 1][E] == e), (tag, n, ref, prev_valued)
                exp = (tv[ref - 1][X] if ref else 0.0) + float(m.group(2))
                assert same(exp, tv[n - 1][X], 0), (tag, n)
                prev_valued = n
            d = per_ent[tag][e]  # final = sum of the entity's own monthly volumes
            own = [src[(e, f[n - 1][G])][0][ph["fluid"]][0] for n in ns]
            tot = sum(v for v in own if v is not None)
            assert d["final"] is None or same(d["final"], tot, 1e-12) and same(d["final"], d["vsum"], 1e-12), e
            last_valued = [tv[n - 1][X] for n in ns if tv[n - 1][X] is not None]
            assert same(last_valued[-1] if last_valued else None, d["final"], 0), e
    assert list(openpyxl.load_workbook(BACKUP, read_only=True)["Mapping_Review"].iter_rows(values_only=True)) == \
        list(openpyxl.load_workbook(tmp, read_only=True)["Mapping_Review"].iter_rows(values_only=True))
    zt.close()
    os.replace(tmp, WB)
    assert sha(SRC) == h_src

    # ---- reports
    with open(REP_CSV, "x", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(["row", "entity", "month", "producing_days", "oil_pd_rate", "water_pd_rate", "np_new", "wp_new",
                    "flags"])
        for n in sorted(flags):
            r = f[n - 1]
            w.writerow([n, r[E], r[G].date().isoformat(), r[K], r[PHASES[0]["O"]], r[PHASES[1]["O"]],
                        res["Np"][n][0], res["Wp"][n][0], "; ".join(flags[n])])
    lines = [f"# v5_DULANG-{PLAT} cumulative Np / Wp recalculation ({STAMP})", "",
             f"- Workbook: `{WB.relative_to(BASE)}` (columns R and S; NPRADIUS cached values); backup: `{BACKUP.relative_to(BASE)}`",
             f"- Monthly volumes: allocated oil / water volume per perforation (= ENTITY) and month from `{SRC.name}`;"
             f" equal to Producing Days x PD rate wherever both exist (max relative difference oil "
             f"{pd_dev['Np']:.1e}, water {pd_dev['Wp']:.1e}). No opening balances documented.",
             f"- Entities: {len(groups)}. Duplicate entity/month: {len(dups) or 'none'} in workbook, "
             f"{len(src_dups) or 'none'} in source. Entities out of month order: {out_of_order or 'none'}.",
             f"- Row flags: `{REP_CSV.relative_to(BASE)}` ({len(flags)} rows)", ""]
    for ph in PHASES:
        tag = ph["tag"]
        pe = per_ent[tag]
        tot = collections.Counter()
        for d in pe.values():
            tot.update(d["miss"])
        lines += [f"## {tag}", "", f"- Entities changed: {sum(1 for d in pe.values() if d['changed'])}; "
                  f"incomplete: {sum(1 for d in pe.values() if d['inc_from'])}",
                  f"- Missing months: {dict(tot) or 'none'}", f"- Allocation issues: {dict(issues[tag]) or 'none'}", "",
                  f"| ENTITY | {tag} before (last row) | {tag} new (final) | Rows changed | Incomplete from |",
                  "|---|---|---|---|---|"]
        lines += [f"| {e} | {fmt(d['old_final'])} | {fmt(d['final'])} | {d['changed']} | "
                  f"{d['inc_from'].strftime('%Y-%m') if d['inc_from'] else ''} |" for e, d in pe.items()]
        lines.append("")
    with open(REP_MD, "x", encoding="utf-8") as fh:
        fh.write("\n".join(lines))

    for ph in PHASES:
        tag = ph["tag"]
        pe = per_ent[tag]
        tot = collections.Counter()
        for d in pe.values():
            tot.update(d["miss"])
        print(f"{tag}: changed {sum(1 for d in pe.values() if d['changed'])}/{len(pe)}, "
              f"incomplete {sum(1 for d in pe.values() if d['inc_from'])}, missing {dict(tot)}, issues "
              f"{dict(issues[tag])}")
        print(f"  unchanged: {[e for e, d in pe.items() if not d['changed']]}")
        print(f"  incomplete: {[e for e, d in pe.items() if d['inc_from']]}")
    print(f"dups wb {len(dups)} src {len(src_dups)} out_of_order {len(out_of_order)}")
    nz = sum(1 for v in npr.values() if v != "")
    oldz = sum(1 for n in npr if isnum(old_v[n - 1][Z]))
    print(f"NPRADIUS: {nz} rows with a value (before {oldz}), {len(npr) - nz} blank "
          f"(Np blank or h/phi/Sw/Bo missing or out of range)")
    print(f"backup {BACKUP}\nreport {REP_MD}\nflags {REP_CSV}")


if __name__ == "__main__":
    main()
