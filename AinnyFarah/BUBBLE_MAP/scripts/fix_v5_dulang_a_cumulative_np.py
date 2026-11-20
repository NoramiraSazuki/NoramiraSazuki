#!/usr/bin/env python3
"""Correct a cumulative-production column in bubble_map_dulang_v5_DULANG-A.xlsx only.

    python3 fix_v5_dulang_a_cumulative_np.py np   -> Cumulative Oil Production, Np (column R)
    python3 fix_v5_dulang_a_cumulative_np.py wp   -> Cumulative Water Production, Wp (column S)

v5 had SUMPRODUCT($K$2:Kn,$O$2:On) / ($K$2:Kn,$P$2:Pn): a running total over every row above, across all entities.
New rule, per ENTITY (the full perforation id = string + reservoir unit, the level the FPC allocation is made at), in
row order (rows of an entity are contiguous and in ascending allocation month - checked):
    monthly volume = 'Alloc. Oil Volume(bbl)' / 'Alloc. Water Volume(bbl)' of that perforation/month in
                     input_FPC_Prod Allocation_11Jun126.xlsx
                     (user decision 3 Nov 2026: use the allocated volume directly. It equals Producing Days x PD Rate
                     wherever both exist - checked - and also covers months allocated with 0 producing days.)
                     Producing Days = 0 with a blank source volume = valid zero month (user decision), flagged.
    cumulative = previous cumulative of the same entity + monthly volume
Months before the entity's first available monthly volume leave the cumulative blank; a month with no volume (blank
Producing Days, or Producing Days > 0 without rate/volume) adds nothing and carries the previous value forward - those
rows and every later row of the entity are reported as incomplete. In-sheet formula, monthly volume as a number:
    available:  =N(IF(An=Ap,Xp,""))+<volume>          missing:  =IF(An=Ap,Xp,"")
Only that column's data cells change (calcChain dropped if present, full recalculation on open). The original is copied
to bubble_map_dulang_sample/_backup/ first; the result is validated before it replaces the file. Flags go to
output_bubble-map/ (no worksheet is added).
"""
import collections
import csv
import datetime as dt
import hashlib
import os
import re
import shutil
import sys
import zipfile
from pathlib import Path

import openpyxl

PHASES = {
    "np": dict(tag="Np", fluid="oil", col="R", idx=17, hdr="Cumulative Oil Production, Np (bbl)",
               wb_rate="OIL PD RATE (bbls/d)", src_vol=(12, "Alloc. Oil Volume(bbl)"),
               src_rate=(19, "Oil PD Rate (bbls/d)")),
    "wp": dict(tag="Wp", fluid="water", col="S", idx=18, hdr="Cumulative Water Production, Wp (bbl)",
               wb_rate="WATER PD RATE (bbls/d)", src_vol=(15, "Alloc. Water Volume(bbl)"),
               src_rate=(22, "Water PD Rate(bbls/d)")),
}
PH = PHASES[sys.argv[1] if len(sys.argv) > 1 else "np"]
TAG, COL = PH["tag"], PH["col"]

BASE = Path(__file__).resolve().parent.parent
WB = BASE / "bubble_map_dulang_sample" / "bubble_map_dulang_v5_DULANG-A.xlsx"
SRC = BASE / "input_FPC_Prod Allocation_11Jun126.xlsx"
STAMP = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
BACKUP = BASE / "bubble_map_dulang_sample" / "_backup" / f"bubble_map_dulang_v5_DULANG-A_before_{TAG}_fix_{STAMP}.xlsx"
REP_CSV = BASE / "output_bubble-map" / f"v5_DULANG-A_{TAG}_fix_flags_{STAMP}.csv"
REP_MD = BASE / "output_bubble-map" / f"v5_DULANG-A_{TAG}_fix_summary_{STAMP}.md"
SHEET = "xl/worksheets/sheet1.xml"


def sha(p):
    return hashlib.sha256(Path(p).read_bytes()).hexdigest()


def isnum(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def same(a, b, tol=1e-9):
    if a in (None, "") or b in (None, ""):
        return a in (None, "") and b in (None, "")
    return abs(a - b) <= tol * max(1.0, abs(a), abs(b))


def fmt(v):
    return "" if v in (None, "") else f"{v:,.0f}"


def sheet_rows(path, data_only=False):
    ws = openpyxl.load_workbook(path, read_only=True, data_only=data_only)["Bubble Map Sample"]
    return [list(r) for r in ws.iter_rows(values_only=True)]


def load_source(entities):
    """(perforation, month) -> [(producing days, allocated volume, PD rate)]"""
    ws = openpyxl.load_workbook(SRC, read_only=True).worksheets[0]
    hdr = [list(r) for r in ws.iter_rows(min_row=1, max_row=2, values_only=True)]
    (vi, vh), (ri, rh) = PH["src_vol"], PH["src_rate"]
    assert hdr[1][3] == "Perforation" and hdr[0][8] == "Alloc. Month" and hdr[0][9] == "Producing Days" \
        and hdr[1][vi] == vh and hdr[1][ri] == rh, hdr
    src = collections.defaultdict(list)
    for r in ws.iter_rows(min_row=3, values_only=True):
        if r[3] in entities:
            src[(r[3], r[8])].append((r[9], r[vi], r[ri]))
    return src


def main():
    h_src = sha(SRC)
    f = sheet_rows(WB)
    old_v = sheet_rows(WB, data_only=True)
    hdr = f[0]
    ix = {h.strip(): i for i, h in enumerate(hdr) if h}
    E, G, K, O, X = (ix[h] for h in ("ENTITY", "ALLOC. DATE", "PRODUCTION DAYS (PD)", PH["wb_rate"], PH["hdr"]))
    assert (E, X) == (0, PH["idx"]), (E, X)  # A and the target column as written in the formula

    # entity groups, order and duplicate checks
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
    out_of_order = [e for e, ns in groups.items() if any(f[b - 1][G] <= f[a - 1][G] for a, b in zip(ns, ns[1:]))]
    src = load_source(set(groups))
    src_dups = {k: v for k, v in src.items() if len(v) > 1}
    assert not src_dups, list(src_dups)[:5]

    # new cumulative per entity (a duplicate entity/month would be counted once - there are none, checked above)
    new, vols, flags, per_ent = {}, {}, [], {}
    issues = collections.Counter()
    pdrate_dev = 0.0
    rate_lbl = f"{PH['fluid'].upper()} PD RATE"
    for e, ns in groups.items():
        cum, started, incomplete_from, nmiss = None, False, None, collections.Counter()
        seen = set()
        for n in ns:
            r = f[n - 1]
            s = src.get((e, r[G]))
            assert s, ("no source record", e, r[G])
            pd, vol, rate = s[0]
            assert (pd == r[K] or same(pd, r[K])) and (rate == r[O] or same(rate, r[O])), (n, pd, r[K], rate, r[O])
            flag = []
            if (e, r[G]) in seen:
                vol_use, flag = None, ["DUPLICATE ENTITY/MONTH - not counted again"]
            elif vol is not None:
                vol_use = float(vol)
                if isnum(r[K]) and isnum(r[O]) and r[K] > 0:
                    pdrate_dev = max(pdrate_dev, abs(r[K] * r[O] - vol) / max(1.0, abs(vol)))
                if vol != 0 and not (isnum(r[K]) and r[K] > 0):
                    flag.append("ALLOCATED VOLUME WITH 0 PRODUCING DAYS AND NO RATE - source volume used")
                    issues["allocated volume with 0 producing days"] += 1
                elif vol != 0 and r[O] is None:
                    flag.append(f"ALLOCATED VOLUME WITHOUT {rate_lbl} - source volume used")
                    issues[f"allocated volume without {rate_lbl.lower()}"] += 1
            elif isnum(r[K]) and r[K] == 0:
                vol_use = 0.0
                flag.append("PRODUCING DAYS 0, SOURCE VOLUME BLANK - treated as valid zero")
                issues["producing days 0 with blank volume (zero)"] += 1
            else:
                vol_use = None
                kind = "PRODUCING DAYS BLANK" if r[K] is None else f"{rate_lbl} / VOLUME BLANK (PD > 0)"
                if started:
                    nmiss[kind] += 1
                    incomplete_from = incomplete_from or r[G]
                    flag.append(f"MISSING INPUT: {kind}; {TAG} carried forward, incomplete")
                else:
                    nmiss["before first volume: " + kind] += 1
                    flag.append(f"MISSING INPUT before first available volume: {kind}; {TAG} blank")
            seen.add((e, r[G]))
            if vol_use is not None:
                cum = (cum or 0.0) + vol_use
                started = True
                if incomplete_from:
                    flag.append("INCOMPLETE: an earlier month of this entity has missing input")
            new[n], vols[n] = cum, vol_use
            if flag:
                flags.append({"row": n, "entity": e, "month": r[G].date().isoformat(), "producing_days": r[K],
                              f"{PH['fluid']}_pd_rate": r[O], f"alloc_{PH['fluid']}_volume": vol,
                              f"{TAG.lower()}_new": cum, "flag": "; ".join(flag)})
        per_ent[e] = dict(rows=len(ns), first=f[ns[0] - 1][G], last=f[ns[-1] - 1][G], final=cum,
                          old_final=old_v[ns[-1] - 1][X], incomplete_from=incomplete_from, missing=nmiss,
                          changed=sum(1 for n in ns if not same(old_v[n - 1][X], new[n])))
    assert pdrate_dev < 1e-6, pdrate_dev

    # backup, then write
    BACKUP.parent.mkdir(exist_ok=True)
    shutil.copy2(WB, BACKUP)
    assert sha(BACKUP) == sha(WB)
    zin = zipfile.ZipFile(BACKUP)
    xml = zin.read(SHEET)
    has_chain = "xl/calcChain.xml" in zin.namelist()

    def cell(m):
        n, attrs = int(m.group(1)), m.group(2).decode()
        if n not in new:
            return m.group(0)
        s = re.search(r'\s+s="\d+"', attrs)
        s = s.group(0) if s else ""
        p = n - 1
        fx = (f'IF(A{n}=A{p},{COL}{p},"")' if vols[n] is None
              else f'N(IF(A{n}=A{p},{COL}{p},""))+{repr(vols[n])}')
        if new[n] is None:
            return f'<c r="{COL}{n}"{s} t="str"><f>{fx}</f><v></v></c>'.encode()
        return f'<c r="{COL}{n}"{s}><f>{fx}</f><v>{repr(new[n])}</v></c>'.encode()

    new_xml = re.sub(rb'<c r="%s(\d+)"((?:\s+[\w:]+="[^"]*")*)\s*(?:/>|>.*?</c>)' % COL.encode(), cell, xml,
                     flags=re.S)
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
    for n, (a, b, av, bv) in enumerate(zip(f, tf, old_v, tv), 1):
        for c in range(max(len(a), len(b))):
            x, y = (a[c] if c < len(a) else None), (b[c] if c < len(b) else None)
            if c == X and n in new:
                assert y.startswith(("=N(IF(A", "=IF(A")) and f"{COL}{n - 1}," in y, (n, y)
                assert same(bv[c], new[n], 0), (n, bv[c], new[n])
            else:
                assert x == y, (n, c, x, y)
                if c < len(av):
                    assert av[c] == bv[c], (n, c)
    # independent check: emulate each formula row by row from the sheet itself (previous row = same ENTITY only)
    for n in new:
        a, p = tf[n - 1], tf[n - 2]
        prev = tv[n - 2][X] if a[E] == p[E] else ""
        m = re.fullmatch(r'=N\(IF\(A\d+=A\d+,%s\d+,""\)\)\+(.+)' % COL, a[X])
        exp = ((prev or 0) + float(m.group(1))) if m else prev
        assert same(exp if exp != "" else None, tv[n - 1][X] if tv[n - 1][X] != "" else None, 0), n
    # each entity's final value = sum of its own source volumes only (plus the 0-PD zero months)
    for e, ns in groups.items():
        own = [src[(e, f[n - 1][G])][0][1] for n in ns]
        tot = sum(v for v in own if v is not None) if any(v is not None for v in own) else None
        if tot is None and per_ent[e]["final"] == 0.0:
            continue
        assert same(tot, per_ent[e]["final"], 1e-12), (e, tot, per_ent[e]["final"])
    assert list(openpyxl.load_workbook(BACKUP, read_only=True)["Mapping_Review"].iter_rows(values_only=True)) == \
        list(openpyxl.load_workbook(tmp, read_only=True)["Mapping_Review"].iter_rows(values_only=True))
    zt.close()
    os.replace(tmp, WB)
    assert sha(SRC) == h_src

    # ---- reports
    with open(REP_CSV, "x", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, list(flags[0]))
        w.writeheader()
        w.writerows(flags)
    corrected = [e for e, d in per_ent.items() if d["changed"]]
    incompl = [e for e, d in per_ent.items() if d["incomplete_from"]]
    lines = [f"# v5_DULANG-A cumulative {TAG} correction ({STAMP})", "",
             f"- Workbook: `{WB.relative_to(BASE)}` (column {COL} only); backup: `{BACKUP.relative_to(BASE)}`",
             f"- {TAG} per ENTITY (perforation = string + reservoir unit, the allocation level) = previous {TAG} of the "
             f"same ENTITY + monthly '{PH['src_vol'][1]}' from `{SRC.name}`; that volume equals Producing Days x "
             f"{PH['fluid']} PD rate wherever both exist (max relative difference {pdrate_dev:.1e}).",
             f"- Entities: {len(groups)}; corrected (any value changed): {len(corrected)}; "
             f"incomplete (missing month after first volume): {len(incompl)}",
             f"- Duplicate ENTITY/month records: {len(dups) or 'none'} in workbook, {len(src_dups) or 'none'} in "
             f"source; entities out of month order: {out_of_order or 'none'}",
             f"- Allocation issues: {dict(issues) or 'none'}", f"- Row flags: `{REP_CSV.relative_to(BASE)}`", "",
             f"| ENTITY | Rows | First month | Last month | {TAG} old (last row) | {TAG} new (last row) | Rows changed | "
             "Incomplete from | Missing inputs |", "|---|---|---|---|---|---|---|---|---|"]
    for e, d in per_ent.items():
        lines.append(f"| {e} | {d['rows']} | {d['first']:%Y-%m} | {d['last']:%Y-%m} | {fmt(d['old_final'])} | "
                     f"{fmt(d['final'])} | {d['changed']} | "
                     f"{d['incomplete_from'].strftime('%Y-%m') if d['incomplete_from'] else ''} | "
                     f"{'; '.join(f'{k}: {v}' for k, v in d['missing'].items())} |")
    with open(REP_MD, "x", encoding="utf-8") as fh:
        fh.write("\n".join(lines) + "\n")

    tot = collections.Counter()
    for d in per_ent.values():
        tot.update(d["missing"])
    print(f"entities {len(groups)} corrected {len(corrected)} unchanged {len(groups) - len(corrected)} "
          f"incomplete {len(incompl)} duplicates {len(dups)} out_of_order {len(out_of_order)}")
    print("allocation issues:", dict(issues))
    print("missing inputs:", dict(tot))
    print("unchanged entities:", [e for e in groups if e not in corrected])
    print("incomplete:", [(e, per_ent[e]["incomplete_from"].strftime("%Y-%m")) for e in incompl])
    vkey = f"alloc_{PH['fluid']}_volume"
    zero_pd = collections.Counter()
    for x in flags:
        if x["flag"].startswith("ALLOCATED VOLUME"):
            zero_pd[x["entity"]] += float(x[vkey])
    print("flagged allocated volume by entity:", {e: round(v, 1) for e, v in zero_pd.most_common()})
    print(f"backup {BACKUP}\nreport {REP_MD}\nflags {REP_CSV} ({len(flags)} rows)")


if __name__ == "__main__":
    main()
