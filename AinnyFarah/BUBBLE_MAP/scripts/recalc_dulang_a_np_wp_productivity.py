#!/usr/bin/env python3
"""bubble_map_DULANG-A.xlsx -> bubble_map_DULANG-A_recalculated.xlsx: rewrite ONLY columns R (Np), S (Wp), T (Productivity)
as live Excel formulas referencing the sheet's own cells.

Per ENTITY (rows contiguous, chronological - checked), row n, previous row p, entity first row s:
  first row of the entity, or no Producing Days recorded yet in the entity (history not started):
      R = IF(ISNUMBER(Kn),IF(Kn=0,0,IF(ISNUMBER(On),On*Kn,"")),"")
  later rows:
      R = IF(COUNT($K$s:Kp)=0, <first-row formula>,
             IF(Rp="","", IF(ISNUMBER(Kn), IF(Kn=0,Rp, IF(ISNUMBER(On),Rp+(On*Kn),"")), "")))
  i.e. PD = 0 keeps the previous total, rate = 0 adds 0, missing PD or missing rate with PD > 0 blanks this and every
  later total of the entity (no restart after a gap). Opening cumulative = 0 (none documented). S: same with P.
  T = Rn/Ln (direct, no IF/IFERROR).
Separator rows stay empty. Column L is checked (not changed). Cached values are computed by evaluating these formulas;
the workbook is set to fully recalculate when opened. The input file is not modified.
"""
import collections
import hashlib
import os
import re
import sys
import zipfile
from pathlib import Path

import openpyxl

BASE = Path(__file__).resolve().parent.parent
SRC = BASE / "bubble_map_dulang_sample" / "bubble_map_DULANG-A.xlsx"
OUT = SRC.with_name(SRC.stem + "_recalculated.xlsx")
REP = BASE / "output_bubble-map" / "bubble_map_DULANG-A_recalculated_summary.md"
SHEET = "xl/worksheets/sheet1.xml"
HDR = {"A": "ENTITY", "G": "ALLOC. DATE", "K": "PRODUCTION DAYS (PD)", "L": "CUMULATIVE PD",
       "O": "OIL PD RATE (bbls/d)", "P": "WATER PD RATE (bbls/d)", "R": "Cumulative Oil Production, Np (bbl)",
       "S": "Cumulative Water Production, Wp (bbl)", "T": "Productivity (bbl/day)"}
IDX = {c: openpyxl.utils.column_index_from_string(c) - 1 for c in HDR}


def sha(p):
    return hashlib.sha256(Path(p).read_bytes()).hexdigest()


def isnum(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def rows(path, data_only=False):
    return [list(r) for r in openpyxl.load_workbook(path, read_only=True, data_only=data_only)["Bubble Map Sample"]
            .iter_rows(values_only=True)]


def first_fx(n, rate):
    return f'IF(ISNUMBER(K{n}),IF(K{n}=0,0,IF(ISNUMBER({rate}{n}),{rate}{n}*K{n},"")),"")'


def cum_fx(col, rate, n, s):
    if n == s:
        return first_fx(n, rate)
    p = n - 1
    return (f'IF(COUNT($K${s}:K{p})=0,{first_fx(n, rate)},IF({col}{p}="","",IF(ISNUMBER(K{n}),IF(K{n}=0,{col}{p},'
            f'IF(ISNUMBER({rate}{n}),{col}{p}+({rate}{n}*K{n}),"")),"")))')


def eval_cum(k, rate, prev, started):
    """Excel semantics of cum_fx; prev is the previous row's value ("" when blank), started = COUNT($K$s:Kp)>0"""
    if not started:
        if not isnum(k):
            return ""
        return 0 if k == 0 else (rate * k if isnum(rate) else "")
    if prev == "":
        return ""
    if not isnum(k):
        return ""
    return prev if k == 0 else (prev + rate * k if isnum(rate) else "")


def main():
    if OUT.exists():
        sys.exit(f"STOP: {OUT.name} exists - not overwritten")
    h0 = sha(SRC)
    f, fv = rows(SRC), rows(SRC, data_only=True)
    for c, h in HDR.items():
        assert (f[0][IDX[c]] or "").strip() == h, (c, f[0][IDX[c]], h)
    A, G, K, L, O, P = (IDX[c] for c in "AGKLOP")

    groups, last, dup = collections.OrderedDict(), None, collections.Counter()
    for n, r in enumerate(f[1:], 2):
        if r[A] is None:
            assert all(v is None for v in r), n
            last = None
            continue
        dup[(r[A], r[G])] += 1
        if r[A] != last:
            assert r[A] not in groups, f"{r[A]} not contiguous"
            groups[r[A]] = []
        groups[r[A]].append(n)
        last = r[A]
    dups = [k for k, c in dup.items() if c > 1]
    unsorted = [e for e, ns in groups.items() if any(f[b - 1][G] <= f[a - 1][G] for a, b in zip(ns, ns[1:]))]
    assert not unsorted, unsorted[:5]  # formulas use the previous row, so rows must already be chronological
    for ns in groups.values():
        assert ns == list(range(ns[0], ns[-1] + 1))

    # formulas + cached values (evaluated with Excel semantics)
    out = {}  # row -> {col: (formula, value)}
    for e, ns in groups.items():
        s = ns[0]
        prev = {"R": "", "S": ""}
        for n in ns:
            r = fv[n - 1]
            started = n > s and any(isnum(fv[m - 1][K]) for m in range(s, n))
            cells = {}
            for col, rate in (("R", O), ("S", P)):
                v = eval_cum(r[K], r[rate], prev[col], started)
                cells[col] = (cum_fx(col, "O" if col == "R" else "P", n, s), v)
                prev[col] = v
            lv = r[L]
            npv = cells["R"][1]
            t = "#VALUE!" if npv == "" else ("#DIV/0!" if lv == 0 else npv / lv)
            cells["T"] = (f"R{n}/L{n}", t)
            out[n] = cells

    # independent check of the cumulative logic: explicit per-entity state machine
    issues = collections.defaultdict(collections.Counter)
    for e, ns in groups.items():
        for col, rate in (("R", O), ("S", P)):
            state, tot = "not started", 0.0
            for n in ns:
                k, q = fv[n - 1][K], fv[n - 1][rate]
                if state == "not started":
                    if isnum(k) and (k == 0 or isnum(q)):
                        state, tot = "running", (0.0 if k == 0 else q * k)
                    elif isnum(k):
                        state = "broken"
                        issues[col]["missing rate on first record"] += 1
                    else:
                        issues[col]["no producing days before history starts"] += 1
                elif state == "running":
                    if not isnum(k):
                        state = "broken"
                        issues[col]["gap: producing days missing"] += 1
                    elif k > 0 and not isnum(q):
                        state = "broken"
                        issues[col]["gap: rate missing with PD > 0"] += 1
                    else:
                        tot += 0 if k == 0 else q * k
                exp = tot if state == "running" else ""
                got = out[n][col][1]
                assert (got == "" and exp == "") or (got != "" and exp != "" and abs(got - exp) <= 1e-9 * max(1, exp)), \
                    (col, e, n, got, exp)
            issues[col]["entities blank from a gap onward" if state == "broken" else "entities complete"] += 1

    # column L check: per-entity running sum of N(K), formula pointing to the previous row of the same entity
    l_bad = []
    for e, ns in groups.items():
        run = 0.0
        for i, n in enumerate(ns):
            run += fv[n - 1][K] if isnum(fv[n - 1][K]) else 0
            fx = f[n - 1][L] or ""
            m = re.fullmatch(r"=IF\(AND\(B(\d+)=B(\d+),D\1=D\2,A\1=A\2\),L\2,0\)\+N\(K\1\)", fx)
            ok_fx = bool(m) and int(m.group(1)) == n and int(m.group(2)) == n - 1
            if not ok_fx or not isnum(fv[n - 1][L]) or abs(fv[n - 1][L] - run) > 1e-9:
                l_bad.append((n, e, fx, fv[n - 1][L], run))

    # write: replace R/S/T cells of data rows only
    zin = zipfile.ZipFile(SRC)
    xml = zin.read(SHEET)
    rst_si = set(re.findall(rb'<c r="[RST]\d+"[^>]*><f t="shared"[^>]*si="(\d+)"', xml))
    other_si = set(re.findall(rb'<c r="(?![RST]\d)[A-Z]+\d+"[^>]*><f t="shared"[^>]*si="(\d+)"', xml))
    assert not rst_si & other_si, "shared formula spans R/S/T and another column"

    def cell(m):
        c, n, attrs = m.group(1).decode(), int(m.group(2)), m.group(3).decode()
        if n not in out:
            return m.group(0)
        s = re.search(r'\s+s="\d+"', attrs)
        s = s.group(0) if s else ""
        fx, v = out[n][c]
        fx = fx.replace("<", "&lt;").replace(">", "&gt;")
        if v == "":
            return f'<c r="{c}{n}"{s} t="str"><f>{fx}</f><v></v></c>'.encode()
        if isinstance(v, str):
            return f'<c r="{c}{n}"{s} t="e"><f>{fx}</f><v>{v}</v></c>'.encode()
        return f'<c r="{c}{n}"{s}><f>{fx}</f><v>{repr(float(v))}</v></c>'.encode()

    new_xml = re.sub(rb'<c r="([RST])(\d+)"((?:\s+[\w:]+="[^"]*")*)\s*(?:/>|>.*?</c>)', cell, xml, flags=re.S)
    # data rows missing an R/S/T cell entirely: insert it in column order
    def add_missing(m):
        n, body = int(m.group(1)), m.group(3)
        if n not in out:
            return m.group(0)
        for c in "RST":
            if f'<c r="{c}{n}"'.encode() in body:
                continue
            new = cell(re.match(rb'<c r="([RST])(\d+)"((?:\s+[\w:]+="[^"]*")*)\s*(/>)', f'<c r="{c}{n}"/>'.encode()))
            later = [body.find(f'<c r="{x}{n}"'.encode()) for x in ("U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC")]
            later = [i for i in later if i >= 0]
            at = min(later) if later else len(body)
            body = body[:at] + new + body[at:]
        return m.group(0)[:m.start(3) - m.start(0)] + body + b"</row>"
    new_xml = re.sub(rb'<row r="(\d+)"([^>]*)>(.*?)</row>', add_missing, new_xml, flags=re.S)

    tmp = OUT.with_name(OUT.stem + ".partial.xlsx")
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
            elif fn == "xl/_rels/workbook.xml.rels":
                data = re.sub(rb'<Relationship Id="rId\d+" Type="[^"]+/calcChain" Target="calcChain.xml"/>', b"", data)
            elif fn == "[Content_Types].xml":
                data = re.sub(rb'<Override PartName="/xl/calcChain.xml"[^>]*/>', b"", data)
            zi = zipfile.ZipInfo(fn, info.date_time)
            zi.compress_type, zi.external_attr = zipfile.ZIP_DEFLATED, info.external_attr
            zout.writestr(zi, data)

    # ---- verify the saved file
    tf, tv = rows(tmp), rows(tmp, data_only=True)
    assert len(tf) == len(f)
    R, S, T = IDX["R"], IDX["S"], IDX["T"]
    nform = collections.Counter()
    for n, (a, b, av, bv) in enumerate(zip(f, tf, fv, tv), 1):
        for c in range(max(len(a), len(b))):
            x, y = (a[c] if c < len(a) else None), (b[c] if c < len(b) else None)
            if c in (R, S, T) and n in out:
                col = "RST"[(R, S, T).index(c)]
                fx, v = out[n][col]
                assert y == "=" + fx, (n, col, y)
                assert re.search(r"[A-Z]\d+", y)  # refers to cells
                got = bv[c]
                assert (got in ("", None) and v == "") or got == v or (isnum(v) and abs(got - v) <= 1e-12 * max(1, abs(v))), \
                    (n, col, got, v)
                nform[col] += 1
            elif c in (R, S, T) and n > 1 and n not in out:
                assert y is None and bv[c] is None, ("separator not blank", n, c)
            else:
                assert x == y, (n, c, x, y)
                if c < len(av) and c not in (R, S, T, 25):  # Z (NPRADIUS) recalculates from R in Excel
                    assert av[c] == bv[c], (n, c)
    for col in "RST":
        assert nform[col] == len(out), (col, nform[col])
    assert list(openpyxl.load_workbook(SRC, read_only=True)["Mapping_Review"].iter_rows(values_only=True)) == \
        list(openpyxl.load_workbook(tmp, read_only=True)["Mapping_Review"].iter_rows(values_only=True))
    os.replace(tmp, OUT)
    assert sha(SRC) == h0

    # ---- summary
    t_err = collections.Counter(c["T"][1] for c in out.values() if isinstance(c["T"][1], str))
    seps = sum(1 for n, r in enumerate(f[1:], 2) if r[A] is None)
    lines = ["# bubble_map_DULANG-A_recalculated", "", f"- Source (unchanged): `{SRC.name}`; output: `{OUT.name}`",
             f"- Entities {len(groups)}, data rows {len(out)}, separator rows {seps} (R/S/T left empty)",
             f"- Formulas written: R {nform['R']}, S {nform['S']}, T {nform['T']}",
             f"- Duplicate entity/month records: {len(dups) or 'none'}; rows out of month order: none",
             "- Opening cumulative volumes: none documented - every entity starts from 0 at its first record, so "
             "totals cover only the history in the sheet.", "",
             "## Np (R)", *[f"- {k}: {v}" for k, v in issues["R"].items()], "",
             "## Wp (S)", *[f"- {k}: {v}" for k, v in issues["S"].items()], "",
             "## Productivity (T = R/L)", f"- error results: {dict(t_err)} (#VALUE! where Np is blank, #DIV/0! where "
             "cumulative PD = 0)", "",
             "## Column L check", f"- rows with a discrepancy: {len(l_bad)}"]
    lines += [f"  - row {n} {e}: formula `{fx}`, value {v}, expected {run}" for n, e, fx, v, run in l_bad[:50]]
    REP.parent.mkdir(exist_ok=True)
    REP.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print("\n".join(lines))
    print(f"\nsaved {OUT}\nreport {REP}")


if __name__ == "__main__":
    main()
