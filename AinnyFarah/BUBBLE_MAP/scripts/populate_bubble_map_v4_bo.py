#!/usr/bin/env python3
"""Create bubble_map_dulang_v4.xlsx as a copy of v2 and fill the 'Bo (rb/stb)' column (Z) from
'input_Bo value per reservoir.xlsx', matched on the reservoir name in 'RESERVOIR UNIT' (E).

Matching (agreed with the user): template names carry a fault-block prefix ('S1-E7', 'N3A-E14B'); only that prefix
(text before the first '-') is removed, and the rest must equal the input name exactly, ignoring leading/trailing
spaces and case. No other equivalences (E14A is not E14). Input names with conflicting Bo values, missing Bo, or no
match leave the template cell as it is.

The copy is made at package level as in the v3 script: every zip part except the 'Bubble Map Sample' sheet XML is
copied byte for byte, and in that sheet the only edit is a new Z cell in matched rows. Inputs are read only; an
existing v4 is never overwritten.
"""
import collections
import hashlib
import re
import sys
import zipfile
from pathlib import Path

import openpyxl

BASE = Path(__file__).resolve().parent.parent
SRC = BASE / "input_Bo value per reservoir.xlsx"
TPL = BASE / "bubble_map_dulang_sample" / "bubble_map_dulang_v2.xlsx"
OUT = BASE / "bubble_map_dulang_sample" / "bubble_map_dulang_v4.xlsx"
REP_MD = BASE / "output_bubble-map" / "bubble_map_dulang_v4_bo_summary.md"
SHEET_PART = "xl/worksheets/sheet1.xml"


def sha(p):
    return hashlib.sha256(Path(p).read_bytes()).hexdigest()


def key(s):
    return str(s).strip().upper()


def strip_block(ru):
    """'S1-E7' -> 'E7'; None when there is no block prefix"""
    m = re.fullmatch(r"\s*([A-Za-z0-9]+)-(.+?)\s*", str(ru or ""))
    return m.group(2) if m else None


def load_input():
    ws = openpyxl.load_workbook(SRC, read_only=True, data_only=True).worksheets[0]
    it = ws.iter_rows(values_only=True)
    hdr = next(it)
    assert [key(h) for h in hdr[:2]] == ["RESERVOIR", "BO VALUE"], hdr
    recs = collections.defaultdict(list)
    for n, r in enumerate(it, start=2):
        if r[0] is None or not str(r[0]).strip():
            continue
        recs[key(r[0])].append((n, str(r[0]), r[1]))
    bo, missing, conflict = {}, {}, {}
    for k, lst in recs.items():
        vals = {v for _, _, v in lst if isinstance(v, (int, float)) and not isinstance(v, bool) and v > 0}
        bad = [v for _, _, v in lst if v is not None and v not in vals]
        if len(vals) == 1 and not bad:
            bo[k] = vals.pop()
        elif not vals and not bad:
            missing[k] = lst
        else:
            conflict[k] = lst
    return bo, missing, conflict, recs


def edit_sheet(xml, fills):
    out, pos = [], 0
    for m in re.finditer(rb'<row r="(\d+)"[^>]*>(.*?)</row>', xml, re.S):
        n = int(m.group(1))
        if n not in fills:
            continue
        body = m.group(2)
        assert b'<c r="Z%d"' % n not in body, n
        cell = b'<c r="Z%d"><v>%s</v></c>' % (n, repr(fills[n]).encode())
        at = min([i for i in (body.find(b'<c r="%s%d"' % (c, n)) for c in (b"AA", b"AB", b"AC")) if i >= 0],
                 default=len(body))
        out += [xml[pos:m.start(2)], body[:at], cell, body[at:]]
        pos = m.end(2)
    out.append(xml[pos:])
    return b"".join(out)


def revert_sheet(xml):
    return re.sub(rb'<c r="Z\d+"><v>[^<]*</v></c>', b"", xml)


def main():
    if OUT.exists():
        sys.exit(f"STOP: {OUT} exists - not overwritten.")
    h0 = {p: sha(p) for p in (SRC, TPL)}
    bo, missing, conflict, recs = load_input()

    ws = openpyxl.load_workbook(TPL, read_only=True)["Bubble Map Sample"]
    it = ws.iter_rows(values_only=True)
    hdr = next(it)
    assert hdr[4] == "RESERVOIR UNIT" and hdr[25] == "Bo (rb/stb)", (hdr[4], hdr[25])
    fills, kept = {}, 0
    per_ru = collections.defaultdict(lambda: [0, None, ""])  # rows, Bo, status
    n_rows = 0
    for n, r in enumerate(it, start=2):
        n_rows += 1
        ru, cur = r[4], r[25]
        name = strip_block(ru)
        k = key(name) if name else None
        d = per_ru[ru]
        d[0] += 1
        if k in bo:
            d[1], d[2] = bo[k], "populated"
            if cur is None or cur == "" or cur == bo[k]:
                if cur is None or cur == "":
                    fills[n] = bo[k]
            else:
                kept += 1
                d[2] = "template value kept (differs from input)"
        elif k in conflict:
            d[2] = "conflicting input Bo - template value kept"
        elif k in missing:
            d[2] = "input Bo missing - template value kept"
        else:
            d[2] = "no matching input reservoir" if name else "no block prefix - not matched"

    zin = zipfile.ZipFile(TPL)
    assert b'name="Bubble Map Sample" sheetId="1" r:id="rId1"' in zin.read("xl/workbook.xml")
    assert b'worksheets/sheet1.xml' in zin.read("xl/_rels/workbook.xml.rels")
    orig = zin.read(SHEET_PART)
    new = edit_sheet(orig, fills)
    assert revert_sheet(new) == orig, "sheet edit is not limited to new Z cells"
    tmp = OUT.with_name(OUT.stem + ".partial.xlsx")
    with open(tmp, "xb") as fh, zipfile.ZipFile(fh, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            zi = zipfile.ZipInfo(info.filename, info.date_time)
            zi.compress_type, zi.external_attr = zipfile.ZIP_DEFLATED, info.external_attr
            zout.writestr(zi, new if info.filename == SHEET_PART else zin.read(info.filename))
    if OUT.exists():
        tmp.unlink()
        sys.exit("STOP: v4 appeared during the run - not overwritten.")
    tmp.rename(OUT)

    # ---- validation (re-read the saved file and compare every cell with v2)
    zo = zipfile.ZipFile(OUT)
    assert [i.filename for i in zo.infolist()] == [i.filename for i in zin.infolist()]
    for i in zin.infolist():
        if i.filename != SHEET_PART:
            assert zo.read(i.filename) == zin.read(i.filename), i.filename
    assert revert_sheet(zo.read(SHEET_PART)) == orig
    a = openpyxl.load_workbook(TPL, read_only=True)
    b = openpyxl.load_workbook(OUT, read_only=True)
    assert a.sheetnames == b.sheetnames
    diffs = 0
    for s in a.sheetnames:
        for n, (ra, rb) in enumerate(zip(a[s].iter_rows(values_only=True), b[s].iter_rows(values_only=True)), 1):
            for c, (va, vb) in enumerate(zip(ra, rb)):
                if va != vb:
                    assert s == "Bubble Map Sample" and c == 25 and va is None and vb == fills[n], (s, n, c, va, vb)
                    diffs += 1
    assert diffs == len(fills), (diffs, len(fills))
    assert {p: sha(p) for p in (SRC, TPL)} == h0, "input changed"

    # ---- report (separate file; nothing added to the workbook)
    used = {key(strip_block(ru)) for ru in per_ru if strip_block(ru)}
    lines = ["# bubble_map_dulang_v4 - Bo (rb/stb) population", "",
             f"- Template: `{TPL.relative_to(BASE)}` (unchanged)", f"- Input: `{SRC.relative_to(BASE)}` (unchanged)",
             f"- Output: `{OUT.relative_to(BASE)}`",
             "- Match: RESERVOIR UNIT (col E) with the fault-block prefix removed ('S1-E7' -> 'E7'), exact name, "
             "case/space-insensitive. Written to Bo (rb/stb) (col Z).", "",
             f"- Template data rows: {n_rows}", f"- Rows populated: {len(fills)}",
             f"- Rows with an existing template Bo that differs from input (kept): {kept}",
             f"- Cells differing from v2: {diffs} (all in col Z, all previously blank)", "",
             "## Template reservoir units", "", "| Reservoir unit | Rows | Bo | Status |", "|---|---|---|---|"]
    for ru, (cnt, v, st) in sorted(per_ru.items(), key=lambda x: str(x[0])):
        lines.append(f"| {ru} | {cnt} | {'' if v is None else v} | {st} |")
    lines += ["", "## Input records", "",
              "- Missing Bo: " + (", ".join(f"{lst[0][1]} (row {lst[0][0]})" for lst in missing.values()) or "none"),
              "- Conflicting Bo: " + ("; ".join(f"{k}: " + ", ".join(f"row {n}={v}" for n, _, v in lst)
                                                for k, lst in conflict.items()) or "none"),
              "- Input reservoirs not used by any template row: "
              + (", ".join(lst[0][1] for k, lst in recs.items() if k not in used) or "none"), ""]
    REP_MD.parent.mkdir(exist_ok=True)
    with open(REP_MD, "x", encoding="utf-8") as fh:
        fh.write("\n".join(lines))

    st = collections.Counter()
    for cnt, _, s in per_ru.values():
        st[s] += cnt
    print(f"v4 written: {OUT}\nrows={n_rows} populated={len(fills)} diffs={diffs} kept={kept}")
    for s, c in st.items():
        print(f"  {s}: {c} rows, {sum(1 for d in per_ru.values() if d[2] == s)} units")
    for ru, (cnt, v, s) in sorted(per_ru.items(), key=lambda x: str(x[0])):
        if s != "populated":
            print(f"  - {ru} ({cnt} rows): {s}")
    print("missing:", list(missing), "conflict:", list(conflict))


if __name__ == "__main__":
    main()
