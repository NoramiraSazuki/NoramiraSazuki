#!/usr/bin/env python3
"""Split bubble_map_dulang_v4.xlsx into one workbook per platform (DULG-A/B/C/D, from WELL NAME in column C).

Each output is a package-level copy of v4: every part is copied byte for byte except
  - 'Bubble Map Sample' sheet XML: header row + that platform's rows only, renumbered from row 2. Row references in
    formulas are shifted the way Excel shifts them when the rows above are deleted; the one reference that points
    above the block (the previous-row test in the cumulative-producing-days formula of the first row) is pointed at
    the header row, exactly as row 2 does in v4. Shared formulas whose master falls outside the block are written
    out explicitly. The cumulative Np/Wp SUMPRODUCTs are anchored at $2 in v4 (running total over the whole sheet);
    they stay anchored at row 2, so in each file they total that platform's rows, and their cached values are
    recomputed accordingly. All other cached values are kept (they depend only on rows within the platform block).
  - dimension / autoFilter / _FilterDatabase ranges resized;
  - calcChain.xml dropped (Excel rebuilds it).
The 'Mapping_Review' sheet, styles, column widths, freeze panes etc. are unchanged. v4 is read only; existing
outputs are never overwritten.
"""
import hashlib
import re
import sys
import zipfile
from pathlib import Path

import openpyxl

BASE = Path(__file__).resolve().parent.parent
SRC = BASE / "bubble_map_dulang_sample" / "bubble_map_dulang_v4.xlsx"
OUT_DIR = BASE / "bubble_map_dulang_sample"
SHEET = "xl/worksheets/sheet1.xml"
REF = re.compile(r"(?<![A-Za-z_\d.$])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\d(A-Za-z_])")
ROW = re.compile(rb'<row r="(\d+)"[^>]*>.*?</row>', re.S)


def sha(p):
    return hashlib.sha256(Path(p).read_bytes()).hexdigest()


def shift(formula, start, off, low_refs):
    """rows >= start move up by off. Above the block: the absolute $2 anchor of the cumulative SUMPRODUCTs stays on
    the first data row (row 2), and the relative previous-row reference (start-1) maps to the header row 1"""
    def sub(m):
        r = int(m.group(4))
        if r >= start:
            r -= off
        elif r == 2 and m.group(3):
            pass
        elif r > 1:
            low_refs.append(r)
            r = 1
        return f"{m.group(1)}{m.group(2)}{m.group(3)}{r}"
    return REF.sub(sub, formula)


def cached(row, col):
    m = re.search(rb'<c r="%s\d+"[^>]*>(?:<f[^>]*/>|<f[^>]*>[^<]*</f>)?<v>([^<]*)</v>' % col, row)
    return m.group(1) if m else None


def num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def recalc_cumulative(row, sums):
    """U = SUMPRODUCT($N$2:N,$R$2:R), V = SUMPRODUCT($N$2:N,$S$2:S) over this file's rows"""
    n = num(cached(row, b"N"))
    sums[0] += n * num(cached(row, b"R"))
    sums[1] += n * num(cached(row, b"S"))
    for col, v in ((b"U", sums[0]), (b"V", sums[1])):
        row, k = re.subn(rb'(<c r="%s\d+"[^>]*><f>SUMPRODUCT\(\$N\$2:[^<]*</f><v>)[^<]*(</v>)' % col,
                         lambda m: m.group(1) + repr(v).encode() + m.group(2), row)
        assert k == 1, row[:80]
    return row


def relmove(formula, d):
    """translate a shared master formula d rows down (relative row refs only)"""
    return REF.sub(lambda m: m.group(0) if m.group(3) else f"{m.group(1)}{m.group(2)}{int(m.group(4)) + d}", formula)


def platform_blocks():
    ws = openpyxl.load_workbook(SRC, read_only=True)["Bubble Map Sample"]
    blocks, prev = {}, None
    for n, (c,) in enumerate(ws.iter_rows(min_row=2, min_col=3, max_col=3, values_only=True), start=2):
        p = str(c)[:6]
        assert re.fullmatch(r"DULG-[A-D]", p), (n, c)
        if p != prev:
            assert p not in blocks, f"{p} rows are not contiguous"
            blocks[p] = [n, n]
            prev = p
        blocks[p][1] = n
    return blocks


def masters(xml):
    """si -> (col, row, formula text, ref) for every shared-formula master"""
    out = {}
    for m in re.finditer(rb'<c r="([A-Z]+)(\d+)"[^>]*><f t="shared" ref="([^"]+)" si="(\d+)">([^<]*)</f>', xml):
        out[m.group(4)] = (m.group(1), int(m.group(2)), m.group(5).decode(), m.group(3).decode())
    return out


def build_sheet(xml, start, end, mst):
    off = start - 2
    head, rest = xml.split(b"<sheetData>", 1)
    body, tail = rest.split(b"</sheetData>", 1)
    rows = list(ROW.finditer(body))
    assert int(rows[0].group(1)) == 1
    low = []
    out = [rows[0].group(0)]

    def fcell(m):
        # cell refs are already renumbered: nr is the new row, r the v4 row
        col, nr, attrs, ftag, ftxt = m.group(1).decode(), int(m.group(2)), m.group(3), m.group(4), m.group(5)
        r = nr + off
        sm = re.fullmatch(rb'<f t="shared"(?: ref="([^"]+)")? si="(\d+)"(/?)>', ftag)
        if sm:
            mc, mr, mtxt, mref = mst[sm.group(2)]
            assert mc.decode() == col, (col, r)
            if mr < start:  # master outside the block -> explicit formula
                txt = relmove(mtxt, r - mr)
                return f'<c r="{col}{nr}"{attrs.decode()}><f>{shift(txt, start, off, low)}</f>'.encode()
            if sm.group(1):  # master in block: clip and renumber its range
                rm = re.fullmatch(r"([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?", sm.group(1).decode())
                a, b = int(rm.group(2)), int(rm.group(4) or rm.group(2))
                ref = f"{col}{a - off}" if a == min(b, end) else f"{col}{a - off}:{col}{min(b, end) - off}"
                return (f'<c r="{col}{nr}"{attrs.decode()}><f t="shared" ref="{ref}" si="{sm.group(2).decode()}">'
                        f'{shift(ftxt.decode(), start, off, low)}</f>').encode()
            return f'<c r="{col}{nr}"{attrs.decode()}><f t="shared" si="{sm.group(2).decode()}"/>'.encode()
        assert ftag == b"<f>", ftag
        return f'<c r="{col}{nr}"{attrs.decode()}><f>{shift(ftxt.decode(), start, off, low)}</f>'.encode()

    cell_f = re.compile(rb'<c r="([A-Z]+)(\d+)"([^>]*)>(<f[^>]*/>|<f[^>]*>)(?:([^<]*)</f>)?')
    sums = [0.0, 0.0]
    for m in rows[1:]:
        n = int(m.group(1))
        if not start <= n <= end:
            continue
        row = m.group(0)
        row = re.sub(rb'<c r="([A-Z]+)(\d+)"', lambda mm: b'<c r="%s%d"' % (mm.group(1), int(mm.group(2)) - off), row)
        row = cell_f.sub(lambda mm: fcell(mm), row)
        row = re.sub(rb'^<row r="\d+"', b'<row r="%d"' % (n - off), row)
        if start > 2:
            row = recalc_cumulative(row, sums)
        out.append(row)
    assert set(low) <= {start - 1}, sorted(set(low))[:10]
    last = end - off
    head = re.sub(rb'<dimension ref="A1:([A-Z]+)\d+"/>', lambda m: b'<dimension ref="A1:%s%d"/>' % (m.group(1), last),
                  head)
    tail = re.sub(rb'<autoFilter ref="A1:([A-Z]+)\d+"', lambda m: b'<autoFilter ref="A1:%s%d"' % (m.group(1), last),
                  tail)
    return head + b"<sheetData>" + b"".join(out) + b"</sheetData>" + tail, last


def main():
    h0 = sha(SRC)
    blocks = platform_blocks()
    outs = {p: OUT_DIR / f"bubble_map_dulang_v4_{p}.xlsx" for p in blocks}
    if any(o.exists() for o in outs.values()):
        sys.exit("STOP: an output already exists - not overwritten: " + ", ".join(o.name for o in outs.values()
                                                                               if o.exists()))
    zin = zipfile.ZipFile(SRC)
    xml = zin.read(SHEET)
    mst = masters(xml)
    for p, (start, end) in blocks.items():
        new_sheet, last = build_sheet(xml, start, end, mst)
        tmp = outs[p].with_name(outs[p].stem + ".partial.xlsx")
        with open(tmp, "xb") as fh, zipfile.ZipFile(fh, "w", zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                if info.filename == "xl/calcChain.xml":
                    continue
                data = zin.read(info.filename)
                if info.filename == SHEET:
                    data = new_sheet
                elif info.filename == "xl/workbook.xml":
                    data, k = re.subn(rb"('Bubble Map Sample'!\$A\$1:\$[A-Z]+\$)\d+",
                                      lambda m: m.group(1) + str(last).encode(), data)
                    assert k == 1
                elif info.filename == "xl/_rels/workbook.xml.rels":
                    data, k = re.subn(rb'<Relationship Id="rId\d+" Type="[^"]+/calcChain" Target="calcChain.xml"/>',
                                      b"", data)
                    assert k == 1
                elif info.filename == "[Content_Types].xml":
                    data, k = re.subn(rb'<Override PartName="/xl/calcChain.xml"[^>]*/>', b"", data)
                    assert k == 1
                zi = zipfile.ZipInfo(info.filename, info.date_time)
                zi.compress_type, zi.external_attr = zipfile.ZIP_DEFLATED, info.external_attr
                zout.writestr(zi, data)
        tmp.rename(outs[p])
        print(f"{p}: rows {start}-{end} ({end - start + 1}) -> {outs[p].name}")

    # ---- validation: every cell value, formula (shifted) and cached value matches v4
    src_f = openpyxl.load_workbook(SRC, read_only=True)["Bubble Map Sample"]
    src_v = openpyxl.load_workbook(SRC, read_only=True, data_only=True)["Bubble Map Sample"]
    sf = list(src_f.iter_rows(values_only=True))
    sv = list(src_v.iter_rows(values_only=True))
    total = 0
    for p, (start, end) in blocks.items():
        off = start - 2
        wf = openpyxl.load_workbook(outs[p], read_only=True)
        assert wf.sheetnames == ["Bubble Map Sample", "Mapping_Review"]
        of = list(wf["Bubble Map Sample"].iter_rows(values_only=True))
        ov = list(openpyxl.load_workbook(outs[p], read_only=True, data_only=True)["Bubble Map Sample"]
                  .iter_rows(values_only=True))
        assert len(of) == end - start + 2 and of[0] == sf[0] and ov[0] == sv[0]
        for k in range(1, len(of)):
            a, b, av, bv = sf[k + off], of[k], sv[k + off], ov[k]
            assert str(b[2]).startswith(p), (p, k)
            assert sf[0][20].startswith("Cumulative Oil") and sf[0][21].startswith("Cumulative Water")
            if start > 2:  # running totals restart at this file's row 2: equal v4 minus v4 at the row above the block
                for c in (20, 21):
                    exp = av[c] - sv[start - 2][c]
                    assert abs(bv[c] - exp) <= 1e-6 * max(1.0, abs(av[c])), (p, k, c, bv[c], exp)
                av, bv = av[:20] + av[22:], bv[:20] + bv[22:]
            assert av == bv, (p, k)
            for c, (x, y) in enumerate(zip(a, b)):
                if isinstance(x, str) and x.startswith("="):
                    assert y == shift(x, start, off, []), (p, k, c, x, y)
                else:
                    assert x == y, (p, k, c, x, y)
        assert list(wf["Mapping_Review"].iter_rows(values_only=True)) == \
            list(openpyxl.load_workbook(SRC, read_only=True)["Mapping_Review"].iter_rows(values_only=True))
        total += len(of) - 1
    assert total == len(sf) - 1, (total, len(sf) - 1)
    assert sha(SRC) == h0, "v4 changed"
    print(f"validated: {total} rows = v4 data rows; v4 unchanged")


if __name__ == "__main__":
    main()
