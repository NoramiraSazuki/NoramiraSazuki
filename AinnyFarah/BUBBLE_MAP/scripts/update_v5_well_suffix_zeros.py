#!/usr/bin/env python3
"""In-place update of the five v5 workbooks: in the WELL NAME column of 'Bubble Map Sample' only, drop leading zeros
from the numbers of S and W suffixes (DULANG-A4S01 -> DULANG-A4S1, DULANG-A1S01W01 -> DULANG-A1S1W1). Other suffixes
(H01) and the rest of the name are untouched.

Package-level edit: the new names are appended to sharedStrings and only the WELL NAME cells are re-pointed to them
(existing strings are not edited, so no other cell can change). Every other zip part is copied byte for byte. Each
result is validated against the original in memory before it replaces the file (atomic rename).
"""
import collections
import html
import os
import re
import zipfile
from pathlib import Path

import openpyxl

DIR = Path(__file__).resolve().parent.parent / "bubble_map_dulang_sample"
FILES = ["bubble_map_dulang_v5.xlsx"] + [f"bubble_map_dulang_v5_DULANG-{p}.xlsx" for p in "ABCD"]
SHEET = "xl/worksheets/sheet1.xml"
SST = "xl/sharedStrings.xml"


def fix(name):
    m = re.fullmatch(r"(DULANG-[A-D]\d+)((?:[A-Z]\d+)*)", name)
    if not m:
        return name
    return m.group(1) + re.sub(r"([SW])0+(\d)", r"\1\2", m.group(2))


def texts(sst_xml):
    return [html.unescape("".join(t.decode() for t in re.findall(rb"<t[^>]*>(.*?)</t>", it, re.S)))
            for it in re.findall(rb"<si>(.*?)</si>", sst_xml, re.S)]


def rows(path):
    return [list(r) for r in openpyxl.load_workbook(path, read_only=True)["Bubble Map Sample"].iter_rows(values_only=True)]


def process(path):
    zin = zipfile.ZipFile(path)
    sst_xml, sheet = zin.read(SST), zin.read(SHEET)
    sst = texts(sst_xml)
    hdr = re.search(rb"<row r=\"1\"[^>]*>(.*?)</row>", sheet, re.S).group(1)
    col = [c.decode() for c, v in re.findall(rb'<c r="([A-Z]+)1"[^>]*t="s"[^>]*><v>(\d+)</v>', hdr)
           if sst[int(v)] == "WELL NAME"]
    assert len(col) == 1, col
    col = col[0].encode()
    index = {t: i for i, t in enumerate(sst)}
    added, changed = {}, collections.Counter()

    def sub(m):
        old = sst[int(m.group(3))]
        new = fix(old)
        if new == old:
            return m.group(0)
        changed[(old, new)] += 1
        if new not in index:
            index[new] = len(sst) + len(added)
            added[new] = index[new]
        return m.group(1) + str(index[new]).encode() + m.group(4)

    new_sheet = re.sub(rb'(<c r="%s(?!1")\d+"[^>]*t="s"[^>]*><v>)((\d+))(</v>)' % col, sub, sheet)
    new_sst = sst_xml.replace(b"</sst>", b"".join(b"<si><t>%s</t></si>" % html.escape(t, quote=False).encode()
                                                for t in added) + b"</sst>")
    new_sst = re.sub(rb'uniqueCount="\d+"', b'uniqueCount="%d"' % (len(sst) + len(added)), new_sst, count=1)

    tmp = path.with_name(path.stem + ".partial.xlsx")
    with open(tmp, "xb") as fh, zipfile.ZipFile(fh, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = {SHEET: new_sheet, SST: new_sst}.get(info.filename) or zin.read(info.filename)
            zi = zipfile.ZipInfo(info.filename, info.date_time)
            zi.compress_type, zi.external_attr = zipfile.ZIP_DEFLATED, info.external_attr
            zout.writestr(zi, data)

    # validate before replacing: same parts, only WELL NAME cells differ, and exactly as intended
    zt = zipfile.ZipFile(tmp)
    assert [i.filename for i in zt.infolist()] == [i.filename for i in zin.infolist()]
    for i in zin.infolist():
        if i.filename not in (SHEET, SST):
            assert zt.read(i.filename) == zin.read(i.filename), i.filename
    a, b = rows(path), rows(tmp)
    wi = a[0].index("WELL NAME")
    assert len(a) == len(b)
    ncells = 0
    for ra, rb_ in zip(a, b):
        for c, (x, y) in enumerate(zip(ra, rb_)):
            if x != y:
                assert c == wi and y == fix(x), (c, x, y)
                ncells += 1
    assert ncells == sum(changed.values())
    for sname in ("Mapping_Review",):
        assert list(openpyxl.load_workbook(path, read_only=True)[sname].iter_rows(values_only=True)) == \
            list(openpyxl.load_workbook(tmp, read_only=True)[sname].iter_rows(values_only=True))
    zt.close()
    os.replace(tmp, path)
    return changed


def main():
    for f in FILES:
        ch = process(DIR / f)
        print(f"{f}: {len(ch)} names updated, {sum(ch.values())} cells; "
              + ", ".join(f"{o}->{n}" for (o, n) in sorted(ch)))


if __name__ == "__main__":
    main()
