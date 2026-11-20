#!/usr/bin/env python3
"""Standardise the WELL column of every Excel workbook in this folder (in place, backups first).

DULG-/DULANG- prefix -> DULANG-; leading zeros dropped from the main number and every numeric suffix; letters kept:
DULG-A001 -> DULANG-A1, DULG-C019S02 -> DULANG-C19S2, DULG-D019H01 -> DULANG-D19H1.
Only cells under a header cell exactly 'WELL' (case/space-insensitive) change; formulas, blanks, already-standard and
unrecognised values are left as they are (unrecognised ones are reported). Package-level edit: new names are appended to
sharedStrings and only those WELL cells are re-pointed, every other part is copied byte for byte. Each result is
validated cell by cell against the backup before it replaces the original.
"""
import collections
import html
import os
import re
import shutil
import sys
import zipfile
from pathlib import Path

import openpyxl

DIR = Path(__file__).resolve().parent
BACKUP = DIR / "backup_before_well_update"
NAME = re.compile(r"(?:DULG|DULANG)-([A-Z])(\d+)((?:[A-Z]+\d+)*)")


def std(v):
    m = NAME.fullmatch(v)
    if not m:
        return None
    sfx = re.sub(r"([A-Z]+)(\d+)", lambda s: s.group(1) + str(int(s.group(2))), m.group(3))
    return f"DULANG-{m.group(1)}{int(m.group(2))}{sfx}"


def col_letter(i):
    return openpyxl.utils.get_column_letter(i + 1)


def sst_texts(xml):
    return [html.unescape("".join(t.decode() for t in re.findall(rb"<t[^>]*>(.*?)</t>", it, re.S)))
            for it in re.findall(rb"<si>(.*?)</si>", xml, re.S)]


def sheet_parts(z):
    wb = z.read("xl/workbook.xml").decode()
    rels = z.read("xl/_rels/workbook.xml.rels").decode()
    tgt = dict(re.findall(r'<Relationship Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
    tgt.update({a: b for b, a in re.findall(r'<Relationship [^>]*Target="([^"]+)"[^>]*Id="([^"]+)"', rels)})
    out = []
    for name, rid in re.findall(r'<sheet name="([^"]+)" sheetId="\d+"[^>]*r:id="([^"]+)"', wb):
        t = tgt[rid].lstrip("/")
        out.append((html.unescape(name), t if t.startswith("xl/") else "xl/" + t))
    return out


def find_well_cols(path, header="WELL"):
    """{sheet: (header_row, col_index)} for header cells exactly `header` within the first 30 rows"""
    found = {}
    wb = openpyxl.load_workbook(path, read_only=True)
    for ws in wb.worksheets:
        hits = [(ri, ci) for ri, r in enumerate(ws.iter_rows(max_row=30, values_only=True), 1)
                for ci, v in enumerate(r) if isinstance(v, str) and v.strip().upper() == header]
        if len(hits) > 1:
            raise ValueError(f"{ws.title}: several {header} headers {hits}")
        if hits:
            found[ws.title] = hits[0]
    return found


def process(path, bpath, header="WELL", conv=std):
    wells = find_well_cols(bpath, header)
    zin = zipfile.ZipFile(bpath)
    names = zin.namelist()
    sst_xml = zin.read("xl/sharedStrings.xml") if "xl/sharedStrings.xml" in names else None
    sst = sst_texts(sst_xml) if sst_xml else []
    index = {t: i for i, t in enumerate(sst)}
    added = {}
    new_parts, report = {}, {}
    for sname, part in sheet_parts(zin):
        if sname not in wells:
            report[sname] = None
            continue
        hrow, ci = wells[sname]
        col = col_letter(ci).encode()
        changed, review, kinds = collections.Counter(), collections.Counter(), collections.Counter()

        def sub(m):
            row, attrs, inner = int(m.group(1)), m.group(2).decode(), m.group(3).decode()
            if row <= hrow:
                return m.group(0)
            if "<f" in inner:
                kinds["formula (left)"] += 1
                return m.group(0)
            t = re.search(r'\st="([^"]+)"', attrs)
            t = t.group(1) if t else None
            if t == "s":
                old = sst[int(re.search(r"<v>(\d+)</v>", inner).group(1))]
            elif t == "inlineStr":
                old = html.unescape("".join(re.findall(r"<t[^>]*>(.*?)</t>", inner, re.S)))
            elif t == "str":
                old = html.unescape(re.search(r"<v>(.*?)</v>", inner, re.S).group(1))
            else:
                if inner.strip():
                    kinds["non-text value (left)"] += 1
                    review[f"non-text: {re.sub(r'<[^>]+>', '', inner)}"] += 1
                return m.group(0)
            new = conv(old)
            if new is None:
                review[old] += 1
                return m.group(0)
            if new == old:
                kinds["already standard"] += 1
                return m.group(0)
            changed[(old, new)] += 1
            if new not in index:
                index[new] = len(sst) + len(added)
                added[new] = index[new]
            a = re.sub(r'\st="[^"]+"', "", attrs)
            return f'<c r="{col.decode()}{row}"{a} t="s"><v>{index[new]}</v></c>'.encode()

        xml = zin.read(part)
        new_xml = re.sub(rb'<c r="%s(\d+)"((?:\s+[\w:]+="[^"]*")*)\s*>(.*?)</c>' % col, sub, xml, flags=re.S)
        new_parts[part] = new_xml
        report[sname] = dict(col=col.decode(), header_row=hrow, changed=changed, review=review, kinds=kinds)

    if not any(r and r["changed"] for r in report.values()):
        return report, False
    if sst_xml is None:
        raise RuntimeError("no sharedStrings part")
    new_sst = sst_xml.replace(b"</sst>", b"".join(b"<si><t>%s</t></si>" % html.escape(t, quote=False).encode()
                                                for t in added) + b"</sst>")
    new_sst = re.sub(rb'uniqueCount="\d+"', b'uniqueCount="%d"' % (len(sst) + len(added)), new_sst, count=1)
    new_parts["xl/sharedStrings.xml"] = new_sst

    tmp = path.with_name(path.stem + ".partial" + path.suffix)
    with open(tmp, "xb") as fh, zipfile.ZipFile(fh, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            zi = zipfile.ZipInfo(info.filename, info.date_time)
            zi.compress_type, zi.external_attr = zipfile.ZIP_DEFLATED, info.external_attr
            zout.writestr(zi, new_parts.get(info.filename, zin.read(info.filename)))

    # validate against the backup: same parts, same sheets, only intended WELL cells differ
    zt = zipfile.ZipFile(tmp)
    assert zt.namelist() == names
    for n in names:
        if n not in new_parts:
            assert zt.read(n) == zin.read(n), n
    a, b = openpyxl.load_workbook(bpath), openpyxl.load_workbook(tmp)  # full load: values, formulas and styles
    assert a.sheetnames == b.sheetnames
    for wa in a.worksheets:
        wb_ = b[wa.title]
        assert (wa.max_row, wa.max_column) == (wb_.max_row, wb_.max_column), wa.title
        info = report.get(wa.title)
        n_changed = 0
        for ra, rb in zip(wa.iter_rows(), wb_.iter_rows()):
            for x, y in zip(ra, rb):
                assert x.style_id == y.style_id and x.number_format == y.number_format, (wa.title, x.coordinate)
                if x.value != y.value:
                    assert info and x.column - 1 == info["header_row"] * 0 + (wells[wa.title][1]) \
                        and x.row > info["header_row"] and y.value == conv(x.value), (wa.title, x.coordinate, x.value,
                                                                                    y.value)
                    n_changed += 1
        assert n_changed == (sum(info["changed"].values()) if info else 0), (wa.title, n_changed)
        assert [c.text for c in []] == []
        # comments survive unchanged
        ca = {(c.row, c.column): c.comment.text for row in wa.iter_rows() for c in row if c.comment}
        cb = {(c.row, c.column): c.comment.text for row in wb_.iter_rows() for c in row if c.comment}
        assert ca == cb, wa.title
    zt.close()
    os.replace(tmp, path)
    return report, True


def main():
    files = sorted(p for p in DIR.iterdir() if p.is_file() and p.suffix.lower() in (".xlsx", ".xlsm")
                   and not p.name.startswith("~$"))
    skipped = [p.name for p in DIR.iterdir() if p.is_file() and p.suffix.lower() in (".xls", ".xlsb", ".ods")]
    BACKUP.mkdir(exist_ok=True)
    for p in files:
        b = BACKUP / p.name
        if b.exists():
            sys.exit(f"STOP: backup {b} already exists - not overwritten")
    for p in files:
        shutil.copy2(p, BACKUP / p.name)
    for p in files:
        try:
            report, saved = process(p, BACKUP / p.name)
        except Exception as e:  # leave the file as it is
            print(f"SKIPPED {p.name}: {e!r}")
            continue
        print(f"== {p.name}: {'saved' if saved else 'no change'}")
        for s, r in report.items():
            if r is None:
                print(f"   sheet '{s}': no WELL column")
                continue
            print(f"   sheet '{s}': WELL = column {r['col']} (header row {r['header_row']}), "
                  f"{len(r['changed'])} names / {sum(r['changed'].values())} cells updated; other: {dict(r['kinds'])}; "
                  f"review: {dict(r['review']) or 'none'}")
            print("     " + ", ".join(f"{o}->{n}" for (o, n) in sorted(r["changed"])))
    if skipped:
        print("not processed (format):", skipped)


if __name__ == "__main__":
    main()
