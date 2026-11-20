#!/usr/bin/env python3
"""Standardise the first seven columns of the well test monitoring sheets into the template order
ORIGINAL WELL NAME | FIELD | PLATFORM | WELL | STRING TYPE | RESERVOIR | DATE, with the Date timestamp split into DATE
(date only) and a new TIME column right after it. Every other column follows unchanged, in its original order.

Usage: standardise_well_tests.py [source folder] [output folder]   (both under all_data_standardisation/)
Defaults: well_test_monitoring_template_DULG_platform_20261015 -> 01_UPDATED/WELL-TEST-MONITORING
Source files are read only; output gets <same name> files plus backup_original/ and _reports/.

  ORIGINAL WELL NAME  Entity (old A) as is
  FIELD               DULANG
  PLATFORM            DULANG-<letter> from the Entity prefix
  WELL                Entity with DULG- -> DULANG- and leading zeros dropped (DULG-A010S02 -> DULANG-A10S2)
  STRING TYPE         blank: the Entity carries no string and the sheet has no string field (Description remarks that
                      mention LS/SS/reservoirs are listed for review, not used)
  RESERVOIR           blank: no reservoir field
  DATE / TIME         old Date (= Start Date) split into its date and its time of day; TIME is left blank when Start
                      Date is blank and Date sits at midnight (the source gave a date only, no time)
Old C..AO move to I..AU. Package-level edit: sheet1, its comments/VML anchor, styles (two formats appended) and the
workbook's filter name change; every other part is copied byte for byte. Validated cell by cell against the source.
"""
import datetime
import re
import shutil
import sys
import zipfile
from pathlib import Path

import openpyxl
from openpyxl.utils import column_index_from_string as cidx, get_column_letter as cl

BASE = Path(__file__).resolve().parent / "all_data_standardisation"
SRC = BASE / (sys.argv[1] if len(sys.argv) > 1 else "well_test_monitoring_template_DULG_platform_20261015")
OUT = BASE / (sys.argv[2] if len(sys.argv) > 2 else "01_UPDATED/WELL-TEST-MONITORING")
BACKUP = OUT / "backup_original"
REPORTS = OUT / "_reports"
SHEET = "Well Test Monitoring"
NEW_HDR = ["ORIGINAL WELL NAME", "FIELD", "PLATFORM", "WELL", "STRING TYPE", "RESERVOIR", "DATE", "TIME"]
WELL_RE = re.compile(r"(?:DULG|DULANG)-([A-D])(\d+)((?:[A-Z]+\d*)*)")
MENTION_RE = re.compile(r"\b(LS|SS|TS|long string|short string|single string|E\d{1,2}[A-Z]?)\b|[A-D]\d{1,2}(LS|SS)\b", re.I)
SHIFT = 6  # old column c >= 3 -> c + 6


def newcol(c):
    return {1: 1, 2: 7}.get(c, c + SHIFT)


def std_well(v):
    m = WELL_RE.fullmatch(v)
    if not m:
        return None
    sfx = re.sub(r"([A-Z]+)(\d*)", lambda s: s.group(1) + (str(int(s.group(2))) if s.group(2) else ""), m.group(3))
    return f"DULANG-{m.group(1)}{int(m.group(2))}{sfx}"


def remap_range(ref):
    a, b = ref.split(":")
    ca, ra = re.fullmatch(r"([A-Z]+)(\d+)", a).groups()
    cb, rb = re.fullmatch(r"([A-Z]+)(\d+)", b).groups()
    cols = [newcol(c) for c in range(cidx(ca), cidx(cb) + 1)] + ([8] if cidx(ca) <= 2 <= cidx(cb) else [])
    return f"{cl(min(cols))}{ra}:{cl(max(cols))}{rb}"


def esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


CELL_RE = re.compile(r'<c r="([A-Z]+)(\d+)"((?:\s+[\w:]+="[^"]*")*)\s*(?:/>|>(.*?)</c>)', re.S)
ROW_RE = re.compile(r'(<row r="(\d+)"[^>]*?)(?:/>|>(.*?)</row>)', re.S)


def build(src, dst):
    wb = openpyxl.load_workbook(src)
    ws = wb[SHEET]
    hdr = [c.value for c in ws[1]]
    assert hdr[:3] == ["Entity", "Date", "Gross (BFPD)"] and hdr[12] == "Start Date", hdr
    ncols = len(hdr)
    plat_file = re.search(r"DULG-([A-D])_", src.name).group(1)
    decisions, review = {}, []
    for r in range(2, ws.max_row + 1):
        ent, dt = ws.cell(r, 1).value, ws.cell(r, 2).value
        if ent is None and dt is None:
            continue
        well = std_well(ent) if isinstance(ent, str) else None
        plat = f"DULANG-{well[7]}" if well else None
        if well is None:
            review.append((src.name, r, ent, "WELL / PLATFORM blank", "Entity not in a recognised format"))
        elif well[7] != plat_file:
            review.append((src.name, r, ent, "review", f"platform {well[7]} differs from the file's platform {plat_file}"))
        start = ws.cell(r, 13).value
        time_known = True
        if not isinstance(dt, datetime.datetime):
            review.append((src.name, r, ent, "DATE / TIME blank", f"Date {dt!r} is not a timestamp"))
            dt = None
        elif start is None:
            if dt.time() == datetime.time(0):
                time_known = False
                review.append((src.name, r, ent, "TIME blank", "Start Date blank and Date has no time (00:00)"))
        elif (start if isinstance(start, datetime.datetime) else str(start)) != \
                (dt if isinstance(start, datetime.datetime) else dt.strftime("%Y-%m-%d %I:%M:%S %p")):
            review.append((src.name, r, ent, "review", f"Date {dt} differs from Start Date {start!r}"))
        desc = ws.cell(r, 41).value
        if isinstance(desc, str) and MENTION_RE.search(desc):
            review.append((src.name, r, ent, "not used (Description)",
                           "remark mentions a string or reservoir; not the record's own designation: " + desc.strip()))
        decisions[r] = dict(well=well, platform=plat, dt=dt, time_known=time_known)

    zin = zipfile.ZipFile(src)
    assert '<sheet name="%s" sheetId="1" r:id="rId1"/>' % SHEET in zin.read("xl/workbook.xml").decode()
    rels = zin.read("xl/_rels/workbook.xml.rels").decode()
    assert re.search(r'<Relationship [^>]*Id="rId1"[^>]*Target="[^"]*worksheets/sheet1.xml"', rels) or \
        re.search(r'<Relationship [^>]*Target="[^"]*worksheets/sheet1.xml"[^>]*Id="rId1"', rels)
    # styles: append a date-only and a time-only format
    styles = zin.read("xl/styles.xml").decode()
    ids = [int(x) for x in re.findall(r'numFmtId="(\d+)"', styles)]
    f_date, f_time = max(ids + [163]) + 1, max(ids + [163]) + 2
    styles = re.sub(r'<numFmts count="(\d+)">(.*?)</numFmts>',
                    lambda m: f'<numFmts count="{int(m.group(1)) + 2}">{m.group(2)}'
                              f'<numFmt numFmtId="{f_date}" formatCode="yyyy-mm-dd"/>'
                              f'<numFmt numFmtId="{f_time}" formatCode="hh:mm:ss"/></numFmts>', styles, flags=re.S)
    nxf = int(re.search(r'<cellXfs count="(\d+)"', styles).group(1))
    s_date, s_time = nxf, nxf + 1
    styles = re.sub(r'<cellXfs count="\d+">(.*?)</cellXfs>',
                    lambda m: f'<cellXfs count="{nxf + 2}">{m.group(1)}'
                              f'<xf numFmtId="{f_date}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
                              f'<xf numFmtId="{f_time}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
                              f'</cellXfs>', styles, flags=re.S)

    def row_sub(mo):
        head, rn, body = mo.group(1), int(mo.group(2)), mo.group(3) or ""
        cells = {}
        for c in CELL_RE.finditer(body):
            assert int(c.group(2)) == rn
            cells[cidx(c.group(1))] = (c.group(3), c.group(4))
        out = {}
        for oc, (attrs, inner) in cells.items():
            if oc == 2:
                continue
            nc = newcol(oc)
            out[nc] = f'<c r="{cl(nc)}{rn}"{attrs}/>' if inner is None else f'<c r="{cl(nc)}{rn}"{attrs}>{inner}</c>'

        def text(c, v, attrs=""):
            out[c] = f'<c r="{cl(c)}{rn}"{attrs} t="str"><v>{esc(v)}</v></c>'
        if rn == 1:
            for c, name in enumerate(NEW_HDR, 1):
                text(c, name, re.sub(r'\st="[^"]+"', "", cells[1 if c == 1 else 2][0]))
        elif rn in decisions:
            d = decisions[rn]
            text(2, "DULANG")
            if d["platform"]:
                text(3, d["platform"])
            if d["well"]:
                text(4, d["well"])
            if d["dt"] is not None:
                dt = d["dt"]
                serial = (dt.date() - datetime.date(1899, 12, 30)).days
                frac = (dt.hour * 3600 + dt.minute * 60 + dt.second) / 86400
                out[7] = f'<c r="G{rn}" s="{s_date}"><v>{serial}</v></c>'
                if d["time_known"]:
                    out[8] = f'<c r="H{rn}" s="{s_time}"><v>{frac!r}</v></c>'
        return f"{head}>{''.join(out[c] for c in sorted(out))}</row>"

    xml = zin.read("xl/worksheets/sheet1.xml").decode()
    pre, rest = xml.split("<sheetData>", 1)
    data, post = rest.split("</sheetData>", 1)
    data = ROW_RE.sub(row_sub, data)
    pre = re.sub(r'<dimension ref="([^"]+)"/>', lambda x: f'<dimension ref="{remap_range(x.group(1))}"/>', pre)
    cols = {}
    for c in re.findall(r"<col [^>]*/>", pre):
        lo, hi = int(re.search(r'min="(\d+)"', c).group(1)), int(re.search(r'max="(\d+)"', c).group(1))
        for k in range(lo, hi + 1):
            targets = [7, 8] if k == 2 else [newcol(k)]
            for t in targets:
                cols[t] = re.sub(r'min="\d+" max="\d+"', f'min="{t}" max="{t}"', c)
    for t in range(2, 7):
        cols[t] = f'<col min="{t}" max="{t}" width="14.83203125" customWidth="1"/>'
    pre = re.sub(r"<cols>.*?</cols>", "<cols>" + "".join(cols[k] for k in sorted(cols)) + "</cols>", pre, flags=re.S)
    assert "<filterColumn" not in post and "<sortState" not in post and "<dataValidation" not in post
    post = re.sub(r'\b(ref|sqref)="([A-Z]+\d+:[A-Z]+\d+(?: [A-Z]+\d+:[A-Z]+\d+)*)"',
                  lambda x: f'{x.group(1)}="{" ".join(remap_range(p) for p in x.group(2).split())}"', post)
    sheet = pre + "<sheetData>" + data + "</sheetData>" + post

    wbx = re.sub(r"(!)\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)",
                 lambda x: x.group(1) + remap_range(f"{x.group(2)}{x.group(3)}:{x.group(4)}{x.group(5)}"),
                 zin.read("xl/workbook.xml").decode())
    parts = {"xl/worksheets/sheet1.xml": sheet, "xl/styles.xml": styles, "xl/workbook.xml": wbx}
    names = zin.namelist()
    if "xl/comments1.xml" in names:
        assert "comments1.xml" in zin.read("xl/worksheets/_rels/sheet1.xml.rels").decode()
        parts["xl/comments1.xml"] = re.sub(r'ref="([A-Z]+)(\d+)"',
                                           lambda x: f'ref="{cl(newcol(cidx(x.group(1))))}{x.group(2)}"',
                                           zin.read("xl/comments1.xml").decode())
        vml = zin.read("xl/drawings/vmlDrawing1.vml").decode()
        vml = re.sub(r"<x:Column>(\d+)</x:Column>",
                     lambda x: f"<x:Column>{newcol(int(x.group(1)) + 1) - 1}</x:Column>", vml)
        vml = re.sub(r"<x:Anchor>\s*(\d+),\s*(\d+),\s*(\d+),\s*(\d+),\s*(\d+),",
                     lambda x: f"<x:Anchor>{int(x.group(1)) + SHIFT},{x.group(2)},{x.group(3)},{x.group(4)},"
                               f"{int(x.group(5)) + SHIFT},", vml)
        parts["xl/drawings/vmlDrawing1.vml"] = vml

    tmp = dst.with_name(dst.stem + ".partial.xlsx")
    with open(tmp, "xb") as fh, zipfile.ZipFile(fh, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            zi = zipfile.ZipInfo(info.filename, info.date_time)
            zi.compress_type, zi.external_attr = zipfile.ZIP_DEFLATED, info.external_attr
            data_ = parts[info.filename].encode() if info.filename in parts else zin.read(info.filename)
            zout.writestr(zi, data_)
    validate(src, tmp, decisions, ncols, parts)
    tmp.rename(dst)
    return decisions, review


def validate(src, tmp, decisions, ncols, parts):
    zs, zt = zipfile.ZipFile(src), zipfile.ZipFile(tmp)
    assert zs.namelist() == zt.namelist()
    for n in zs.namelist():
        if n not in parts:
            assert zs.read(n) == zt.read(n), n
    a, b = openpyxl.load_workbook(src), openpyxl.load_workbook(tmp)
    assert a.sheetnames == b.sheetnames
    for name in a.sheetnames[1:]:
        assert [[c.value for c in r] for r in a[name].iter_rows()] == [[c.value for c in r] for r in b[name].iter_rows()]
    wa, wb_ = a[SHEET], b[SHEET]
    assert wa.max_row == wb_.max_row and wb_.max_column == ncols + SHIFT
    assert [c.value for c in wb_[1]][:8] == NEW_HDR
    for r in range(1, wa.max_row + 1):
        assert wa.row_dimensions[r].height == wb_.row_dimensions[r].height, r
        for oc in [1] + list(range(3, ncols + 1)):
            x, y = wa.cell(r, oc), wb_.cell(r, newcol(oc))
            assert x.number_format == y.number_format and repr(x.font) == repr(y.font) \
                and repr(x.fill) == repr(y.fill) and repr(x.alignment) == repr(y.alignment), (r, oc)
            if r > 1 or oc != 1:
                assert x.value == y.value, (r, oc, x.value, y.value)
            assert (x.comment.text if x.comment else None) == (y.comment.text if y.comment else None), (r, oc)
        if r in decisions:
            d = decisions[r]
            assert [wb_.cell(r, c).value for c in (2, 3, 4, 5, 6)] == ["DULANG", d["platform"], d["well"], None, None]
            g, h = wb_.cell(r, 7).value, wb_.cell(r, 8).value
            if d["dt"] is None:
                assert g is None and h is None
            else:
                gd = g.date() if isinstance(g, datetime.datetime) else g
                assert gd == wa.cell(r, 2).value.date() and wb_.cell(r, 7).number_format == "yyyy-mm-dd", r
                if not d["time_known"]:
                    assert h is None, r
                else:
                    assert wb_.cell(r, 8).number_format == "hh:mm:ss"
                    ht = h if isinstance(h, datetime.time) else h.time()
                    assert ht.replace(microsecond=0) == wa.cell(r, 2).value.time(), (r, h)
    assert wb_.auto_filter.ref == f"A1:{cl(ncols + SHIFT)}{wa.max_row}"
    widths = {}
    for dim in wa.column_dimensions.values():
        if dim.min:
            for k in range(dim.min, dim.max + 1):
                widths[k] = dim.width
    for k in range(1, ncols + 1):
        assert wb_.column_dimensions[cl(newcol(k))].width == widths.get(k, 13), k
    zs.close()
    zt.close()


def main():
    files = sorted(p for p in SRC.glob("well_test_*.xlsx") if not p.name.startswith("~$"))
    for p in files:
        for d in (OUT / p.name, BACKUP / p.name):
            if d.exists():
                sys.exit(f"STOP: {d} already exists - not overwritten")
    BACKUP.mkdir(parents=True, exist_ok=True)
    REPORTS.mkdir(exist_ok=True)
    allrev, summary = [], []
    for p in files:
        shutil.copy2(p, BACKUP / p.name)
        decisions, review = build(BACKUP / p.name, OUT / p.name)
        allrev += review
        wells = sorted({d["well"] for d in decisions.values() if d["well"]})
        summary.append([p.name, len(decisions), len(wells), ", ".join(wells)])
        print(f"== {p.name}: {len(decisions)} rows, {len(wells)} wells; review {len(review)}")
    out = openpyxl.Workbook()
    ws = out.active
    ws.title = "Unresolved"
    ws.append(["File", "Row", "Original Well Name", "Issue", "Detail"])
    for r in allrev:
        ws.append(list(r))
    ws2 = out.create_sheet("Summary")
    ws2.append(["File", "Data rows", "Distinct WELL", "WELL values"])
    for s in summary:
        ws2.append(s)
    rp = REPORTS / f"well_test_standardisation_review_{re.search(r"_(\d{8}[^_]*)$", SRC.name).group(1)}.xlsx"
    out.save(rp)
    print("review list:", rp)


if __name__ == "__main__":
    main()
