#!/usr/bin/env python3
"""Standardise the first seven columns of the platform allocated-production workbooks into the template order
ORIGINAL WELL NAME | FIELD | PLATFORM | WELL | STRING TYPE | RESERVOIR | DATE, with the Date split into DATE (date only)
and a new TIME column right after it. Every other column follows unchanged, in its original order.

Usage: standardise_allocated_production.py <source folder> [output folder]   (both under all_data_standardisation/)
e.g.   standardise_allocated_production.py DULG-A  ->  01_UPDATED/DULG-A/<same name>  (+ backup_original/, _reports/)

The Entity (old A) is WELL-STRING-FAULTBLOCK-RESERVOIR, e.g. DULG-A002-LS-S3-E12/13:
  ORIGINAL WELL NAME  Entity as is
  FIELD               DULANG
  PLATFORM            DULANG-<letter> from the Entity
  WELL                wellbore part, DULG- -> DULANG-, leading zeros dropped (DULG-A010S02 -> DULANG-A10S2)
  STRING TYPE         TS / LS / SS as written in the Entity
  RESERVOIR           last Entity part exactly as written (E12/13, E34A, MIDB, LD1)
  DATE / TIME         Date split into its date and time of day; a date-only text Date gets a blank TIME
Old C.. move six columns right. Package-level edit: the data sheet, styles (two formats appended) and the workbook's
filter name change; every other part is copied byte for byte. Validated cell by cell against the source.
"""
import collections
import datetime
import re
import shutil
import sys
import zipfile
from pathlib import Path

import openpyxl
from openpyxl.utils import column_index_from_string as cidx, get_column_letter as cl

BASE = Path(__file__).resolve().parent / "all_data_standardisation"
SRC = BASE / sys.argv[1]
OUT = BASE / (sys.argv[2] if len(sys.argv) > 2 else f"01_UPDATED/{sys.argv[1]}")
BACKUP = OUT / "backup_original"
REPORTS = OUT / "_reports"
NEW_HDR = ["ORIGINAL WELL NAME", "FIELD", "PLATFORM", "WELL", "STRING TYPE", "RESERVOIR", "DATE", "TIME"]
ENT_RE = re.compile(r"(?:DULG|DULANG)-([A-D])(\d+)((?:[A-Z]+\d+)*)-(TS|LS|SS)-([A-Z0-9]+)-([^-]+)")
SHIFT = 6  # old column c >= 3 -> c + 6


def newcol(c):
    return {1: 1, 2: 7}.get(c, c + SHIFT)


def remap_range(ref):
    def one(a):
        return re.fullmatch(r"(\$?)([A-Z]+)(\$?)(\d+)", a).groups()
    if ":" not in ref:
        d1, c, d2, r = one(ref)
        return f"{d1}{cl(newcol(cidx(c)))}{d2}{r}"
    (a1, ca, a2, ra), (b1, cb, b2, rb) = (one(p) for p in ref.split(":"))
    lo, hi = cidx(ca), cidx(cb)
    cols = [newcol(c) for c in range(lo, hi + 1)] + ([8] if lo <= 2 <= hi else [])
    return f"{a1}{cl(min(cols))}{a2}{ra}:{b1}{cl(max(cols))}{b2}{rb}"


def parse_entity(ent):
    m = ENT_RE.fullmatch(ent) if isinstance(ent, str) else None
    if not m:
        return None
    sfx = re.sub(r"([A-Z]+)(\d+)", lambda s: s.group(1) + str(int(s.group(2))), m.group(3))
    return dict(platform=f"DULANG-{m.group(1)}", well=f"DULANG-{m.group(1)}{int(m.group(2))}{sfx}",
                string=m.group(4), block=m.group(5), reservoir=m.group(6))


def esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


CELL_RE = re.compile(r'<c\s+r="([A-Z]+)(\d+)"((?:\s+[\w:]+="[^"]*")*)\s*(?:/>|>(.*?)</c>)', re.S)
ROW_RE = re.compile(r'(<row\s+r="(\d+)"[^>]*?)\s*(?:/>|>(.*?)</row>)', re.S)


def sheet1_part(z):
    wb = z.read("xl/workbook.xml").decode()
    rels = z.read("xl/_rels/workbook.xml.rels").decode()
    name, rid = re.search(r'<sheet\s+name="([^"]+)"\s+sheetId="\d+"\s+r:id="([^"]+)"', wb).groups()
    tgt = next(t for r in re.findall(r"<Relationship\b[^>]*>", rels)
               if f'Id="{rid}"' in r for t in re.findall(r'Target="([^"]+)"', r))
    assert tgt.lstrip("/").replace("xl/", "") == "worksheets/sheet1.xml", tgt
    return name


def build(src, dst, plat_file):
    z = zipfile.ZipFile(src)
    sname = sheet1_part(z)
    wb = openpyxl.load_workbook(src)
    ws = wb[sname]
    assert wb.sheetnames[0] == sname
    hdr = [c.value for c in ws[1]]
    assert hdr[:2] == ["Entity", "Date"], hdr
    ncols = ws.max_column
    decisions, review = {}, []
    for r in range(2, ws.max_row + 1):
        ent, dv = ws.cell(r, 1).value, ws.cell(r, 2).value
        if ent is None and dv is None:
            continue
        d = parse_entity(ent) or dict(platform=None, well=None, string=None, block=None, reservoir=None)
        if d["well"] is None:
            review.append((src.name, r, ent, "WELL/STRING/RESERVOIR blank", "Entity not in WELL-STRING-BLOCK-RESERVOIR form"))
        elif d["platform"][-1] != plat_file:
            review.append((src.name, r, ent, "review", f"platform differs from the folder's platform {plat_file}"))
        elif re.fullmatch(r"E0\d.*|E\d{3}.*", d["reservoir"]):
            review.append((src.name, r, ent, "review", f"unusual reservoir code {d['reservoir']!r}, kept as written"))
        date = time = None
        if isinstance(dv, datetime.datetime):
            date, time = dv.date(), dv.time()
        elif isinstance(dv, str) and re.fullmatch(r"\d{4}-\d{2}-\d{2}", dv.strip()):
            date = datetime.date.fromisoformat(dv.strip())
        else:
            review.append((src.name, r, ent, "DATE/TIME blank", f"Date {dv!r} not recognised"))
        d.update(date=date, time=time)
        decisions[r] = d

    styles = z.read("xl/styles.xml").decode()
    ids = [int(x) for x in re.findall(r'numFmtId="(\d+)"', styles)]
    f_date, f_time = max(ids + [163]) + 1, max(ids + [163]) + 2
    fmts = f'<numFmt numFmtId="{f_date}" formatCode="yyyy-mm-dd"/><numFmt numFmtId="{f_time}" formatCode="hh:mm:ss"/>'
    if re.search(r"<numFmts\b", styles):
        styles = re.sub(r'<numFmts\s+count="(\d+)"\s*>(.*?)</numFmts>',
                        lambda m: f'<numFmts count="{int(m.group(1)) + 2}">{m.group(2)}{fmts}</numFmts>', styles, flags=re.S)
    else:
        styles = re.sub(r"(<styleSheet\b[^>]*>)", lambda m: m.group(1) + f'<numFmts count="2">{fmts}</numFmts>', styles,
                        count=1)
    nxf = int(re.search(r'<cellXfs\s+count="(\d+)"', styles).group(1))
    s_date, s_time = nxf, nxf + 1
    styles = re.sub(r'<cellXfs\s+count="\d+"\s*>(.*?)</cellXfs>',
                    lambda m: f'<cellXfs count="{nxf + 2}">{m.group(1)}'
                              f'<xf numFmtId="{f_date}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
                              f'<xf numFmtId="{f_time}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
                              f'</cellXfs>', styles, flags=re.S)
    assert len(re.findall(r"<xf\b", re.search(r"<cellXfs.*?</cellXfs>", styles, re.S).group(0))) == nxf + 2

    def row_sub(mo):
        head, rn, body = mo.group(1), int(mo.group(2)), mo.group(3) or ""
        head = re.sub(r'spans="1:\d+"', f'spans="1:{ncols + SHIFT}"', head)
        cells = {}
        for c in CELL_RE.finditer(body):
            assert int(c.group(2)) == rn
            cells[cidx(c.group(1))] = (c.group(3), c.group(4))
        out = {}
        for oc, (attrs, inner) in cells.items():
            if oc == 2 and rn > 1:
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
            for c, k in ((3, "platform"), (4, "well"), (5, "string"), (6, "reservoir")):
                if d[k]:
                    text(c, d[k])
            if d["date"] is not None:
                out[7] = f'<c r="G{rn}" s="{s_date}"><v>{(d["date"] - datetime.date(1899, 12, 30)).days}</v></c>'
            if d["time"] is not None:
                t = d["time"]
                out[8] = f'<c r="H{rn}" s="{s_time}"><v>{(t.hour * 3600 + t.minute * 60 + t.second) / 86400!r}</v></c>'
        return f"{head}>{''.join(out[c] for c in sorted(out))}</row>"

    xml = z.read("xl/worksheets/sheet1.xml").decode()
    pre, rest = re.split(r"<sheetData\s*>", xml, maxsplit=1)
    data, post = rest.split("</sheetData>", 1)
    assert not re.search(r"<(mergeCell|dataValidation|conditionalFormatting|hyperlink|tablePart|legacyDrawing|drawing|"
                         r"filterColumn)\b", post), post[:300]
    assert "<f>" not in data and "<f " not in data
    data = ROW_RE.sub(row_sub, data)
    pre = re.sub(r'<dimension\s+ref="([^"]+)"\s*/>', lambda x: f'<dimension ref="{remap_range(x.group(1))}"/>', pre)
    if re.search(r"<cols\b", pre):
        cols = {}
        for c in re.findall(r"<col\b[^>]*/>", pre):
            lo, hi = int(re.search(r'min="(\d+)"', c).group(1)), int(re.search(r'max="(\d+)"', c).group(1))
            for k in range(lo, hi + 1):
                for t in ([7, 8] if k == 2 else [newcol(k)]):
                    cols[t] = re.sub(r'max="\d+"', f'max="{t}"', re.sub(r'min="\d+"', f'min="{t}"', c))
        for t in range(2, 7):
            cols[t] = f'<col min="{t}" max="{t}" width="14.83203125" customWidth="1"/>'
        pre = re.sub(r"<cols\b[^>]*>.*?</cols>", "<cols>" + "".join(cols[k] for k in sorted(cols)) + "</cols>", pre,
                     flags=re.S)
    post = re.sub(r'\b(ref|sqref)="([^"]+)"', lambda x: f'{x.group(1)}="{" ".join(remap_range(p) for p in x.group(2).split())}"',
                  post)
    sheet = pre + "<sheetData>" + data + "</sheetData>" + post
    wbx = re.sub(r"(!)(\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?\d+)", lambda x: x.group(1) + remap_range(x.group(2)),
                 z.read("xl/workbook.xml").decode())
    parts = {"xl/worksheets/sheet1.xml": sheet, "xl/styles.xml": styles, "xl/workbook.xml": wbx}

    tmp = dst.with_name(dst.stem + ".partial.xlsx")
    with open(tmp, "xb") as fh, zipfile.ZipFile(fh, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in z.infolist():
            zi = zipfile.ZipInfo(info.filename, info.date_time)
            zi.compress_type, zi.external_attr = zipfile.ZIP_DEFLATED, info.external_attr
            zout.writestr(zi, parts[info.filename].encode() if info.filename in parts else z.read(info.filename))
    validate(src, tmp, sname, decisions, ncols, parts)
    tmp.rename(dst)
    return decisions, review


def validate(src, tmp, sname, decisions, ncols, parts):
    zs, zt = zipfile.ZipFile(src), zipfile.ZipFile(tmp)
    assert zs.namelist() == zt.namelist()
    for n in zs.namelist():
        if n not in parts:
            assert zs.read(n) == zt.read(n), n
    a, b = openpyxl.load_workbook(src), openpyxl.load_workbook(tmp)
    assert a.sheetnames == b.sheetnames
    for name in a.sheetnames:
        if name != sname:
            assert [[c.value for c in r] for r in a[name].iter_rows()] == [[c.value for c in r] for r in b[name].iter_rows()]
    wa, wb_ = a[sname], b[sname]
    assert wa.max_row == wb_.max_row and wb_.max_column == ncols + SHIFT
    assert [c.value for c in wb_[1]][:8] == NEW_HDR
    for r, (ra, rb) in enumerate(zip(wa.iter_rows(), wb_.iter_rows()), 1):
        assert wa.row_dimensions[r].height == wb_.row_dimensions[r].height, r
        for oc in [1] + list(range(3, ncols + 1)):
            x, y = ra[oc - 1], rb[newcol(oc) - 1]
            assert x.number_format == y.number_format and repr(x.font) == repr(y.font) \
                and repr(x.fill) == repr(y.fill) and repr(x.alignment) == repr(y.alignment), (r, oc)
            if r > 1 or oc != 1:
                assert x.value == y.value, (r, oc, x.value, y.value)
        if r in decisions:
            d = decisions[r]
            assert [rb[c].value for c in range(1, 6)] == ["DULANG", d["platform"], d["well"], d["string"], d["reservoir"]]
            g, h = rb[6].value, rb[7].value
            src_d = ra[1].value
            if d["date"] is None:
                assert g is None
            else:
                assert g.date() == (src_d.date() if isinstance(src_d, datetime.datetime)
                                    else datetime.date.fromisoformat(src_d.strip())) and rb[6].number_format == "yyyy-mm-dd"
            if d["time"] is None:
                assert h is None
            else:
                ht = h if isinstance(h, datetime.time) else h.time()
                assert ht.replace(microsecond=0) == src_d.time() and rb[7].number_format == "hh:mm:ss", (r, h)
            assert ra[0].value.startswith(d["well"].replace("DULANG-", "DULG-")[:6])
    widths = {}
    for dim in wa.column_dimensions.values():
        if dim.min and dim.width:
            for k in range(dim.min, dim.max + 1):
                widths[k] = dim.width
    for k, w in widths.items():
        for t in ([7, 8] if k == 2 else [newcol(k)]):
            assert wb_.column_dimensions[cl(t)].width == w, k
    if wa.auto_filter.ref:
        assert wb_.auto_filter.ref == remap_range(wa.auto_filter.ref)
    assert wb_.freeze_panes == wa.freeze_panes
    zs.close()
    zt.close()


def main():
    plat_file = re.search(r"DULG-([A-D])$", SRC.name).group(1)
    files = sorted(p for p in SRC.glob("*.xlsx") if not p.name.startswith("~$"))
    for p in files:
        for d in (OUT / p.name, BACKUP / p.name):
            if d.exists():
                sys.exit(f"STOP: {d} already exists - not overwritten")
    BACKUP.mkdir(parents=True, exist_ok=True)
    REPORTS.mkdir(exist_ok=True)
    allrev, summary, seen = [], [], collections.defaultdict(list)
    for p in files:
        shutil.copy2(p, BACKUP / p.name)
        decisions, review = build(BACKUP / p.name, OUT / p.name, plat_file)
        allrev += review
        if not decisions:
            allrev.append((p.name, None, None, "review", "no data rows (header only)"))
        ents = sorted({openpyxl.load_workbook(OUT / p.name, read_only=True).worksheets[0].cell(r, 1).value
                       for r in decisions}) if len(decisions) < 1000 else None
        times = collections.Counter(d["time"] for d in decisions.values())
        if ents and len(ents) == 1:
            seen[ents[0]].append(p.name)
        summary.append([p.name, len(decisions), len({d["well"] for d in decisions.values()}),
                        ", ".join(sorted({d["reservoir"] or "" for d in decisions.values()})),
                        ", ".join(f"{t or 'blank'} x{n}" for t, n in times.items())])
        print(f"== {p.name}: {len(decisions)} rows; times {dict(times)}; review {len(review)}")
    for ent, names in seen.items():
        if len(names) > 1:
            allrev.append((" + ".join(names), None, ent, "review", "same Entity in more than one file"))
    out = openpyxl.Workbook()
    ws = out.active
    ws.title = "Unresolved"
    ws.append(["File", "Row", "Original Well Name", "Issue", "Detail"])
    for r in allrev:
        ws.append(list(r))
    ws2 = out.create_sheet("Summary")
    ws2.append(["File", "Data rows", "Distinct WELL", "RESERVOIR values", "TIME values"])
    for s in summary:
        ws2.append(s)
    rp = REPORTS / f"allocated_production_standardisation_review_{SRC.name}.xlsx"
    out.save(rp)
    print("review list:", rp, "| unresolved:", len(allrev))


if __name__ == "__main__":
    main()
