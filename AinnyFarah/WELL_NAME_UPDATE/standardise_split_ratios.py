#!/usr/bin/env python3
"""Add the seven template columns in front of the FPC Perforation Split Ratio sheets:
ORIGINAL WELL NAME | FIELD | PLATFORM | WELL | STRING TYPE | RESERVOIR | DATE, all taken from the same row.
The 15 original columns (group label, Well, ..., Perf. Code, ..., Start Ratio Date, ratios) move unchanged to H..V.

Source (read only): all_data_standardisation/Perforation-Split-Ratios/*.xlsx
Output:             all_data_standardisation/01_UPDATED/PERFORATION-SPLIT-RATIOS/<same name>  (+ backup_original/, _reports/)

  ORIGINAL WELL NAME  Perf. Code as is (e.g. DULG-A002-LS-S3-E12/13); Perf. Code also stays in its own column
  FIELD               DULANG
  PLATFORM            DULANG-<letter> from Perf. Code
  WELL                wellbore part of Perf. Code, DULG- -> DULANG-, leading zeros dropped (DULG-A010S02 -> DULANG-A10S2)
  STRING TYPE         TS / LS / SS as written in Perf. Code
  RESERVOIR           last Perf. Code part exactly as written (E12/13, E11A/E11B, MIDB)
  DATE                Start Ratio Date ('21-Sep-2018' text) as a real date; the source has no end date and no time
"Well: ..." group rows (merged across the old columns) are not records: their new cells stay blank in the group shading.
Package-level edit: only the sheet XML and styles change (one date format appended); every other part is copied byte
for byte. The result is validated cell by cell against the source before it is written.
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
SRC = BASE / "Perforation-Split-Ratios"
OUT = BASE / "01_UPDATED" / "PERFORATION-SPLIT-RATIOS"
BACKUP = OUT / "backup_original"
REPORTS = OUT / "_reports"
NEW_HDR = ["ORIGINAL WELL NAME", "FIELD", "PLATFORM", "WELL", "STRING TYPE", "RESERVOIR", "DATE"]
OLD_HDR = [None, "Well", "Well Type", "Well Status", "Perf. Code", "Resv. Block Formation", "Resv. Block",
           "Resv. Formation", "Start Ratio Date"]
PERF, DATE_COL = 5, 9
SHIFT = 7
ENT_RE = re.compile(r"(?:DULG|DULANG)-([A-D])(\d+)((?:[A-Z]+\d+)*)-(TS|LS|SS)-([A-Z0-9]+)-([^-]+)")
CELL_RE = re.compile(r'<c\s+r="([A-Z]+)(\d+)"((?:\s+[\w:]+="[^"]*")*)\s*(?:/>|>(.*?)</c>)', re.S)
ROW_RE = re.compile(r'(<row\s+r="(\d+)"[^>]*?)\s*(?:/>|>(.*?)</row>)', re.S)


def parse_perf(code):
    m = ENT_RE.fullmatch(code) if isinstance(code, str) else None
    if not m:
        return None
    sfx = re.sub(r"([A-Z]+)(\d+)", lambda s: s.group(1) + str(int(s.group(2))), m.group(3))
    return dict(platform=f"DULANG-{m.group(1)}", well=f"DULANG-{m.group(1)}{int(m.group(2))}{sfx}",
                string=m.group(4), reservoir=m.group(6))


def shift_ref(ref):
    return re.sub(r"(\$?)([A-Z]+)(\$?\d+)", lambda x: f"{x.group(1)}{cl(cidx(x.group(2)) + SHIFT)}{x.group(3)}", ref)


def esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def build(src, dst):
    z = zipfile.ZipFile(src)
    wb = openpyxl.load_workbook(src)
    assert wb.sheetnames == ["Sheet1"]
    assert re.search(r'Target="/?(?:xl/)?worksheets/sheet1.xml"', z.read("xl/_rels/workbook.xml.rels").decode())
    ws = wb["Sheet1"]
    assert [c.value for c in ws[1]][:9] == OLD_HDR, [c.value for c in ws[1]]
    ncols = ws.max_column
    decisions, group_rows, review = {}, set(), []
    label = None
    for r in range(2, ws.max_row + 1):
        a, well, code, dv = ws.cell(r, 1).value, ws.cell(r, 2).value, ws.cell(r, PERF).value, ws.cell(r, DATE_COL).value
        if isinstance(a, str) and a.startswith("Well:"):
            group_rows.add(r)
            label = a.split(":", 1)[1].strip()
            assert all(ws.cell(r, c).value is None for c in range(2, ncols + 1)), r
            continue
        if code is None and all(ws.cell(r, c).value is None for c in range(1, ncols + 1)):
            continue
        d = parse_perf(code)
        if d is None:
            review.append((src.name, r, code, "blank", "Perf. Code not in WELL-STRING-BLOCK-RESERVOIR form"))
            d = dict(platform=None, well=None, string=None, reservoir=None)
        else:
            if not (isinstance(well, str) and code.startswith(well + "-")):
                review.append((src.name, r, code, "review", f"Perf. Code does not start with Well {well!r}"))
            if well != label:
                review.append((src.name, r, code, "review", f"row sits under group label {label!r}"))
            if re.fullmatch(r"E0\d.*|E\d{3}.*", d["reservoir"]):
                review.append((src.name, r, code, "review", f"unusual reservoir code {d['reservoir']!r}, kept as written"))
        try:
            d["date"] = datetime.datetime.strptime(dv.strip(), "%d-%b-%Y").date()
        except (AttributeError, ValueError):
            d["date"] = None
            review.append((src.name, r, code, "DATE blank", f"Start Ratio Date {dv!r} not recognised"))
        d["code"] = code
        decisions[r] = d

    styles = z.read("xl/styles.xml").decode()
    ids = [int(x) for x in re.findall(r'numFmtId="(\d+)"', styles)]
    f_date = max(ids + [163]) + 1
    fmt = f'<numFmt numFmtId="{f_date}" formatCode="yyyy-mm-dd"/>'
    if re.search(r"<numFmts\b", styles):
        styles = re.sub(r'<numFmts\s+count="(\d+)"\s*>(.*?)</numFmts>',
                        lambda m: f'<numFmts count="{int(m.group(1)) + 1}">{m.group(2)}{fmt}</numFmts>', styles, flags=re.S)
    else:
        styles = re.sub(r"(<styleSheet\b[^>]*>)", lambda m: m.group(1) + f'<numFmts count="1">{fmt}</numFmts>', styles,
                        count=1)
    nxf = int(re.search(r'<cellXfs\s+count="(\d+)"', styles).group(1))
    s_date = nxf
    styles = re.sub(r'<cellXfs\s+count="\d+"\s*>(.*?)</cellXfs>',
                    lambda m: f'<cellXfs count="{nxf + 1}">{m.group(1)}'
                              f'<xf numFmtId="{f_date}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
                              f'</cellXfs>', styles, flags=re.S)
    assert len(re.findall(r"<xf\b", re.search(r"<cellXfs.*?</cellXfs>", styles, re.S).group(0))) == nxf + 1

    def row_sub(mo):
        head, rn, body = mo.group(1), int(mo.group(2)), mo.group(3) or ""
        head = re.sub(r'spans="1:\d+"', f'spans="1:{ncols + SHIFT}"', head)
        cells = {}
        for c in CELL_RE.finditer(body):
            assert int(c.group(2)) == rn
            cells[cidx(c.group(1))] = (c.group(3), c.group(4))
        out = {}
        for oc, (attrs, inner) in cells.items():
            nc = oc + SHIFT
            out[nc] = f'<c r="{cl(nc)}{rn}"{attrs}/>' if inner is None or not inner.strip() and "<" not in inner \
                else f'<c r="{cl(nc)}{rn}"{attrs}>{inner}</c>'

        def style_of(oc):
            m = re.search(r'\ss="(\d+)"', cells.get(oc, ("", None))[0])
            return f' s="{m.group(1)}"' if m else ""

        def text(c, v, s=""):
            out[c] = f'<c r="{cl(c)}{rn}"{s} t="str"><v>{esc(v)}</v></c>'
        if rn == 1:
            for c, name in enumerate(NEW_HDR, 1):
                text(c, name, style_of(2))
        elif rn in group_rows:
            for c in range(1, SHIFT + 1):
                out[c] = f'<c r="{cl(c)}{rn}"{style_of(1)}/>'
        elif rn in decisions:
            d = decisions[rn]
            text(1, d["code"])
            text(2, "DULANG")
            for c, k in ((3, "platform"), (4, "well"), (5, "string"), (6, "reservoir")):
                if d[k]:
                    text(c, d[k])
            if d["date"] is not None:
                out[7] = f'<c r="G{rn}" s="{s_date}"><v>{(d["date"] - datetime.date(1899, 12, 30)).days}</v></c>'
        return f"{head}>{''.join(out[c] for c in sorted(out))}</row>"

    xml = z.read("xl/worksheets/sheet1.xml").decode()
    pre, rest = re.split(r"<sheetData\s*>", xml, maxsplit=1)
    data, post = rest.split("</sheetData>", 1)
    assert not re.search(r"<(dataValidation|conditionalFormatting|hyperlink|tablePart|legacyDrawing|drawing|autoFilter|"
                         r"sortState)\b", post), post[:300]
    assert "<f>" not in data and "<f " not in data
    data = ROW_RE.sub(row_sub, data)
    pre = re.sub(r'<dimension\s+ref="A1:([A-Z]+)(\d+)"\s*/>',
                 lambda x: f'<dimension ref="A1:{cl(cidx(x.group(1)) + SHIFT)}{x.group(2)}"/>', pre)
    cols = {}
    for c in re.findall(r"<col\b[^>]*/>", pre):
        lo, hi = int(re.search(r'min="(\d+)"', c).group(1)), int(re.search(r'max="(\d+)"', c).group(1))
        cols[lo + SHIFT] = re.sub(r'max="\d+"', f'max="{hi + SHIFT}"', re.sub(r'min="\d+"', f'min="{lo + SHIFT}"', c))
    widths = {1: re.search(r'width="([\d.]+)"', next(c for k, c in cols.items() if k == PERF + SHIFT)).group(1)}
    for t in range(1, SHIFT + 1):
        cols[t] = f'<col min="{t}" max="{t}" width="{widths.get(t, "14.83203125")}" customWidth="1"/>'
    pre = re.sub(r"<cols\b[^>]*>.*?</cols>", "<cols>" + "".join(cols[k] for k in sorted(cols)) + "</cols>", pre, flags=re.S)
    post = re.sub(r'(<mergeCell\s+ref=")([^"]+)(")', lambda x: x.group(1) + shift_ref(x.group(2)) + x.group(3), post)
    sheet = pre + "<sheetData>" + data + "</sheetData>" + post
    parts = {"xl/worksheets/sheet1.xml": sheet, "xl/styles.xml": styles}

    tmp = dst.with_name(dst.stem + ".partial.xlsx")
    with open(tmp, "xb") as fh, zipfile.ZipFile(fh, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in z.infolist():
            zi = zipfile.ZipInfo(info.filename, info.date_time)
            zi.compress_type, zi.external_attr = zipfile.ZIP_DEFLATED, info.external_attr
            zout.writestr(zi, parts[info.filename].encode() if info.filename in parts else z.read(info.filename))
    validate(src, tmp, decisions, group_rows, ncols, parts)
    tmp.rename(dst)
    return decisions, group_rows, review


def validate(src, tmp, decisions, group_rows, ncols, parts):
    zs, zt = zipfile.ZipFile(src), zipfile.ZipFile(tmp)
    assert zs.namelist() == zt.namelist()
    for n in zs.namelist():
        if n not in parts:
            assert zs.read(n) == zt.read(n), n
    a, b = openpyxl.load_workbook(src)["Sheet1"], openpyxl.load_workbook(tmp)["Sheet1"]
    assert a.max_row == b.max_row and b.max_column == ncols + SHIFT
    assert [c.value for c in b[1]][:SHIFT] == NEW_HDR
    for r, (ra, rb) in enumerate(zip(a.iter_rows(), b.iter_rows()), 1):
        assert a.row_dimensions[r].height == b.row_dimensions[r].height, r
        for oc in range(1, ncols + 1):
            x, y = ra[oc - 1], rb[oc - 1 + SHIFT]
            assert x.value == y.value, (r, oc, x.value, y.value)
            assert x.number_format == y.number_format and repr(x.font) == repr(y.font) \
                and repr(x.fill) == repr(y.fill) and repr(x.alignment) == repr(y.alignment), (r, oc)
        new = [c.value for c in rb[:SHIFT]]
        if r in group_rows:
            assert new == [None] * SHIFT and all(repr(c.fill) == repr(ra[0].fill) for c in rb[:SHIFT]), r
        elif r in decisions:
            d = decisions[r]
            g = new[6].date() if isinstance(new[6], datetime.datetime) else new[6]
            assert new[:6] == [d["code"], "DULANG", d["platform"], d["well"], d["string"], d["reservoir"]], (r, new)
            assert new[0] == ra[PERF - 1].value, r
            assert g == d["date"] and (g is None or g.strftime("%d-%b-%Y") == ra[DATE_COL - 1].value), (r, g)
            assert d["date"] is None or rb[6].number_format == "yyyy-mm-dd"
    ma = sorted(str(m) for m in a.merged_cells.ranges)
    mb = sorted(str(m) for m in b.merged_cells.ranges)
    assert sorted(shift_ref(m) for m in ma) == mb
    def widths(ws):
        w = {}
        for dim in ws.column_dimensions.values():
            if dim.min and dim.width:
                for c in range(dim.min, dim.max + 1):
                    w[c] = dim.width
        return w
    wa_, wb2 = widths(a), widths(b)
    assert all(wb2.get(c + SHIFT) == w for c, w in wa_.items()), (wa_, wb2)
    assert b.freeze_panes == a.freeze_panes
    zs.close()
    zt.close()


def main():
    files = sorted(p for p in SRC.glob("*.xlsx") if not p.name.startswith("~$"))
    for p in files:
        for d in (OUT / p.name, BACKUP / p.name):
            if d.exists():
                sys.exit(f"STOP: {d} already exists - not overwritten")
    BACKUP.mkdir(parents=True, exist_ok=True)
    REPORTS.mkdir(exist_ok=True)
    allrev, summary = [], []
    for p in files:
        shutil.copy2(p, BACKUP / p.name)
        decisions, group_rows, review = build(BACKUP / p.name, OUT / p.name)
        allrev += review
        k = collections.Counter
        summary.append([p.name, len(decisions), len(group_rows), len({d["well"] for d in decisions.values()}),
                        str(dict(k(d["platform"] for d in decisions.values()))),
                        str(dict(k(d["string"] for d in decisions.values())))])
        print(f"== {p.name}: {len(decisions)} perforation rows, {len(group_rows)} group rows; review {len(review)}")
    out = openpyxl.Workbook()
    ws = out.active
    ws.title = "Unresolved"
    ws.append(["File", "Row", "Perf. Code", "Issue", "Detail"])
    for r in allrev:
        ws.append(list(r))
    ws2 = out.create_sheet("Summary")
    ws2.append(["File", "Perforation rows", "Well group rows", "Distinct WELL", "PLATFORM counts", "STRING TYPE counts"])
    for s in summary:
        ws2.append(s)
    rp = REPORTS / "split_ratio_standardisation_review.xlsx"
    out.save(rp)
    print("review list:", rp, "| unresolved:", len(allrev))


if __name__ == "__main__":
    main()
