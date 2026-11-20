#!/usr/bin/env python3
"""Remove the merged "Well: ..." group rows from the standardised Perforation Split Ratio outputs.

A row is removed only when it holds nothing but the group label (old column A, now H), the label equals the Well
value of every perforation row beneath it, and at least one perforation row follows it, so no well name is lost.
The remaining rows keep their order, values and formatting and are renumbered; the merges (all on group rows) go.
Current outputs are copied to backup_before_group_row_removal/ first; each result is validated before it replaces
the output.
"""
import os
import re
import shutil
import sys
import zipfile
from pathlib import Path

import openpyxl
from openpyxl.utils import column_index_from_string as cidx

OUT = Path(__file__).resolve().parent / "all_data_standardisation" / "01_UPDATED" / "PERFORATION-SPLIT-RATIOS"
BACKUP = OUT / "backup_before_group_row_removal"
LABEL_COL, WELL_COL = 8, 9  # H = old group label column, I = Well
ROW_RE = re.compile(r'<row\s+r="(\d+)"[^>]*?(?:/>|>.*?</row>)', re.S)


def group_rows(ws):
    rows, label, seen = [], None, {}
    for r in range(2, ws.max_row + 1):
        vals = [c.value for c in ws[r]]
        lab = vals[LABEL_COL - 1]
        if isinstance(lab, str) and lab.startswith("Well:"):
            others = [v for i, v in enumerate(vals) if i != LABEL_COL - 1 and v is not None]
            assert not others, (r, others)
            label = lab.split(":", 1)[1].strip()
            rows.append(r)
            seen[r] = 0
            continue
        assert lab is None, (r, lab)
        assert vals[WELL_COL - 1] == label, (r, vals[WELL_COL - 1], label)
        seen[rows[-1]] += 1
    empty = [r for r, n in seen.items() if n == 0]
    assert not empty, f"group rows with no perforation row beneath: {empty}"
    return rows


def process(path):
    ws = openpyxl.load_workbook(path)["Sheet1"]
    drop = set(group_rows(ws))
    merges = sorted(str(m) for m in ws.merged_cells.ranges)
    assert all(int(re.match(r"[A-Z]+(\d+)", m).group(1)) in drop for m in merges) and len(merges) == len(drop)

    z = zipfile.ZipFile(path)
    xml = z.read("xl/worksheets/sheet1.xml").decode()
    pre, rest = re.split(r"<sheetData\s*>", xml, maxsplit=1)
    data, post = rest.split("</sheetData>", 1)
    renum, kept = {}, 0
    for m in ROW_RE.finditer(data):
        r = int(m.group(1))
        if r not in drop:
            kept += 1
            renum[r] = kept

    def row_sub(m):
        r = int(m.group(1))
        if r in drop:
            return ""
        n = renum[r]
        s = m.group(0)
        s = re.sub(r'(<row\s+r=")\d+"', rf'\g<1>{n}"', s, count=1)
        return re.sub(r'(<c\s+r="[A-Z]+)\d+"', rf'\g<1>{n}"', s)
    data = ROW_RE.sub(row_sub, data)
    pre = re.sub(r'(<dimension\s+ref="A1:[A-Z]+)\d+"', rf'\g<1>{kept}"', pre)
    post = re.sub(r"<mergeCells\b[^>]*>.*?</mergeCells>|<mergeCells\b[^>]*/>", "", post, flags=re.S)
    assert "<mergeCell" not in post and not re.search(r'\b(ref|sqref)="', post), post[:300]
    sheet = pre + "<sheetData>" + data + "</sheetData>" + post

    tmp = path.with_name(path.stem + ".partial.xlsx")
    with open(tmp, "xb") as fh, zipfile.ZipFile(fh, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in z.infolist():
            zi = zipfile.ZipInfo(info.filename, info.date_time)
            zi.compress_type, zi.external_attr = zipfile.ZIP_DEFLATED, info.external_attr
            zout.writestr(zi, sheet.encode() if info.filename == "xl/worksheets/sheet1.xml" else z.read(info.filename))

    # validate: the new sheet is exactly the old one with the group rows taken out
    zt = zipfile.ZipFile(tmp)
    for n in z.namelist():
        if n != "xl/worksheets/sheet1.xml":
            assert z.read(n) == zt.read(n), n
    b = openpyxl.load_workbook(tmp)["Sheet1"]
    keep = [r for r in range(1, ws.max_row + 1) if r not in drop]
    assert b.max_row == len(keep) and b.max_column == ws.max_column and not b.merged_cells.ranges
    for new_r, old_r in enumerate(keep, 1):
        assert ws.row_dimensions[old_r].height == b.row_dimensions[new_r].height
        for c in range(1, ws.max_column + 1):
            x, y = ws.cell(old_r, c), b.cell(new_r, c)
            assert x.value == y.value and x.number_format == y.number_format and repr(x.font) == repr(y.font) \
                and repr(x.fill) == repr(y.fill), (old_r, c)
    assert all(b.cell(r, LABEL_COL).value is None for r in range(2, b.max_row + 1))
    assert b.freeze_panes == ws.freeze_panes
    zt.close()
    z.close()
    os.replace(tmp, path)
    return len(drop), len(keep) - 1


def main():
    files = sorted(p for p in OUT.glob("*.xlsx") if not p.name.startswith("~$"))
    for p in files:
        if (BACKUP / p.name).exists():
            sys.exit(f"STOP: {BACKUP / p.name} already exists - not overwritten")
    BACKUP.mkdir(exist_ok=True)
    for p in files:
        shutil.copy2(p, BACKUP / p.name)
    for p in files:
        removed, rows = process(p)
        print(f"{p.name}: {removed} group rows removed, {rows} perforation rows kept")


if __name__ == "__main__":
    main()
