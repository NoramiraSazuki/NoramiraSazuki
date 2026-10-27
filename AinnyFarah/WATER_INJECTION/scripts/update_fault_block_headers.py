#!/usr/bin/env python3
"""Re-order the blue identification columns of every water_injection_<FB>.xlsx in water_injection_by_fault_block.

New blue order: ORIGINAL WELL NAME | FIELD | PLATFORM | WELL | STRING TYPE | RESERVOIR | DATE | FAULT BLOCK, then every
purple column in its existing order. Columns are matched by header text (not position); each cell moves with its value
and full style. ORIGINAL WELL NAME is added (blank) when absent — never built from WELL. FIELD 'DULG' -> 'DULANG' only.
Column widths, freeze panes and the auto-filter follow the moved columns. Backups go to backup_before_header_update/.
"""
import copy
import shutil
import sys
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

FOLDER = Path(__file__).resolve().parent.parent / "output_water-injection" / "water_injection_by_fault_block"
BACKUP = FOLDER / "backup_before_header_update"
BLUE = "FF2D7DCE"
NEW_BLUE = ["ORIGINAL WELL NAME", "FIELD", "PLATFORM", "WELL", "STRING TYPE", "RESERVOIR", "DATE", "FAULT BLOCK"]


def norm(v):
    return " ".join(str(v).split()).upper() if v is not None else ""


def is_blue(c):
    return c.fill is not None and c.fill.fill_type == "solid" and c.fill.fgColor.rgb == BLUE


def plan(ws):
    """Return (mapping new_col -> old_col or None, notes) or raise ValueError if the sheet cannot be updated safely."""
    hdr = {c.column: c for c in ws[1] if c.value is not None}
    blue = {norm(c.value): col for col, c in hdr.items() if is_blue(c)}
    purple = [col for col, c in sorted(hdr.items()) if not is_blue(c)]
    need = NEW_BLUE[1:]
    missing = [h for h in need if h not in blue]
    if missing:
        raise ValueError(f"blue header(s) not found: {missing}")
    extra = [h for h in blue if h not in NEW_BLUE]
    if extra:
        raise ValueError(f"unexpected blue header(s): {extra}")
    if max(blue.values()) > min(purple):
        raise ValueError("blue and purple header blocks are interleaved")
    order = [blue.get("ORIGINAL WELL NAME")] + [blue[h] for h in need] + purple
    return order


def update(path):
    wb = load_workbook(path)
    report = []
    for ws in wb.worksheets:
        try:
            order = plan(ws)
        except ValueError as e:
            report.append((ws.title, "NOT UPDATED", str(e)))
            continue
        maxr, maxc = ws.max_row, ws.max_column
        field_old = [i for i, c in enumerate(ws[1], 1) if norm(c.value) == "FIELD"][0]
        snap = {}                                            # (row, old col) -> (value, style, number_format, hyperlink, comment)
        for row in ws.iter_rows(min_row=1, max_row=maxr, max_col=maxc):
            for c in row:
                snap[(c.row, c.column)] = (c.value, copy.copy(c._style), c.hyperlink, c.comment)
        widths = {col: ws.column_dimensions[get_column_letter(col)].width for col in range(1, maxc + 1)}
        dulg = other = 0
        others = {}
        for r in range(1, maxr + 1):
            for new_col, old_col in enumerate(order, 1):
                tgt = ws.cell(r, new_col)
                if old_col is None:                          # new ORIGINAL WELL NAME column
                    v, st, hl, cm = snap[(r, field_old)]     # take the FIELD cell's formatting for this row
                    tgt._style = copy.copy(st)
                    tgt.value = "ORIGINAL WELL NAME" if r == 1 else None
                    tgt.hyperlink = None; tgt.comment = None
                    continue
                v, st, hl, cm = snap[(r, old_col)]
                if r == 1 and new_col <= len(NEW_BLUE):
                    v = NEW_BLUE[new_col - 1]                # exact uppercase header text
                if r > 1 and old_col == field_old and isinstance(v, str):
                    if v.strip().upper() == "DULG":
                        v = "DULANG"; dulg += 1
                    elif v != "DULANG":
                        others[v] = others.get(v, 0) + 1
                tgt.value = v
                tgt._style = copy.copy(st)
                tgt.hyperlink = hl
                tgt.comment = cm
        for new_col, old_col in enumerate(order, 1):
            w = widths.get(old_col if old_col else field_old)
            if w:
                ws.column_dimensions[get_column_letter(new_col)].width = w
        if ws.auto_filter.ref:
            ws.auto_filter.ref = f"A1:{get_column_letter(len(order))}{maxr}"
        report.append((ws.title, "UPDATED", {"rows": maxr - 1, "DULG->DULANG": dulg, "other FIELD values": others,
                                              "ORIGINAL WELL NAME": "added (blank — not in file)" if order[0] is None else "existing, kept"}))
    wb.save(path)
    return report


def main():
    files = sorted(p for p in FOLDER.glob("water_injection_*.xlsx") if not p.name.startswith("~$"))
    locks = sorted(FOLDER.glob("~$*"))
    if locks:
        raise SystemExit(f"files open in Excel: {[l.name for l in locks]} — close them first")
    if BACKUP.exists():
        raise SystemExit(f"{BACKUP} already exists — not overwriting earlier backups")
    BACKUP.mkdir()
    for p in files:
        shutil.copy2(p, BACKUP / p.name)
    print(f"backed up {len(files)} files -> {BACKUP}")
    for p in files:
        for sheet, status, info in update(p):
            print(f"{p.name:36s} [{sheet}] {status}: {info}")


if __name__ == "__main__":
    main()
