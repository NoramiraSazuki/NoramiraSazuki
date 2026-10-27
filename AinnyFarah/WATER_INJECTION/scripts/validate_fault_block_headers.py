#!/usr/bin/env python3
"""Validate the header update against backup_before_header_update/: header order + colour, every value under the same
header in the same row, DULG -> DULANG in FIELD, FAULT BLOCK retained, purple headers/data/styles unchanged."""
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

FOLDER = Path(__file__).resolve().parent.parent / "output_water-injection" / "water_injection_by_fault_block"
BACKUP = FOLDER / "backup_before_header_update"
BLUE = "FF2D7DCE"
NEW_BLUE = ["ORIGINAL WELL NAME", "FIELD", "PLATFORM", "WELL", "STRING TYPE", "RESERVOIR", "DATE", "FAULT BLOCK"]


def table(ws):
    rows = list(ws.iter_rows())
    hdr = [c.value for c in rows[0]]
    return hdr, rows


def style(c):
    return (c.number_format, c.font.b, c.font.color.rgb if c.font.color else None, c.fill.fill_type, c.fill.fgColor.rgb,
            c.alignment.horizontal, c.alignment.wrap_text, c.border.left.style, c.border.bottom.style)


def main():
    allok = True
    for b in sorted(BACKUP.glob("water_injection_*.xlsx")):
        new = FOLDER / b.name
        wo, wn = load_workbook(b), load_workbook(new)
        res = []
        same_sheets = wo.sheetnames == wn.sheetnames
        for so, sn in zip(wo.worksheets, wn.worksheets):
            ho, ro = table(so)
            hn, rn = table(sn)
            col_o = {" ".join(str(h).split()).upper(): i for i, h in enumerate(ho) if h is not None}
            col_n = {h: i for i, h in enumerate(hn) if h is not None}
            purple_o = [i for i, c in enumerate(ro[0]) if c.fill.fgColor.rgb != BLUE]
            purple_n = [i for i, c in enumerate(rn[0]) if c.fill.fgColor.rgb != BLUE]
            checks = {}
            checks["8 blue headers in order"] = hn[:8] == NEW_BLUE and all(rn[0][i].fill.fgColor.rgb == BLUE for i in range(8))
            checks["FAULT BLOCK right after DATE"] = hn.index("FAULT BLOCK") == hn.index("DATE") + 1
            checks["new header formatted like FIELD header"] = style(rn[0][0]) == style(rn[0][1])
            checks["row count unchanged"] = len(ro) == len(rn)
            moved_ok = True
            for name in NEW_BLUE[1:]:
                io, inew = col_o[name], col_n[name]
                for r in range(1, len(ro)):
                    vo = ro[r][io].value
                    if name == "FIELD" and isinstance(vo, str) and vo.strip().upper() == "DULG":
                        vo = "DULANG"
                    if vo != rn[r][inew].value or style(ro[r][io]) != style(rn[r][inew]):
                        moved_ok = False
            checks["blue values + styles under same header and row"] = moved_ok
            checks["no DULG left in FIELD"] = all(rn[r][col_n["FIELD"]].value != "DULG" for r in range(1, len(rn)))
            checks["FAULT BLOCK values retained"] = [r[col_o["FAULT BLOCK"]].value for r in ro[1:]] == [r[col_n["FAULT BLOCK"]].value for r in rn[1:]]
            checks["ORIGINAL WELL NAME blank"] = all(rn[r][0].value is None for r in range(1, len(rn)))
            checks["purple headers same text/order/style"] = [ho[i] for i in purple_o] == [hn[i] for i in purple_n] and \
                all(style(ro[0][i]) == style(rn[0][j]) for i, j in zip(purple_o, purple_n))
            checks["purple data values + styles unchanged"] = all(ro[r][i].value == rn[r][j].value and style(ro[r][i]) == style(rn[r][j])
                                                                  for r in range(1, len(ro)) for i, j in zip(purple_o, purple_n))
            checks["column widths followed columns"] = all(
                so.column_dimensions[get_column_letter(i + 1)].width == sn.column_dimensions[get_column_letter(j + 1)].width
                for i, j in zip(purple_o, purple_n))
            checks["freeze panes kept"] = so.freeze_panes == sn.freeze_panes
            checks["auto-filter covers all columns"] = sn.auto_filter.ref == f"A1:{get_column_letter(len(hn))}{len(rn)}"
            res.append((sn.title, checks))
        for title, checks in res:
            bad = [k for k, v in checks.items() if not v]
            allok &= not bad and same_sheets
            print(f"{b.name:36s} [{title}] {'ALL PASS' if not bad and same_sheets else 'FAIL: ' + str(bad)}")
    print("OVERALL:", "ALL PASS" if allok else "FAILURES — see above")


if __name__ == "__main__":
    main()
