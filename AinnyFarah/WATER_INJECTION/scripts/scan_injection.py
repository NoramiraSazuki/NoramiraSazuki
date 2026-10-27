#!/usr/bin/env python3
"""Read-only scan of every worksheet in the 47 Dulang Unit MBAL workbooks for water-injection content.

Files are read from the NAS into memory (os.open/os.read — nothing written to the NAS). A worksheet is a hit when
its name or any cell text matches inject / injection / inj / WI / CUM_WTR.INJ style terms. Results are written to
scan_results.json (in the output folder) for the extraction step.
"""
import io
import json
import os
import re
import sys
from pathlib import Path

import xlrd
from openpyxl import load_workbook

AREA = os.environ.get("DULANG_AREA", "Dulang Unit")
NAS = Path("/Volumes/homes/public/DulangOSS/20261015/MBAL/03 Others/04 Dulang Network (MBAL)/Dulang MBAL") / AREA
TAG = "" if AREA == "Dulang Unit" else "_" + AREA.replace(" ", "_")
OUT = Path(__file__).resolve().parent.parent / "output_water-injection"
PAT = re.compile(r"inject|\binj\b|\binj[\s._]|w\.?\s*inj|wtr\.?\s*inj|water\s*inj|\bW_?I\b|\bWinj\b|\bCWI\b|cum.?w.?i\b", re.I)


def read(p):
    fd = os.open(p, os.O_RDONLY)
    parts = []
    try:
        while True:
            b = os.read(fd, 1 << 20)
            if not b:
                break
            parts.append(b)
    finally:
        try:
            os.close(fd)
        except OSError:
            pass
    return b"".join(parts)


def sheets(data):
    """yield (sheet name, state, n_rows, n_cols, rows as lists) for xlsx or xls bytes."""
    if data[:2] == b"PK":
        wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
        for ws in wb.worksheets:
            rows = [list(r) for r in ws.iter_rows(values_only=True)]
            yield ws.title, ws.sheet_state, rows
    else:
        wb = xlrd.open_workbook(file_contents=data)
        for s in wb.sheets():
            rows = [s.row_values(i) for i in range(s.nrows)]
            yield s.name, ("visible" if s.visibility == 0 else "hidden"), rows


def main():
    files = sorted(p for p in NAS.rglob("*") if p.is_file() and p.suffix.lower() in (".xlsx", ".xls") and not p.name.startswith("~$"))
    res = []
    for p in files:
        rel = str(p.relative_to(NAS))
        try:
            data = read(p)
            for name, state, rows in sheets(data):
                hits = []
                if PAT.search(name):
                    hits.append(("(sheet name)", name))
                for i, r in enumerate(rows):
                    for j, v in enumerate(r):
                        if isinstance(v, str) and PAT.search(v):
                            hits.append((f"R{i + 1}C{j + 1}", v.strip()[:80]))
                nonempty = sum(1 for r in rows for v in r if v not in (None, ""))
                res.append(dict(file=rel, size=len(data), fmt="xlsx" if data[:2] == b"PK" else "xls", sheet=name, state=state,
                                rows=len(rows), cols=max((len(r) for r in rows), default=0), cells=nonempty, hits=hits[:40], n_hits=len(hits)))
        except Exception as e:  # noqa: BLE001
            res.append(dict(file=rel, sheet=None, error=f"{type(e).__name__}: {e}"))
    OUT.mkdir(exist_ok=True)
    (OUT / f"scan_results{TAG}.json").write_text(json.dumps(res, indent=1, default=str))
    nf = len(files)
    hit_sheets = [r for r in res if r.get("n_hits")]
    print(f"workbooks {nf}, worksheets {sum(1 for r in res if r.get('sheet'))}, errors {sum(1 for r in res if 'error' in r)}, sheets with hits {len(hit_sheets)}")
    for r in res:
        if "error" in r:
            print("ERROR", r["file"], r["error"])
    for r in hit_sheets:
        print(f"\n## {r['file']} :: [{r['sheet']}] {r['state']} {r['rows']}x{r['cols']} hits={r['n_hits']}")
        for h in r["hits"][:8]:
            print("   ", h)


if __name__ == "__main__":
    main()
