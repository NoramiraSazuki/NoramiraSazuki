#!/usr/bin/env python3
"""Rename the platform in the PLATFORM column of every Excel workbook in this folder (in place, backups first):
DULG-A -> DULANG-A (B, C, D likewise). Only cells under a header cell exactly 'PLATFORM' change; same package-level
edit and cell-by-cell validation as standardise_well_names.py. Workbooks open in Excel (a ~$ lock file exists) are
skipped.
"""
import re
import shutil
import sys

from standardise_well_names import DIR, process

BACKUP = DIR / "backup_before_platform_update"


def conv(v):
    m = re.fullmatch(r"(?:DULG|DULANG)-([A-Z])", v)
    return f"DULANG-{m.group(1)}" if m else None


def main():
    files = sorted(p for p in DIR.iterdir() if p.is_file() and p.suffix.lower() in (".xlsx", ".xlsm")
                   and not p.name.startswith("~$"))
    open_now = [p for p in files if (p.parent / ("~$" + p.name)).exists()]
    files = [p for p in files if p not in open_now]
    BACKUP.mkdir(exist_ok=True)
    for p in files:
        if (BACKUP / p.name).exists():
            sys.exit(f"STOP: backup {BACKUP / p.name} already exists - not overwritten")
    for p in files:
        shutil.copy2(p, BACKUP / p.name)
    for p in files:
        try:
            report, saved = process(p, BACKUP / p.name, "PLATFORM", conv)
        except Exception as e:  # leave the file as it is
            print(f"SKIPPED {p.name}: {e!r}")
            continue
        print(f"== {p.name}: {'saved' if saved else 'no change'}")
        for s, r in report.items():
            if r is None:
                print(f"   sheet '{s}': no PLATFORM column")
                continue
            print(f"   sheet '{s}': PLATFORM = column {r['col']} (header row {r['header_row']}), "
                  f"{sum(r['changed'].values())} cells updated; other: {dict(r['kinds'])}; "
                  f"review: {dict(r['review']) or 'none'}")
            print("     " + ", ".join(f"{o}->{n} ({c})" for (o, n), c in sorted(r["changed"].items())))
    for p in open_now:
        print(f"SKIPPED {p.name}: open in Excel (~$ lock file present)")


if __name__ == "__main__":
    main()
