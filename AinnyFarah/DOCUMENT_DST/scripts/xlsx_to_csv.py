#!/usr/bin/env python3
"""Convert the Metadata sheet of DST_Report_Metadata.xlsx to CSV (local only, prints no content)."""
import csv
from pathlib import Path

from openpyxl import load_workbook

OUT_ROOT = Path("/Users/farahyusof/Documents/GitHub/NoramiraSazuki/AinnyFarah/DOCUMENT_DST")
XLSX = OUT_ROOT / "DST_Report_Metadata.xlsx"
CSV = OUT_ROOT / "DST_Report_Metadata.csv"

ws = load_workbook(XLSX, read_only=True)["Metadata"]
rows = [["" if v is None else v for v in r] for r in ws.iter_rows(values_only=True)]

# utf-8-sig so Excel opens it with the right encoding; fields with commas are quoted.
with open(CSV, "w", newline="", encoding="utf-8-sig") as f:
    csv.writer(f, quoting=csv.QUOTE_MINIMAL).writerows(rows)

print(f"Saved: {CSV} ({len(rows) - 1} data rows, {len(rows[0])} columns)")
