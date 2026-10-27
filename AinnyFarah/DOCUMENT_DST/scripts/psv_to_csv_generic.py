#!/usr/bin/env python3
"""Convert a staged pipe-separated transcription (.psv) to CSV with a given header.

Usage: psv_to_csv_generic.py IN.psv OUT.csv "Col1|Col2|..."
Each data line must have exactly as many '|' fields as the header. Lines starting
with '#' are comments. Rows with a wrong field count are reported and the run fails.
"""
import csv
import sys

src, out, header = sys.argv[1], sys.argv[2], sys.argv[3].split("|")
rows, bad = [], []
for n, line in enumerate(open(src, encoding="utf-8"), 1):
    line = line.rstrip("\n")
    if not line.strip() or line.startswith("#"):
        continue
    f = line.split("|")
    # Tolerate only trailing EMPTY cells being short/extra; anything else is an error.
    while len(f) > len(header) and f[-1] == "":
        f.pop()
    if len(f) < len(header):
        f += [""] * (len(header) - len(f))
    (rows if len(f) == len(header) else bad).append(f if len(f) == len(header) else (n, len(f), line[:80]))
if bad:
    for b in bad:
        print("FIELD COUNT", b)
    sys.exit(1)
with open(out, "w", newline="", encoding="utf-8") as fh:
    w = csv.writer(fh)
    w.writerow(header)
    w.writerows(rows)
print(len(rows), "rows ->", out)
