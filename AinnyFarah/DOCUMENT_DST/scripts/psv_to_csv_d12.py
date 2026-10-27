#!/usr/bin/env python3
"""Convert a staged D-12 CRG printout transcription (.psv) into a CSV.

Staging format: '#SECTION|<section title>' lines set the current section;
data lines are 'page|date|time|bhct|bhp|event'.
Usage: psv_to_csv_d12.py IN.psv OUT.csv
"""
import csv
import sys

src, out = sys.argv[1], sys.argv[2]
section = ""
rows = []
for line in open(src, encoding="utf-8"):
    line = line.rstrip("\n")
    if not line.strip():
        continue
    if line.startswith("#SECTION|"):
        section = line.split("|", 1)[1]
        continue
    page, date, time, bhct, bhp, event = line.split("|", 5)
    rows.append([page, section, date, time, bhct, bhp, event])

with open(out, "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f)
    w.writerow(["Source Page", "Section (as printed)", "Date", "Time", "BHCT (degF)", "BHP (PSIA)", "Event / Remark"])
    w.writerows(rows)

# Sanity checks: printed for review, not applied.
prev = None
for r in rows:
    if r[4] and r[5]:
        t, p = float(r[4]), float(r[5])
        if not (150 <= t <= 260 and 500 <= p <= 3000):
            print("RANGE?", r)
        if prev and r[1] == prev[1] and abs(p - prev[2]) > 60 and not any(x[6] for x in rows[max(0, rows.index(r) - 3):rows.index(r)]):
            print("JUMP?", prev[0], "->", r[0], r[3], prev[2], "->", p)
        prev = (r[0], r[1], p)
print(len(rows), "rows")
