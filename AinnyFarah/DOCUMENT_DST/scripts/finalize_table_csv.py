#!/usr/bin/env python3
"""Finalize table_csv/: fill row counts in the index, write page coverage, validate, add UTF-8 BOM."""
import csv
from pathlib import Path

D = Path("/Users/farahyusof/Documents/GitHub/NoramiraSazuki/AinnyFarah/DOCUMENT_DST/table_csv")

def read(p):
    return list(csv.reader(open(p, encoding="utf-8-sig", newline="")))

def write(p, rows):
    with open(p, "w", encoding="utf-8-sig", newline="") as f:
        csv.writer(f).writerows(rows)

tables = sorted(p for p in D.glob("T*.csv"))
counts, pages_used = {}, {}
for t in tables:
    rows = read(t)
    w = len(rows[0])
    bad = [i for i, r in enumerate(rows) if len(r) != w]
    assert not bad, f"{t.name}: ragged rows {bad[:5]}"
    counts[t.name] = len(rows) - 1
    for r in rows[1:]:
        pages_used.setdefault(int(r[0]), set()).add(t.name)

idx = read(D / "_table_index.csv")
hdr, body = idx[0], idx[1:]
seen = set()
out = [hdr]
for r in body:
    if r[0] in seen:
        continue
    seen.add(r[0])
    r[4] = str(counts[r[0]])
    out.append(r)
missing = set(counts) - seen
assert not missing, f"not indexed: {missing}"
write(D / "_table_index.csv", out)

nontable = {
    1: "Cover", 2: "Transmittal memo (text)", 3: "Title page", 4: "Table of contents", 5: "Summary (text)",
    6: "Summary (text)", 7: "Discussion (text)", 8: "Discussion (text)", 9: "Test procedure (text)",
    10: "Test equipment / discussion (text)", 11: "Discussion of results (text)", 12: "Discussion of results (text)",
    13: "Divider: TABLES", 18: "Divider: FIGURES", 19: "Figure 1: tail pipe assembly sketch",
    20: "Figure 2: surface equipment layout sketch", 21: "Figure 3: bottomhole pressure gauge assembly sketch",
    22: "Figure 4: Horner plot (chart; header parameters duplicated in Table 3)",
    23: "Figure 5: type-curve plot (chart; header parameters duplicated in Table 4)", 24: "Divider: APPENDIX A",
    61: "Divider: APPENDIX B", 83: "Divider: APPENDIX C", 109: "Divider: APPENDIX D",
}
cov = [["Page", "Status", "CSV file(s) / Description"]]
for p in range(1, 124):
    if p in pages_used:
        cov.append([p, "Table extracted", "; ".join(sorted(pages_used[p]))])
    elif p in nontable:
        cov.append([p, "No table", nontable[p]])
    else:
        cov.append([p, "NOT EXTRACTED - CHECK", ""])
write(D / "_page_coverage.csv", cov)

flags = read(D / "_review_flags.csv")
write(D / "_review_flags.csv", flags)
for t in tables:
    write(t, read(t))

print(len(tables), "table CSVs;", sum(counts.values()), "data rows;", len(flags) - 1, "review flags")
print("pages with tables:", len(pages_used), "| no-table pages:", len(nontable),
      "| unaccounted:", [r[0] for r in cov[1:] if r[1].startswith("NOT")])
