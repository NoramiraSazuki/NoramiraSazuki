#!/usr/bin/env python3
"""Group EMF text records (emf/imageN.tsv: x, y, text) into rows by y: emf_rows.py emf/imageN.tsv [ytol]"""
import sys
tol = float(sys.argv[2]) if len(sys.argv) > 2 else 6
recs = []
for line in open(sys.argv[1]):
    p = line.rstrip("\n").split("\t")
    if len(p) >= 3 and p[2].strip():
        recs.append((float(p[1]), float(p[0]), p[2].strip()))
recs.sort()
rows = []
for y, x, t in recs:
    if rows and y - rows[-1][0] <= tol:
        rows[-1][1].append((x, t))
    else:
        rows.append([y, [(x, t)]])
for y, items in rows:
    print(f"{y:6.0f} " + " | ".join(f"[{x:.0f}]{t}" for x, t in sorted(items)))
