#!/usr/bin/env python3
"""Group tesseract TSV words into visual lines: prints y, x and text (source-image pixel coordinates). Usage: ocr_lines.py TSV SCALE"""
import csv, sys
rows = [r for r in csv.DictReader(open(sys.argv[1]), delimiter="\t", quoting=csv.QUOTE_NONE) if r["text"].strip() and float(r["conf"]) > 0]
s = float(sys.argv[2])
rows.sort(key=lambda r: (int(r["top"]) + int(r["height"]) / 2))
lines = []
for r in rows:
    yc = (int(r["top"]) + int(r["height"]) / 2) / s
    if lines and abs(lines[-1][0] - yc) < 5:
        lines[-1][1].append(r)
    else:
        lines.append([yc, [r]])
for yc, ws in lines:
    ws.sort(key=lambda r: int(r["left"]))
    segs, cur = [], [ws[0]]
    for w in ws[1:]:
        prev = cur[-1]
        if int(w["left"]) - (int(prev["left"]) + int(prev["width"])) > 40 * s / 3:
            segs.append(cur); cur = [w]
        else:
            cur.append(w)
    segs.append(cur)
    print(f"{yc:6.0f} " + " || ".join(f"[{int(int(g[0]['left'])/s)}] " + " ".join(w['text'] for w in g) for g in segs))
