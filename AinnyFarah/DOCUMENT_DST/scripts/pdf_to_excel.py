#!/usr/bin/env python3
"""Convert one scanned DST report PDF into an Excel workbook using local Tesseract.

- Renders each page (from the rotation-corrected searchable PDF) with Ghostscript.
- Runs Tesseract in TSV mode to get every word with its position and confidence.
- Rebuilds lines and columns from word positions, so tables land in separate cells.
- Writes one sheet per page plus an "All_Text" sheet (Page, Line, Text, Confidence).

Fully offline. Prints only progress counts, never document text.
"""
import csv
import io
import os
import re
import shutil
import subprocess
import sys
from concurrent.futures import ProcessPoolExecutor
from pathlib import Path
from statistics import median

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

OUT_ROOT = Path("/Users/farahyusof/Documents/GitHub/NoramiraSazuki/AinnyFarah/DOCUMENT_DST")
NAME = "DULANG-7_EPMI Well Test Report DST 1,1A,2,2A_ 01 Dec. 1990"
SRC_PDF = OUT_ROOT / "searchable_pdf" / f"{NAME}_OCR.pdf"      # rotation-corrected copy
OUT_XLSX = OUT_ROOT / f"{NAME}.xlsx"
TMP_DIR = OUT_ROOT / "tmp" / "pdf_to_excel"
DPI = 300

LOW_CONF = 60          # flag lines whose mean word confidence is below this
NUM_RE = re.compile(r"^-?(?:0|[1-9]\d*)(?:\.\d+)?$")


def page_count(pdf):
    out = subprocess.run(["qpdf", "--show-npages", str(pdf)], capture_output=True, text=True, check=True)
    return int(out.stdout.strip())


def ocr_page(page_no):
    """Render one page and return its words as dicts (text, left, top, width, height, conf)."""
    env = dict(os.environ, TMPDIR=str(TMP_DIR))
    png = TMP_DIR / f"p{page_no:04d}.png"
    subprocess.run(["gs", "-q", "-dNOPAUSE", "-dBATCH", "-dSAFER", "-sDEVICE=pnggray", f"-r{DPI}",
                    f"-dFirstPage={page_no}", f"-dLastPage={page_no}", f"-sOutputFile={png}", str(SRC_PDF)],
                   check=True, env=env, capture_output=True)
    # psm 6 = treat the page as one uniform block; keeps table rows together better than auto layout.
    tsv = subprocess.run(["tesseract", str(png), "stdout", "-l", "eng", "--psm", "6", "tsv"],
                         capture_output=True, text=True, env=dict(env, OMP_THREAD_LIMIT="1")).stdout
    png.unlink(missing_ok=True)
    words = []
    for r in csv.DictReader(io.StringIO(tsv), delimiter="\t", quoting=csv.QUOTE_NONE):
        if r.get("level") != "5" or not (r.get("text") or "").strip():
            continue
        words.append({"text": r["text"].strip(), "left": int(r["left"]), "top": int(r["top"]),
                      "width": int(r["width"]), "height": int(r["height"]), "conf": float(r["conf"])})
    return page_no, words


def build_rows(words):
    """Group words into lines by vertical position, then split lines into cells and align columns."""
    if not words:
        return [], []
    h = median(w["height"] for w in words)
    words = sorted(words, key=lambda w: (w["top"] + w["height"] / 2, w["left"]))

    # 1) Lines: words whose vertical centres are within ~half a text height.
    lines = []
    for w in words:
        cy = w["top"] + w["height"] / 2
        if lines and abs(cy - lines[-1]["cy"]) <= 0.5 * h:
            ln = lines[-1]
            ln["words"].append(w)
            ln["cy"] = sum(x["top"] + x["height"] / 2 for x in ln["words"]) / len(ln["words"])
        else:
            lines.append({"cy": cy, "words": [w]})

    # 2) Cells: split a line where the horizontal gap is wider than ~1.5 text heights.
    gap = 1.5 * h
    line_cells = []
    for ln in lines:
        ws = sorted(ln["words"], key=lambda w: w["left"])
        cells, cur = [], [ws[0]]
        for prev, w in zip(ws, ws[1:]):
            if w["left"] - (prev["left"] + prev["width"]) > gap:
                cells.append(cur)
                cur = [w]
            else:
                cur.append(w)
        cells.append(cur)
        line_cells.append([{"left": c[0]["left"], "text": " ".join(x["text"] for x in c),
                            "conf": [x["conf"] for x in c]} for c in cells])

    # 3) Columns: cluster cell left edges across the page.
    lefts = sorted(c["left"] for cells in line_cells for c in cells)
    tol = 2 * h
    anchors = []
    for x in lefts:
        if anchors and x - anchors[-1][-1] <= tol:
            anchors[-1].append(x)
        else:
            anchors.append([x])
    col_x = [median(a) for a in anchors]

    rows, confs = [], []
    for cells in line_cells:
        row = [""] * len(col_x)
        for c in cells:
            idx = min(range(len(col_x)), key=lambda i: abs(col_x[i] - c["left"]))
            # Never overwrite a filled cell; shift right to the next free column.
            while idx < len(row) and row[idx] != "":
                idx += 1
            if idx == len(row):
                row.append("")
            row[idx] = c["text"]
        while row and row[-1] == "":
            row.pop()
        rows.append(row)
        allc = [x for c in cells for x in c["conf"] if x >= 0]
        confs.append(round(sum(allc) / len(allc), 1) if allc else None)
    return rows, confs


def as_value(s):
    """Plain numbers become numeric cells; everything else stays text exactly as OCR'd."""
    return float(s) if "." in s and NUM_RE.match(s) else int(s) if NUM_RE.match(s) else s


def main():
    if not SRC_PDF.exists():
        sys.exit(f"Missing searchable PDF: {SRC_PDF}")
    TMP_DIR.mkdir(parents=True, exist_ok=True)
    n = page_count(SRC_PDF)
    print(f"Pages: {n}")

    results = {}
    with ProcessPoolExecutor(max_workers=os.cpu_count() or 2) as ex:
        for i, (pno, words) in enumerate(ex.map(ocr_page, range(1, n + 1)), 1):
            results[pno] = words
            if i % 20 == 0 or i == n:
                print(f"  OCR done: {i}/{n}")

    wb = Workbook()
    summary = wb.active
    summary.title = "All_Text"
    summary.append(["Page", "Line", "Text", "Mean OCR Confidence", "Low Confidence (<%d)" % LOW_CONF])
    header_fill = PatternFill("solid", fgColor="DDEBF7")
    low_fill = PatternFill("solid", fgColor="FFF2CC")

    empty_pages, total_lines, low_lines = [], 0, 0
    for pno in range(1, n + 1):
        rows, confs = build_rows(results[pno])
        ws = wb.create_sheet(f"Page {pno:03d}")
        if not rows:
            empty_pages.append(pno)
            ws.append(["(no text recognised on this page)"])
        for row, conf in zip(rows, confs):
            ws.append([as_value(v) for v in row])
            low = conf is not None and conf < LOW_CONF
            # Line number is filled in after all pages are written.
            summary.append([pno, None, " | ".join(v for v in row if v), conf, "Yes" if low else ""])
            total_lines += 1
            low_lines += low
            if low:
                for cell in ws[ws.max_row]:
                    cell.fill = low_fill
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = min(
                max(len(str(c.value)) for c in col if c.value is not None) + 2 if any(c.value is not None for c in col) else 10, 60)

    # Line numbers per page in All_Text.
    counter = {}
    for r in summary.iter_rows(min_row=2):
        counter[r[0].value] = counter.get(r[0].value, 0) + 1
        r[1].value = counter[r[0].value]
        if r[4].value == "Yes":
            for cell in r:
                cell.fill = low_fill

    for cell in summary[1]:
        cell.font = Font(bold=True)
        cell.fill = header_fill
    for col, w in zip("ABCDE", [7, 7, 120, 12, 14]):
        summary.column_dimensions[col].width = w
    for r in summary.iter_rows(min_row=2, min_col=3, max_col=3):
        r[0].alignment = Alignment(wrap_text=True, vertical="top")
    summary.freeze_panes = "A2"
    summary.auto_filter.ref = summary.dimensions

    wb.save(OUT_XLSX)
    shutil.rmtree(TMP_DIR, ignore_errors=True)
    try:
        TMP_DIR.parent.rmdir()
    except OSError:
        pass
    print(f"Lines: {total_lines} (low confidence: {low_lines}); pages with no text: {empty_pages or 'none'}")
    print(f"Saved: {OUT_XLSX}")


if __name__ == "__main__":
    main()
