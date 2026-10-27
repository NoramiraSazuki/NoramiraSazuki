#!/usr/bin/env python3
"""v3 of 'ARPR 1.1.2025 YEAR Presentation Pack_Dulang - for AOF' — tables reproduced as they look on the slide.

Base     : the user-saved v2 workbook (copied; v2 itself is not modified). All original sheets are kept as-is
           (Table_002 stays removed). The 7 v2 extraction sheets + 'Update Log v2'/'Update Review v2' are replaced.
Source   : current NAS PDF (63 pages, physical page numbers from 1) — read-only, verified by sha256.
Tables   : cell geometry from the PDF itself —
             * vector slides: each cell = drawn cell rectangle / ruled line grid (merged cells = larger rectangles),
               fill colour from the drawing;
             * picture slides: grid lines and merged cells detected from image pixels, fill sampled from the image;
             * text-only slides: word-layout tables (flagged).
           Cell text: native PDF words inside the cell, else Tesseract OCR of that single cell at 400 dpi.
Layout   : each slide is laid onto a fine Excel column grid (1 column = UNIT pt) so column widths, row heights
           and merged cells keep the slide's proportions; fills/borders reproduce the table appearance.
Edits    : manually edited values in existing table sheets for the requested slides are carried into the new
           tables (never overwritten by OCR) and flagged.
Console output: counts/status only.
"""
import hashlib
import json
import math
import os
import re
import shutil
import subprocess
import sys
from collections import Counter, defaultdict
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import numpy as np
import pymupdf
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from PIL import Image

sys.path.insert(0, str(Path(__file__).parent))
import pdf_to_excel as P  # noqa: E402
import update_presentation_pack_v2 as U  # noqa: E402  (line/text helpers)

STEM = "ARPR 1.1.2025 YEAR Presentation Pack_Dulang - for AOF"
BASE_WB = P.OUT / f"{STEM}_v2.xlsx"
OUT_WB = P.OUT / f"{STEM}_v3.xlsx"
W = P.WORK / "_update_v3"
SRC_PDF = P.SRC / f"{STEM}.pdf"
ORIG_63 = P.WORK / STEM / "_working_copy.pdf"
REGEN_OLD = P.WORK / "_regen_old" / "regen_old.xlsx"
V2_SHEETS = ["Action Items S5-7", "Field Resources AG S20", "Project Resources S23", "Slide 24 Tables",
             "Slides 49-50", "Slides 52-55", "Slides 57-61", "Update Log v2", "Update Review v2"]
GROUPS = [
    ("Action Items S5-7", [5, 6, 7], "Summary for Action Item", "combine"),
    ("Field Resources AG S20", [20], "Field Resources Summary (AG)", "single"),
    ("Project Resources S23", [23], "Project Resources Summary", "single"),
    ("Slide 24 Tables", [24], None, "all"),
    ("Slides 49-50", [49, 50], None, "all"),
    ("Slides 52-55", [52, 53, 54, 55], None, "all"),
    ("Slides 57-61", [57, 58, 59, 60, 61], None, "all"),
]
LOG_SHEET, REVIEW_SHEET = "Update Log v3", "Update Review v3"
UNIT = 6.0                 # pt per Excel grid column
COL0 = 3                   # A = source page, B = item type, C.. = slide grid
CELL_DPI = 400
SNAP = 2.5
THIN = Side(style="thin", color="7F7F7F")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def sha_bytes(b):
    return hashlib.sha256(b).hexdigest()


# --------------------------------------------------------------------------- geometry helpers
def snap_values(vals, tol=SNAP):
    vals = sorted(vals)
    groups = []
    for v in vals:
        if groups and v - groups[-1][-1] <= tol:
            groups[-1].append(v)
        else:
            groups.append([v])
    return [sum(g) / len(g) for g in groups]


def idx_of(bounds, v):
    return min(range(len(bounds)), key=lambda i: abs(bounds[i] - v))


def iou(a, b):
    i = P.inter(a, b)
    u = P.rect_area(a) + P.rect_area(b) - i
    return i / u if u else 0


def rgb_hex(c):
    if c is None:
        return None
    r, g, b = [max(0, min(255, int(round(x * 255 if max(c) <= 1 else x)))) for x in c[:3]]
    return f"{r:02X}{g:02X}{b:02X}"


def luminance(hexrgb):
    if not hexrgb:
        return 1.0
    r, g, b = (int(hexrgb[i:i + 2], 16) / 255 for i in (0, 2, 4))
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


# --------------------------------------------------------------------------- vector tables
def vector_cells(page):
    """Cell boxes from filled rectangles + ruled-line table cells."""
    W_, H_ = page.rect.width, page.rect.height
    boxes = []
    for d in page.get_drawings():
        fill = d.get("fill")
        for it in d["items"]:
            if it[0] not in ("re", "qu"):
                continue
            r = it[1] if it[0] == "re" else it[1].rect
            if 4 < r.width < 0.95 * W_ and 4 < r.height < 0.9 * H_:
                boxes.append({"bbox": (r.x0, r.y0, r.x1, r.y1), "fill": rgb_hex(fill) if fill else None, "src": "rect"})
    try:
        for t in page.find_tables(strategy="lines").tables:
            for c in t.cells:
                if c and (c[2] - c[0]) > 3 and (c[3] - c[1]) > 3:
                    boxes.append({"bbox": tuple(c), "fill": None, "src": "lines"})
    except Exception:
        pass
    # dedupe near-identical boxes (keep the one with a fill)
    out = []
    for b in sorted(boxes, key=lambda b: (b["fill"] is None, P.rect_area(b["bbox"]))):
        if any(iou(b["bbox"], o["bbox"]) > 0.9 for o in out):
            continue
        out.append(b)
    # drop containers holding >= 2 other boxes (table background, frames)
    keep = []
    for b in out:
        inner = sum(1 for o in out if o is not b and P.frac_inside(o["bbox"], b["bbox"]) > 0.9
                    and P.rect_area(o["bbox"]) < 0.9 * P.rect_area(b["bbox"]))
        if inner < 2:
            keep.append(b)
    return keep


def cluster_tables(boxes):
    n = len(boxes)
    parent = list(range(n))

    def f(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i
    for i in range(n):
        for j in range(i + 1, n):
            if P.near(boxes[i]["bbox"], boxes[j]["bbox"], 2.0):
                parent[f(i)] = f(j)
    groups = defaultdict(list)
    for i in range(n):
        groups[f(i)].append(boxes[i])
    return [g for g in groups.values() if len(g) >= 4]


def grid_from_boxes(group):
    xs = snap_values([b["bbox"][0] for b in group] + [b["bbox"][2] for b in group])
    ys = snap_values([b["bbox"][1] for b in group] + [b["bbox"][3] for b in group])
    if len(xs) < 3 or len(ys) < 3:
        return None
    occ = {}
    cells = []
    for b in sorted(group, key=lambda b: P.rect_area(b["bbox"])):
        c0, c1 = idx_of(xs, b["bbox"][0]), idx_of(xs, b["bbox"][2]) - 1
        r0, r1 = idx_of(ys, b["bbox"][1]), idx_of(ys, b["bbox"][3]) - 1
        if c1 < c0 or r1 < r0:
            continue
        pos = [(r, c) for r in range(r0, r1 + 1) for c in range(c0, c1 + 1)]
        if sum(1 for p in pos if p in occ) > 0.5 * len(pos):
            continue
        for p in pos:
            occ.setdefault(p, len(cells))
        cells.append({"r0": r0, "r1": r1, "c0": c0, "c1": c1, "bbox": (xs[c0], ys[r0], xs[c1 + 1], ys[r1 + 1]),
                      "fill": b["fill"]})
    for r in range(len(ys) - 1):                       # uncovered grid positions -> plain cells
        for c in range(len(xs) - 1):
            if (r, c) not in occ:
                cells.append({"r0": r, "r1": r, "c0": c, "c1": c, "bbox": (xs[c], ys[r], xs[c + 1], ys[r + 1]), "fill": None,
                              "gap": True})
    return {"xs": xs, "ys": ys, "cells": cells, "bbox": (xs[0], ys[0], xs[-1], ys[-1]), "method": "vector cell geometry"}


# --------------------------------------------------------------------------- raster tables
def longest_runs(mask):
    out = np.zeros(mask.shape[0], dtype=int)
    for i, row in enumerate(mask):
        if not row.any():
            continue
        d = np.diff(np.concatenate(([0], row.astype(np.int8), [0])))
        s, e = np.flatnonzero(d == 1), np.flatnonzero(d == -1)
        out[i] = int((e - s).max())
    return out


def bounds_from(runs, thr, gap=3):
    idx = np.flatnonzero(runs >= thr)
    if not len(idx):
        return []
    groups, cur = [], [idx[0]]
    for v in idx[1:]:
        if v - cur[-1] <= gap:
            cur.append(v)
        else:
            groups.append(cur)
            cur = [v]
    groups.append(cur)
    return [float(np.mean(g)) for g in groups]


def whitespace_columns(rgb, y0, y1, x0, x1):
    """Column separators from vertical white-space gaps shared by the rows of the region."""
    reg = rgb[y0:y1, x0:x1]
    if reg.size == 0:
        return []
    bg = np.median(reg, axis=1, keepdims=True)                  # per-row background (handles colour bands)
    ink = (np.abs(reg - bg).sum(axis=2) > 90)
    occ = ink.mean(axis=0)
    gap = occ < 0.004
    d = np.diff(np.concatenate(([0], gap.astype(np.int8), [0])))
    st, en = np.flatnonzero(d == 1), np.flatnonzero(d == -1)
    inkcols = np.flatnonzero(~gap)
    if not len(inkcols):
        return []
    left, right = inkcols[0], inkcols[-1]
    seps = [x0 + (a + b) / 2 for a, b in zip(st, en) if b - a >= 14 and a > left and b < right]
    return [x0 + max(0, left - 4)] + seps + [x0 + min(x1 - x0, right + 4)]


def raster_grid(page, bbox, dpi=200, ink_merge=True):
    pix = page.get_pixmap(dpi=dpi, clip=pymupdf.Rect(bbox) & page.rect)
    rgb = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)[:, :, :3].astype(np.int16)
    g = rgb.mean(axis=2)
    Hh, Ww = g.shape
    dy = np.abs(np.diff(g, axis=0)) > 12            # horizontal edges (faint rules included)
    dx = np.abs(np.diff(g, axis=1)) > 12            # vertical edges
    hb = bounds_from(longest_runs(dy), 0.2 * Ww)
    if len(hb) < 3:
        return None
    y0, y1 = int(hb[0]), int(hb[-1])
    vb = bounds_from(longest_runs(dx[y0:y1 + 1].T), 0.45 * (y1 - y0))
    cols_from = "lines"
    if len(vb) < 3:
        vb = whitespace_columns(rgb, y0, y1, 0, Ww)
        cols_from = "white-space gaps"
    if len(vb) < 3:
        # colour/edge changes shared by many rows, kept only where they do not cut through words
        cand = bounds_from(dx[y0:y1].mean(axis=0), 0.2, gap=6)
        png = W / f"_grid_probe_{abs(hash(bbox)) % 10**8}.png"
        Image.fromarray(rgb.astype(np.uint8)).save(png)
        words = [w for w in P.ocr_tsv(png, psm=11) if w["conf"] >= 0]
        nrows = max(1, len(hb) - 1)
        keep = []
        for x in cand:
            cuts = sum(1 for w in words if w["left"] + 2 < x < w["left"] + w["w"] - 2 and y0 <= w["top"] <= y1)
            if cuts <= 0.35 * nrows:          # spanning rows may cross a real boundary (become merged cells)
                keep.append(x)
        if words:
            xs_ = [w["left"] for w in words if y0 <= w["top"] <= y1] + [w["left"] + w["w"] for w in words if y0 <= w["top"] <= y1]
            if xs_:
                keep += [max(0, min(xs_) - 6), min(Ww - 1, max(xs_) + 6)]
        vb = snap_values(keep, 10)
        cols_from = "colour changes validated against word positions"
        if len(vb) < 3:
            return None
    x0, x1 = int(vb[0]), int(vb[-1])
    if cols_from == "lines":
        hb = bounds_from(longest_runs(dy[:, x0:x1 + 1]), 0.45 * (x1 - x0))
        hb = [h for h in hb if y0 - 3 <= h <= y1 + 3]
        if len(hb) < 3:
            return None
    hb, vb = snap_values(hb, 4), snap_values(vb, 4)
    nr, nc = len(hb) - 1, len(vb) - 1
    # merge detection: shared boundary with little edge evidence => merged
    parent = {(r, c): (r, c) for r in range(nr) for c in range(nc)}

    def f(k):
        while parent[k] != k:
            k = parent[k]
        return k
    for r in range(nr):
        for c in range(nc - 1):
            x = int(vb[c + 1]); ya, yb = int(hb[r]) + 2, int(hb[r + 1]) - 2
            if yb <= ya:
                continue
            if cols_from != "white-space gaps":
                seg = dx[ya:yb, max(0, x - 2):min(dx.shape[1], x + 3)].any(axis=1).mean()
                if seg < 0.4:
                    parent[f((r, c + 1))] = f((r, c))
            else:
                band = rgb[ya:yb, max(0, x - 3):x + 4]
                bg = np.median(rgb[ya:yb], axis=(0, 1))
                if (np.abs(band - bg).sum(axis=2) > 90).mean() > 0.02:      # text runs across the gap -> merged
                    parent[f((r, c + 1))] = f((r, c))
    for r in range(nr - 1):
        for c in range(nc):
            if cols_from != "lines":
                continue                                   # row rules are drawn; keep rows separate
            y = int(hb[r + 1]); xa, xb = int(vb[c]) + 2, int(vb[c + 1]) - 2
            if xb <= xa:
                continue
            seg = dy[max(0, y - 2):min(dy.shape[0], y + 3), xa:xb].any(axis=0).mean()
            if seg < 0.4:
                parent[f((r + 1, c))] = f((r, c))
    groups = defaultdict(list)
    for k in parent:
        groups[f(k)].append(k)
    s = 72 / dpi
    xs = [bbox[0] + v * s for v in vb]
    ys = [bbox[1] + v * s for v in hb]
    cells, rect_ok = [], True
    for members in groups.values():
        rs = [m[0] for m in members]; cs = [m[1] for m in members]
        r0, r1, c0, c1 = min(rs), max(rs), min(cs), max(cs)
        if len(members) != (r1 - r0 + 1) * (c1 - c0 + 1):
            rect_ok = False
            for m in members:
                cells.append(_rcell(m[0], m[0], m[1], m[1], xs, ys, rgb, hb, vb))
            continue
        cells.append(_rcell(r0, r1, c0, c1, xs, ys, rgb, hb, vb))
    return {"xs": xs, "ys": ys, "cells": cells, "bbox": (xs[0], ys[0], xs[-1], ys[-1]),
            "method": f"grid detected from rendered pixels (rows: rules; columns: {cols_from})"
                      + ("" if rect_ok else " (irregular merges split)") + " — verify"}


def _rcell(r0, r1, c0, c1, xs, ys, rgb, hb, vb):
    ya, yb, xa, xb = int(hb[r0]), int(hb[r1 + 1]), int(vb[c0]), int(vb[c1 + 1])
    py, px = max(1, (yb - ya) // 5), max(1, (xb - xa) // 5)
    patch = rgb[ya + py:yb - py, xa + px:xb - px].reshape(-1, 3)
    fill = None
    if len(patch):
        med = np.median(patch, axis=0)
        fill = "{:02X}{:02X}{:02X}".format(*[int(v) for v in med])
        if fill in ("FFFFFF", "FEFEFE", "FDFDFD"):
            fill = None
    return {"r0": r0, "r1": r1, "c0": c0, "c1": c1, "bbox": (xs[c0], ys[r0], xs[c1 + 1], ys[r1 + 1]), "fill": fill}


# --------------------------------------------------------------------------- cell text
def ocr_cell(args):
    png, = args
    im = Image.open(png).convert("L")
    if np.asarray(im).mean() < 128:                 # white text on dark fill
        im = Image.fromarray(255 - np.asarray(im))
    im.save(png)
    env = dict(os.environ, OMP_THREAD_LIMIT="1")
    best = ("", -1.0)
    for psm in (6, 7):
        r = subprocess.run(["tesseract", str(png), "stdout", "-l", "eng", "--psm", str(psm), "tsv"],
                           capture_output=True, text=True, env=env)
        words, confs = defaultdict(list), []
        for line in r.stdout.splitlines()[1:]:
            f = line.split("\t")
            if len(f) == 12 and f[0] == "5" and f[11].strip():
                words[(int(f[2]), int(f[3]), int(f[4]))].append(f[11])
                confs.append(float(f[10]))
        text = "\n".join(" ".join(v) for _, v in sorted(words.items()))
        mean = sum(confs) / len(confs) if confs else -1
        if text and mean > best[1]:
            best = (text, mean)
        if text and mean >= 70:
            break
    return best


def fill_text(page, tables, pno, native):
    jobs, targets = [], []
    pix = page.get_pixmap(dpi=CELL_DPI)
    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    s = CELL_DPI / 72
    for ti, t in enumerate(tables):
        for ci, c in enumerate(t["cells"]):
            bb = c["bbox"]
            inside = [w for w in native if bb[0] - 1 <= (w[0] + w[2]) / 2 <= bb[2] + 1 and bb[1] - 1 <= (w[1] + w[3]) / 2 <= bb[3] + 1]
            if inside:
                c["text"], c["conf"], c["src"] = U.text_of([(w[0], w[1], w[2], w[3], w[4], 100, "native") for w in inside]), 100.0, "native"
                continue
            if c.get("gap") and (bb[2] - bb[0] < 4 or bb[3] - bb[1] < 4):
                c["text"], c["conf"], c["src"] = "", 100.0, "none"
                continue
            pad = 1.5
            crop = img.crop((int((bb[0] + pad) * s), int((bb[1] + pad) * s), int((bb[2] - pad) * s), int((bb[3] - pad) * s)))
            if crop.width < 8 or crop.height < 8:
                c["text"], c["conf"], c["src"] = "", 100.0, "none"
                continue
            png = W / f"p{pno:02d}_t{ti}_c{ci}.png"
            crop.save(png)
            jobs.append((png,))
            targets.append(c)
    with ThreadPoolExecutor(max_workers=os.cpu_count() or 4) as ex:
        for c, (text, conf) in zip(targets, ex.map(ocr_cell, jobs)):
            c["text"], c["conf"], c["src"] = text, conf, "ocr"


def table_rows_text(t):
    rows = defaultdict(list)
    for c in t["cells"]:
        rows[c["r0"]].append((c["c0"], c.get("text", "")))
    return [" | ".join(x for _, x in sorted(v)) for _, v in sorted(rows.items())]


# --------------------------------------------------------------------------- page analysis
def analyse_page(doc, pno, analysed, review, name):
    page = doc[pno - 1]
    Wp, Hp = page.rect.width, page.rect.height
    native = [w for w in page.get_text("words") if w[4].strip()]
    tables = []
    for g in cluster_tables(vector_cells(page)):
        t = grid_from_boxes(g)
        if t and len(t["xs"]) >= 3 and len(t["ys"]) >= 3:
            tables.append(t)
    regions = [tuple(i["bbox"]) for i in page.get_image_info() if P.rect_area(i["bbox"]) >= 0.08 * Wp * Hp]
    hl = []
    for d in page.get_drawings():
        for it in d["items"]:
            if it[0] == "l" and abs(it[1].y - it[2].y) < 1 and abs(it[1].x - it[2].x) >= 40:
                hl.append((min(it[1].x, it[2].x), it[1].y - 1, max(it[1].x, it[2].x), it[1].y + 1))
    for c in P.cluster(hl, 40):
        if c[3] - c[1] > 20:
            regions.append((c[0] - 4, c[1] - 16, c[2] + 4, c[3] + 16))
    for bb in regions:
        if P.rect_area(bb) < 0.02 * Wp * Hp or any(P.inter(bb, t["bbox"]) > 0.3 * P.rect_area(t["bbox"]) for t in tables):
            continue
        t = raster_grid(page, (max(0, bb[0]), max(0, bb[1]), min(Wp, bb[2]), min(Hp, bb[3])))
        if t and (len(t["xs"]) - 1) * (len(t["ys"]) - 1) >= 4:
            tables.append(t)
        elif P.rect_area(bb) > 0.3 * Wp * Hp:
            review.append((name, pno, "Table in picture not detected",
                           "A large picture on this slide has no detectable cell grid; content captured as OCR text + slide image."))
    tables.sort(key=lambda t: (t["bbox"][1], t["bbox"][0]))
    for t in tables:
        t["page"] = pno
    fill_text(page, tables, pno, native)
    # text-only tables on native pages (no drawn grid)
    words, mode, rot = U.page_words(page, W, pno) if len(native) >= 40 else (None, None, 0)
    text_items, lines = [], []
    if len(native) >= 40:
        lines = U.to_lines([(w[0], w[1], w[2], w[3], w[4], 100.0, "native") for w in native])
        for lt in U.layout_tables(lines, [t["bbox"] for t in tables]):
            cells = []
            for r, row in enumerate(lt["cells"]):
                for c, v in enumerate(row):
                    cells.append({"r0": r, "r1": r, "c0": c, "c1": c, "text": v or "", "conf": 100.0, "src": "native",
                                  "fill": None})
            n = lt["ncol"]
            x0, x1 = lt["bbox"][0], lt["bbox"][2]
            xs = [x0 + (x1 - x0) * k / n for k in range(n + 1)]
            h = (lt["bbox"][3] - lt["bbox"][1]) / max(1, lt["nrow"])
            ys = [lt["bbox"][1] + h * k for k in range(lt["nrow"] + 1)]
            for c in cells:
                c["bbox"] = (xs[c["c0"]], ys[c["r0"]], xs[c["c1"] + 1], ys[c["r1"] + 1])
            tables.append({"xs": xs, "ys": ys, "cells": cells, "bbox": lt["bbox"], "page": pno,
                           "method": "word layout on native text (no drawn grid) — verify"})
            review.append((name, pno, "Uncertain table structure", "Table without drawn grid rebuilt from word positions; verify."))
        mode = "native text"
        free = [ln for ln in lines if not any(P.frac_inside(ln["bbox"], t["bbox"]) > 0.6 for t in tables)]
        for ln in free:
            text_items.append({"bbox": ln["bbox"], "text": " ".join(w[4] for w in ln["w"]), "conf": 100.0, "src": "native"})
    else:
        mode = "OCR (text drawn as vectors/pictures)"
        png = W / f"p{pno:02d}_page.png"
        pix = page.get_pixmap(dpi=300)
        im = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        from PIL import ImageDraw
        dr = ImageDraw.Draw(im)
        s = 300 / 72
        for t in tables:
            b = t["bbox"]
            dr.rectangle([b[0] * s - 2, b[1] * s - 2, b[2] * s + 2, b[3] * s + 2], fill="white")
        im.save(png)
        rot = P.osd_rotation(png)
        if rot:
            review.append((name, pno, "Orientation", f"Slide appears rotated {rot}° — verify."))
        for ob in P.ocr_blocks(P.ocr_tsv(png, psm=3)):
            bb = tuple(v / s for v in ob["bbox_px"])
            if ob["text"].strip():
                text_items.append({"bbox": bb, "text": ob["text"], "conf": ob["mean_conf"], "src": "ocr"})
        lines = U.to_lines([(w[0], w[1], w[2], w[3], w[4], 100.0, "native") for w in native])
    # merge consecutive native lines into paragraphs
    paras = []
    for it in sorted(text_items, key=lambda x: (x["bbox"][1], x["bbox"][0])):
        if paras and it["src"] == "native" and paras[-1]["src"] == "native" and it["bbox"][1] - paras[-1]["bbox"][3] < 6 \
                and abs(it["bbox"][0] - paras[-1]["bbox"][0]) < 30 and not re.match(r"^\s*(\d{1,2}[.)]|[•▪◦\-–])\s", it["text"]):
            paras[-1]["text"] += "\n" + it["text"]
            paras[-1]["bbox"] = P.union(paras[-1]["bbox"], it["bbox"])
        else:
            paras.append(dict(it))
    figs = [f for f in analysed[pno - 1]["figures"]
            if not any(P.inter(f["bbox"], t["bbox"]) > 0.3 * P.rect_area(f["bbox"]) for t in tables)
            and P.rect_area(f["bbox"]) < 0.7 * Wp * Hp]
    pn = U.printed_number(U.to_lines([(w[0], w[1], w[2], w[3], w[4], 100.0, "native") for w in native]) or [], Hp)
    return {"tables": tables, "text": paras, "figs": figs, "mode": mode, "printed": pn, "lines": lines}


# --------------------------------------------------------------------------- continuation
def same_columns(a, b):
    return len(a["xs"]) == len(b["xs"]) and all(abs(x - y) <= 3 for x, y in zip(a["xs"], b["xs"]))


def merge_continuations(pages_info, pnos, rule, name, review):
    seq = [t for p in pnos for t in pages_info[p]["tables"]]
    logical = []
    for t in seq:
        prev = logical[-1] if logical else None
        cont = (prev is not None and rule != "single" and t["page"] == prev["pages"][-1] + 1 and same_columns(prev["parts"][-1], t)
                and t is pages_info[t["page"]]["tables"][0] and prev["parts"][-1] is pages_info[prev["pages"][-1]]["tables"][-1])
        if cont:
            prev["parts"].append(t)
            prev["pages"].append(t["page"])
        else:
            if prev is not None and rule == "combine" and t["page"] == prev["pages"][-1] + 1:
                review.append((name, t["page"], "Possible continuation",
                               f"Table on p{t['page']} does not share the previous slide's column layout — kept as a separate table."))
            logical.append({"parts": [t], "pages": [t["page"]]})
    # drop repeated header rows in continuation parts
    for lt in logical:
        if len(lt["parts"]) < 2:
            continue
        head_rows = table_rows_text(lt["parts"][0])[:2]
        for part in lt["parts"][1:]:
            rows = table_rows_text(part)
            drop = set()
            for r, txt in enumerate(rows[:2]):
                if any(U.header_sim([txt], [h]) >= 0.8 for h in head_rows):
                    drop.add(r)
                else:
                    break
            if drop:
                part["cells"] = [c for c in part["cells"] if c["r0"] not in drop]
                part["dropped_header_rows"] = len(drop)
    return logical


# --------------------------------------------------------------------------- manual edits
def carried_edits(doc, wb_base, gen_wb, pages_wanted):
    """Edited cells in existing Table_### sheets (base workbook vs unedited regeneration) mapped to PDF bboxes."""
    old_pages = P.analyse_pdf(ORIG_63, P.WORK / "_regen_old")[1]
    P.logical_tables(old_pages)
    out = []
    for s in wb_base.sheetnames:
        if not s.startswith("Table_") or s not in gen_wb.sheetnames:
            continue
        src = [int(x) for x in str(wb_base[s]["B3"].value).split(",") if x.strip().isdigit()]
        if not any(p in pages_wanted for p in src):
            continue
        a, b = wb_base[s], gen_wb[s]
        va = {(c.row, c.column): c.value for row in a.iter_rows() for c in row if c.value is not None}
        vb = {(c.row, c.column): c.value for row in b.iter_rows() for c in row if c.value is not None}
        edits = {k: va.get(k) for k in set(va) | set(vb) if va.get(k) != vb.get(k)}
        merges = sorted(map(str, a.merged_cells.ranges)) != sorted(map(str, b.merged_cells.ranges))
        if not edits and not merges:
            continue
        for p in src:
            tabs = [t for t in old_pages[p - 1]["tables"] if t.get("id") == s]
            if not tabs:
                continue
            pg = doc[p - 1]
            ft = [t for t in pg.find_tables(strategy="lines").tables]
            k = old_pages[p - 1]["tables"].index(tabs[0])
            if k >= len(ft):
                continue
            ftab = ft[k]
            for (r, c), v in edits.items():
                ri, ci = r - 8, c - 1
                if 0 <= ri < ftab.row_count and 0 <= ci < ftab.col_count:
                    cb = ftab.rows[ri].cells[ci] if ci < len(ftab.rows[ri].cells) else None
                    out.append({"sheet": s, "page": p, "row": r, "col": c, "value": v, "bbox": cb})
                else:
                    out.append({"sheet": s, "page": p, "row": r, "col": c, "value": v, "bbox": None})
            if merges:
                out.append({"sheet": s, "page": p, "merges_changed": True})
    return out


# --------------------------------------------------------------------------- writing
def write_cell_range(ws, r0, r1, c0, c1, value, fill, bold=False, size=9):
    if value not in (None, ""):
        P.put(ws, r0, c0, value)
    for r in range(r0, r1 + 1):
        for c in range(c0, c1 + 1):
            cell = ws.cell(r, c)
            cell.border = BOX
            if fill:
                cell.fill = PatternFill("solid", fgColor=fill)
    top = ws.cell(r0, c0)
    top.alignment = Alignment(wrap_text=True, vertical="center", horizontal="left")
    top.font = Font(size=size, bold=bold, color="FFFFFF" if fill and luminance(fill) < 0.45 else "000000")
    if r1 > r0 or c1 > c0:
        ws.merge_cells(start_row=r0, start_column=c0, end_row=r1, end_column=c1)


def write_group(wb, name, pnos, title, rule, doc, analysed, edits, review, log):
    ws = wb.create_sheet(name)
    page_w = doc[0].rect.width
    ncols = int(math.ceil(page_w / UNIT))
    ws.column_dimensions["A"].width, ws.column_dimensions["B"].width = 9, 16
    for c in range(COL0, COL0 + ncols + 1):
        ws.column_dimensions[get_column_letter(c)].width = 1.15
    last = get_column_letter(COL0 + ncols - 1)
    infos = {p: analyse_page(doc, p, analysed, review, name) for p in pnos}
    logical = merge_continuations(infos, pnos, rule, name, review)
    hdr = [("Worksheet", name), ("Source physical PDF pages", ", ".join(map(str, pnos))),
           ("Printed slide numbers", "; ".join(f"p{p}: {infos[p]['printed'] or 'none found'}" for p in pnos)),
           ("Requested", title or "all content"), ("Method", "; ".join(f"p{p}: {infos[p]['mode']}" for p in pnos))]
    for i, (k, v) in enumerate(hdr, 1):
        P.put(ws, i, 1, k, bold=True)
        P.put(ws, i, COL0, v)
        ws.cell(i, 1).fill = P.LABEL_FILL
    P.link(ws, 6, 1, "◀ Update Log", f"#'{LOG_SHEET}'!A1")
    for p in pnos:
        if infos[p]["printed"] and infos[p]["printed"] != str(p):
            review.append((name, p, "Slide number", f"Printed slide number {infos[p]['printed']} differs from physical page {p}."))
    if title:
        for p in pnos:
            hit = U.find_title(infos[p]["lines"] or [], title)
            hit_ocr = any(U.norm(title) in U.norm(t["text"]) or
                          U.difflib.SequenceMatcher(None, U.norm(title), U.norm(t["text"])[:len(title) + 10]).ratio() >= 0.75
                          for t in infos[p]["text"])
            log["titles"].append((name, title, p, "found" if (hit or hit_ocr) else "not found"))
    r = 8
    placed, figures = set(), []
    tno = 0
    for p in pnos:
        info = infos[p]
        P.put(ws, r, 1, f"p{p}", bold=True); P.put(ws, r, 2, "Slide section", bold=True)
        P.put(ws, r, COL0, f"Physical PDF page {p}", bold=True)
        for c in range(1, COL0 + ncols):
            ws.cell(r, c).fill = P.HDR_FILL
        r += 2
        entries = [("text", t["bbox"][1], t) for t in info["text"]]
        entries += [("table", lt["parts"][0]["bbox"][1], lt) for lt in logical if lt["pages"][0] == p]
        entries += [("figure", f["bbox"][1], f) for f in info["figs"]]
        entries.sort(key=lambda e: e[1])
        if not entries:
            P.put(ws, r, 1, f"p{p}"); P.put(ws, r, 2, "No content")
            P.put(ws, r, COL0, "(nothing detected — see slide image)")
            review.append((name, p, "Possible missing content", "Nothing extracted; check slide image."))
            r += 2
        for kind, _, obj in entries:
            if kind == "text":
                low = obj["src"] == "ocr" and obj["conf"] < P.LOW_CONF
                P.put(ws, r, 1, f"p{p}"); P.put(ws, r, 2, "Text" + (" (OCR)" if obj["src"] == "ocr" else "") + (" — review" if low else ""))
                c0 = COL0 + max(0, int(obj["bbox"][0] / UNIT))
                c1 = min(COL0 + ncols - 1, max(c0 + 10, COL0 + int(obj["bbox"][2] / UNIT)))
                P.put(ws, r, c0, obj["text"])
                ws.cell(r, c0).alignment = Alignment(wrap_text=True, vertical="top")
                ws.merge_cells(start_row=r, start_column=c0, end_row=r, end_column=c1)
                cpl = max(10, (c1 - c0 + 1) * UNIT / 5.2)
                n = sum(max(1, math.ceil(len(x) / cpl)) for x in obj["text"].split("\n"))
                ws.row_dimensions[r].height = min(15 * n + 3, 400)
                if low:
                    review.append((name, p, "Uncertain OCR text", f"Row {r}: mean OCR confidence {obj['conf']:.0f}."))
                r += 1
            elif kind == "table":
                lt = obj
                tno += 1
                P.put(ws, r, 1, "/".join(f"p{x}" for x in lt["pages"]), bold=True); P.put(ws, r, 2, f"Table {tno}", bold=True)
                P.put(ws, r, COL0, f"Table {tno} — source page(s) {', '.join(map(str, lt['pages']))}"
                      + (" — continued across slides, merged (same column layout; repeated header rows removed)" if len(lt["pages"]) > 1 else "")
                      + f" — {lt['parts'][0]['method']}", bold=True)
                r += 1
                for k, part in enumerate(lt["parts"]):
                    if k:
                        P.put(ws, r, 1, f"p{part['page']}"); P.put(ws, r, 2, "(continued)")
                    xs, ys = part["xs"], part["ys"]
                    rows_used = sorted({c["r0"] for c in part["cells"]} | {c["r1"] for c in part["cells"]})
                    if not rows_used:
                        continue
                    rmin = min(c["r0"] for c in part["cells"])
                    row_map, rr = {}, r
                    for gr in range(rmin, len(ys) - 1):
                        if any(c["r0"] <= gr <= c["r1"] for c in part["cells"]):
                            row_map[gr] = rr
                            hpt = ys[gr + 1] - ys[gr]
                            ws.row_dimensions[rr].height = max(15, min(hpt, 200))
                            P.put(ws, rr, 1, f"p{part['page']}")
                            rr += 1
                    for c in part["cells"]:
                        if c["r0"] not in row_map or c["r1"] not in row_map:
                            continue
                        xc0 = COL0 + int(round(xs[c["c0"]] / UNIT))
                        xc1 = COL0 + int(round(xs[c["c1"] + 1] / UNIT)) - 1
                        xc1 = max(xc0, xc1)
                        text = c.get("text", "")
                        edited = None
                        for e in edits:
                            if e.get("bbox") and e["page"] == part["page"] and c["bbox"][0] - 1 <= (e["bbox"][0] + e["bbox"][2]) / 2 <= c["bbox"][2] + 1 \
                                    and c["bbox"][1] - 1 <= (e["bbox"][1] + e["bbox"][3]) / 2 <= c["bbox"][3] + 1:
                                edited = e
                        if edited is not None:
                            text = edited["value"] if edited["value"] is not None else ""
                            edited["used"] = True
                            review.append((name, part["page"], "Manual edit carried over",
                                           f"{get_column_letter(xc0)}{row_map[c['r0']]}: value from {edited['sheet']}!"
                                           f"{get_column_letter(edited['col'])}{edited['row']} (manual edit) used instead of extraction."))
                        elif c.get("src") == "ocr" and text and c.get("conf", 100) < P.LOW_CONF:
                            review.append((name, part["page"], "Uncertain OCR table cell",
                                           f"{get_column_letter(xc0)}{row_map[c['r0']]}: OCR confidence {c['conf']:.0f}."))
                        elif c.get("src") == "ocr" and not text and c.get("fill"):
                            pass
                        bold = c["r0"] == min(cc["r0"] for cc in part["cells"]) and k == 0
                        write_cell_range(ws, row_map[c["r0"]], row_map[c["r1"]], xc0, xc1, text, c.get("fill"), bold=bold)
                        # fit row height to wrapped text
                        cpl = max(4, (xc1 - xc0 + 1) * UNIT / 5.0)
                        need = sum(max(1, math.ceil(len(x) / cpl)) for x in str(text).split("\n")) * 12 + 4 if text else 0
                        span_rows = list(range(row_map[c["r0"]], row_map[c["r1"]] + 1))
                        have = sum(ws.row_dimensions[x].height or 15 for x in span_rows)
                        if need > have:
                            ws.row_dimensions[span_rows[-1]].height = min((ws.row_dimensions[span_rows[-1]].height or 15) + need - have, 300)
                    r = rr
                    if part.get("dropped_header_rows"):
                        review.append((name, part["page"], "Continuation", f"{part['dropped_header_rows']} repeated header row(s) not repeated as data."))
                log["tables"].append((name, tno, lt["pages"], lt["parts"][0]["method"]))
                if "verify" in lt["parts"][0]["method"] or "picture" in lt["parts"][0]["method"]:
                    review.append((name, lt["pages"][0], "Verify table", f"Table {tno}: {lt['parts'][0]['method']}; compare with slide image."))
                r += 1
            else:
                figures.append((p, obj))
                P.put(ws, r, 1, f"p{p}"); P.put(ws, r, 2, "Figure")
                P.put(ws, r, COL0, f"Figure p{p}-{sum(1 for x in figures if x[0] == p)} — embedded below (Figures on these slides)")
                r += 1
        r += 1
    # images below all content
    r += 1
    if figures:
        P.put(ws, r, 1, "Figures on these slides", bold=True)
        r += 1
        for k, (p, f) in enumerate(figures, 1):
            crop = P.render_crop(doc[p - 1], f["bbox"], W / f"{name.replace(' ', '_')}_fig{k}.png")
            P.put(ws, r, 1, f"p{p}"); P.put(ws, r, 2, f"Figure p{p}-{k}")
            used, _, _ = P.add_image(ws, crop, COL0, r + 1, 900)
            r += used + 1
            log["figures"].append((name, p))
    P.put(ws, r, 1, "Slide images (verification)", bold=True)
    r += 1
    for p in pnos:
        jp = W / f"slide_{p:02d}.jpg"
        pix = doc[p - 1].get_pixmap(dpi=120)
        Image.frombytes("RGB", (pix.width, pix.height), pix.samples).save(jp, quality=82)
        P.put(ws, r, 1, f"p{p}", bold=True); P.put(ws, r, 2, f"Slide (physical page {p})")
        used, _, _ = P.add_image(ws, jp, COL0, r + 1, 1000)
        r += used + 1
    ws.freeze_panes = "C8"
    ws.sheet_view.zoomScale = 90
    return infos


def snapshot(ws):
    return U.snapshot(ws)


def main():
    W.mkdir(parents=True, exist_ok=True)
    assert not OUT_WB.exists(), f"refusing to overwrite {OUT_WB}"
    base_hash = P.sha(BASE_WB)
    pdf = P.read_bytes(SRC_PDF)
    pdf_hash = sha_bytes(pdf)
    assert pdf_hash == sha_bytes(ORIG_63.read_bytes()), "NAS PDF is not the 63-page version expected"
    local = W / "_source_63.pdf"
    local.write_bytes(pdf)
    doc = pymupdf.open(local)
    assert len(doc) == 63
    _, analysed, _ = P.analyse_pdf(local, W)

    base = load_workbook(BASE_WB)
    gen = load_workbook(REGEN_OLD)
    wanted = sorted({p for _, ps, _, _ in GROUPS for p in ps})
    edits = carried_edits(doc, base, gen, wanted)

    shutil.copy2(BASE_WB, W / "_base_copy.xlsx")
    wb = load_workbook(W / "_base_copy.xlsx")
    keep = [s for s in wb.sheetnames if s not in V2_SHEETS]
    before = {s: snapshot(wb[s]) for s in keep}
    for s in V2_SHEETS:
        if s in wb.sheetnames:
            del wb[s]
    review, log = [], {"titles": [], "tables": [], "figures": []}
    for name, pnos, title, rule in GROUPS:
        write_group(wb, name, pnos, title, rule, doc, analysed, edits, review, log)
    for e in edits:
        if e.get("merges_changed"):
            review.append(("(existing)", e["page"], "Manual merge edits",
                           f"{e['sheet']} has manually merged cells; new table uses the slide's own cell geometry — compare."))
        elif not e.get("used"):
            review.append(("(existing)", e["page"], "Manual edit not placed",
                           f"{e['sheet']}!{get_column_letter(e['col'])}{e['row']}: could not be matched to a cell in the new table; "
                           f"{e['sheet']} is preserved unchanged."))

    lg = wb.create_sheet(LOG_SHEET)
    P.head(lg, ["Item", "Detail"], [40, 110])
    rows = [("Base workbook (copied, unchanged)", str(BASE_WB)), ("Output workbook", str(OUT_WB)),
            ("Source PDF", f"{SRC_PDF} — 63 physical pages (current NAS version); sha256 {pdf_hash[:16]}…"),
            ("Page numbering", "Physical PDF page positions from 1; blank pages counted in position"),
            ("Replaced from v2", ", ".join(V2_SHEETS)),
            ("Kept unchanged", "all other sheets (Table_002 remains removed); verified cell-by-cell"),
            ("Table method", "cell geometry from the PDF (drawn cell rectangles / ruled lines / picture grid lines); "
                             "each cell's text from native PDF text or OCR of that cell; merged cells and fills reproduced")]
    for i, (k, v) in enumerate(rows, 2):
        P.put(lg, i, 1, k, bold=True); P.put(lg, i, 2, v, wrap=True)
    r = len(rows) + 3
    P.put(lg, r, 1, "Worksheets added", bold=True)
    for name, pnos, title, _ in GROUPS:
        r += 1
        P.link(lg, r, 1, name, f"#'{name}'!A1"); P.put(lg, r, 2, f"pages {', '.join(map(str, pnos))}" + (f" — {title}" if title else " — all content"))
    r += 2
    P.put(lg, r, 1, "Requested titles", bold=True)
    for t in log["titles"]:
        r += 1
        P.put(lg, r, 1, f"{t[0]} p{t[2]}"); P.put(lg, r, 2, f"'{t[1]}': {t[3]}")
    rv = wb.create_sheet(REVIEW_SHEET)
    P.head(rv, ["Worksheet", "Source page", "Issue", "Detail"], [24, 11, 30, 110])
    for i, x in enumerate(review, 2):
        for c, v in enumerate(x, 1):
            P.put(rv, i, c, v, wrap=c == 4)
    wb.save(OUT_WB)

    # validation
    out = load_workbook(OUT_WB)
    problems = []
    for s, snap in before.items():
        if s not in out.sheetnames:
            problems.append(f"missing {s}")
            continue
        a = snapshot(out[s])
        for k in snap:
            if a[k] != snap[k]:
                problems.append(f"{s}: {k} changed")
    if "Table_002" in out.sheetnames:
        problems.append("Table_002 present")
    for name, pnos, _, _ in GROUPS:
        ws = out[name]
        present = {str(v) for (v,) in ws.iter_rows(min_row=8, max_col=1, values_only=True) if v}
        for p in pnos:
            if f"p{p}" not in present:
                problems.append(f"{name}: p{p} missing")
        last_cell_row = max((c.row for row in ws.iter_rows(min_col=COL0) for c in row if c.value is not None
                             and not str(c.value).startswith(("Figure p", "Slide (physical"))), default=0)
        first_img = min((i.anchor._from.row + 1 for i in ws._images), default=10 ** 9)
        if first_img <= last_cell_row and not any(str(ws.cell(first_img - 1, 2).value or "").startswith(("Figure", "Slide")) for _ in [0]):
            problems.append(f"{name}: image may overlap content")
        for row in ws.iter_rows():
            for c in row:
                if c.data_type == "f":
                    problems.append(f"{name}: formula {c.coordinate}")
    assert P.sha(BASE_WB) == base_hash, "v2 changed"
    assert sha_bytes(P.read_bytes(SRC_PDF)) == pdf_hash, "source PDF changed during run"
    summary = {"problems": problems, "tables": len(log["tables"]),
               "tables_per_sheet": dict(Counter(t[0] for t in log["tables"])),
               "methods": dict(Counter(t[3].split(" —")[0] for t in log["tables"])),
               "multi_page_tables": [(t[0], t[2]) for t in log["tables"] if len(t[2]) > 1],
               "figures": len(log["figures"]), "review": dict(Counter(x[2] for x in review)),
               "titles": [(t[0], t[2], t[3]) for t in log["titles"]],
               "edits_found": sum(1 for e in edits if not e.get("merges_changed")),
               "edits_carried": sum(1 for e in edits if e.get("used"))}
    (P.LOGS / f"{STEM}_v3_update.json").write_text(json.dumps(summary, indent=2, default=str))
    print(json.dumps(summary, indent=1, default=str))


if __name__ == "__main__":
    main()
