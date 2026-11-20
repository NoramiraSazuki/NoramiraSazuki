#!/usr/bin/env python3
"""Convert the complete contents of PDF files into Excel workbooks (local processing only).

SOURCE  : /Volumes/homes/public/DulangOSS/20261006/ARPR/  (read-only; all subfolders)
OUTPUT  : /Users/farahyusof/Documents/GitHub/NoramiraSazuki/AinnyFarah/ARPR/
          output/<relative path>.xlsx  (+ _PageRef_partN.xlsx if a workbook would be too large)
          output/ARPR_Processing_Summary.xlsx
          work/<pdf>/   renders, crops, OCR output        logs/   per-PDF JSON + run log
          state/manifest.json  (resume: unchanged + completed PDFs are skipped)

Tools: PyMuPDF (text, tables, drawings, rendering) and Tesseract (OCR, orientation detection).
Console output is limited to counts and statuses; no document content is printed.

Usage: pdf_to_excel.py [--only <substring>] [--force]
"""
import argparse
import hashlib
import json
import math
import re
import subprocess
import sys
import time
import traceback
from collections import Counter, defaultdict
from pathlib import Path

import pymupdf
from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from PIL import Image

SRC = Path("/Volumes/homes/public/DulangOSS/20261006/ARPR")
ROOT = Path("/Users/farahyusof/Documents/GitHub/NoramiraSazuki/AinnyFarah/ARPR")
OUT, WORK, LOGS, STATE = ROOT / "output", ROOT / "work", ROOT / "logs", ROOT / "state"
MANIFEST = STATE / "manifest.json"
assert SRC not in ROOT.parents and ROOT != SRC, "OUTPUT_FOLDER must be outside SOURCE_FOLDER"

PAGE_DPI = 110                 # Page Reference images
FIG_DPI = 150                  # figure / table crops
OCR_DPI = 300
CELL_LIMIT = 30000             # < Excel 32,767 char limit
MAX_PART_IMAGE_BYTES = 120 * 1024 * 1024
ROW_PX = 20                    # default row height (15 pt)
LOW_CONF = 60
PROGRAM_VERSION = "1.2"

pymupdf.TOOLS.mupdf_display_errors(False)
pymupdf.TOOLS.mupdf_display_warnings(False)


# --------------------------------------------------------------------------- helpers
def read_bytes(p, tries=5):
    """Read a file read-only via a raw descriptor. On this SMB mount close() can fail with
    EBADF after a successful read; that close error is ignored (nothing is written)."""
    import os
    for k in range(tries):
        try:
            fd = os.open(p, os.O_RDONLY)
            try:
                parts = []
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
            data = b"".join(parts)
            if len(data) != os.stat(p).st_size:
                raise OSError("short read")
            return data
        except OSError:
            if k == tries - 1:
                raise
            time.sleep(1.5 * (k + 1))


def sha(p):
    return hashlib.sha256(read_bytes(p)).hexdigest()


def safe(s):
    """Excel-safe text: remove illegal control characters (counted by caller)."""
    return ILLEGAL_CHARACTERS_RE.sub("�", s) if isinstance(s, str) else s


def put(ws, row, col, value, bold=False, wrap=False):
    """Write a value as TEXT (never a formula, date or number conversion)."""
    c = ws.cell(row=row, column=col)
    if value is None:
        return c
    v = safe(str(value))
    c.value = v
    c.data_type = "s"                       # '=' prefixed strings stay text
    c.number_format = "@"
    if bold:
        c.font = Font(bold=True)
    if wrap:
        c.alignment = Alignment(wrap_text=True, vertical="top")
    return c


def link(ws, row, col, text, target):
    c = put(ws, row, col, text)
    c.hyperlink = target
    c.font = Font(color="0563C1", underline="single")
    return c


def chunks(text, n=CELL_LIMIT):
    if len(text) <= n:
        return [text]
    out, cur = [], ""
    for part in re.split(r"(\s+)", text):
        if len(cur) + len(part) > n:
            out.append(cur)
            cur = ""
        while len(part) > n:
            out.append(part[:n])
            part = part[n:]
        cur += part
    if cur:
        out.append(cur)
    return out


def rect_area(r):
    return max(0, r[2] - r[0]) * max(0, r[3] - r[1])


def inter(a, b):
    return rect_area((max(a[0], b[0]), max(a[1], b[1]), min(a[2], b[2]), min(a[3], b[3])))


def frac_inside(inner, outer):
    a = rect_area(inner)
    return inter(inner, outer) / a if a else 0


def union(a, b):
    return (min(a[0], b[0]), min(a[1], b[1]), max(a[2], b[2]), max(a[3], b[3]))


def near(a, b, gap):
    return not (a[2] + gap < b[0] or b[2] + gap < a[0] or a[3] + gap < b[1] or b[3] + gap < a[1])


def cluster(rects, gap):
    rects = [tuple(r) for r in rects]
    changed = True
    while changed:
        changed = False
        out = []
        while rects:
            r = rects.pop()
            i = 0
            while i < len(rects):
                if near(r, rects[i], gap):
                    r = union(r, rects.pop(i))
                    changed = True
                else:
                    i += 1
            out.append(r)
        rects = out
    return rects


NUM_RE = re.compile(r"[-−–]?\d[\d,]*\.?\d*%?")
MATH_RE = re.compile(r"[=±×÷√∑∫∂≈≤≥∆Δπσμ^]")
CAPTION_RE = re.compile(r"^\s*(figure|fig\.?|chart|graph|plot|photo(graph)?|exhibit|diagram|map)\s*[\dA-Z]", re.I)
TABLE_CAP_RE = re.compile(r"^\s*table\s*[\dA-Z]", re.I)
FIG_NUM_RE = re.compile(r"\b(?:figure|fig\.?)\s*([\dA-Z]+(?:[.\-]\d+)?)", re.I)
TAB_NUM_RE = re.compile(r"\btable\s*([\dA-Z]+(?:[.\-]\d+)?)", re.I)
NOTE_RE = re.compile(r"^\s*(note|notes|source|sources|remark|\*|†|‡|¹|²|³|\(\d\)|\[\d\])", re.I)


# --------------------------------------------------------------------------- OCR
def osd_rotation(img_path):
    r = subprocess.run(["tesseract", str(img_path), "stdout", "--psm", "0"], capture_output=True, text=True)
    m = re.search(r"Rotate:\s*(\d+)", r.stdout)
    c = re.search(r"Orientation confidence:\s*([\d.]+)", r.stdout)
    if m and c and float(c.group(1)) >= 2.0:
        return int(m.group(1))
    return 0


def ocr_tsv(img_path, psm=3):
    r = subprocess.run(["tesseract", str(img_path), "stdout", "-l", "eng", "--psm", str(psm), "tsv"],
                       capture_output=True, text=True)
    words = []
    for line in r.stdout.splitlines()[1:]:
        f = line.split("\t")
        if len(f) == 12 and f[0] == "5" and f[11].strip():
            words.append({"block": int(f[2]), "par": int(f[3]), "line": int(f[4]), "left": int(f[6]),
                          "top": int(f[7]), "w": int(f[8]), "h": int(f[9]), "conf": float(f[10]), "text": f[11]})
    return words


def ocr_blocks(words):
    """Group OCR words into paragraph blocks (Tesseract layout order)."""
    paras = defaultdict(list)
    for w in words:
        paras[(w["block"], w["par"])].append(w)
    out = []
    for key in sorted(paras, key=lambda k: (min(w["top"] for w in paras[k]), min(w["left"] for w in paras[k]))):
        ws_ = paras[key]
        lines = defaultdict(list)
        for w in ws_:
            lines[w["line"]].append(w)
        text = "\n".join(" ".join(x["text"] for x in sorted(lines[l], key=lambda x: x["left"])) for l in sorted(lines))
        confs = [w["conf"] for w in ws_ if w["conf"] >= 0]
        low = [w["text"] for w in ws_ if 0 <= w["conf"] < LOW_CONF]
        out.append({"text": text, "mean_conf": sum(confs) / len(confs) if confs else 0, "low": low,
                    "n": len(ws_), "bbox_px": (min(w["left"] for w in ws_), min(w["top"] for w in ws_),
                                               max(w["left"] + w["w"] for w in ws_), max(w["top"] + w["h"] for w in ws_))})
    return out


# --------------------------------------------------------------------------- per-page analysis
def analyse_pdf(pdf_path, wdir):
    doc = pymupdf.open(pdf_path)
    pages = []
    for pno, page in enumerate(doc, 1):
        W, H = page.rect.width, page.rect.height
        parea = W * H
        info = {"pno": pno, "W": W, "H": H, "blocks": [], "tables": [], "figures": [], "issues": [], "ocr": "none"}
        tdict = page.get_text("dict", sort=True)
        text_blocks = []
        for b in tdict["blocks"]:
            if b.get("type") != 0:
                continue
            lines, sizes, bold = [], [], 0
            for l in b["lines"]:
                spans = [s for s in l["spans"] if s["text"].strip()]
                if not spans:
                    continue
                lines.append("".join(s["text"] for s in l["spans"]).rstrip())
                sizes += [s["size"] for s in spans]
                bold += sum(1 for s in spans if "bold" in s["font"].lower() or s["flags"] & 16)
            if lines:
                text_blocks.append({"bbox": tuple(b["bbox"]), "text": "\n".join(lines),
                                    "size": max(sizes) if sizes else 0, "bold": bold > 0, "src": "native"})
        native_chars = sum(len(b["text"].strip()) for b in text_blocks)
        imgs = [tuple(i["bbox"]) for i in page.get_image_info() if rect_area(i["bbox"]) > 0]
        img_cov = min(1.0, sum(inter(r, (0, 0, W, H)) for r in imgs) / parea) if parea else 0
        unusable = native_chars > 0 and sum(ch == "�" for b in text_blocks for ch in b["text"]) / max(1, native_chars) > 0.2

        # ---- tables (native only)
        tabs = []
        if native_chars > 30:
            try:
                for t in page.find_tables(strategy="lines").tables:
                    data = t.extract()
                    nrow, ncol = len(data), max((len(r) for r in data), default=0)
                    filled = sum(1 for r in data for c in r if c not in (None, "") and str(c).strip())
                    if nrow >= 2 and ncol >= 2 and filled / max(1, nrow * ncol) >= 0.3:
                        hdr = [str(c).strip() if c else "" for c in (t.header.names if t.header else data[0])]
                        tabs.append({"bbox": tuple(t.bbox), "data": data, "nrow": nrow, "ncol": ncol,
                                     "fill": filled / (nrow * ncol), "header": hdr, "page": pno})
            except Exception as e:
                info["issues"].append(("Table", f"Table detection error ({type(e).__name__}); page text kept in Content."))
        # text belonging to tables is represented in the table sheets
        for b in text_blocks:
            b["in_table"] = next((i for i, t in enumerate(tabs) if frac_inside(b["bbox"], t["bbox"]) > 0.6), None)

        # ---- figures: cluster raster images + vector drawings outside tables
        graphics = [r for r in imgs]
        try:
            for d in page.get_drawings():
                r = tuple(d["rect"])
                if rect_area(r) > 0.9 * parea:
                    continue                                   # page background
                if any(frac_inside(r, t["bbox"]) > 0.8 for t in tabs):
                    continue                                   # table rules
                if (r[2] - r[0]) < 1 and (r[3] - r[1]) < 1:
                    continue
                graphics.append(r)
        except Exception:
            info["issues"].append(("Figure", "Vector drawings could not be read on this page; see Page Reference."))
        clusters = [c for c in cluster(graphics, 8)
                    if rect_area(c) >= 0.02 * parea and (c[2] - c[0]) >= 40 and (c[3] - c[1]) >= 30]
        full_page_scan = False
        figs = []
        for c in clusters:
            c = (max(0, c[0]), max(0, c[1]), min(W, c[2]), min(H, c[3]))
            if rect_area(c) >= 0.85 * parea:
                full_page_scan = True
                continue
            if any(frac_inside(c, t["bbox"]) > 0.7 for t in tabs):
                continue
            # include text labels overlapping the graphic (legend, axes)
            for b in text_blocks:
                if b["in_table"] is None and inter(b["bbox"], c) > 0.5 * rect_area(b["bbox"]):
                    c = union(c, b["bbox"])
            figs.append({"bbox": c, "page": pno})
        figs = [{"bbox": r, "page": pno} for r in cluster([f["bbox"] for f in figs], 2)]

        # ---- OCR decisions
        scanned = (native_chars < 30 and img_cov >= 0.5) or unusable
        blocks = []
        if scanned:
            info["ocr"] = "full page"
            png = wdir / f"ocr_p{pno:04d}.png"
            page.get_pixmap(dpi=OCR_DPI).save(png)
            rot = osd_rotation(png)
            if rot:
                Image.open(png).rotate(-rot, expand=True).save(png)
                info["issues"].append(("Page", f"OCR working copy rotated {rot}° (orientation detected)."))
            words = ocr_tsv(png)
            scale = 72 / OCR_DPI
            for ob in ocr_blocks(words):
                bb = tuple(v * scale for v in ob["bbox_px"])
                blk = {"bbox": bb, "text": ob["text"], "size": 0, "bold": False, "src": "ocr", "in_table": None,
                       "conf": ob["mean_conf"], "low": ob["low"]}
                blocks.append(blk)
            if unusable:
                info["issues"].append(("Text", "Native text unusable (replacement characters); page OCR'd instead."))
            if sum(len(b["text"]) for b in blocks) < 20:
                info["issues"].append(("Page", "Scanned page produced little or no OCR text — possible missing content; see Page Reference."))
            info["issues"].append(("Page", "Scanned page: tables and figures inside the scan are not reconstructed; "
                                           "OCR text in Content, full page image in Page Reference."))
            figs = []                                           # whole page is an image; not cropped
        else:
            blocks = [b for b in text_blocks]
            # mixed page: OCR image regions that carry no native text
            for f in figs:
                has_text = any(inter(b["bbox"], f["bbox"]) > 0.3 * rect_area(b["bbox"]) for b in text_blocks)
                img_in = sum(inter(r, f["bbox"]) for r in imgs) / max(1, rect_area(f["bbox"]))
                if not has_text and img_in > 0.5:
                    png = wdir / f"ocr_p{pno:04d}_r{len(blocks)}.png"
                    page.get_pixmap(dpi=OCR_DPI, clip=pymupdf.Rect(f["bbox"])).save(png)
                    words = ocr_tsv(png, psm=11)
                    good = [w for w in words if w["conf"] >= 0]
                    mean = sum(w["conf"] for w in good) / len(good) if good else 0
                    if len(good) >= 3 and mean < 40:
                        info["issues"].append(("Figure", f"OCR of an image region gave very low confidence (mean {mean:.0f}, "
                                                         f"{len(good)} fragments) — likely a photo/graphic; text not captured, "
                                                         "image kept as a figure."))
                    elif len(good) >= 3:
                        scale = 72 / OCR_DPI
                        for ob in ocr_blocks(words):
                            bb = (f["bbox"][0] + ob["bbox_px"][0] * scale, f["bbox"][1] + ob["bbox_px"][1] * scale,
                                  f["bbox"][0] + ob["bbox_px"][2] * scale, f["bbox"][1] + ob["bbox_px"][3] * scale)
                            blocks.append({"bbox": bb, "text": ob["text"], "size": 0, "bold": False, "src": "ocr-region",
                                           "in_table": None, "conf": ob["mean_conf"], "low": ob["low"], "fig_text": True})
                        info["ocr"] = "image regions"
            if full_page_scan and native_chars >= 30:
                info["issues"].append(("Page", "Page has a full-page image behind native text; image content may not be captured as text."))
        info["blocks"], info["tables"], info["figures"] = blocks, tabs, figs
        info["native_chars"], info["img_cov"], info["scanned"] = native_chars, img_cov, scanned
        info["native_text_raw"] = page.get_text() if not scanned else ""
        pages.append(info)
    warn = pymupdf.TOOLS.mupdf_warnings()
    return doc, pages, (len(warn.splitlines()) if warn else 0)


# --------------------------------------------------------------------------- classification
def classify(pages):
    """Assign content types; detect running headers/footers."""
    sizes = [b["size"] for p in pages for b in p["blocks"] if b["src"] == "native" and b["size"]]
    med = sorted(sizes)[len(sizes) // 2] if sizes else 10
    edge_txt = Counter()
    for p in pages:
        for b in p["blocks"]:
            if b["bbox"][3] < 0.08 * p["H"] or b["bbox"][1] > 0.92 * p["H"]:
                edge_txt[re.sub(r"\d+", "#", b["text"].strip().lower())] += 1
    repeat = {k for k, v in edge_txt.items() if v >= max(3, 0.3 * len(pages))}
    for p in pages:
        for b in p["blocks"]:
            t = b["text"].strip()
            key = re.sub(r"\d+", "#", t.lower())
            top, bot = b["bbox"][3] < 0.08 * p["H"], b["bbox"][1] > 0.92 * p["H"]
            if b.get("fig_text"):
                ty = "Figure text (OCR)"
            elif b["in_table"] is not None:
                ty = "Table text"
            elif (top or bot) and (key in repeat or re.fullmatch(r"(page\s*)?\d+(\s*(of|/)\s*\d+)?", t, re.I)):
                ty = "Header" if top else "Footer"
            elif CAPTION_RE.match(t) or TABLE_CAP_RE.match(t):
                ty = "Caption"
            elif b["src"] == "native" and len(MATH_RE.findall(t)) >= 2 and len(re.findall(r"[A-Za-z]{4,}", t)) <= 3:
                ty = "Equation"
            elif b["src"] == "native" and b["size"] >= med * 1.25 and len(t) <= 200:
                ty = "Heading"
            elif b["src"] == "native" and b["size"] and b["size"] <= med * 0.8 and b["bbox"][1] > 0.75 * p["H"]:
                ty = "Footnote"
            elif re.match(r"^\s*([•▪◦\-–•]|\d+[.)])\s", t):
                ty = "List item"
            else:
                ty = "Paragraph"
            if b["src"] == "ocr":
                ty += " (OCR)" if not ty.endswith("(OCR)") else ""
            b["type"] = ty
        # figure text: native blocks inside a figure crop
        for f in p["figures"]:
            for b in p["blocks"]:
                if b["src"] == "native" and b["in_table"] is None and frac_inside(b["bbox"], f["bbox"]) > 0.6 \
                        and b["type"] not in ("Caption",):
                    b["type"] = "Figure text"


# --------------------------------------------------------------------------- tables, figures, relations
def logical_tables(pages):
    """Merge tables continued across consecutive pages only when column count and header row match exactly."""
    out = []
    for p in pages:
        for i, t in enumerate(p["tables"]):
            t["id"] = None
            prev = out[-1] if out else None
            cont = (prev is not None and i == 0 and prev["pages"][-1] == p["pno"] - 1 and prev["ncol"] == t["ncol"]
                    and any(h for h in t["header"]) and [h.lower() for h in prev["header"]] == [h.lower() for h in t["header"]]
                    and prev["last_index_on_page"] == len(pages[prev["pages"][-1] - 1]["tables"]) - 1)
            if cont:
                prev["parts"].append(t)
                prev["pages"].append(p["pno"])
                prev["last_index_on_page"] = i
            else:
                out.append({"parts": [t], "pages": [p["pno"]], "ncol": t["ncol"], "header": t["header"],
                            "last_index_on_page": i})
    for n, lt in enumerate(out, 1):
        lt["id"] = f"Table_{n:03d}"
        for t in lt["parts"]:
            t["id"] = lt["id"]
    return out


def attach_text(pages, ltables):
    """Title above / notes below each table part; captions for figures."""
    for p in pages:
        free = [b for b in p["blocks"] if b["in_table"] is None]
        for t in p["tables"]:
            tb = t["bbox"]
            above = [b for b in free if b["bbox"][3] <= tb[1] + 2 and tb[1] - b["bbox"][3] < 40
                     and inter((tb[0], 0, tb[2], 1e9), (b["bbox"][0], 0, b["bbox"][2], 1e9)) > 0]
            below = [b for b in free if b["bbox"][1] >= tb[3] - 2 and b["bbox"][1] - tb[3] < 30 and NOTE_RE.match(b["text"])]
            t["title_block"] = max(above, key=lambda b: b["bbox"][3]) if above else None
            t["note_blocks"] = sorted(below, key=lambda b: b["bbox"][1])
            for b in ([t["title_block"]] if t["title_block"] else []) + t["note_blocks"]:
                b["belongs_to"] = t["id"]
        for f in p["figures"]:
            fb = f["bbox"]
            cands = [b for b in free if (0 <= b["bbox"][1] - fb[3] < 30 or 0 <= fb[1] - b["bbox"][3] < 30)
                     and inter((fb[0], 0, fb[2], 1e9), (b["bbox"][0], 0, b["bbox"][2], 1e9)) > 0
                     and CAPTION_RE.match(b["text"])]
            f["caption_block"] = min(cands, key=lambda b: abs(b["bbox"][1] - fb[3])) if cands else None
            f["caption"] = f["caption_block"]["text"] if f["caption_block"] else ""
            f["inner_text"] = [b for b in p["blocks"] if b.get("type") in ("Figure text", "Figure text (OCR)")
                               and frac_inside(b["bbox"], fb) > 0.6]


def relate(pages, ltables, figures):
    """Figure <-> table only by explicit numbered references (caption/title/notes); never proximity alone."""
    tab_nums = {}
    for lt in ltables:
        texts = " ".join((t["title_block"]["text"] if t.get("title_block") else "") for t in lt["parts"])
        m = TAB_NUM_RE.search(texts)
        if m:
            tab_nums[m.group(1).lower()] = lt
        lt["fig_refs"] = {x.lower() for t in lt["parts"] for b in ([t.get("title_block")] + t.get("note_blocks", []))
                          if b for x in FIG_NUM_RE.findall(b["text"])}
    issues = []
    for f in figures:
        f["tables"] = []
        cap = f["caption"]
        for m in TAB_NUM_RE.findall(cap):
            if m.lower() in tab_nums:
                f["tables"].append(tab_nums[m.lower()]["id"])
        fm = FIG_NUM_RE.search(cap)
        if fm:
            for lt in ltables:
                if fm.group(1).lower() in lt["fig_refs"] and lt["id"] not in f["tables"]:
                    f["tables"].append(lt["id"])
        if not f["tables"]:
            same = [t["id"] for t in pages[f["page"] - 1]["tables"]]
            if same:
                issues.append((f["page"], f["id"], "Figure is on the same page as " + ", ".join(sorted(set(same))) +
                               " but no explicit caption/reference links them — kept separate (relationship uncertain)."))
    return issues


# --------------------------------------------------------------------------- workbook writing
def img_size_px(path, max_w):
    with Image.open(path) as im:
        w, h = im.size
    if w > max_w:
        h = int(h * max_w / w)
        w = max_w
    return w, h


PLACEMENTS = []      # (workbook file, sheet, col, first_row, last_row) recorded at write time


def add_image(ws, path, anchor_col, anchor_row, max_w):
    w, h = img_size_px(path, max_w)
    img = XLImage(str(path))
    img.width, img.height = w, h
    ws.add_image(img, f"{get_column_letter(anchor_col)}{anchor_row}")
    rows = math.ceil(h / ROW_PX)
    PLACEMENTS.append((id(ws.parent), ws.title, anchor_col, anchor_row, anchor_row + rows - 1, w))
    return rows + 2, w, h


def render_crop(page, bbox, path, dpi=FIG_DPI, pad=4):
    r = pymupdf.Rect(bbox[0] - pad, bbox[1] - pad, bbox[2] + pad, bbox[3] + pad) & page.rect
    page.get_pixmap(dpi=dpi, clip=r).save(path)
    return path


THIN = Side(style="thin", color="808080")
GRID = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HDR_FILL = PatternFill("solid", fgColor="D9E1F2")
LABEL_FILL = PatternFill("solid", fgColor="F2F2F2")
NUMLIKE = re.compile(r"^[\s(]*[-−–+]?[\d,]*\.?\d+\s*[%)]?\s*$")


def style_table(ws, first_row, last_row, ncol, header_rows):
    """Proper table view: grid on every cell (blanks included), shaded bold header row(s),
    wrapped top-aligned text, numeric-looking text right-aligned (still stored as text)."""
    for r in range(first_row, last_row + 1):
        for c in range(1, ncol + 1):
            cell = ws.cell(r, c)
            cell.border = GRID
            v = cell.value
            if r in header_rows:
                cell.font = Font(bold=True)
                cell.fill = HDR_FILL
                cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
            else:
                right = isinstance(v, str) and bool(NUMLIKE.match(v))
                cell.alignment = Alignment(wrap_text=True, vertical="top", horizontal="right" if right else "left")


def fit_table(ws, ranges, ncol):
    """Column widths from content (capped) and row heights for wrapped text."""
    widths = {}
    for first, last in ranges:
        for r in range(first, last + 1):
            for c in range(1, ncol + 1):
                v = ws.cell(r, c).value
                if v:
                    longest = max(len(x) for x in str(v).split("\n"))
                    widths[c] = max(widths.get(c, 0), longest)
    colw = {}
    for c in range(1, ncol + 1):
        colw[c] = min(max(widths.get(c, 0) + 2, 9), 50)
        ws.column_dimensions[get_column_letter(c)].width = max(colw[c], ws.column_dimensions[get_column_letter(c)].width or 0
                                                               if c == 1 else colw[c])
    for first, last in ranges:
        for r in range(first, last + 1):
            lines = 1
            for c in range(1, ncol + 1):
                v = ws.cell(r, c).value
                if v:
                    w = ws.column_dimensions[get_column_letter(c)].width or colw[c]
                    lines = max(lines, sum(max(1, math.ceil(len(x) / max(1, w - 1))) for x in str(v).split("\n")))
            ws.row_dimensions[r].height = min(15 * lines + 2, 300)


def head(ws, hdr, widths):
    for i, (h, w) in enumerate(zip(hdr, widths), 1):
        put(ws, 1, i, h, bold=True)
        ws.cell(1, i).fill = PatternFill("solid", fgColor="DDEBF7")
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"


def build_workbook(pdf, rel, doc, pages, mupdf_warn, wdir, out_path):
    classify(pages)
    ltables = logical_tables(pages)
    attach_text(pages, ltables)
    figures = []
    for p in pages:
        for f in p["figures"]:
            f["id"] = f"F{len(figures) + 1:03d}"
            figures.append(f)
    rel_issues = relate(pages, ltables, figures)
    review = []                                      # (page, item, issue, reason)
    for p in pages:
        for kind, msg in p["issues"]:
            review.append((p["pno"], kind, msg.split(";")[0], msg))
    for pg, fid, msg in rel_issues:
        review.append((pg, f"Figure {fid}", "Uncertain table–figure relationship", msg))

    wb = Workbook()
    ix = wb.active
    ix.title = "Index"
    content = wb.create_sheet("Content")
    head(content, ["PDF Page", "Block Order", "Content Type", "Extracted Text", "Review Status", "Link"], [9, 10, 18, 100, 34, 22])

    # ---- table sheets
    fig_loc = {}
    tsheets = {}
    for lt in ltables:
        ws = wb.create_sheet(lt["id"])
        tsheets[lt["id"]] = ws
        title = next((t["title_block"]["text"] for t in lt["parts"] if t.get("title_block")), "")
        put(ws, 1, 1, "Table ID", bold=True); put(ws, 1, 2, lt["id"])
        put(ws, 2, 1, "Title (as printed)", bold=True); put(ws, 2, 2, title or "(no title detected)", wrap=False)
        put(ws, 3, 1, "Source PDF pages", bold=True); put(ws, 3, 2, ", ".join(map(str, lt["pages"])))
        put(ws, 4, 1, "Continued across pages", bold=True)
        put(ws, 4, 2, "Yes — merged: identical column count and header row on consecutive pages" if len(lt["pages"]) > 1 else "No")
        link(ws, 5, 1, "◀ Index", "#'Index'!A1")
        for rr in range(1, 5):
            ws.cell(rr, 1).fill = LABEL_FILL
        r = 7
        lt["anchor_rows"] = []
        ranges, header_rows = [], set()
        unreliable = False
        for k, t in enumerate(lt["parts"]):
            put(ws, r, 1, f"Source page {t['page']}" + (" (continued)" if k else ""), bold=True)
            r += 1
            start = r
            rows = t["data"] if k == 0 else t["data"][1:] if [str(c or "").strip().lower() for c in t["data"][0]] == \
                [h.lower() for h in lt["header"]] else t["data"]
            first_is_header = [str(c or "").strip().lower() for c in rows[0]] == [h.lower() for h in lt["header"]] \
                if rows else False
            for row in rows:
                for c, val in enumerate(row, 1):
                    if val not in (None, ""):
                        put(ws, r, c, val, wrap=True)
                r += 1
            if rows:
                ranges.append((start, r - 1))
                if first_is_header:
                    header_rows.add(start)
            lt["anchor_rows"].append(start)
            for nb in t.get("note_blocks", []):
                put(ws, r, 1, "Note (as printed)", bold=True); put(ws, r, 2, nb["text"], wrap=True)
                r += 1
            if t["fill"] < 0.5 or t["ncol"] > 15:
                unreliable = True
            r += 1
        ws.column_dimensions["A"].width = 24
        for first, last in ranges:
            style_table(ws, first, last, lt["ncol"], header_rows)
        fit_table(ws, ranges, lt["ncol"])
        if ranges:
            ws.freeze_panes = f"A{min(header_rows) + 1}" if header_rows else f"A{ranges[0][0]}"
        lt["ranges"], lt["header_rows"] = ranges, header_rows
        # images below the table: reference crops + related figures (never over cells)
        r += 1
        put(ws, r, 1, "Table reference image(s) for verification", bold=True)
        r += 1
        for t in lt["parts"]:
            crop = render_crop(doc[t["page"] - 1], t["bbox"], wdir / f"{lt['id']}_p{t['page']}.png")
            put(ws, r, 1, f"Source page {t['page']} — table region")
            used, _, _ = add_image(ws, crop, 1, r + 1, 900)
            r += used + 1
        lt["img_row_end"] = r
        lt["unreliable"] = unreliable
        if unreliable:
            review.append((lt["pages"][0], lt["id"], "Uncertain table structure",
                           "Low cell fill or very wide table — editable extraction may be incomplete; compare with the reference image."))
        lt["next_row"] = r + 1

    # ---- figures: related -> table sheet; others -> Figures sheet
    figs_ws = wb.create_sheet("Figures")
    head(figs_ws, ["Figure ID", "Caption (as printed)", "Source PDF Page", "Destination Worksheet", "Location", "Notes"],
         [10, 50, 14, 22, 18, 50])
    fig_img_col, fig_row = 8, 2
    for f in figures:
        page = doc[f["page"] - 1]
        crop = render_crop(page, f["bbox"], wdir / f"{f['id']}_p{f['page']}.png")
        f["crop"] = crop
        if f["tables"]:
            primary = f["tables"][0]
            ws = tsheets[primary]
            lt = next(x for x in ltables if x["id"] == primary)
            r = lt["next_row"]
            put(ws, r, 1, f"Related figure {f['id']} (source page {f['page']})", bold=True)
            put(ws, r + 1, 1, "Caption", bold=True); put(ws, r + 1, 2, f["caption"] or "(none)", wrap=True)
            used, _, _ = add_image(ws, crop, 1, r + 2, 900)
            fig_loc[f["id"]] = (primary, r)
            lt["next_row"] = r + 2 + used + 1
            for other in f["tables"][1:]:
                ows = tsheets[other]
                olt = next(x for x in ltables if x["id"] == other)
                link(ows, olt["next_row"], 1, f"Related figure {f['id']} → embedded in {primary}", f"#'{primary}'!A{r}")
                olt["next_row"] += 2
        else:
            used, _, _ = add_image(figs_ws, crop, fig_img_col, fig_row, 700)
            fig_loc[f["id"]] = ("Figures", fig_row)
            f["fig_sheet_rows"] = (fig_row, used)
            fig_row += used + 1
    for i, f in enumerate(figures, 2):
        dest, row = fig_loc[f["id"]]
        put(figs_ws, i, 1, f["id"]); put(figs_ws, i, 2, f["caption"] or "(no caption detected)", wrap=True)
        put(figs_ws, i, 3, f["page"]); put(figs_ws, i, 4, dest)
        link(figs_ws, i, 5, "Go to image", f"#'{dest}'!{'H' if dest == 'Figures' else 'A'}{row}")
        notes = []
        if f.get("inner_text"):
            notes.append(f"{len(f['inner_text'])} text block(s) inside figure listed in Content")
        if not f["caption"]:
            notes.append("no caption found")
        put(figs_ws, i, 6, "; ".join(notes), wrap=True)

    # ---- Content in reading order, with table/figure references at their positions
    order = 0
    for p in pages:
        items = []
        for b in p["blocks"]:
            if b["in_table"] is not None:
                continue
            items.append(("block", b["bbox"][1], b["bbox"][0], b))
        seen_t = set()
        for t in p["tables"]:
            if t["id"] not in seen_t:
                items.append(("table", t["bbox"][1], t["bbox"][0], t))
                seen_t.add(t["id"])
        for f in p["figures"]:
            items.append(("figure", f["bbox"][1], f["bbox"][0], f))
        if not p["scanned"]:
            items.sort(key=lambda x: (round(x[1] / 4), x[2]))         # native: y then x (PyMuPDF order kept within band)
        if not items:
            order += 1
            r = content.max_row + 1
            put(content, r, 1, p["pno"]); put(content, r, 2, order); put(content, r, 3, "No content detected")
            put(content, r, 4, "(no extractable text, table or figure on this page — see Page Reference)")
            put(content, r, 5, "Review: possible missing content")
            link(content, r, 6, "Page image", f"#'Page Reference'!A{1}")
            review.append((p["pno"], "Page", "No content detected", "Page produced no text/table/figure; check page image."))
        for kind, _, _, obj in items:
            r = content.max_row + 1
            order += 1
            if kind == "table":
                put(content, r, 1, p["pno"]); put(content, r, 2, order); put(content, r, 3, "Table reference")
                put(content, r, 4, f"[{obj['id']}] table on this page — values in worksheet {obj['id']}")
                link(content, r, 6, obj["id"], f"#'{obj['id']}'!A1")
                continue
            if kind == "figure":
                dest, row = fig_loc[obj["id"]]
                put(content, r, 1, p["pno"]); put(content, r, 2, order); put(content, r, 3, "Figure reference")
                put(content, r, 4, f"[{obj['id']}] figure" + (f" — {obj['caption'][:200]}" if obj["caption"] else ""))
                link(content, r, 6, f"{obj['id']} in {dest}", f"#'{dest}'!{'H' if dest == 'Figures' else 'A'}{row}")
                continue
            b = obj
            status = ""
            if b["src"].startswith("ocr"):
                lowfrac = len(b.get("low", [])) / max(1, len(b["text"].split()))
                if b.get("conf", 100) < LOW_CONF or lowfrac > 0.3:
                    status = f"Review: low OCR confidence (mean {b.get('conf', 0):.0f})"
                    review.append((p["pno"], f"Content row {r}", "Uncertain OCR text",
                                   f"Mean confidence {b.get('conf', 0):.0f}; {len(b.get('low', []))} low-confidence word(s)."))
                else:
                    status = f"OCR (mean conf {b.get('conf', 0):.0f})"
            if b["type"] == "Equation":
                status = "Review: equation — verify symbols against page image"
                review.append((p["pno"], f"Content row {r}", "Complex equation/symbols", "Verify against Page Reference."))
            if "�" in b["text"]:
                status = (status + "; " if status else "") + "Review: unreadable characters (�)"
                review.append((p["pno"], f"Content row {r}", "Unclear characters", "Replacement characters in extracted text."))
            if b.get("belongs_to"):
                status = (status + "; " if status else "") + f"also in {b['belongs_to']}"
            parts = chunks(b["text"])
            for k, part in enumerate(parts):
                rr = content.max_row + 1 if k else r
                put(content, rr, 1, p["pno"]); put(content, rr, 2, order)
                put(content, rr, 3, b["type"] + (f" (part {k + 1}/{len(parts)})" if len(parts) > 1 else ""))
                put(content, rr, 4, part, wrap=True); put(content, rr, 5, status)
    content.auto_filter.ref = content.dimensions

    # ---- Page Reference (may move to part workbooks if too large)
    page_imgs = []
    for p in pages:
        jp = wdir / f"page_{p['pno']:04d}.jpg"
        pix = doc[p["pno"] - 1].get_pixmap(dpi=PAGE_DPI)
        Image.frombytes("RGB", (pix.width, pix.height), pix.samples).save(jp, quality=72)
        page_imgs.append(jp)
    fixed_bytes = sum(Path(x).stat().st_size for x in wdir.glob("*.png") if not x.name.startswith("ocr_"))
    page_bytes = sum(x.stat().st_size for x in page_imgs)
    parts = []
    if fixed_bytes + page_bytes <= MAX_PART_IMAGE_BYTES:
        groups = [("main", list(range(len(page_imgs))))]
    else:
        per = max(1, int(len(page_imgs) * MAX_PART_IMAGE_BYTES / max(1, page_bytes)))
        groups = [(f"part{i + 1}", list(range(s, min(s + per, len(page_imgs)))))
                  for i, s in enumerate(range(0, len(page_imgs), per))]
    pr = wb.create_sheet("Page Reference")
    head(pr, ["PDF Page", "Page image"], [10, 12])
    page_row = {}

    def fill_pageref(ws, idxs):
        r = 2
        for i in idxs:
            put(ws, r, 1, f"Page {i + 1}", bold=True)
            page_row[i + 1] = (ws.title, r)
            used, _, _ = add_image(ws, page_imgs[i], 2, r, 1000)
            r += used + 1

    if groups[0][0] == "main":
        fill_pageref(pr, groups[0][1])
    else:
        put(pr, 2, 1, "Page images are in separate part workbooks (size limit):", bold=True)
        for k, (name, idxs) in enumerate(groups):
            part_path = out_path.with_name(f"{out_path.stem}_PageRef_{name}.xlsx")
            put(pr, 3 + k, 1, f"{part_path.name}: pages {idxs[0] + 1}-{idxs[-1] + 1}")
            parts.append((part_path, idxs))

    # ---- possible omissions: numbers in native page text not present in Content/table values
    out_nums = defaultdict(Counter)
    for p in pages:
        for b in p["blocks"]:
            if b["in_table"] is None:
                out_nums[p["pno"]].update(NUM_RE.findall(b["text"]))
        for t in p["tables"]:
            for row in t["data"]:
                for c in row:
                    if c:
                        out_nums[p["pno"]].update(NUM_RE.findall(str(c)))
    for p in pages:
        if p["scanned"]:
            continue
        miss = Counter(NUM_RE.findall(p["native_text_raw"])) - out_nums[p["pno"]]
        if miss:
            review.append((p["pno"], "Page text", "Possible omission (numbers)",
                           f"{sum(miss.values())} numeric token(s) in the PDF text layer not found in Content/table cells "
                           "(often text inside table cells split differently or hidden layers); compare with Page Reference."))

    # ---- Review Required
    rv = wb.create_sheet("Review Required")
    head(rv, ["Source Page", "Affected Item", "Issue", "Reason / Detail"], [11, 26, 40, 100])
    for k, (pg, item, issue, reason) in enumerate(sorted(review, key=lambda x: (x[0], x[1])), 2):
        put(rv, k, 1, pg); put(rv, k, 2, item); put(rv, k, 3, issue, wrap=True); put(rv, k, 4, reason, wrap=True)

    # ---- Index
    n_ocr = sum(1 for p in pages if p["ocr"] != "none")
    n_text = sum(1 for p in pages if not p["scanned"])
    rows = [("Source filename", pdf.name), ("Source path", str(pdf)), ("Relative path", str(rel)),
            ("Total PDF pages", len(pages)), ("Pages with reliable existing text", n_text),
            ("Pages requiring OCR (full page)", sum(1 for p in pages if p["ocr"] == "full page")),
            ("Pages with OCR of image regions (mixed pages)", sum(1 for p in pages if p["ocr"] == "image regions")),
            ("Logical tables", len(ltables)), ("Figures", len(figures)), ("Review items", len(review)),
            ("MuPDF parser warnings (structure)", mupdf_warn), ("Processing status", "Completed — see Review Required"),
            ("Program version", PROGRAM_VERSION), ("Processed", time.strftime("%Y-%m-%d %H:%M"))]
    for i, (k, v) in enumerate(rows, 1):
        put(ix, i, 1, k, bold=True); put(ix, i, 2, v)
    r = len(rows) + 2
    put(ix, r, 1, "Navigation", bold=True)
    for name in ["Content", "Figures", "Page Reference", "Review Required"]:
        r += 1
        link(ix, r, 1, name, f"#'{name}'!A1")
    for lt in ltables:
        r += 1
        link(ix, r, 1, lt["id"], f"#'{lt['id']}'!A1")
        put(ix, r, 2, f"pages {', '.join(map(str, lt['pages']))}")
    for part_path, idxs in parts:
        r += 1
        put(ix, r, 1, "Page Reference part"); put(ix, r, 2, f"{part_path.name} (pages {idxs[0] + 1}-{idxs[-1] + 1})")
    ix.column_dimensions["A"].width, ix.column_dimensions["B"].width = 44, 90

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    for part_path, idxs in parts:
        pw = Workbook()
        ws = pw.active
        ws.title = "Page Reference"
        head(ws, ["PDF Page", "Page image"], [10, 12])
        put(ws, 1, 3, f"Part of {out_path.name}")
        fill_pageref(ws, idxs)
        pw.save(part_path)

    stats = {"_wb_id": id(wb), "pages": len(pages), "text_pages": n_text, "ocr_full": sum(1 for p in pages if p["ocr"] == "full page"),
             "ocr_regions": sum(1 for p in pages if p["ocr"] == "image regions"), "tables": len(ltables),
             "tables_multi_page": sum(1 for lt in ltables if len(lt["pages"]) > 1), "figures": len(figures),
             "figures_in_table_sheets": sum(1 for f in figures if f["tables"]), "review": len(review),
             "review_by_issue": Counter(x[2] for x in review), "parts": [str(p_[0].name) for p_ in parts],
             "mupdf_warnings": mupdf_warn}
    return stats, ltables, figures, pages


# --------------------------------------------------------------------------- validation (reopen)
def validate(out_path, stats, ltables, figures, pages):
    problems = []
    wb = load_workbook(out_path)
    names = wb.sheetnames
    for s in ["Index", "Content", "Figures", "Page Reference", "Review Required"]:
        if s not in names:
            problems.append(f"missing sheet {s}")
    content = list(wb["Content"].iter_rows(min_row=2, values_only=True))
    pages_in_content = {int(r[0]) for r in content if r[0]}
    if pages_in_content != set(range(1, stats["pages"] + 1)):
        problems.append(f"pages missing from Content: {sorted(set(range(1, stats['pages'] + 1)) - pages_in_content)[:10]}")
    pr_pages = set()
    for ws in [wb["Page Reference"]] + [load_workbook(out_path.with_name(n))["Page Reference"] for n in stats["parts"]]:
        pr_pages |= {int(str(v).split()[1]) for (v,) in ws.iter_rows(min_row=2, max_col=1, values_only=True)
                     if isinstance(v, str) and v.startswith("Page ")}
    if pr_pages != set(range(1, stats["pages"] + 1)):
        problems.append("Page Reference does not contain every page")
    # formulas must not exist
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if c.data_type == "f":
                    problems.append(f"formula cell in {ws.title}!{c.coordinate}")
    # table dimensions + images below data
    for lt in ltables:
        ws = wb[lt["id"]]
        exp = sum(len(t["data"]) for t in lt["parts"])
        got_rows = sum(1 for _ in lt["parts"])
        last_data_row = max((c.row for row in ws.iter_rows(min_row=7) for c in row if c.value is not None
                             and not str(c.value).startswith(("Table reference", "Source page", "Related figure", "Caption"))),
                            default=7)
        for img in ws._images:
            if img.anchor._from.row + 1 <= min(lt["anchor_rows"]) + exp - 1 and img.anchor._from.row + 1 >= min(lt["anchor_rows"]):
                problems.append(f"{lt['id']}: image anchored inside table data rows")
    # proper table view: every cell in each table range has a grid border; header rows bold
    for lt in ltables:
        ws = wb[lt["id"]]
        for first, last in lt.get("ranges", []):
            for r in range(first, last + 1):
                for c in range(1, lt["ncol"] + 1):
                    if ws.cell(r, c).border.left.style != "thin":
                        problems.append(f"{lt['id']}: cell {get_column_letter(c)}{r} missing table border")
                        break
        for hr in lt.get("header_rows", []):
            if not ws.cell(hr, 1).font.bold:
                problems.append(f"{lt['id']}: header row {hr} not bold")
    # figures related to tables are in that sheet
    for f in figures:
        if f["tables"] and not any(True for _ in wb[f["tables"][0]]._images):
            problems.append(f"{f['id']}: related figure not embedded in {f['tables'][0]}")
    # images in Figures sheet stay right of the inventory columns
    for img in wb["Figures"]._images:
        if img.anchor._from.col < 7:
            problems.append("Figures sheet image overlaps inventory columns")
    # overlap: recorded placements (exact pixel sizes at write time) must not intersect,
    # and every reopened image must sit at a recorded anchor
    rec = defaultdict(list)
    for wbid, sheet, col, r0, r1, w in PLACEMENTS:
        if wbid == stats["_wb_id"]:
            rec[sheet].append((col, r0, r1, w))
    for sheet, spans in rec.items():
        spans.sort(key=lambda x: (x[1], x[0]))
        for i, a in enumerate(spans):
            for b in spans[i + 1:]:
                if b[1] > a[2]:
                    break
                cols_a = (a[0], a[0] + a[3] / 64)        # ~64 px per default column width
                cols_b = (b[0], b[0] + b[3] / 64)
                if cols_a[0] < cols_b[1] and cols_b[0] < cols_a[1]:
                    problems.append(f"{sheet}: overlapping images at rows {a[1]}-{b[1]}")
        anchors = {(img.anchor._from.col + 1, img.anchor._from.row + 1) for img in wb[sheet]._images}
        if anchors != {(c, r0) for c, r0, _, _ in spans}:
            problems.append(f"{sheet}: reopened image anchors differ from placements")
    # numbers present on native pages vs output
    out_nums = defaultdict(Counter)
    for r in content:
        if r[0] and r[3]:
            out_nums[int(r[0])].update(NUM_RE.findall(str(r[3])))
    for lt in ltables:
        for t in lt["parts"]:
            for row in t["data"]:
                for c in row:
                    if c:
                        out_nums[t["page"]].update(NUM_RE.findall(str(c)))
    missing = {}
    for p in pages:
        if p["scanned"]:
            continue
        src = Counter(NUM_RE.findall(p["native_text_raw"]))
        miss = sum((src - out_nums[p["pno"]]).values())
        if miss:
            missing[p["pno"]] = miss
    stats.pop("_wb_id", None)
    return problems, missing


# --------------------------------------------------------------------------- driver
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--only")
    ap.add_argument("--force", action="store_true")
    a = ap.parse_args()
    for d in (OUT, WORK, LOGS, STATE):
        d.mkdir(parents=True, exist_ok=True)
    manifest = json.loads(MANIFEST.read_text()) if MANIFEST.exists() else {}
    pdfs = sorted(p for p in SRC.rglob("*") if p.is_file() and p.suffix.lower() == ".pdf")
    if a.only:
        pdfs = [p for p in pdfs if a.only in p.name]
    for pdf in pdfs:
        rel = pdf.relative_to(SRC)
        out_path = OUT / rel.with_suffix(".xlsx")
        key = str(rel)
        h_before = sha(pdf)
        m = manifest.get(key)
        if m and not a.force and m.get("sha256") == h_before and m.get("status") == "ok" and Path(m["output"]).exists() \
                and m.get("version") == PROGRAM_VERSION:
            print(f"SKIP (unchanged, completed): file #{pdfs.index(pdf) + 1}")
            continue
        t0 = time.time()
        wdir = WORK / rel.with_suffix("")
        wdir.mkdir(parents=True, exist_ok=True)
        for old in wdir.glob("*"):
            if old.is_file():
                old.unlink()
        try:
            # Work on a local copy (verified identical) so the NAS original is only read once.
            copy = wdir / "_working_copy.pdf"
            data = read_bytes(pdf)
            assert hashlib.sha256(data).hexdigest() == h_before
            copy.write_bytes(data)
            doc, pages, warn = analyse_pdf(copy, wdir)
            stats, ltables, figures, pages = build_workbook(pdf, rel, doc, pages, warn, wdir, out_path)
            problems, missing = validate(out_path, stats, ltables, figures, pages)
            h_after = sha(pdf)
            assert h_after == h_before, "SOURCE PDF CHANGED"
            stats.update({"validation_problems": problems, "pages_with_numbers_not_found_in_output": missing,
                          "seconds": round(time.time() - t0, 1)})
            status = "ok" if not problems else "ok-with-validation-issues"
            manifest[key] = {"sha256": h_before, "status": "ok", "output": str(out_path), "version": PROGRAM_VERSION,
                             "stats": {k: (dict(v) if isinstance(v, Counter) else v) for k, v in stats.items()},
                             "validation": status}
            (LOGS / f"{rel.with_suffix('').as_posix().replace('/', '__')}.json").write_text(
                json.dumps(manifest[key], indent=2, default=str))
            print(f"OK file #{pdfs.index(pdf) + 1}: pages={stats['pages']} text={stats['text_pages']} "
                  f"ocr_full={stats['ocr_full']} ocr_regions={stats['ocr_regions']} tables={stats['tables']} "
                  f"(multi-page {stats['tables_multi_page']}) figures={stats['figures']} "
                  f"(in table sheets {stats['figures_in_table_sheets']}) review={stats['review']} "
                  f"num_check_pages={len(missing)} validation_problems={len(problems)} parts={len(stats['parts'])} "
                  f"time={stats['seconds']}s")
        except Exception as e:
            manifest[key] = {"sha256": h_before, "status": "failed", "output": str(out_path), "version": PROGRAM_VERSION,
                             "error": f"{type(e).__name__}: {e}"}
            (LOGS / f"{rel.with_suffix('').as_posix().replace('/', '__')}_error.txt").write_text(traceback.format_exc())
            print(f"FAILED file #{pdfs.index(pdf) + 1}: {type(e).__name__} (details in logs/)")
        MANIFEST.write_text(json.dumps(manifest, indent=2, default=str))
    write_summary(manifest)


def write_summary(manifest):
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    hdr = ["Source (relative)", "Status", "Validation", "Pages", "Pages with reliable text", "Pages OCR (full)",
           "Pages OCR (image regions)", "Logical tables", "Multi-page tables", "Figures", "Figures in table sheets",
           "Review items", "Pages with numbers not found in output", "Validation problems", "Page-reference parts",
           "Output workbook", "Error"]
    head(ws, hdr, [44, 10, 22, 7, 10, 9, 10, 9, 9, 8, 10, 9, 14, 12, 12, 70, 40])
    for r, (k, m) in enumerate(sorted(manifest.items()), 2):
        s = m.get("stats", {})
        vals = [k, m["status"], m.get("validation", ""), s.get("pages"), s.get("text_pages"), s.get("ocr_full"),
                s.get("ocr_regions"), s.get("tables"), s.get("tables_multi_page"), s.get("figures"),
                s.get("figures_in_table_sheets"), s.get("review"), len(s.get("pages_with_numbers_not_found_in_output", {}) or {}),
                len(s.get("validation_problems", []) or []), ", ".join(s.get("parts", []) or []), m.get("output"), m.get("error", "")]
        for c, v in enumerate(vals, 1):
            put(ws, r, c, v)
        link(ws, r, 16, m.get("output", ""), Path(m.get("output", "")).name)
    rv = wb.create_sheet("Review By Issue")
    head(rv, ["Source (relative)", "Issue", "Count"], [44, 60, 8])
    r = 2
    for k, m in sorted(manifest.items()):
        for issue, n in (m.get("stats", {}).get("review_by_issue") or {}).items():
            put(rv, r, 1, k); put(rv, r, 2, issue); put(rv, r, 3, n)
            r += 1
    wb.save(OUT / "ARPR_Processing_Summary.xlsx")


if __name__ == "__main__":
    main()
