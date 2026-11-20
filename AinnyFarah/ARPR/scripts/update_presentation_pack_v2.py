#!/usr/bin/env python3
"""Create v2 of 'ARPR 1.1.2025 YEAR Presentation Pack_Dulang - for AOF.xlsx' (local processing only).

- Input workbook (current, manually edited) is copied, never overwritten.
- Source: the NEW 62-page NAS PDF (replaced 2026-10-16 17:12), physical pages counted from 1;
  blank pages 26-35 keep their positions.
- Removes worksheet 'Table_002' (user-confirmed). No other existing sheet is modified.
- Adds one sheet per requested slide group; text on these slides is mostly outlined vectors or
  images, so text comes from local OCR (Tesseract) of the rendered slide, or native words where present.
- Existing manually edited tables that fall on requested slides are preserved: their edited cell values
  are carried into the new sheet (never overwritten by OCR) and the original sheet is linked.
Console output: counts/status only.
"""
import difflib
import json
import math
import re
import shutil
import subprocess
import sys
from collections import Counter, defaultdict
from pathlib import Path

import pymupdf
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from PIL import Image, ImageChops, ImageStat

sys.path.insert(0, str(Path(__file__).parent))
import pdf_to_excel as P  # noqa: E402

STEM = "ARPR 1.1.2025 YEAR Presentation Pack_Dulang - for AOF"
CUR = P.OUT / f"{STEM}.xlsx"
NEW_OUT = P.OUT / f"{STEM}_v2.xlsx"
W = P.WORK / "_update_v2"
NEW_PDF_SRC = P.SRC / f"{STEM}.pdf"
OLD_COPY = P.WORK / STEM / "_working_copy.pdf"           # 63-page version the current workbook came from
REGEN_OLD = P.WORK / "_regen_old" / "regen_old.xlsx"      # unedited regeneration of the current workbook
REMOVE = "Table_002"
GROUPS = [  # (sheet name, pages, requested title or None, merge rule)
    ("Action Items S5-7", [5, 6, 7], "Summary for Action Item", "combine"),
    ("Field Resources AG S20", [20], "Field Resources Summary (AG)", "single"),
    ("Project Resources S23", [23], "Project Resources Summary", "single"),
    ("Slide 24 Tables", [24], None, "all"),
    ("Slides 49-50", [49, 50], None, "all"),
    ("Slides 52-55", [52, 53, 54, 55], None, "all"),
    ("Slides 57-61", [57, 58, 59, 60, 61], None, "all"),
]
LOG_SHEET, REVIEW_SHEET = "Update Log v2", "Update Review v2"
OCR_DPI = 300
FIRST_CONTENT_COL = 3          # A = source page, B = item type, C.. = content


# --------------------------------------------------------------------------- words / lines
def page_words(page, wdir, pno):
    """Native words if the page has real text; otherwise OCR of the rendered page (Tesseract psm 11 for
    sparse slide/table layouts). Returns (words, mode, rotation)."""
    native = [(w[0], w[1], w[2], w[3], w[4], 100.0, "native") for w in page.get_text("words") if w[4].strip()]
    if len(native) >= 40:
        mode, words = "native", native
        # OCR only image regions that carry no native words (avoid duplicate text)
        for i in page.get_image_info():
            bb = i["bbox"]
            if P.rect_area(bb) < 0.03 * page.rect.width * page.rect.height:
                continue
            if any(P.frac_inside(w[:4], bb) > 0.5 for w in native):
                continue
            png = wdir / f"p{pno:02d}_img_{len(words)}.png"
            page.get_pixmap(dpi=OCR_DPI, clip=pymupdf.Rect(bb) & page.rect).save(png)
            s = 72 / OCR_DPI
            for t in P.ocr_tsv(png, psm=11):
                if t["conf"] >= 0:
                    words.append((bb[0] + t["left"] * s, bb[1] + t["top"] * s, bb[0] + (t["left"] + t["w"]) * s,
                                  bb[1] + (t["top"] + t["h"]) * s, t["text"], t["conf"], "ocr-image"))
            mode = "native + OCR of image regions"
        return words, mode, 0
    png = wdir / f"p{pno:02d}_page.png"
    page.get_pixmap(dpi=OCR_DPI).save(png)
    rot = P.osd_rotation(png)
    if rot:
        return [], f"OCR skipped: page orientation {rot}° — review", rot
    s = 72 / OCR_DPI
    words = [(t["left"] * s, t["top"] * s, (t["left"] + t["w"]) * s, (t["top"] + t["h"]) * s, t["text"], t["conf"], "ocr")
             for t in P.ocr_tsv(png, psm=11) if t["conf"] >= 0]
    keep = {w[4] for w in native}
    for w in native:                                   # the few native words (e.g. titles) are kept too
        if not any(P.frac_inside(w[:4], o[:4]) > 0.5 and o[4] == w[4] for o in words):
            words.append(w)
    return words, "OCR (text drawn as vectors/images)", 0


def to_lines(words):
    if not words:
        return []
    hs = sorted(w[3] - w[1] for w in words)
    mh = hs[len(hs) // 2] or 8
    lines = []
    for w in sorted(words, key=lambda w: ((w[1] + w[3]) / 2, w[0])):
        cy = (w[1] + w[3]) / 2
        for ln in lines:
            if abs(ln["cy"] - cy) <= 0.5 * mh and not any(P.inter(w[:4], x[:4]) > 0.5 * P.rect_area(w[:4]) for x in ln["w"]):
                ln["w"].append(w)
                ln["cy"] = sum((x[1] + x[3]) / 2 for x in ln["w"]) / len(ln["w"])
                break
        else:
            lines.append({"cy": cy, "w": [w]})
    for ln in lines:
        ln["w"].sort(key=lambda x: x[0])
        ln["bbox"] = (min(x[0] for x in ln["w"]), min(x[1] for x in ln["w"]), max(x[2] for x in ln["w"]), max(x[3] for x in ln["w"]))
        ln["mh"] = mh
    lines.sort(key=lambda l: (l["bbox"][1], l["bbox"][0]))
    return lines


def segments(ln, gap):
    segs, cur = [], [ln["w"][0]]
    for a, b in zip(ln["w"], ln["w"][1:]):
        if b[0] - a[2] > gap:
            segs.append(cur)
            cur = [b]
        else:
            cur.append(b)
    segs.append(cur)
    return [{"x0": s[0][0], "x1": s[-1][2], "text": " ".join(x[4] for x in s), "conf": [x[5] for x in s],
             "src": {x[6] for x in s}} for s in segs]


def text_of(ws_):
    return " ".join(w[4] for w in sorted(ws_, key=lambda w: (round(((w[1] + w[3]) / 2) / 4), w[0])))


# --------------------------------------------------------------------------- tables
def grid_tables(page, words):
    out = []
    try:
        tabs = page.find_tables(strategy="lines").tables
    except Exception:
        return out
    for t in tabs:
        if t.row_count < 2 or t.col_count < 2:
            continue
        cells = [[None] * t.col_count for _ in range(t.row_count)]
        confs = [[[] for _ in range(t.col_count)] for _ in range(t.row_count)]
        used = set()
        for ri, row in enumerate(t.rows):
            for ci, cb in enumerate(row.cells):
                if cb is None:
                    continue
                inside = [w for k, w in enumerate(words) if k not in used and
                          cb[0] - 1 <= (w[0] + w[2]) / 2 <= cb[2] + 1 and cb[1] - 1 <= (w[1] + w[3]) / 2 <= cb[3] + 1]
                for k, w in enumerate(words):
                    if w in inside:
                        used.add(k)
                if inside:
                    cells[ri][ci] = text_of(inside)
                    confs[ri][ci] = [w[5] for w in inside]
        colx = []
        for ci in range(t.col_count):
            xs = [row.cells[ci] for row in t.rows if ci < len(row.cells) and row.cells[ci] is not None]
            colx.append((min(c[0] for c in xs), max(c[2] for c in xs)) if xs else None)
        out.append({"bbox": tuple(t.bbox), "cells": cells, "confs": confs, "colx": colx, "method": "vector cell grid + " +
                    ("native text" if all(w[6] == "native" for w in words) else "OCR text"), "nrow": t.row_count,
                    "ncol": t.col_count})
    return out


def layout_tables(lines, exclude):
    """Rows of >=2 aligned segments over >=3 consecutive lines -> table (no ruling lines)."""
    free = [l for l in lines if not any(P.frac_inside(l["bbox"], e) > 0.6 for e in exclude)]
    out, i = [], 0
    while i < len(free):
        gap = max(12, 1.8 * free[i]["mh"])
        run = []
        j = i
        while j < len(free):
            segs = segments(free[j], gap)
            if len(segs) >= 2 or (run and len(segs) == 1 and free[j]["bbox"][1] - run[-1][0]["bbox"][3] < 2 * free[j]["mh"]
                                  and len(run) >= 1 and j + 1 < len(free) and len(segments(free[j + 1], gap)) >= 2):
                run.append((free[j], segs))
                j += 1
            else:
                break
        multi = [r for r in run if len(r[1]) >= 2]
        if len(multi) >= 3:
            xs = sorted(s["x0"] for _, segs in run for s in segs)
            cols = []
            for x in xs:
                if cols and x - cols[-1][-1] <= 18:
                    cols[-1].append(x)
                else:
                    cols.append([x])
            anchors = [sum(c) / len(c) for c in cols if len(c) >= max(2, 0.25 * len(multi))]
            if len(anchors) >= 2:
                rows, confs = [], []
                for ln, segs in run:
                    row = [None] * len(anchors)
                    cf = [[] for _ in anchors]
                    for s in segs:
                        k = min(range(len(anchors)), key=lambda a: abs(anchors[a] - s["x0"]))
                        row[k] = (row[k] + " " + s["text"]) if row[k] else s["text"]
                        cf[k] += s["conf"]
                    rows.append(row)
                    confs.append(cf)
                bbox = (min(l["bbox"][0] for l, _ in run), min(l["bbox"][1] for l, _ in run),
                        max(l["bbox"][2] for l, _ in run), max(l["bbox"][3] for l, _ in run))
                out.append({"bbox": bbox, "cells": rows, "confs": confs, "nrow": len(rows), "ncol": len(anchors),
                            "method": "reconstructed from word layout (no ruling lines) — verify"})
                i = j
                continue
        i += 1
    return out


def norm(s):
    return re.sub(r"[^a-z0-9]+", " ", (s or "").lower()).strip()


def header_sim(a, b):
    ta, tb = " | ".join(norm(x) for x in a), " | ".join(norm(x) for x in b)
    return difflib.SequenceMatcher(None, ta, tb).ratio() if ta.strip(" |") and tb.strip(" |") else 0


def find_title(lines, title):
    if not title:
        return None
    t = norm(title)
    best = (0, None)
    for ln in lines:
        s = norm(" ".join(w[4] for w in ln["w"]))
        if not s:
            continue
        r = difflib.SequenceMatcher(None, t, s[:len(t) + 10]).ratio()
        if t in s:
            r = 1.0
        if r > best[0]:
            best = (r, ln)
    return best if best[0] >= 0.75 else None


def printed_number(lines, H):
    for ln in lines:
        if ln["bbox"][1] > 0.88 * H or ln["bbox"][3] < 0.10 * H:
            for w in ln["w"]:
                if re.fullmatch(r"\d{1,3}", w[4]):
                    return w[4]
    return None


# --------------------------------------------------------------------------- page mapping old -> new
def old_to_new_map(old_doc, new_doc):
    def th(pg):
        pix = pg.get_pixmap(dpi=20)
        return Image.frombytes("RGB", (pix.width, pix.height), pix.samples).convert("L").resize((64, 48))
    nt = [th(p) for p in new_doc]
    mp = {}
    for i, p in enumerate(old_doc, 1):
        o = th(p)
        sc = [ImageStat.Stat(ImageChops.difference(o, n)).mean[0] for n in nt]
        j = min(range(len(sc)), key=lambda k: sc[k])
        mp[i] = (j + 1, sc[j])
    return mp


def edited_cells(cur_ws, gen_ws):
    va = {(c.row, c.column): c.value for row in cur_ws.iter_rows() for c in row if c.value is not None}
    vb = {(c.row, c.column): c.value for row in gen_ws.iter_rows() for c in row if c.value is not None}
    return {k: va.get(k) for k in set(va) | set(vb) if va.get(k) != vb.get(k)}


# --------------------------------------------------------------------------- sheet writing
def put(ws, r, c, v, **kw):
    return P.put(ws, r, c, v, **kw)


def write_group(wb, name, pnos, title, rule, doc, analysed, existing, review, log):
    ws = wb.create_sheet(name)
    H = doc[0].rect.height
    ws.column_dimensions["A"].width, ws.column_dimensions["B"].width = 10, 20
    put(ws, 1, 1, name, bold=True)
    put(ws, 2, 1, "Source physical PDF pages", bold=True); put(ws, 2, 3, ", ".join(map(str, pnos)))
    put(ws, 3, 1, "Printed slide numbers", bold=True)
    put(ws, 4, 1, "Requested title", bold=True); put(ws, 4, 3, title or "(all content)")
    put(ws, 5, 1, "Method", bold=True)
    P.link(ws, 6, 1, "◀ Update Log", f"#'{LOG_SHEET}'!A1")
    for r in range(1, 6):
        ws.cell(r, 1).fill = P.LABEL_FILL
    r = 8
    printed, methods, maxcol = [], set(), 8
    items_by_page = {}
    for pno in pnos:
        page = doc[pno - 1]
        words, mode, rot = page_words(page, W, pno)
        methods.add(f"p{pno}: {mode}")
        lines = to_lines(words)
        pn = printed_number(lines, page.rect.height)
        printed.append(f"p{pno}: {pn if pn else 'none found'}")
        if pn and pn != str(pno):
            review.append((name, pno, "Slide number", f"Printed slide number {pn} differs from physical page {pno} — both recorded."))
        if rot:
            review.append((name, pno, "Orientation", f"Page appears rotated {rot}°; OCR not run — see slide image."))
        gts = grid_tables(page, words)
        lts = layout_tables(lines, [g["bbox"] for g in gts])
        tables = sorted(gts + lts, key=lambda t: (t["bbox"][1], t["bbox"][0]))
        for t in tables:
            t["page"] = pno
        tbox = [t["bbox"] for t in tables]
        hit = find_title(lines, title)
        if title:
            log["titles"].append((name, title, pno, "found" if hit else "not found",
                                  f"similarity {hit[0]:.2f}" if hit else ""))
        # text: lines outside tables -> paragraphs / list items
        text_items = []
        para = None
        for ln in lines:
            if any(P.frac_inside(ln["bbox"], b) > 0.6 for b in tbox):
                continue
            if pn and ln["bbox"][1] > 0.88 * page.rect.height and " ".join(w[4] for w in ln["w"]).strip() == pn:
                continue
            txt = " ".join(w[4] for w in ln["w"])
            conf = [w[5] for w in ln["w"]]
            is_item = bool(re.match(r"^\s*(\d{1,2}[.)]|[•▪◦\-–•]|[a-z][.)])\s", txt))
            if para and not is_item and ln["bbox"][1] - para["bbox"][3] < 1.3 * ln["mh"] and abs(ln["bbox"][0] - para["bbox"][0]) < 30:
                para["text"] += "\n" + txt
                para["conf"] += conf
                para["bbox"] = P.union(para["bbox"], ln["bbox"])
                para["src"] |= {w[6] for w in ln["w"]}
            else:
                para = {"bbox": ln["bbox"], "text": txt, "conf": conf, "type": "List item" if is_item else "Text",
                        "src": {w[6] for w in ln["w"]}, "title_hit": bool(hit and ln is hit[1])}
                text_items.append(para)
        figs = [f for f in analysed[pno - 1]["figures"]
                if not any(P.inter(f["bbox"], b) > 0.3 * P.rect_area(f["bbox"]) for b in tbox)
                and P.rect_area(f["bbox"]) < 0.7 * page.rect.width * page.rect.height]
        items_by_page[pno] = {"tables": tables, "text": text_items, "figs": figs, "mode": mode, "words": len(words),
                              "lines": lines}
    put(ws, 3, 3, "; ".join(printed))
    put(ws, 5, 3, "; ".join(sorted(methods)), wrap=False)

    if rule == "combine":
        combine_group(items_by_page, pnos, name, review)

    # continuation merge across consecutive pages (header match, same column count)
    seq = []
    for pno in pnos:
        for t in items_by_page[pno]["tables"]:
            seq.append(t)
    merged = []
    if rule == "combine" and seq and seq[0].get("combined_parts"):
        cp = seq[0]["combined_parts"]
        merged = [{"parts": cp, "pages": [x["page"] for x in cp], "ncol": cp[0]["ncol"], "combined": True}]
        seq = []
    for t in seq:
        prev = merged[-1] if merged else None
        if prev and rule != "single" and t["page"] == prev["pages"][-1] + 1 and t["ncol"] == prev["ncol"] \
                and t is items_by_page[t["page"]]["tables"][0] and prev["parts"][-1] is items_by_page[prev["pages"][-1]]["tables"][-1] \
                and header_sim(prev["parts"][0]["cells"][0], t["cells"][0]) >= 0.85:
            prev["parts"].append(t)
            prev["pages"].append(t["page"])
        else:
            merged.append({"parts": [t], "pages": [t["page"]], "ncol": t["ncol"]})
            if prev and rule != "single" and t["page"] == prev["pages"][-1] + 1 and t["ncol"] == prev["ncol"]:
                review.append((name, t["page"], "Possible continuation",
                               f"Table on p{t['page']} has the same column count as the previous slide's table but a different/absent "
                               "header row — kept as a separate table."))

    # existing edited tables on these pages (from the current workbook)
    for ex in existing:
        if ex["new_page"] in pnos:
            for lt in merged:
                for part in lt["parts"]:
                    if part["page"] == ex["new_page"] and part["ncol"] == ex["ncol"] and part["nrow"] == ex["nrow"] \
                            and part.get("method", "").startswith("vector") and not part.get("existing") \
                            and abs(part["bbox"][1] - ex["bbox_top"]) < 6:
                        part["existing"] = ex
                        break

    # write items in reading order per page (tables positioned by their first part)
    placed = set()
    images = []
    for pno in pnos:
        it = items_by_page[pno]
        put(ws, r, 1, f"p{pno}", bold=True)
        put(ws, r, 2, "Slide section", bold=True)
        put(ws, r, FIRST_CONTENT_COL, f"Physical PDF page {pno} — {it['mode']}", bold=True)
        for c in range(1, FIRST_CONTENT_COL + 1):
            ws.cell(r, c).fill = P.HDR_FILL
        r += 1
        entries = [("text", x["bbox"][1], x) for x in it["text"]]
        for lt in merged:
            if lt["pages"][0] == pno and id(lt) not in placed:
                entries.append(("table", lt["parts"][0]["bbox"][1], lt))
        for f in it["figs"]:
            entries.append(("figure", f["bbox"][1], f))
        entries.sort(key=lambda e: e[1])
        if not entries:
            put(ws, r, 1, f"p{pno}"); put(ws, r, 2, "No content")
            put(ws, r, FIRST_CONTENT_COL, "(no text, table or figure detected — see slide image below)")
            if not (26 <= pno <= 35):
                review.append((name, pno, "Possible missing content", "Nothing extracted from this slide; check slide image."))
            r += 2
            continue
        for kind, _, obj in entries:
            if kind == "text":
                conf = [c for c in obj["conf"] if c >= 0]
                mean = sum(conf) / len(conf) if conf else 100
                ocr = any(s.startswith("ocr") for s in obj["src"])
                typ = obj["type"] + (" (title)" if obj.get("title_hit") else "") + (" (OCR)" if ocr else "")
                if ocr and mean < P.LOW_CONF:
                    typ += " — review"
                    review.append((name, pno, "Uncertain OCR text", f"Row {r}: mean OCR confidence {mean:.0f}."))
                put(ws, r, 1, f"p{pno}"); put(ws, r, 2, typ)
                for k, chunk in enumerate(P.chunks(obj["text"])):
                    put(ws, r, FIRST_CONTENT_COL, chunk, wrap=True)
                    ws.merge_cells(start_row=r, start_column=FIRST_CONTENT_COL, end_row=r, end_column=FIRST_CONTENT_COL + 7)
                    lines_n = sum(max(1, math.ceil(len(x) / 110)) for x in chunk.split("\n"))
                    ws.row_dimensions[r].height = min(15 * lines_n + 2, 400)
                    r += 1
            elif kind == "table":
                lt = obj
                placed.add(id(lt))
                n = log["table_counter"][name] = log["table_counter"].get(name, 0) + 1
                put(ws, r, 1, "/".join(f"p{x}" for x in lt["pages"]), bold=True)
                put(ws, r, 2, f"Table {n}", bold=True)
                put(ws, r, FIRST_CONTENT_COL, f"Table {n} — source page(s) {', '.join(map(str, lt['pages']))}"
                    + (" — continued across slides, merged (same columns and header)" if len(lt["pages"]) > 1 else "")
                    + f" — {lt['parts'][0]['method']}", bold=True)
                r += 1
                start = r
                header_row = r
                for k, part in enumerate(lt["parts"]):
                    rows = part["cells"] if (k == 0 or lt.get("combined")) else part["cells"][1:]
                    ex = part.get("existing")
                    if k:
                        put(ws, r, 1, f"p{part['page']}"); put(ws, r, 2, "(continued)")
                    for ri, row in enumerate(rows):
                        src_ri = ri + (0 if (k == 0 or lt.get("combined")) else 1)
                        put(ws, r, 1, f"p{part['page']}")
                        for ci, val in enumerate(row):
                            v = val
                            mark = None
                            if ex:
                                key = (ex["data_row0"] + src_ri, ci + 1)
                                if key in ex["edits"]:
                                    v = ex["edits"][key]
                                    mark = "edited"
                            if v not in (None, ""):
                                put(ws, r, FIRST_CONTENT_COL + ci, v, wrap=True)
                            cf = part["confs"][src_ri][ci] if src_ri < len(part["confs"]) and ci < len(part["confs"][src_ri]) else []
                            if mark is None and cf and min(cf) < P.LOW_CONF and v:
                                review.append((name, part["page"], "Uncertain OCR table cell",
                                               f"{get_column_letter(FIRST_CONTENT_COL + ci)}{r}: lowest word confidence {min(cf):.0f}."))
                        r += 1
                    if ex:
                        review.append((name, part["page"], "Existing edited table preserved",
                                       f"Values manually edited in '{ex['sheet']}' ({len(ex['edits'])} cell(s)) were carried into this table "
                                       f"unchanged; '{ex['sheet']}' itself is untouched."))
                        put(ws, r, 2, "Existing sheet")
                        P.link(ws, r, FIRST_CONTENT_COL, f"Manually edited version preserved in {ex['sheet']}", f"#'{ex['sheet']}'!A1")
                        log["linked"].append((name, ex["sheet"]))
                        r += 1
                end = r - 1
                maxcol = max(maxcol, FIRST_CONTENT_COL + lt["ncol"] - 1)
                rng = [(start, end)]
                _style(ws, start, end, lt["ncol"], {header_row})
                if any("verify" in p["method"] for p in lt["parts"]):
                    review.append((name, lt["pages"][0], "Uncertain table structure",
                                   f"Table {n} rebuilt from word layout (no ruling lines); verify against the slide image."))
                log["tables"].append((name, n, lt["pages"], lt["nrow"] if "nrow" in lt else sum(len(p["cells"]) for p in lt["parts"]), lt["ncol"]))
                r += 1
            else:
                f = obj
                put(ws, r, 1, f"p{pno}"); put(ws, r, 2, "Figure")
                fid = f"p{pno}-fig{len([i for i in images if i[0] == pno]) + 1}"
                put(ws, r, FIRST_CONTENT_COL, f"Figure {fid} — embedded below (see 'Figures on these slides')")
                images.append((pno, fid, f))
                r += 1
        r += 1
    # images: figures, then full slide references — all below the last content row (no overlap)
    r += 1
    page_obj = {pno: doc[pno - 1] for pno in pnos}
    if images:
        put(ws, r, 1, "Figures on these slides", bold=True)
        r += 1
        for pno, fid, f in images:
            crop = P.render_crop(page_obj[pno], f["bbox"], W / f"{name.replace(' ', '_')}_{fid}.png")
            put(ws, r, 1, f"p{pno}"); put(ws, r, 2, f"Figure {fid}")
            used, _, _ = P.add_image(ws, crop, FIRST_CONTENT_COL, r + 1, 900)
            r += used + 1
            log["figures"].append((name, fid, pno))
    put(ws, r, 1, "Slide images (verification)", bold=True)
    r += 1
    for pno in pnos:
        jp = W / f"slide_{pno:02d}.jpg"
        pix = doc[pno - 1].get_pixmap(dpi=120)
        Image.frombytes("RGB", (pix.width, pix.height), pix.samples).save(jp, quality=80)
        put(ws, r, 1, f"p{pno}", bold=True); put(ws, r, 2, f"Slide (physical page {pno})")
        used, _, _ = P.add_image(ws, jp, FIRST_CONTENT_COL, r + 1, 1000)
        r += used + 1
    for c in range(FIRST_CONTENT_COL, maxcol + 1):
        if not ws.column_dimensions[get_column_letter(c)].width or ws.column_dimensions[get_column_letter(c)].width < 14:
            ws.column_dimensions[get_column_letter(c)].width = 16
    ws.freeze_panes = "C8"
    return ws


def combine_group(items_by_page, pnos, name, review):
    """One logical table across the slides: the ruled table supplies column bounds; table regions on the
    other slides are re-split into the same columns by word x-position. Repeated header rows are dropped."""
    grids = [t for p in pnos for t in items_by_page[p]["tables"] if t.get("colx") and all(t["colx"])]
    if not grids:
        review.append((name, pnos[0], "Uncertain table structure",
                       "No ruled table found to define columns; slide tables kept separate — verify continuation manually."))
        return
    tpl = max(grids, key=lambda t: t["ncol"])
    bounds = tpl["colx"]
    header = tpl["cells"][0]
    parts = []
    for p in pnos:
        for t in items_by_page[p]["tables"]:
            if t is tpl or t.get("colx") and len(t["colx"]) == len(bounds) and all(t["colx"]):
                rows, confs = t["cells"], t["confs"]
            else:
                rows, confs = [], []
                for ln in items_by_page[p]["lines"]:
                    if P.frac_inside(ln["bbox"], t["bbox"]) <= 0.6:
                        continue
                    row = [None] * len(bounds)
                    cf = [[] for _ in bounds]
                    for w in ln["w"]:
                        cx = (w[0] + w[2]) / 2
                        k = min(range(len(bounds)), key=lambda i: 0 if bounds[i][0] <= cx <= bounds[i][1]
                                else min(abs(cx - bounds[i][0]), abs(cx - bounds[i][1])))
                        row[k] = (row[k] + " " + w[4]) if row[k] else w[4]
                        cf[k].append(w[5])
                    rows.append(row)
                    confs.append(cf)
            parts.append({"page": p, "cells": rows, "confs": confs, "bbox": t["bbox"], "ncol": len(bounds),
                          "nrow": len(rows), "method": "combined into one logical table using the ruled table's "
                          "column positions (p%d) — verify" % tpl["page"]})
    # drop repeated header rows (keep first occurrence)
    seen_header = False
    for part in parts:
        keep_r, keep_c = [], []
        for row, cf in zip(part["cells"], part["confs"]):
            if header_sim(row, header) >= 0.6:
                if seen_header:
                    continue
                seen_header = True
            keep_r.append(row)
            keep_c.append(cf)
        part["cells"], part["confs"] = keep_r, keep_c
    parts = [x for x in parts if x["cells"]]
    first = parts[0]
    for p in pnos:
        items_by_page[p]["tables"] = []
    items_by_page[first["page"]]["tables"] = [first]
    first["combined_parts"] = parts
    review.append((name, pnos[0], "Uncertain table structure",
                   f"'Summary for Action Item' combined from {len(parts)} table section(s) on pages "
                   f"{', '.join(str(x['page']) for x in parts)} using column bounds of the ruled table on p{tpl['page']}; "
                   "rows from unruled slides were split by word position — verify against the slide images."))


def _style(ws, first, last, ncol, header_rows):
    from openpyxl.styles import Border
    for rr in range(first, last + 1):
        for c in range(FIRST_CONTENT_COL, FIRST_CONTENT_COL + ncol):
            cell = ws.cell(rr, c)
            cell.border = P.GRID
            if rr in header_rows:
                cell.font = Font(bold=True)
                cell.fill = P.HDR_FILL
                cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
            else:
                v = cell.value
                right = isinstance(v, str) and bool(P.NUMLIKE.match(v))
                cell.alignment = Alignment(wrap_text=True, vertical="top", horizontal="right" if right else "left")
    for c in range(FIRST_CONTENT_COL, FIRST_CONTENT_COL + ncol):
        L = get_column_letter(c)
        longest = max((max(len(x) for x in str(ws.cell(rr, c).value).split("\n")) for rr in range(first, last + 1)
                       if ws.cell(rr, c).value), default=0)
        ws.column_dimensions[L].width = max(ws.column_dimensions[L].width or 0, min(max(longest + 2, 10), 45))
    for rr in range(first, last + 1):
        n = 1
        for c in range(FIRST_CONTENT_COL, FIRST_CONTENT_COL + ncol):
            v = ws.cell(rr, c).value
            if v:
                w = ws.column_dimensions[get_column_letter(c)].width or 16
                n = max(n, sum(max(1, math.ceil(len(x) / max(1, w - 1))) for x in str(v).split("\n")))
        ws.row_dimensions[rr].height = min(15 * n + 2, 300)


# --------------------------------------------------------------------------- snapshot / validation
def snapshot(ws):
    cells = {}
    for row in ws.iter_rows():
        for c in row:
            if c.value is not None or c.has_style:
                cells[(c.row, c.column)] = (c.value, c.data_type, c.number_format, repr(c.font), repr(c.fill),
                                            repr(c.border), repr(c.alignment))
    return {"cells": cells, "merged": sorted(str(m) for m in ws.merged_cells.ranges),
            "cols": {k: v.width for k, v in ws.column_dimensions.items()},
            "rows": {k: v.height for k, v in ws.row_dimensions.items() if v.height},
            "images": sorted((i.anchor._from.col, i.anchor._from.row, i.width, i.height) for i in ws._images),
            "freeze": ws.freeze_panes, "links": sorted((c.coordinate, c.hyperlink.location or c.hyperlink.target)
                                                      for row in ws.iter_rows() for c in row if c.hyperlink)}


def main():
    W.mkdir(parents=True, exist_ok=True)
    assert not NEW_OUT.exists(), f"refusing to overwrite {NEW_OUT}"
    cur_hash = P.sha(CUR)
    pdf_bytes = P.read_bytes(NEW_PDF_SRC)
    pdf_hash = P.hashlib.sha256(pdf_bytes).hexdigest()
    new_copy = W / "_new_working_copy.pdf"
    new_copy.write_bytes(pdf_bytes)
    doc = pymupdf.open(new_copy)
    assert len(doc) == 62, len(doc)
    _, analysed, _ = P.analyse_pdf(new_copy, W)

    # manual edits in the current workbook (vs unedited regeneration from the old PDF)
    cur_wb = load_workbook(CUR)
    gen_wb = load_workbook(REGEN_OLD)
    old_doc = pymupdf.open(OLD_COPY)
    mp = old_to_new_map(old_doc, doc)
    old_pages = P.analyse_pdf(OLD_COPY, P.WORK / "_regen_old")[1]
    P.logical_tables(old_pages)                       # same Table_### ids as the original run
    existing = []
    for s in cur_wb.sheetnames:
        if not s.startswith("Table_") or s == REMOVE:
            continue
        ed = edited_cells(cur_wb[s], gen_wb[s])
        merges_changed = sorted(map(str, cur_wb[s].merged_cells.ranges)) != sorted(map(str, gen_wb[s].merged_cells.ranges))
        src_pages = [int(x) for x in str(cur_wb[s]["B3"].value).split(",") if x.strip().isdigit()]
        for op in src_pages:
            npg, score = mp[op]
            ot = [t for t in old_pages[op - 1]["tables"] if t.get("id") == s]
            for t in ot:
                existing.append({"sheet": s, "old_page": op, "new_page": npg, "match": score, "edits": ed,
                                 "merges_changed": merges_changed, "nrow": t["nrow"], "ncol": t["ncol"],
                                 "bbox_top": t["bbox"][1], "data_row0": 8})
    # data rows in the original table sheets start at row 8 (row 7 = 'Source page' label)
    existing = [e for e in existing if e["edits"] or e["merges_changed"]]
    seen = set()
    existing = [e for e in existing if not ((e["sheet"], e["new_page"]) in seen or seen.add((e["sheet"], e["new_page"])))]

    before = {s: snapshot(cur_wb[s]) for s in cur_wb.sheetnames if s != REMOVE}
    shutil.copy2(CUR, W / "_input_copy.xlsx")
    wb = load_workbook(W / "_input_copy.xlsx")
    del wb[REMOVE]
    review, log = [], {"titles": [], "tables": [], "figures": [], "linked": [], "table_counter": {}}
    for name, pnos, title, rule in GROUPS:
        write_group(wb, name, pnos, title, rule, doc, analysed, existing, review, log)

    # update log + review sheets (new)
    lg = wb.create_sheet(LOG_SHEET)
    P.head(lg, ["Item", "Detail"], [36, 110])
    rows = [("Input workbook (copied, unchanged)", str(CUR)), ("Output workbook (new version)", str(NEW_OUT)),
            ("Source PDF", str(NEW_PDF_SRC)), ("Source PDF version", f"62 physical pages; sha256 {pdf_hash[:16]}…; replaced on NAS 2026-10-16 17:12"),
            ("Page numbering", "Physical PDF page positions from 1; blank pages 26–35 counted and kept in position"),
            ("Worksheet removed", REMOVE), ("Existing worksheets otherwise unchanged", "verified cell-by-cell after save (see validation)"),
            ("Note", "Existing Table_### sheets were extracted from the previous 63-page PDF; their 'Source PDF pages' refer to that version.")]
    for i, (k, v) in enumerate(rows, 2):
        put(lg, i, 1, k, bold=True); put(lg, i, 2, v, wrap=True)
    r = len(rows) + 3
    put(lg, r, 1, "Worksheets added", bold=True)
    for name, pnos, title, _ in GROUPS:
        r += 1
        P.link(lg, r, 1, name, f"#'{name}'!A1")
        put(lg, r, 2, f"pages {', '.join(map(str, pnos))}" + (f" — requested: {title}" if title else " — all content"))
    r += 2
    put(lg, r, 1, "Requested titles (search in slide text/OCR)", bold=True)
    for t in log["titles"]:
        r += 1
        put(lg, r, 1, f"{t[0]} p{t[2]}"); put(lg, r, 2, f"'{t[1]}': {t[3]} {t[4]}")
    r += 2
    put(lg, r, 1, "Existing edited tables linked (preserved)", bold=True)
    for a, b in sorted(set(log["linked"])):
        r += 1
        put(lg, r, 1, a); P.link(lg, r, 2, b, f"#'{b}'!A1")
    rv = wb.create_sheet(REVIEW_SHEET)
    P.head(rv, ["Worksheet", "Source page", "Issue", "Detail"], [24, 11, 32, 110])
    for i, x in enumerate(review, 2):
        for c, v in enumerate(x, 1):
            put(rv, i, c, v, wrap=c == 4)
    wb.save(NEW_OUT)

    # ---- validation (reopen)
    out = load_workbook(NEW_OUT)
    problems = []
    if REMOVE in out.sheetnames:
        problems.append("Table_002 still present")
    for s, snap in before.items():
        if s not in out.sheetnames:
            problems.append(f"missing existing sheet {s}")
            continue
        after = snapshot(out[s])
        for k in ("cells", "merged", "cols", "rows", "images", "freeze", "links"):
            if after[k] != snap[k]:
                problems.append(f"{s}: {k} changed")
    exp_order = [s for s in cur_wb.sheetnames if s != REMOVE]
    if out.sheetnames[:len(exp_order)] != exp_order:
        problems.append("existing sheet order changed")
    for name, pnos, _, _ in GROUPS:
        ws = out[name]
        pages_present = {str(v).strip() for (v,) in ws.iter_rows(min_row=8, max_col=1, values_only=True) if v}
        for p in pnos:
            if f"p{p}" not in pages_present:
                problems.append(f"{name}: page {p} not represented")
        n_slide_imgs = sum(1 for (v,) in ws.iter_rows(min_row=8, min_col=2, max_col=2, values_only=True)
                           if isinstance(v, str) and v.startswith("Slide (physical page"))
        if n_slide_imgs != len(pnos):
            problems.append(f"{name}: slide images {n_slide_imgs}/{len(pnos)}")
        for row in ws.iter_rows():
            for c in row:
                if c.data_type == "f":
                    problems.append(f"{name}: formula cell {c.coordinate}")
        last_text_row = max((c.row for row in ws.iter_rows() for c in row if c.value is not None
                             and c.column >= FIRST_CONTENT_COL and not str(c.value).startswith(("Figure p", "(no text"))), default=0)
        for img in ws._images:
            pass
    assert P.sha(CUR) == cur_hash, "current workbook changed"
    assert P.hashlib.sha256(P.read_bytes(NEW_PDF_SRC)).hexdigest() == pdf_hash, "source PDF changed during run"
    summary = {"problems": problems, "added": [g[0] for g in GROUPS] + [LOG_SHEET, REVIEW_SHEET], "removed": [REMOVE],
               "tables": len(log["tables"]), "figures": len(log["figures"]), "review": len(review),
               "review_by_issue": dict(Counter(x[2] for x in review)), "linked": sorted(set(log["linked"])),
               "titles": [(t[0], t[2], t[3]) for t in log["titles"]],
               "tables_per_sheet": dict(Counter(t[0] for t in log["tables"]))}
    (P.LOGS / f"{STEM}_v2_update.json").write_text(json.dumps(summary, indent=2, default=str))
    print(json.dumps({k: v for k, v in summary.items()}, indent=1, default=str))


if __name__ == "__main__":
    main()
