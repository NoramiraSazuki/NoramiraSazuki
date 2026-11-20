#!/usr/bin/env python3
"""v4 of 'ARPR 1.1.2025 YEAR Presentation Pack_Dulang - for AOF' — copy/paste-friendly tables.

Base   : v3 workbook (copied; v3 itself untouched). All original sheets kept exactly; Index navigation updated
         (links to added sheets; dead 'Table_002' link marked as removed). The 7 extraction sheets and the v3
         log/review sheets are rebuilt.
Tables : same cell geometry and cell text as v3 (vector cells / picture grid / word layout, native text or per-cell
         OCR), but written as a PROPER table — one Excel column per table column and one Excel row per table row,
         no merged cells. A value spanning several slide cells sits in the first (top-left) cell; the spanned cells
         keep the same fill and are listed in the review sheet. Fill colours, borders and the header row are kept.
Edits  : manual edits in Table_010/012/013 carried over as in v3.
Console output: counts/status only.
"""
import json
import math
import shutil
import sys
from collections import Counter
from pathlib import Path

import pymupdf
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from PIL import Image

sys.path.insert(0, str(Path(__file__).parent))
import pdf_to_excel as P  # noqa: E402
import update_presentation_pack_v2 as U  # noqa: E402
import update_presentation_pack_v3 as V  # noqa: E402

STEM = V.STEM
BASE_WB = P.OUT / f"{STEM}_v3.xlsx"
OUT_WB = P.OUT / f"{STEM}_v4.xlsx"
W = P.WORK / "_update_v4"
V.W = W                                   # crops/OCR working files for this run
GROUP_SHEETS = [g[0] for g in V.GROUPS]
REBUILD = GROUP_SHEETS + ["Update Log v3", "Update Review v3"]
LOG_SHEET, REVIEW_SHEET = "Update Log v4", "Update Review v4"
V.LOG_SHEET = LOG_SHEET
COL0 = 3                                  # A = source page, B = item type, C.. = table columns


def compress(part):
    """Keep only grid columns/rows where at least one cell starts -> proper table indices."""
    cols = sorted({c["c0"] for c in part["cells"]})
    rows = sorted({c["r0"] for c in part["cells"]})
    cmap = {c: i for i, c in enumerate(cols)}
    rmap = {r: i for i, r in enumerate(rows)}

    def last_start(starts, end):
        cand = [s for s in starts if s <= end]
        return cand[-1] if cand else starts[0]
    cells = []
    for c in part["cells"]:
        nc0, nr0 = cmap[c["c0"]], rmap[c["r0"]]
        nc1, nr1 = cmap[last_start(cols, c["c1"])], rmap[last_start(rows, c["r1"])]
        cells.append({**c, "C0": nc0, "C1": max(nc0, nc1), "R0": nr0, "R1": max(nr0, nr1)})
    xs = part["xs"]
    widths = [xs[min(len(xs) - 1, (cols[i + 1] if i + 1 < len(cols) else len(xs) - 1))] - xs[cols[i]] for i in range(len(cols))]
    return cells, len(cols), len(rows), widths


def write_table(ws, r, lt, tno, edits, review, name):
    parts = lt["parts"]
    base_cells, ncol, _, widths = compress(parts[0])
    header_written = False
    fills_used = 0
    for k, part in enumerate(parts):
        cells, nc, nr, _ = compress(part) if k else (base_cells, ncol, None, widths)
        if nc != ncol:
            review.append((name, part["page"], "Continuation layout differs",
                           f"Table {tno}: continued part on p{part['page']} has {nc} columns vs {ncol}; columns aligned left."))
        nr = max(c["R1"] for c in cells) + 1 if cells else 0
        for gr in range(nr):
            P.put(ws, r + gr, 1, f"p{part['page']}")
            P.put(ws, r + gr, 2, "Header" if (k == 0 and gr == 0) else ("(continued)" if gr == 0 and k else ""))
        for c in cells:
            text = c.get("text", "") or ""
            for e in edits:
                if e.get("bbox") and e["page"] == part["page"] and c["bbox"][0] - 1 <= (e["bbox"][0] + e["bbox"][2]) / 2 <= c["bbox"][2] + 1 \
                        and c["bbox"][1] - 1 <= (e["bbox"][1] + e["bbox"][3]) / 2 <= c["bbox"][3] + 1:
                    text = e["value"] if e["value"] is not None else ""
                    e["used"] = True
                    review.append((name, part["page"], "Manual edit carried over",
                                   f"{get_column_letter(COL0 + c['C0'])}{r + c['R0']}: value from {e['sheet']}!"
                                   f"{get_column_letter(e['col'])}{e['row']} used instead of extraction."))
            rr, cc = r + c["R0"], COL0 + c["C0"]
            if text:
                P.put(ws, rr, cc, text)
            fill = c.get("fill")
            dark = fill and V.luminance(fill) < 0.45
            for y in range(c["R0"], c["R1"] + 1):
                for x in range(c["C0"], c["C1"] + 1):
                    cell = ws.cell(r + y, COL0 + x)
                    cell.border = V.BOX
                    if fill:
                        cell.fill = PatternFill("solid", fgColor=fill)
                        fills_used += 1
                    cell.alignment = Alignment(wrap_text=True, vertical="top",
                                               horizontal="right" if isinstance(text, str) and P.NUMLIKE.match(text or "x") else "left")
                    cell.font = Font(size=10, bold=(k == 0 and c["R0"] == 0), color="FFFFFF" if dark else "000000")
            if c["C1"] > c["C0"] or c["R1"] > c["R0"]:
                review.append((name, part["page"], "Spanning cell (no merge)",
                               f"{get_column_letter(cc)}{rr}: on the slide this cell spans "
                               f"{c['C1'] - c['C0'] + 1} column(s) × {c['R1'] - c['R0'] + 1} row(s); value kept in the first cell."))
            if c.get("src") == "ocr" and text and c.get("conf", 100) < P.LOW_CONF:
                review.append((name, part["page"], "Uncertain OCR table cell",
                               f"{get_column_letter(cc)}{rr}: OCR confidence {c['conf']:.0f}."))
        # also mark header row fill if none
        if k == 0:
            for x in range(ncol):
                h = ws.cell(r, COL0 + x)
                if h.fill is None or h.fill.fgColor.rgb in (None, "00000000"):
                    h.fill = P.HDR_FILL
            header_written = True
        r += nr
    return r, ncol


def fit(ws, first_row, last_row, ncol):
    for x in range(ncol):
        L = get_column_letter(COL0 + x)
        longest = 0
        for rr in range(first_row, last_row + 1):
            v = ws.cell(rr, COL0 + x).value
            if v:
                longest = max(longest, max(len(s) for s in str(v).split("\n")))
        ws.column_dimensions[L].width = max(ws.column_dimensions[L].width or 0, min(max(longest + 2, 10), 48))
    for rr in range(first_row, last_row + 1):
        n = 1
        for x in range(ncol):
            v = ws.cell(rr, COL0 + x).value
            if v:
                w = ws.column_dimensions[get_column_letter(COL0 + x)].width or 12
                n = max(n, sum(max(1, math.ceil(len(s) / max(1, w - 1))) for s in str(v).split("\n")))
        ws.row_dimensions[rr].height = min(15 * n + 2, 300)


def write_group(wb, name, pnos, title, rule, doc, analysed, edits, review, log):
    ws = wb.create_sheet(name)
    ws.column_dimensions["A"].width, ws.column_dimensions["B"].width = 9, 16
    infos = {p: V.analyse_page(doc, p, analysed, review, name) for p in pnos}
    logical = V.merge_continuations(infos, pnos, rule, name, review)
    hdr = [("Worksheet", name), ("Source physical PDF pages", ", ".join(map(str, pnos))),
           ("Printed slide numbers", "; ".join(f"p{p}: {infos[p]['printed'] or 'none found'}" for p in pnos)),
           ("Requested", title or "all content"), ("Method", "; ".join(f"p{p}: {infos[p]['mode']}" for p in pnos))]
    for i, (k, v) in enumerate(hdr, 1):
        P.put(ws, i, 1, k, bold=True)
        P.put(ws, i, COL0, v)
        ws.cell(i, 1).fill = P.LABEL_FILL
    P.link(ws, 6, 1, "◀ Index", "#'Index'!A1")
    P.link(ws, 6, 2, "Update Log", f"#'{LOG_SHEET}'!A1")
    for p in pnos:
        if infos[p]["printed"] and infos[p]["printed"] != str(p):
            review.append((name, p, "Slide number", f"Printed slide number {infos[p]['printed']} differs from physical page {p}."))
    if title:
        for p in pnos:
            hit = U.find_title(infos[p]["lines"] or [], title)
            hit_ocr = any(U.norm(title) in U.norm(t["text"]) for t in infos[p]["text"])
            log["titles"].append((name, title, p, "found" if (hit or hit_ocr) else "not found"))
    r = 8
    tno, figures, maxcol = 0, [], 3
    for p in pnos:
        info = infos[p]
        P.put(ws, r, 1, f"p{p}", bold=True); P.put(ws, r, 2, "Slide section", bold=True)
        P.put(ws, r, COL0, f"Physical PDF page {p}", bold=True)
        for c in range(1, COL0 + 1):
            ws.cell(r, c).fill = P.HDR_FILL
        r += 2
        entries = [("text", t["bbox"][1], t) for t in info["text"]]
        entries += [("table", lt["parts"][0]["bbox"][1], lt) for lt in logical if lt["pages"][0] == p]
        entries += [("figure", f["bbox"][1], f) for f in info["figs"]]
        entries.sort(key=lambda e: e[1])
        if not entries:
            P.put(ws, r, 1, f"p{p}"); P.put(ws, r, 2, "No content"); P.put(ws, r, COL0, "(nothing detected — see slide image)")
            review.append((name, p, "Possible missing content", "Nothing extracted; check slide image."))
            r += 2
        for kind, _, obj in entries:
            if kind == "text":
                low = obj["src"] == "ocr" and obj["conf"] < P.LOW_CONF
                P.put(ws, r, 1, f"p{p}")
                P.put(ws, r, 2, "Text" + (" (OCR)" if obj["src"] == "ocr" else "") + (" — review" if low else ""))
                P.put(ws, r, COL0, obj["text"])
                ws.cell(r, COL0).alignment = Alignment(wrap_text=False, vertical="top")
                if low:
                    review.append((name, p, "Uncertain OCR text", f"Row {r}: mean OCR confidence {obj['conf']:.0f}."))
                r += 1
            elif kind == "table":
                tno += 1
                lt = obj
                r += 1
                P.put(ws, r, 1, "/".join(f"p{x}" for x in lt["pages"]), bold=True)
                P.put(ws, r, 2, f"Table {tno}", bold=True)
                P.put(ws, r, COL0, f"Table {tno} — source page(s) {', '.join(map(str, lt['pages']))}"
                      + (" — continued across slides, combined (repeated header rows removed)" if len(lt["pages"]) > 1 else ""),
                      bold=True)
                P.put(ws, r + 1, COL0, f"Method: {lt['parts'][0]['method']}")
                ws.cell(r + 1, COL0).font = Font(italic=True, size=9, color="595959")
                r += 2
                start = r
                r, ncol = write_table(ws, r, lt, tno, edits, review, name)
                fit(ws, start, r - 1, ncol)
                maxcol = max(maxcol, COL0 + ncol - 1)
                log["tables"].append((name, tno, lt["pages"], ncol, r - start, f"{name}!{get_column_letter(COL0)}{start}"))
                if "verify" in lt["parts"][0]["method"] or "pixels" in lt["parts"][0]["method"]:
                    review.append((name, lt["pages"][0], "Verify table", f"Table {tno}: {lt['parts'][0]['method']}; compare with slide image."))
                r += 1
            else:
                figures.append((p, obj))
                P.put(ws, r, 1, f"p{p}"); P.put(ws, r, 2, "Figure")
                P.put(ws, r, COL0, f"Figure p{p}-{sum(1 for x in figures if x[0] == p)} — embedded below")
                r += 1
        r += 1
    r += 1
    if figures:
        P.put(ws, r, 1, "Figures on these slides", bold=True)
        r += 1
        for k, (p, f) in enumerate(figures, 1):
            crop = P.render_crop(doc[p - 1], f["bbox"], W / f"{name.replace(' ', '_')}_fig{k}.png")
            P.put(ws, r, 1, f"p{p}"); P.put(ws, r, 2, f"Figure p{p}-{k}")
            used, _, _ = P.add_image(ws, crop, COL0, r + 1, 900)
            r += used + 1
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


def update_index(ix, log):
    """Index navigation: mark the removed sheet; add links to the added sheets and their tables."""
    for row in ix.iter_rows():
        for c in row:
            if c.hyperlink and "Table_002" in str(c.hyperlink.location or c.hyperlink.target):
                c.hyperlink = None
                P.put(ix, c.row, c.column, "Table_002 (removed in v2)")
                c.font = Font(color="808080", italic=True)
                P.put(ix, c.row, 2, "removed at user request; see Update Log")
    r = ix.max_row + 2
    P.put(ix, r, 1, "Added worksheets (slide extractions)", bold=True)
    ix.cell(r, 1).fill = P.HDR_FILL
    for name, pnos, title, _ in V.GROUPS:
        r += 1
        P.link(ix, r, 1, name, f"#'{name}'!A1")
        P.put(ix, r, 2, f"physical pages {', '.join(map(str, pnos))}" + (f" — {title}" if title else " — all content"))
        for t in [t for t in log["tables"] if t[0] == name]:
            r += 1
            P.link(ix, r, 1, f"    {name} — Table {t[1]}", f"#'{name}'!{t[5].split('!')[1]}")
            P.put(ix, r, 2, f"pages {', '.join(map(str, t[2]))}; {t[4]} rows × {t[3]} columns")
    for s in (LOG_SHEET, REVIEW_SHEET):
        r += 1
        P.link(ix, r, 1, s, f"#'{s}'!A1")


def main():
    W.mkdir(parents=True, exist_ok=True)
    assert not OUT_WB.exists(), f"refusing to overwrite {OUT_WB}"
    base_hash = P.sha(BASE_WB)
    pdf = P.read_bytes(V.SRC_PDF)
    pdf_hash = V.sha_bytes(pdf)
    assert pdf_hash == V.sha_bytes(V.ORIG_63.read_bytes()), "NAS PDF is not the 63-page version"
    local = W / "_source_63.pdf"
    local.write_bytes(pdf)
    doc = pymupdf.open(local)
    _, analysed, _ = P.analyse_pdf(local, W)
    base = load_workbook(BASE_WB)
    gen = load_workbook(V.REGEN_OLD)
    wanted = sorted({p for _, ps, _, _ in V.GROUPS for p in ps})
    edits = V.carried_edits(doc, base, gen, wanted)

    shutil.copy2(BASE_WB, W / "_base_copy.xlsx")
    wb = load_workbook(W / "_base_copy.xlsx")
    keep = [s for s in wb.sheetnames if s not in REBUILD and s != "Index"]
    before = {s: V.snapshot(wb[s]) for s in keep}
    index_before = V.snapshot(wb["Index"])
    for s in REBUILD:
        if s in wb.sheetnames:
            del wb[s]
    review, log = [], {"titles": [], "tables": [], "figures": []}
    for name, pnos, title, rule in V.GROUPS:
        write_group(wb, name, pnos, title, rule, doc, analysed, edits, review, log)
    for e in edits:
        if e.get("merges_changed"):
            review.append(("(existing)", e["page"], "Manual merge edits",
                           f"{e['sheet']} has manually merged cells; the new table is written without merges — compare."))
        elif not e.get("used"):
            review.append(("(existing)", e["page"], "Manual edit not placed",
                           f"{e['sheet']}!{get_column_letter(e['col'])}{e['row']} could not be matched; {e['sheet']} preserved."))

    lg = wb.create_sheet(LOG_SHEET)
    P.head(lg, ["Item", "Detail"], [40, 110])
    rows = [("Base workbook (copied, unchanged)", str(BASE_WB)), ("Output workbook", str(OUT_WB)),
            ("Source PDF", f"{V.SRC_PDF} — 63 physical pages; sha256 {pdf_hash[:16]}…"),
            ("Change in v4", "Tables rewritten as proper tables: one Excel column per table column, one row per table "
                             "row, no merged cells (spanning values in the first cell; listed in Update Review v4)."),
            ("Index", "Navigation updated: links to added sheets and their tables; Table_002 link marked as removed"),
            ("Kept unchanged", "all other original sheets (verified cell-by-cell); Table_002 remains removed"),
            ("Rebuilt", ", ".join(REBUILD))]
    for i, (k, v) in enumerate(rows, 2):
        P.put(lg, i, 1, k, bold=True); P.put(lg, i, 2, v, wrap=True)
    r = len(rows) + 3
    P.link(lg, r, 1, "◀ Index", "#'Index'!A1")
    r += 1
    P.put(lg, r, 1, "Tables", bold=True)
    for t in log["tables"]:
        r += 1
        P.link(lg, r, 1, f"{t[0]} — Table {t[1]}", f"#'{t[0]}'!{t[5].split('!')[1]}")
        P.put(lg, r, 2, f"pages {', '.join(map(str, t[2]))}; {t[4]} rows × {t[3]} columns")
    r += 2
    P.put(lg, r, 1, "Requested titles", bold=True)
    for t in log["titles"]:
        r += 1
        P.put(lg, r, 1, f"{t[0]} p{t[2]}"); P.put(lg, r, 2, f"'{t[1]}': {t[3]}")
    rv = wb.create_sheet(REVIEW_SHEET)
    P.head(rv, ["Worksheet", "Source page", "Issue", "Detail"], [24, 11, 30, 110])
    P.link(rv, 1, 6, "◀ Index", "#'Index'!A1")
    for i, x in enumerate(review, 2):
        for c, v in enumerate(x, 1):
            P.put(rv, i, c, v, wrap=c == 4)
    update_index(wb["Index"], log)
    wb.save(OUT_WB)

    # ---- validation (reopen)
    out = load_workbook(OUT_WB)
    problems = []
    for s, snap in before.items():
        a = V.snapshot(out[s])
        for k in snap:
            if a[k] != snap[k]:
                problems.append(f"{s}: {k} changed")
    ia = V.snapshot(out["Index"])
    changed = [k for k in index_before["cells"] if index_before["cells"][k] != ia["cells"].get(k)]
    unexpected = [k for k in changed if not str(ia["cells"].get(k, ("",))[0] or "").startswith(("Table_002 (removed", "removed at user"))]
    if unexpected:
        problems.append(f"Index: unexpected changes in existing cells {unexpected[:5]}")
    names = set(out.sheetnames)
    for ws in out.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if c.data_type == "f":
                    problems.append(f"{ws.title}: formula {c.coordinate}")
                if c.hyperlink and c.hyperlink.location:
                    tgt = c.hyperlink.location.split("!")[0].strip("'")
                    if tgt not in names:
                        problems.append(f"{ws.title}!{c.coordinate}: link to missing sheet {tgt}")
    for name, pnos, _, _ in V.GROUPS:
        ws = out[name]
        if ws.merged_cells.ranges:
            problems.append(f"{name}: has merged cells")
        present = {str(v) for (v,) in ws.iter_rows(min_row=8, max_col=1, values_only=True) if v}
        for p in pnos:
            if f"p{p}" not in present:
                problems.append(f"{name}: p{p} missing")
        last_val = max((c.row for row in ws.iter_rows(min_col=COL0) for c in row if c.value is not None
                        and not str(c.value).startswith(("Figure p", "Slide (physical"))), default=0)
        imgs = sorted(i.anchor._from.row + 1 for i in ws._images)
        if imgs and imgs[0] <= last_val and not str(ws.cell(imgs[0] - 1, 2).value or "").startswith(("Figure", "Slide")):
            problems.append(f"{name}: image above last content row")
    assert P.sha(BASE_WB) == base_hash, "v3 changed"
    summary = {"problems": problems, "tables": len(log["tables"]),
               "tables_per_sheet": dict(Counter(t[0] for t in log["tables"])),
               "review": dict(Counter(x[2] for x in review)),
               "edits_carried": sum(1 for e in edits if e.get("used")), "edits_found": sum(1 for e in edits if not e.get("merges_changed")),
               "index_links_added": sum(1 for row in out["Index"].iter_rows() for c in row if c.hyperlink) - 16}
    (P.LOGS / f"{STEM}_v4_update.json").write_text(json.dumps(summary, indent=2, default=str))
    print(json.dumps(summary, indent=1, default=str))


if __name__ == "__main__":
    main()
