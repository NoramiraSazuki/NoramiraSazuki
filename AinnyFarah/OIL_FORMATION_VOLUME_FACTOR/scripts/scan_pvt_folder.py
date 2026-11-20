#!/usr/bin/env python3
"""Locate (not extract) Bo / oil volume at reservoir conditions / oil volume at stock-tank conditions in every file of the
'PVT Fluid Analysis Document' folder.

Each source file is copied once to _work_pvt_scan/src (sources are read only and re-hashed at the end). Text is gathered
from worksheet cells (.xls/.xlsx), text boxes (.xls TXO records, .xlsx drawings), .docx paragraphs/tables/text boxes,
and OCR of every embedded picture (PNG/JPEG/DIB/TIFF; EMF/WMF via native text, embedded bitmaps and a vector render).
Cached values of external-workbook links are NOT read (they belong to other workbooks).

Usage: python scan_pvt_folder.py   (run with the DOCUMENT_PVT_Fluid_Analysis venv; needs /usr/local/bin/tesseract)
"""
import csv
import hashlib
import io
import json
import re
import shutil
import struct
import subprocess
import sys
import zipfile
import zlib
from pathlib import Path
from xml.etree import ElementTree as ET

import olefile
import openpyxl
import xlrd
from PIL import Image

SRC_DIR = Path("/Volumes/homes/public/DulangOSS/20260925/DOCUMENTS/PVT Fluid Analysis Document")
BASE = Path(__file__).resolve().parent.parent
WORK = BASE / "_work_pvt_scan"
TESS = "/usr/local/bin/tesseract"
sys.path.insert(0, str(Path(__file__).resolve().parent))

# ---- search patterns (label text only; values are not read out)
PATTERNS = {
    "Bo (oil formation volume factor)": re.compile(
        r"formation\s*volume\s*factor|\bF\.?\s?V\.?\s?F\b|\bB\s?o\s?[sdbf]{0,2}\b(?!\w)|\bBo\s*[\(=]|\bBob\b|\bBofb\b|\bBodb?\b|"
        r"\bB\s?t\b(?=.*(volume|factor))|vol(ume)?\.?\s*factor|shrinkage\s*factor|\bshrinkage\b", re.I),
    "Oil volume at reservoir conditions": re.compile(
        r"reservoir\s*(condition|barrel|bbl|volume|temperature\s*and\s*pressure)|res\.?\s*bbl|\brb\b|\brbbl|bbl\s*res|"
        r"volume\s*of\s*(saturated\s*)?oil\s*at\s*(indicated|saturation|bubble|reservoir|\d)|oil\s*volume\s*at\s*(reservoir|\d)|"
        r"relative\s*(oil\s*)?volume|V\s*/\s*V\s*(sat|b)|vol(ume)?\s*at\s*(saturation|bubble\s*point)|saturated\s*oil", re.I),
    "Oil volume at stock-tank conditions": re.compile(
        r"stock[\s\-]*tank|stocktank|\bS\.?\s?T\.?\s?O\b|\bSTB\b|\bSTBO\b|\bstk\.?\s*tk|tank\s*oil|"
        r"per\s*(barrel|volume|bbl)\s*of\s*(residual|stock|tank)|residual\s*oil\s*at\s*60|\bbbl/stb|/STB|separator\s*oil", re.I),
}
# words that make a hit irrelevant (e.g. 'stocktank water')
NEGATIVE = re.compile(r"stock[\s\-]*tank\s*water|stocktank\s*water", re.I)


def sha(p):
    h = hashlib.sha256()
    with open(p, "rb") as fh:
        for b in iter(lambda: fh.read(1 << 20), b""):
            h.update(b)
    return h.hexdigest()


def ocr(img_path):
    try:
        r = subprocess.run([TESS, str(img_path), "-", "--psm", "3"], capture_output=True, text=True, timeout=180)
        return r.stdout
    except subprocess.TimeoutExpired:
        return ""


def to_png(data, out_png):
    """bytes of PNG/JPEG/TIFF/BMP/DIB -> PNG (upscaled when small for OCR); returns path or None"""
    try:
        im = Image.open(io.BytesIO(data))
    except Exception:
        try:  # raw DIB (BITMAPINFOHEADER without file header)
            hsz = struct.unpack_from("<I", data, 0)[0]
            bpp = struct.unpack_from("<H", data, 14)[0]
            ncol = struct.unpack_from("<I", data, 32)[0] or (1 << bpp if bpp <= 8 else 0)
            off = 14 + hsz + 4 * ncol
            im = Image.open(io.BytesIO(b"BM" + struct.pack("<IHHI", 14 + len(data), 0, 0, off) + data))
        except Exception:
            return None
    im.load()
    if im.mode not in ("RGB", "L"):
        im = im.convert("RGB")
    if im.width < 1800:
        f = max(2, round(1800 / max(1, im.width)))
        im = im.resize((im.width * f, im.height * f), Image.LANCZOS)
    im.save(out_png)
    return out_png


# ---- EMF handling
def emf_records(b):
    i = 0
    while i + 8 <= len(b):
        t, sz = struct.unpack_from("<II", b, i)
        if sz < 8:
            break
        yield t, b[i:i + sz]
        i += sz
        if t == 14:
            break


def emf_native_text(b):
    out = []
    for t, r in emf_records(b):
        if t in (83, 84):
            base = 8 + 16 + 12
            try:
                x, y, n, off = struct.unpack_from("<iiII", r, base)
                s = r[off:off + (2 * n if t == 84 else n)].decode("utf-16-le" if t == 84 else "latin-1", "replace")
                out.append((y, x, s))
            except struct.error:
                pass
    out.sort()
    lines, cur, cy = [], [], None
    for y, x, s in out:
        if cy is not None and abs(y - cy) > 2:
            lines.append(" ".join(cur)); cur = []
        cur.append(s); cy = y
    if cur:
        lines.append(" ".join(cur))
    return "\n".join(lines)


def emf_bitmaps(b):
    res = []
    for t, r in emf_records(b):
        if t in (81, 80, 76):
            o = {81: 48, 80: 48, 76: 84}[t]
            try:
                obmi, cbmi, obits, cbits = struct.unpack_from("<IIII", r, o)
            except struct.error:
                continue
            if cbmi and cbits:
                bmi, bits = r[obmi:obmi + cbmi], r[obits:obits + cbits]
                res.append(b"BM" + struct.pack("<IHHI", 14 + len(bmi) + len(bits), 0, 0, 14 + len(bmi)) + bmi + bits)
    return res


def emf_vector_png(b, out_png):
    from PIL import ImageDraw
    shapes, cur, brush, pen, objs = [], (0, 0), (255, 255, 255), (0, 0, 0), {}
    for t, r in emf_records(b):
        d = r[8:]
        try:
            if t == 39:
                ih, style, col = struct.unpack_from("<III", d, 0)
                objs[ih] = ("b", None if style == 1 else (col & 255, col >> 8 & 255, col >> 16 & 255))
            elif t in (38, 95):
                ih = struct.unpack_from("<I", d, 0)[0]; col = struct.unpack_from("<I", d, 16 if t == 38 else 24)[0]
                objs[ih] = ("p", (col & 255, col >> 8 & 255, col >> 16 & 255))
            elif t == 37:
                ih = struct.unpack_from("<I", d, 0)[0]
                if ih & 0x80000000:
                    s = ih & 0xFF
                    brush = {0: (255,) * 3, 1: (192,) * 3, 2: (128,) * 3, 3: (64,) * 3, 4: (0,) * 3, 5: None}.get(s, brush)
                    pen = {7: (0, 0, 0), 8: None}.get(s, pen)
                elif ih in objs:
                    k, v = objs[ih]
                    if k == "b": brush = v
                    else: pen = v
            elif t in (86, 87, 85, 88, 3, 4, 2, 5, 6):
                n = struct.unpack_from("<I", d, 16)[0]
                w = 4 if t >= 85 else 8
                p = [struct.unpack_from("<hh" if w == 4 else "<ii", d, 20 + w * k) for k in range(n)]
                if t in (88, 5, 6): p = [cur] + p
                shapes.append(("poly" if t in (86, 3) else "line", p, brush, pen))
                if p: cur = p[-1]
            elif t in (91, 8, 90, 7):
                npoly = struct.unpack_from("<I", d, 16)[0]
                counts = struct.unpack_from("<%dI" % npoly, d, 24)
                w = 4 if t >= 90 else 8
                off = 24 + 4 * npoly
                for c in counts:
                    p = [struct.unpack_from("<hh" if w == 4 else "<ii", d, off + w * k) for k in range(c)]
                    off += w * c
                    shapes.append(("poly" if t in (91, 8) else "line", p, brush, pen))
            elif t == 27:
                cur = struct.unpack_from("<ii", d, 0)
            elif t == 54:
                p = struct.unpack_from("<ii", d, 0); shapes.append(("line", [cur, p], brush, pen)); cur = p
        except struct.error:
            continue
    pts = [q for s in shapes for q in s[1]]
    if len(shapes) < 5 or not pts:
        return None
    x0, x1 = min(p[0] for p in pts), max(p[0] for p in pts)
    y0, y1 = min(p[1] for p in pts), max(p[1] for p in pts)
    if x1 <= x0 or y1 <= y0:
        return None
    sc = 3000 / (x1 - x0)
    h = int((y1 - y0) * sc) + 40
    if h > 12000:
        return None
    im = Image.new("RGB", (3040, h), "white"); dr = ImageDraw.Draw(im)
    T = lambda q: ((q[0] - x0) * sc + 20, (q[1] - y0) * sc + 20)
    for kind, p, br, pn in shapes:
        if kind == "poly" and len(p) > 2:
            dr.polygon([T(q) for q in p], fill=br, outline=pn)
        elif len(p) > 1 and pn:
            dr.line([T(q) for q in p], fill=pn, width=2)
    im.save(out_png)
    return out_png


def picture_texts(data, ext, stem, imgdir):
    """-> list of (method, text) for one embedded picture"""
    out = []
    if ext in ("emf", "wmf"):
        if ext == "emf":
            nt = emf_native_text(data)
            if nt.strip():
                out.append(("EMF native text", nt))
            for k, bmp in enumerate(emf_bitmaps(data), 1):
                p = to_png(bmp, imgdir / f"{stem}_bmp{k}.png")
                if p: out.append((f"OCR (bitmap {k} inside EMF)", ocr(p)))
            p = emf_vector_png(data, imgdir / f"{stem}_vector.png")
            if p: out.append(("OCR (EMF vector render)", ocr(p)))
        else:
            out.append(("WMF not rendered", ""))
        return out
    p = to_png(data, imgdir / f"{stem}.png")
    if p:
        out.append(("OCR", ocr(p)))
    return out


# ---- .xls (BIFF8)
def xls_units(path, imgdir):
    units = []
    wb = xlrd.open_workbook(str(path), formatting_info=False, on_demand=True)
    for sh in wb.sheets():
        vis = {0: "", 1: " (hidden sheet)", 2: " (very hidden sheet)"}.get(sh.visibility, "")
        for r in range(sh.nrows):
            for c in range(sh.ncols):
                v = sh.cell_value(r, c)
                if isinstance(v, str) and v.strip():
                    units.append(dict(where=f"sheet '{sh.name}'{vis} cell {xlrd.colname(c)}{r + 1}", how="cell text", text=v))
    data = olefile.OleFileIO(str(path)).openstream("Workbook").read()
    recs, i = [], 0
    while i + 4 <= len(data):
        t, n = struct.unpack_from("<HH", data, i); recs.append((t, data[i + 4:i + 4 + n])); i += 4 + n
    sheets = [s.name for s in wb.sheets()]
    grp, bof, j = b"", 0, 0
    while j < len(recs):
        t, b = recs[j]
        if t == 0x809:
            bof += 1
        if t == 0xEB:
            grp += b
            while j + 1 < len(recs) and recs[j + 1][0] == 0x3C:
                j += 1; grp += recs[j][1]
        if t == 0x1B6 and len(b) >= 12:
            cch = struct.unpack_from("<H", b, 10)[0]
            s, got, jj = "", 0, j + 1
            while jj < len(recs) and recs[jj][0] == 0x3C and got < cch:
                c = recs[jj][1]; part = c[1:].decode("utf-16le", "replace") if c[0] & 1 else c[1:].decode("latin-1")
                s += part; got += len(part); jj += 1
            sname = sheets[bof - 2] if 0 <= bof - 2 < len(sheets) else "chart/other substream"
            if s.strip():
                units.append(dict(where=f"text box on '{sname}'", how="text box", text=s[:cch]))
        j += 1
    blips = []

    def walk(buf, off, end):
        while off + 8 <= end:
            vi, rt, ln = struct.unpack_from("<HHI", buf, off)
            if vi & 0xF == 0xF:
                walk(buf, off + 8, min(end, off + 8 + ln))
            elif rt == 0xF007:
                walk(buf, off + 8 + 36, min(end, off + 8 + ln))
            elif 0xF018 <= rt <= 0xF117:
                blips.append((rt, vi >> 4, buf[off + 8:off + 8 + ln]))
            off += 8 + ln
    walk(grp, 0, len(grp))
    for n, (rt, inst, body) in enumerate(blips, 1):
        stem = f"pic{n:02d}"
        if rt in (0xF01A, 0xF01B):
            raw = None
            for h in (50, 66):
                try:
                    raw = zlib.decompress(body[h:]); break
                except Exception:
                    pass
            if raw is None:
                continue
            ext = "emf" if rt == 0xF01A else "wmf"
        else:
            raw = None
            for sig in (b"\x89PNG", b"\xff\xd8", b"II*\x00", b"MM\x00*"):
                p = body.find(sig, 0, 80)
                if p >= 0:
                    raw = body[p:]; break
            if raw is None:  # DIB: 17 or 33 byte header
                raw = body[33:] if struct.unpack_from("<I", body, 33)[0] in (12, 40, 108, 124) else body[17:]
            ext = "img"
        for how, txt in picture_texts(raw, ext, stem, imgdir):
            units.append(dict(where=f"embedded picture {n} ({ext})", how=how, text=txt))
    return units


# ---- .xlsx / .docx (OOXML)
def ooxml_pictures(z, imgdir):
    units = []
    media = [n for n in z.namelist() if "/media/" in n]
    owners = {}
    for rel in [n for n in z.namelist() if n.endswith(".rels")]:
        for m in re.finditer(r'Target="([^"]*media/[^"]+)"', z.read(rel).decode("utf-8", "replace")):
            owners.setdefault(Path(m.group(1)).name, set()).add(rel.split("/_rels/")[-1].replace(".rels", ""))
    for n, name in enumerate(media, 1):
        ext = Path(name).suffix.lower().lstrip(".")
        ext = ext if ext in ("emf", "wmf") else "img"
        own = ",".join(sorted(owners.get(Path(name).name, []))) or "?"
        for how, txt in picture_texts(z.read(name), ext, f"media{n:02d}_{Path(name).stem}", imgdir):
            units.append(dict(where=f"picture {Path(name).name} (used in {own})", how=how, text=txt))
    return units


def xlsx_units(path, imgdir):
    units = []
    wb = openpyxl.load_workbook(io.BytesIO(Path(path).read_bytes()), read_only=False, data_only=False)
    for ws in wb.worksheets:
        vis = "" if ws.sheet_state == "visible" else f" ({ws.sheet_state} sheet)"
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.strip() and not c.value.startswith("="):
                    units.append(dict(where=f"sheet '{ws.title}'{vis} cell {c.coordinate}", how="cell text", text=c.value))
                if c.comment:
                    units.append(dict(where=f"sheet '{ws.title}' comment {c.coordinate}", how="comment", text=c.comment.text))
    z = zipfile.ZipFile(path)
    for n in z.namelist():
        if n.startswith("xl/drawings/") and n.endswith(".xml"):
            t = " ".join(re.findall(r"<a:t>([^<]*)</a:t>", z.read(n).decode("utf-8", "replace")))
            if t.strip():
                units.append(dict(where=f"text box/shape in {n}", how="drawing text", text=t))
    return units + ooxml_pictures(z, imgdir)


def docx_units(path, imgdir):
    units = []
    z = zipfile.ZipFile(path)
    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    for part in [n for n in z.namelist() if re.match(r"word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$", n)]:
        root = ET.fromstring(z.read(part))
        for k, p in enumerate(root.iter(W + "p"), 1):
            t = "".join(x.text or "" for x in p.iter(W + "t"))
            if t.strip():
                units.append(dict(where=f"{part} paragraph {k}", how="document text", text=t))
    return units + ooxml_pictures(z, imgdir)


def find_hits(units):
    hits = []
    for u in units:
        txt = re.sub(r"[ \t]+", " ", u["text"])
        for line in txt.splitlines():
            if not line.strip() or NEGATIVE.search(line) and not any(
                    p.search(NEGATIVE.sub("", line)) for p in PATTERNS.values()):
                continue
            for item, rx in PATTERNS.items():
                m = rx.search(NEGATIVE.sub("", line))
                if m:
                    hits.append(dict(item=item, where=u["where"], how=u["how"], matched=m.group(0),
                                     label=line.strip()[:160]))
    return hits


def main():
    files = sorted(p for p in SRC_DIR.iterdir() if p.is_file() and not p.name.startswith((".", "._")))
    (WORK / "src").mkdir(parents=True, exist_ok=True)
    before = {p.name: sha(p) for p in files}
    results = []
    for p in files:
        rec = dict(file=p.name, type=p.suffix.lower(), size=p.stat().st_size, sha256=before[p.name], status="", units=0, hits=[])
        if p.name.startswith("~$"):
            rec["status"] = "Office lock file (owner-name placeholder, not a document) - not scanned"
            results.append(rec); continue
        local = WORK / "src" / p.name
        if not local.exists():
            shutil.copy2(p, local)
        assert sha(local) == before[p.name]
        imgdir = WORK / "img" / p.stem
        imgdir.mkdir(parents=True, exist_ok=True)
        try:
            is_zip = local.read_bytes()[:4] == b"PK\x03\x04"
            rec["container"] = "OOXML zip" if is_zip else "OLE2/BIFF"
            if p.suffix.lower() in (".xls", ".xlsx") and is_zip:  # several '.xls' files are really .xlsx
                units = xlsx_units(local, imgdir)
            elif p.suffix.lower() == ".xls":
                units = xls_units(local, imgdir)
            elif p.suffix.lower() == ".docx":
                units = docx_units(local, imgdir)
            else:
                units = []
            rec["status"] = "scanned"
        except Exception as e:  # report, keep going
            units, rec["status"] = [], f"ERROR {type(e).__name__}: {e}"
        rec["units"] = len(units)
        rec["n_pictures"] = len({u["where"] for u in units if "picture" in u["where"]})
        rec["ocr_chars"] = sum(len(u["text"]) for u in units if u["how"].startswith(("OCR", "EMF")))
        rec["hits"] = find_hits(units)
        (WORK / "text").mkdir(exist_ok=True)
        (WORK / "text" / (p.stem + ".json")).write_text(json.dumps(units, ensure_ascii=False, indent=0))
        results.append(rec)
        print(f"{p.name}: {rec['status']} units={len(units)} pictures={rec['n_pictures']} hits={len(rec['hits'])}", flush=True)
    after = {p.name: sha(p) for p in files}
    assert before == after, "a source file changed"
    (WORK / "scan_results.json").write_text(json.dumps(results, ensure_ascii=False, indent=1))
    print("sources unchanged:", before == after)


if __name__ == "__main__":
    main()
