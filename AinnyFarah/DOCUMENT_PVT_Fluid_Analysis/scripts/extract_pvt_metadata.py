#!/usr/bin/env python3
"""Local metadata extraction for the Dulang PVT / fluid-analysis reports on the NAS.

Source  : /Volumes/homes/public/DulangOSS/20260925/DOCUMENTS/PVT Fluid Analysis Document (read-only)
Outputs : AinnyFarah/DOCUMENT_PVT_Fluid_Analysis/{work,output,logs}

The source files are Excel/Word, not PDF. Units used as "pages":
  - .xlsx / .xls : one worksheet = one page (workbook order, hidden sheets included and marked)
  - .docx        : pages counted from rendered/explicit page breaks; total from docProps/app.xml
A page has extractable text if it has cell text, text-box text (xlsx drawings) or paragraph/table text.
Pages without text that hold raster images are OCR-selected; images on the first 3 pages (cover area)
are OCR'd too. OCR uses local Tesseract only. Nothing is sent anywhere; this script prints counts only.

"PDF Creation Date" is reported as the source file's own document-property creation date
(OOXML docProps/core.xml dcterms:created, or OLE2 SummaryInformation create_time).
"""
import hashlib
import io
import json
import re
import subprocess
import zipfile
from collections import Counter, OrderedDict, defaultdict
from datetime import datetime
from pathlib import Path

import olefile
import openpyxl
import xlrd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from PIL import Image

SRC = Path("/Volumes/homes/public/DulangOSS/20260925/DOCUMENTS/PVT Fluid Analysis Document")
ROOT = Path("/Users/farahyusof/Documents/GitHub/NoramiraSazuki/AinnyFarah/DOCUMENT_PVT_Fluid_Analysis")
WORK, OUT, LOGS = ROOT / "work", ROOT / "output", ROOT / "logs"
FRONT = 3                     # "cover and first few pages"
NF, REVIEW = "Not found", "Needs review"

# --------------------------------------------------------------------------- patterns
MONTHS = (r"Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|June?|July?|Aug(?:ust)?|"
          r"Sep(?:t(?:ember)?)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?")
DATE_RE = re.compile(
    rf"\b(?:\d{{1,2}}(?:st|nd|rd|th)?[\s.\-/]*(?:{MONTHS})\.?[\s,.\-/]*(?:19|20)\d\d"
    rf"|(?:{MONTHS})\.?\s+\d{{1,2}}(?:st|nd|rd|th)?,?\s+(?:19|20)\d\d"
    rf"|(?:{MONTHS})\.?[\s,]+(?:19|20)\d\d"
    rf"|\d{{1,2}}[/.\-]\d{{1,2}}[/.\-](?:19|20)\d\d|(?:19|20)\d\d-\d\d-\d\d)\b", re.I)
REPORT_DATE_LABEL = re.compile(r"\b(report\s*date|date\s*of\s*report|date\s*issued|issue\s*date|date\s*reported|"
                               r"report\s*issued|date\s*completed|completion\s*date|date)\b\s*[:.\-]?", re.I)
NOT_REPORT_DATE = re.compile(r"sampl|test\s*date|collect|receiv|recover|drill|run|perforat|production|spud|dst|"
                             r"bottle|transfer|flash|separat|shipp|arriv|dispatch|validat", re.I)
TITLE_KEYS = re.compile(r"\b(reservoir\s+fluids?\s+stud(?:y|ies)|pvt(?:\s+stud(?:y|ies))?|fluid\s+stud(?:y|ies)|"
                        r"fluid\s+analysis|compositional\s+analysis|water\s+analysis|crude\s+oil\s+analysis|"
                        r"bubble\s+point|sample\s+information|oil\s+analysis|gas\s+analysis|"
                        r"partial\s+reservoir\s+fluid|study\s+report|analysis\s+report)\b", re.I)
TITLE_REJECT = re.compile(r"\b(table|figure|fig\.|page\s*\d|prepared|customer|client|contractor|operator|"
                          r"well\s*:|field\s*:|date\s*:)\b|^\s*\d", re.I)

BY_LABELS = (r"prepared\s+by|performed\s+by|conducted\s+by|analy[sz]ed\s+by|submitted\s+by|issued\s+by|"
             r"reported\s+by|contractor|service\s+company|laboratory|lab\.?\s*name")
FOR_LABELS = r"prepared\s+for|performed\s+for|submitted\s+to|report\s+to|customer|client|operator|company\s*name|oil\s+company"
# A label counts only when it is the whole cell ('Client' | value) or is followed by ':' ('Client: value').
BY_CELL = re.compile(rf"^\s*({BY_LABELS})\s*[:.\-]?\s*$", re.I)
BY_INLINE = re.compile(rf"^\s*({BY_LABELS})\s*:\s*(.+)$", re.I)
FOR_CELL = re.compile(rf"^\s*({FOR_LABELS})\s*[:.\-]?\s*$", re.I)
FOR_INLINE = re.compile(rf"^\s*({FOR_LABELS})\s*:\s*(.+)$", re.I)
PERSON = re.compile(r"^(?:(?:Mr|Ms|Mrs|Dr|Ir)\.?\s+)?[A-Z][a-z]+(?:\s+(?:bin|binti|b\.|bt\.|[A-Z][a-z.]+)){1,4}$")
COMPANY_SUFFIX = re.compile(r"\b(sdn\.?\s*bhd\.?|berhad|inc\.?|incorporated|ltd\.?|limited|corp(?:oration)?\.?|"
                            r"plc|b\.v\.|s\.a\.|gmbh|llc|laboratories)\b", re.I)
KNOWN_ORG = re.compile(r"\b(esso(?:\s+production\s+malaysia(?:\s+inc\.?)?)?|epmi|exxon\w*|petronas(?:\s+carigali)?|"
                       r"carigali|core\s*lab(?:oratories)?|corelab|schlumberger|oilphase|flopetrol(?:\s+johnston)?|"
                       r"expro|weatherford|intertek|halliburton|baker\s+hughes|petronas\s+research[\w\s&]*|prss)\b", re.I)
ORG_STOP = {"company", "client", "customer", "operator", "contractor", "laboratory", "well", "field", "date", "report",
            "page", "sample", "samples", "test", "n/a", "na", "-", "nil", "none", "tbc"}

WELL_RE = re.compile(r"\bDULANG[\s\-_#]*(?:no\.?\s*)?([A-Z]?\s?-?\d{1,3}[A-Z]?(?:\s?[-_/]\s?\d{1,2}[A-Z]?){0,3})", re.I)
WELL_OCR_RE = re.compile(r"\b[DO0][UV][LI1|][A4][NM][G6C][\s\-_#]*([A-Z]?\s?-?\d{1,3}[A-Z]?(?:\s?[-_/]\s?\d{1,2}[A-Z]?){0,3})", re.I)
WELL_LABEL = re.compile(r"\bwell(?:\s*(?:name|no\.?|number))?\s*[:.\-]\s*(.+)", re.I)
FIELD_RE = re.compile(r"\b([A-Z][A-Za-z\-]{2,})\s+(?:oil\s+|gas\s+)?field\b", re.I)
FIELD_LABEL = re.compile(r"\bfield(?:\s*name)?\s*[:.\-]\s*([A-Za-z][A-Za-z\-\s]{1,30})", re.I)
FIELD_STOP = {"OIL", "GAS", "THE", "THIS", "THAT", "IN", "OF", "OFFSHORE", "ONSHORE", "MAGNETIC", "ELECTRIC", "FLOW",
              "TEST", "NEAR", "SAME", "MAIN", "EACH", "WHICH", "FROM", "AND", "FOR", "BY", "ITS", "TO", "AT", "ON",
              "EXPLORATION", "DEVELOPMENT", "PRODUCTION", "PRESSURE", "DATA", "SAMPLE", "FREE", "NAME", "WELL",
              "OILFIELD", "OTHER", "NEW", "FIELD", "FLASH", "SEPARATOR", "RESERVOIR", "TOTAL", "STOCK", "TANK"}


def sha(p):
    return hashlib.sha256(p.read_bytes()).hexdigest()


def clean(s):
    return re.sub(r"\s+", " ", str(s)).strip(" \t:;,.|-")


# --------------------------------------------------------------------------- readers
def file_kind(b):
    if b.startswith(b"\xd0\xcf\x11\xe0"):
        return "ole"
    if b.startswith(b"PK"):
        names = zipfile.ZipFile(io.BytesIO(b)).namelist()
        if any(n.startswith("xl/") for n in names):
            return "xlsx"
        if any(n.startswith("word/") for n in names):
            return "docx"
    return "unknown"


def ooxml_props(z):
    names = z.namelist()
    props = {}
    if "docProps/core.xml" in names:
        x = z.read("docProps/core.xml").decode("utf-8", "ignore")
        for tag, key in (("dcterms:created", "created"), ("dcterms:modified", "modified"), ("dc:creator", "author"),
                         ("cp:lastModifiedBy", "last_modified_by"), ("dc:title", "title")):
            m = re.search(rf"<{tag}[^>]*>(.*?)</{tag}>", x, re.S)
            if m and m.group(1).strip():
                props[key] = m.group(1).strip()
    if "docProps/app.xml" in names:
        x = z.read("docProps/app.xml").decode("utf-8", "ignore")
        for tag, key in (("Company", "company"), ("Pages", "pages"), ("Application", "application")):
            m = re.search(rf"<{tag}>(.*?)</{tag}>", x, re.S)
            if m and m.group(1).strip():
                props[key] = m.group(1).strip()
    return props


def read_xlsx(b):
    z = zipfile.ZipFile(io.BytesIO(b))
    names = z.namelist()
    wb = openpyxl.load_workbook(io.BytesIO(b), read_only=True, data_only=True)
    # sheet order/name/state from workbook.xml to map sheetN.xml targets
    wbx = z.read("xl/workbook.xml").decode("utf-8", "ignore")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8", "ignore")
    rid_target = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
    rid_target.update({a: b_ for b_, a in re.findall(r'Target="([^"]+)"[^>]*Id="([^"]+)"', rels)})
    sheets = re.findall(r"<sheet\b([^>]*)/>", wbx)
    pages = []
    for idx, (attrs, ws) in enumerate(zip(sheets, wb.worksheets), 1):
        rid = re.search(r'r:id="([^"]+)"', attrs).group(1)
        state = (re.search(r'state="([^"]+)"', attrs) or [None, "visible"])[1]
        target = rid_target.get(rid, "")
        path = "xl/" + target.lstrip("/").replace("xl/", "") if not target.startswith("/xl/") else target.lstrip("/")
        lines = []
        for row in ws.iter_rows(values_only=True):
            vals = [clean(v) for v in row if v not in (None, "") and clean(v)]
            if vals:
                lines.append(" | ".join(vals))
        images, box_text = [], []
        srel = path.replace("worksheets/", "worksheets/_rels/") + ".rels"
        if srel in names:
            for d in re.findall(r'Target="\.\./drawings/(drawing\d+\.xml)"', z.read(srel).decode("utf-8", "ignore")):
                dpath = f"xl/drawings/{d}"
                if dpath in names:
                    dx = z.read(dpath).decode("utf-8", "ignore")
                    box_text += [clean(t) for t in re.findall(r"<a:t>(.*?)</a:t>", dx, re.S) if clean(t)]
                drel = f"xl/drawings/_rels/{d}.rels"
                if drel in names:
                    for m in re.findall(r'Target="\.\./media/([^"]+)"', z.read(drel).decode("utf-8", "ignore")):
                        images.append(("xl/media/" + m, z.read("xl/media/" + m)))
        if box_text:
            lines += ["[textbox] " + t for t in box_text]
        pages.append({"unit": idx, "name": ws.title, "state": state, "lines": lines, "images": images})
    return pages, ooxml_props(z)


def read_ole_xls(b):
    wb = xlrd.open_workbook(file_contents=b, on_demand=True)
    pages = []
    for i in range(wb.nsheets):
        s = wb.sheet_by_index(i)
        lines = []
        for r in range(s.nrows):
            vals = []
            for c in range(s.ncols):
                cell = s.cell(r, c)
                v = cell.value
                if cell.ctype == xlrd.XL_CELL_DATE:
                    try:
                        v = xlrd.xldate_as_datetime(v, wb.datemode).strftime("%d %b %Y")
                    except Exception:
                        pass
                v = clean(v)
                if v:
                    vals.append(v)
            if vals:
                lines.append(" | ".join(vals))
        vis = {0: "visible", 1: "hidden", 2: "veryHidden"}.get(getattr(s, "visibility", 0), "visible")
        pages.append({"unit": i + 1, "name": s.name, "state": vis, "lines": lines, "images": []})
    ole = olefile.OleFileIO(b)
    raw = ole.openstream("Workbook" if ole.exists("Workbook") else "Book").read()
    # BIFF8 splits the drawing group (0x00EB) over CONTINUE (0x003C) records; join the payloads
    # so embedded pictures are contiguous again.
    stream, pos, prev = bytearray(), 0, None
    while pos + 4 <= len(raw):
        rtype, rlen = int.from_bytes(raw[pos:pos + 2], "little"), int.from_bytes(raw[pos + 2:pos + 4], "little")
        payload = raw[pos + 4:pos + 4 + rlen]
        if rtype == 0x00EB or (rtype == 0x003C and prev == 0x00EB):
            stream += payload
            prev = 0x00EB
        else:
            prev = rtype
        pos += 4 + rlen
    stream = bytes(stream)
    imgs = []
    for m in re.finditer(rb"\x89PNG\r\n\x1a\n", stream):
        end = stream.find(b"IEND", m.start())
        if end > 0:
            imgs.append((f"ole_png_{len(imgs) + 1}.png", stream[m.start():end + 8]))
    for m in re.finditer(rb"\xff\xd8\xff[\xe0-\xef]", stream):
        end = stream.find(b"\xff\xd9", m.start())
        if end > 0:
            imgs.append((f"ole_jpg_{len(imgs) + 1}.jpg", stream[m.start():end + 2]))
    meta = ole.get_metadata()
    props = {}
    for k, attr in (("created", "create_time"), ("modified", "last_saved_time"), ("author", "author"),
                    ("last_modified_by", "last_saved_by"), ("title", "title"), ("company", "company")):
        v = getattr(meta, attr, None)
        if v:
            props[k] = v.isoformat() if isinstance(v, datetime) else (v.decode("latin-1", "ignore") if isinstance(v, bytes) else str(v))
    # OLE images cannot be tied to a sheet without parsing BIFF drawings: attach to sheet 1 (cover) and flag.
    if pages and imgs:
        pages[0]["images"] = imgs
        pages[0]["ole_images_unassigned"] = True
    return pages, props


def read_docx(b):
    z = zipfile.ZipFile(io.BytesIO(b))
    x = z.read("word/document.xml").decode("utf-8", "ignore")
    rels = z.read("word/_rels/document.xml.rels").decode("utf-8", "ignore")
    rid_media = dict(re.findall(r'Id="([^"]+)"[^>]*Target="(media/[^"]+)"', rels))
    rid_media.update({a: b_ for b_, a in re.findall(r'Target="(media/[^"]+)"[^>]*Id="([^"]+)"', rels)})
    body = re.search(r"<w:body>(.*)</w:body>", x, re.S).group(1)
    page, pages = 1, defaultdict(lambda: {"lines": [], "images": []})
    for para in re.findall(r"<w:p[ >].*?</w:p>", body, re.S):
        brk = len(re.findall(r'<w:lastRenderedPageBreak/>|<w:br [^>]*w:type="page"', para))
        text = clean("".join(re.findall(r"<w:t[^>]*>(.*?)</w:t>", para, re.S)))
        if brk and re.search(r'^(?:(?!<w:t[ >]).)*(<w:lastRenderedPageBreak/>|<w:br [^>]*w:type="page")', para, re.S):
            page += brk
            brk = 0
        if text:
            pages[page]["lines"].append(text)
        for rid in re.findall(r'r:embed="([^"]+)"', para):
            if rid in rid_media:
                pages[page]["images"].append(("word/" + rid_media[rid], z.read("word/" + rid_media[rid])))
        page += brk
    props = ooxml_props(z)
    total = int(props.get("pages", max(pages) if pages else 1))
    out = []
    for u in range(1, max(total, max(pages) if pages else 1) + 1):
        p = pages.get(u, {"lines": [], "images": []})
        out.append({"unit": u, "name": f"page {u}", "state": "visible", "lines": p["lines"], "images": p["images"]})
    return out, props


# --------------------------------------------------------------------------- OCR
def ocr_image(name, data, dest):
    ext = Path(name).suffix.lower()
    if ext not in (".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff"):
        return None, f"unsupported image type {ext}"
    try:
        im = Image.open(io.BytesIO(data))
        im.load()
    except Exception as e:
        return None, f"unreadable image ({type(e).__name__})"
    if im.mode not in ("L", "RGB"):
        im = im.convert("RGB")
    if max(im.size) < 1600:
        f = 1600 / max(im.size)
        im = im.resize((int(im.width * f), int(im.height * f)))
    dest.parent.mkdir(parents=True, exist_ok=True)
    im.save(dest)
    r = subprocess.run(["tesseract", str(dest), "stdout", "-l", "eng", "--psm", "3"], capture_output=True, text=True)
    return r.stdout, "ok" if r.returncode == 0 else f"tesseract rc={r.returncode}"


# --------------------------------------------------------------------------- extraction
def plausible_org(v):
    v = clean(v)
    if not (3 <= len(v) <= 90) or v.lower() in ORG_STOP:
        return False
    if DATE_RE.fullmatch(v) or re.fullmatch(r"[\d\s.,/\-:]+", v):
        return False
    if sum(c.isalpha() for c in v) < 3:
        return False
    return bool(COMPANY_SUFFIX.search(v) or KNOWN_ORG.search(v))


def labelled_values(line, next_lines, cell_re, inline_re):
    """(label, value) pairs: 'Label: value' in one cell, or a label cell followed by the next
    non-empty cell in the same row, or the first cell of the next row."""
    cells = [clean(c) for c in line.split("|")]
    out = []
    for j, c in enumerate(cells):
        m = inline_re.match(c)
        if m:
            out.append((clean(m.group(1)), clean(m.group(2))))
            continue
        m = cell_re.match(c)
        if m:
            val = next((x for x in cells[j + 1:] if x), "")
            if not val and next_lines:
                val = clean(next_lines[0].split("|")[0])
            if val:
                out.append((clean(m.group(1)), val))
    return out


def classify_party(v):
    """organisation / person / possible organisation / reject."""
    if plausible_org(v):
        return "organisation", "high"
    if PERSON.match(v) and len(v) <= 40:
        return "person", "review"
    if re.fullmatch(r"[A-Z][A-Z&.,'\-]*(?:\s+[A-Z][A-Z&.,'\-]*){0,4}", v) and 3 <= len(v) <= 45 \
            and v.lower() not in ORG_STOP and not DATE_RE.search(v):
        return "possible organisation (no company suffix)", "review"
    return None, None


def extract(pages, fname):
    front = [p for p in pages if p["unit"] <= FRONT]
    ev = defaultdict(list)

    def add(field, value, unit, method, source, conf):
        ev[field].append({"value": value, "unit": unit, "method": method, "source": source, "confidence": conf})

    for p in front:
        for src_name, lines in (("text", p["lines"]), ("ocr", p.get("ocr_lines", []))):
            for i, line in enumerate(lines):
                nxt = lines[i + 1:i + 3]
                # title
                if TITLE_KEYS.search(line) and not TITLE_REJECT.search(line) and 8 <= len(line) <= 140:
                    add("Report Title", clean(line.split("|")[0] if len(line.split("|")[0]) > 8 else line), p["unit"],
                        "keyword line", src_name, "medium")
                # created by / for
                for field, cre, ire in (("Created By", BY_CELL, BY_INLINE), ("Created For", FOR_CELL, FOR_INLINE)):
                    for lab, val in labelled_values(line, nxt, cre, ire):
                        kind, conf = classify_party(val)
                        if kind:
                            add(field, val if kind == "organisation" else f"{val} [{kind}]", p["unit"],
                                f"label '{lab}'", src_name, conf)
                        else:
                            add(field + " (rejected)", val[:80], p["unit"], f"label '{lab}' value not a party name",
                                src_name, "rejected")
                for m in KNOWN_ORG.finditer(line):
                    add("Organisation mentioned", clean(m.group(0)), p["unit"], "known organisation name", src_name, "low")
                for seg in line.split("|"):
                    if COMPANY_SUFFIX.search(seg) and plausible_org(seg) and len(clean(seg)) <= 70 and \
                            re.search(r"(sdn|bhd|inc|ltd|limited|corp|laborator)", seg, re.I):
                        add("Organisation mentioned", clean(seg), p["unit"], "company suffix", src_name, "low")
                # report date
                for m in REPORT_DATE_LABEL.finditer(line):
                    if NOT_REPORT_DATE.search(line):
                        continue
                    tail = line[m.end():]
                    dm = DATE_RE.search(tail) or (DATE_RE.search(nxt[0]) if nxt and not clean(tail) else None)
                    if dm:
                        add("Report Date", clean(dm.group(0)), p["unit"], f"label '{clean(m.group(1))}'", src_name,
                            "high" if "report" in m.group(1).lower() or "issue" in m.group(1).lower() else "medium")
                if p["unit"] == 1 and not NOT_REPORT_DATE.search(line):
                    for dm in DATE_RE.finditer(line):
                        add("Cover date (unlabelled)", clean(dm.group(0)), p["unit"], "date on first page", src_name, "low")
                # field
                for m in FIELD_LABEL.finditer(line):
                    v = clean(m.group(1)).upper().split()[0]
                    if v not in FIELD_STOP:
                        add("Field Name", v, p["unit"], "label 'Field:'", src_name, "high")
                for m in FIELD_RE.finditer(line):
                    w = m.group(1).upper()
                    if w not in FIELD_STOP and m.group(1)[0].isupper():
                        add("Field Name", w, p["unit"], "'<Name> Field' phrase", src_name, "medium")
    # wells: whole document, text + OCR
    wells = Counter()
    well_pages = defaultdict(set)
    ocr_variants = Counter()
    for p in pages:
        for src_name, lines in (("text", p["lines"]), ("ocr", p.get("ocr_lines", []))):
            blob = "\n".join(lines)
            for m in WELL_RE.finditer(blob):
                raw = clean(m.group(0))
                wells[raw] += 1
                well_pages[raw].add(p["unit"])
            if src_name == "ocr":
                for m in WELL_OCR_RE.finditer(blob):
                    raw = clean(m.group(0))
                    if not re.match(r"dulang", raw, re.I):
                        ocr_variants[raw] += 1
                        well_pages[raw].add(p["unit"])
            for i, line in enumerate(lines):
                if p["unit"] <= FRONT:
                    for m in WELL_LABEL.finditer(line):
                        v = clean(m.group(1).split("|")[0])
                        if v and re.search(r"\d", v) and len(v) <= 30:
                            add("Well (labelled)", v, p["unit"], "label 'Well:'", src_name, "high")
    return ev, wells, well_pages, ocr_variants


def canon_well(raw):
    s = re.sub(r"[\s_#]+", "", raw.upper())
    s = re.sub(r"^[DO0][UV][LI1|][A4][NM][G6C]", "DULANG", s)
    s = re.sub(r"^DULANG-?(NO\.?)?", "DULANG-", s)
    return s


def choose(items, prefer=("high", "medium", "review", "low")):
    for conf in prefer:
        vals = [e for e in items if e["confidence"] == conf]
        if vals:
            uniq = list(OrderedDict.fromkeys(e["value"] for e in vals))
            return uniq, vals
    return [], []


def parse_date(s):
    s2 = re.sub(r"(st|nd|rd|th)\b", "", s, flags=re.I).replace(",", " ").replace(".", " ")
    s2 = re.sub(r"\s+", " ", s2).strip()
    for fmt in ("%d %B %Y", "%d %b %Y", "%B %d %Y", "%b %d %Y", "%B %Y", "%b %Y", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y",
                "%d-%b-%Y", "%d-%B-%Y", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%dT%H:%M:%SZ"):
        try:
            return datetime.strptime(s2.replace("Sept", "Sep"), fmt)
        except ValueError:
            pass
    return None


# --------------------------------------------------------------------------- main
def main():
    for d in (WORK, OUT, LOGS):
        d.mkdir(exist_ok=True)
    files = sorted(p for p in SRC.iterdir() if p.is_file() and not p.name.startswith(("~$", ".")))
    hashes_before = {p: sha(p) for p in files}

    inventory, meta_rows, evidence_rows, variant_rows, review_rows = [], [], [], [], []
    totals = Counter()

    for n, f in enumerate(files, 1):
        b = f.read_bytes()
        kind = file_kind(b)
        try:
            pages, props = {"xlsx": read_xlsx, "ole": read_ole_xls, "docx": read_docx}[kind](b)
        except Exception as e:
            review_rows.append([f.name, "File", "unreadable", f"{type(e).__name__}: {e}"[:200]])
            print(f"{n:2} {kind:5} ERROR {type(e).__name__}")
            continue
        fdir = WORK / f"{n:02d}"
        fdir.mkdir(exist_ok=True)
        ocr_units = 0
        for p in pages:
            has_text = bool(p["lines"])
            select = (not has_text and p["images"]) or (p["unit"] <= FRONT and p["images"])
            p["ocr_lines"], statuses = [], []
            if select:
                for k, (iname, data) in enumerate(p["images"], 1):
                    txt, st = ocr_image(iname, data, fdir / f"unit{p['unit']:03d}_img{k}{Path(iname).suffix.lower() or '.png'}")
                    statuses.append(st)
                    if txt:
                        p["ocr_lines"] += [clean(l) for l in txt.splitlines() if clean(l)]
                if any(s == "ok" for s in statuses):
                    ocr_units += 1
            (fdir / f"unit{p['unit']:03d}.txt").write_text("\n".join(p["lines"] + ["--- OCR ---"] + p["ocr_lines"]))
            inventory.append([f.name, kind, p["unit"], p["name"], p["state"], len(p["lines"]), "Yes" if has_text else "No",
                              len(p["images"]), "Yes" if select else "No",
                              "; ".join(sorted(set(statuses))) if statuses else "", len(" ".join(p["ocr_lines"])),
                              "OLE images attached to sheet 1 (sheet unknown)" if p.get("ole_images_unassigned") else ""])
            totals["pages"] += 1
            totals["text"] += has_text
            totals["notext"] += not has_text
            totals["ocr_selected"] += bool(select)
            totals["ocr_selected_notext"] += bool(select and not has_text)
            totals["images"] += len(p["images"])

        ev, wells, well_pages, ocr_variants = extract(pages, f.name)
        for field, items in ev.items():
            for e in items:
                evidence_rows.append([f.name, field, e["value"], e["unit"], pages[e["unit"] - 1]["name"] if e["unit"] <= len(pages) else "",
                                      e["method"], e["source"], e["confidence"]])

        def pick(field, multi_ok=False):
            vals, items = choose(ev.get(field, []))
            if not vals:
                return NF, ""
            pages_ = sorted({e["unit"] for e in items})
            evid = f"pp. {', '.join(map(str, pages_))} ({items[0]['method']}, {items[0]['source']})"
            if len(vals) > 1 and not multi_ok:
                return f"{REVIEW}: " + " / ".join(vals[:4]), evid
            conf = items[0]["confidence"]
            v = "; ".join(vals[:4])
            return (v if conf in ("high",) or multi_ok else f"{v} ({REVIEW}: {conf} confidence)"), evid

        tc = ev.get("Report Title", [])
        cover = [e for e in tc if e["unit"] == min((x["unit"] for x in tc), default=0)]
        if not tc:
            title, title_ev = NF, ""
        else:
            vals = list(OrderedDict.fromkeys(e["value"] for e in cover))
            first_page = cover[0]["unit"]
            title = " — ".join(vals[:2])
            title_ev = f"p. {first_page} ({cover[0]['source']}, keyword lines on first page with a title)"
            if first_page != 1:
                title += f" ({REVIEW}: no title-like line on page 1; taken from page {first_page})"
            elif len(vals) > 2:
                title += f" ({REVIEW}: {len(vals)} title-like lines on page 1)"
        by, by_ev = pick("Created By", multi_ok=True)
        fr, fr_ev = pick("Created For", multi_ok=True)
        orgs = list(OrderedDict.fromkeys(e["value"] for e in ev.get("Organisation mentioned", [])))
        if by == NF and orgs:
            by = f"{REVIEW}: no 'Prepared by'-type label; organisations mentioned: {', '.join(orgs[:4])}"
        if fr == NF and orgs:
            fr = f"{REVIEW}: no 'Prepared for'/'Client' label; organisations mentioned: {', '.join(orgs[:4])}"
        rd, rd_ev = pick("Report Date")
        if rd == NF:
            cov = list(OrderedDict.fromkeys(e["value"] for e in ev.get("Cover date (unlabelled)", [])))
            if len(cov) == 1:
                rd, rd_ev = f"{cov[0]} ({REVIEW}: unlabelled date on first page)", "p. 1 (unlabelled)"
            elif cov:
                rd, rd_ev = f"{REVIEW}: several unlabelled dates on first page ({', '.join(cov[:4])})", "p. 1"
        created = props.get("created", "")
        rd_dt = parse_date(re.sub(r"\s*\(.*$", "", rd)) if not rd.startswith((NF, REVIEW)) else None
        cr_dt = parse_date(created[:19]) if created else None
        if rd_dt and cr_dt:
            delta = (cr_dt - rd_dt).days
            cmp_ = f"File created {delta:+d} days vs report date" + ("" if abs(delta) <= 31 else f" — {REVIEW}: differs by >1 month (file likely re-keyed/converted later)")
        else:
            cmp_ = "Not comparable (" + ("no report date" if not rd_dt else "no file creation date") + ")"

        # wells
        canon = Counter()
        for raw, c in wells.items():
            canon[canon_well(raw)] += c
        labelled = list(OrderedDict.fromkeys(e["value"] for e in ev.get("Well (labelled)", [])))
        if canon:
            top, topc = canon.most_common(1)[0]
            others = [w for w in canon if w != top]
            well = top if not others or topc >= 2 * max(canon[w] for w in others) else f"{REVIEW}: {', '.join(w for w, _ in canon.most_common(4))}"
        elif labelled:
            well = f"{labelled[0]} ({REVIEW}: from 'Well:' label only)"
        else:
            well = NF
        for raw, c in sorted(wells.items(), key=lambda x: -x[1]):
            variant_rows.append([f.name, canon_well(raw), raw, c, ", ".join(map(str, sorted(well_pages[raw]))), "text/OCR exact"])
        for raw, c in sorted(ocr_variants.items(), key=lambda x: -x[1]):
            variant_rows.append([f.name, canon_well(raw), raw, c, ", ".join(map(str, sorted(well_pages[raw]))), "OCR variant (misread of DULANG)"])
        fname_well = re.search(r"DULANG[\s\-_]*[A-Z0-9][A-Z0-9\-_ ]*?(?=_| PVT|_PVT)", f.name, re.I)
        fname_well = canon_well(fname_well.group(0)) if fname_well else ""
        well_check = ("" if not fname_well or well.startswith((NF, REVIEW)) or canon_well(well) == fname_well
                      else f"{REVIEW}: filename suggests {fname_well}")
        field, field_ev = pick("Field Name")

        meta_rows.append([f.name, kind, len(pages), sum(1 for p in pages if p["lines"]), sum(1 for p in pages if not p["lines"]),
                          ocr_units, title, title_ev, by, by_ev, fr, fr_ev, rd, rd_ev, created or NF, props.get("modified", ""),
                          cmp_, well, "; ".join(f"{w} (x{c})" for w, c in canon.most_common(6)),
                          "; ".join(f"{w} (x{c})" for w, c in ocr_variants.most_common(4)), well_check, field, field_ev,
                          props.get("author", ""), props.get("company", "")])
        for col, val in (("Report Title", title), ("Created By", by), ("Created For", fr), ("Report Date", rd),
                         ("Well Name", well), ("Field Name", field)):
            if val.startswith(NF) or REVIEW in val:
                review_rows.append([f.name, col, "Not found" if val.startswith(NF) else "Needs review", "see Metadata / Evidence"])
        if well_check:
            review_rows.append([f.name, "Well Name", "Needs review", "differs from filename"])
        if "Needs review" in cmp_:
            review_rows.append([f.name, "Report Date vs File Creation Date", "Needs review", "differs by more than one month"])
        status = lambda v: "NF" if v.startswith(NF) else "REV" if REVIEW in v else "ok"
        print(f"{n:2} {kind:5} pages={len(pages):3} text={sum(1 for p in pages if p['lines']):3} "
              f"ocr={ocr_units:2} | title={status(title):3} by={status(by):3} for={status(fr):3} "
              f"date={status(rd):3} well={status(well):3} field={status(field):3}")

    write_book(meta_rows, inventory, evidence_rows, variant_rows, review_rows, totals, len(files))
    assert all(sha(p) == h for p, h in hashes_before.items()), "SOURCE FILE CHANGED"
    (LOGS / "run_summary.json").write_text(json.dumps({"files": len(files), **totals,
                                                         "run": datetime.now().isoformat(timespec="seconds")}, indent=2))
    print(f"\nTOTAL pages={totals['pages']} with_text={totals['text']} without_text={totals['notext']} "
          f"ocr_selected={totals['ocr_selected']} (of which no-text pages={totals['ocr_selected_notext']}) images={totals['images']}")
    print("source files unchanged (sha256 verified)")
    print(OUT / "PVT_Report_Metadata.xlsx")


def write_book(meta, inv, evid, var, rev, totals, nfiles):
    wb = openpyxl.Workbook()
    s = wb.active
    s.title = "Test Summary"
    for row in [["Item", "Value", "Note"],
                ["Files reviewed", nfiles, "Lock files (~$...) excluded"],
                ["Total file pages", totals["pages"], "xlsx/xls: worksheets; docx: rendered pages"],
                ["Pages with extractable text", totals["text"], "cell, text-box or paragraph text present"],
                ["Pages without extractable text", totals["notext"], ""],
                ["OCR selected (pages)", totals["ocr_selected"],
                 f"{totals['ocr_selected_notext']} text-less pages with images + image-bearing pages within the first {FRONT}"],
                ["Embedded images found", totals["images"], "OCR: local Tesseract (eng)"],
                ["Run time", datetime.now().strftime("%Y-%m-%d %H:%M"), "No file contents sent to external services"]]:
        s.append(row)
    sheets = [
        ("Metadata", ["File Name", "Actual Format", "Total Pages", "Pages With Text", "Pages Without Text", "OCR Pages",
                      "Report Title", "Title Evidence", "Created By", "Created By Evidence", "Created For",
                      "Created For Evidence", "Report Date", "Report Date Evidence", "File Creation Date (document properties)",
                      "File Modified Date", "Report Date vs File Creation Date", "Well Name", "Well Name Variants (canonical)",
                      "OCR Variants", "Well vs Filename Check", "Field Name", "Field Evidence", "Doc Property Author",
                      "Doc Property Company"], meta),
        ("Page Inventory", ["File Name", "Format", "Page/Unit", "Sheet / Page Name", "Visibility", "Text Lines",
                            "Extractable Text", "Images", "OCR Selected", "OCR Status", "OCR Characters", "Note"], inv),
        ("Evidence", ["File Name", "Field", "Value", "Page/Unit", "Sheet / Page Name", "Method", "Source (text/ocr)",
                      "Confidence"], evid),
        ("Well Variants", ["File Name", "Canonical", "As Found", "Count", "Pages", "Type"], var),
        ("Review Flags", ["File Name", "Field", "Status", "Note"], rev),
    ]
    for title, hdr, rows in sheets:
        ws = wb.create_sheet(title)
        ws.append(hdr)
        for r in rows:
            ws.append(r)
    for ws in wb.worksheets:
        for c in ws[1]:
            c.font = Font(bold=True)
            c.fill = PatternFill("solid", fgColor="DDEBF7")
            c.alignment = Alignment(wrap_text=True, vertical="top")
        for i in range(1, ws.max_column + 1):
            ws.column_dimensions[get_column_letter(i)].width = 14 if ws.title != "Test Summary" else [34, 14, 70][i - 1]
        ws.freeze_panes = "A2"
    for col, w in zip("ABCDEFGHIJKLMNOPQRSTUVWXY", [60, 8, 8, 8, 8, 8, 50, 30, 45, 30, 45, 30, 28, 30, 22, 22, 40, 26, 40, 30, 30, 14, 30, 20, 20]):
        wb["Metadata"].column_dimensions[col].width = w
    wb.save(OUT / "PVT_Report_Metadata.xlsx")


if __name__ == "__main__":
    main()
