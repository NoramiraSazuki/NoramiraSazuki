#!/usr/bin/env python3
"""Extract DST report metadata from local OCR text into an Excel workbook.

Inputs : OCR sidecar TXT files (one page per form-feed) from run_ocr.py,
         original PDFs (read-only) for the physical page count.
Output : DOCUMENT_DST/DST_Report_Metadata.xlsx
         - "Metadata": one row per source PDF (7 requested columns)
         - "Extraction_Notes": evidence (page numbers, rule used) for review

Runs fully offline. Prints only found/not-found status, never document text.
"""
import re
import subprocess
from collections import Counter, OrderedDict
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

SRC_ROOT = Path("/Volumes/homes/public/DulangOSS/20260925/DOCUMENTS/DST Document")
OUT_ROOT = Path("/Users/farahyusof/Documents/GitHub/NoramiraSazuki/AinnyFarah/DOCUMENT_DST")
TXT_DIR = OUT_ROOT / "ocr_text"
PDF_DIR = OUT_ROOT / "searchable_pdf"
TMP_DIR = OUT_ROOT / "tmp"
XLSX = OUT_ROOT / "DST_Report_Metadata.xlsx"

FRONT_PAGES = 5          # "cover and first few pages"
NOT_FOUND = "Not found"
REVIEW = "Needs review"

# ---------------------------------------------------------------- helpers

MONTHS = (r"Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|June?|July?|"
          r"Aug(?:ust)?|Sep(?:t(?:ember)?)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?")
DATE_RE = re.compile(
    rf"\b(?:\d{{1,2}}(?:st|nd|rd|th)?[\s.\-/]*(?:{MONTHS})\.?[\s,.\-/]*(?:19|20)\d\d"   # 1 Sept. 1990
    rf"|(?:{MONTHS})\.?\s+\d{{1,2}}(?:st|nd|rd|th)?,?\s+(?:19|20)\d\d"                  # Sept 1, 1990
    rf"|(?:{MONTHS})\.?[\s,]+(?:19|20)\d\d"                                             # September 1990
    rf"|\d{{1,2}}[/.\-]\d{{1,2}}[/.\-](?:19|20)?\d\d)\b",                               # 01/09/1990
    re.IGNORECASE)
DATE_LABEL_RE = re.compile(
    r"\b(report\s+date|date\s+of\s+(?:report|issue)|issue\s+date|date\s+issued|issued(?:\s+on)?|"
    r"prepared\s+on|revision\s+date|date)\b\s*[:.\-]?", re.IGNORECASE)
DATE_EXCLUDE_RE = re.compile(r"test|spud|drill|perforat|run\b|flow|shut|sampl|logg|complet|gaug",
                             re.IGNORECASE)

ROLE_RE = re.compile(r"\b(prepared|submitted|presented|compiled|written|issued)\s+(by|for|to)\b\s*[:.\-]?\s*(.*)",
                     re.IGNORECASE)
COMPANY_SUFFIX_RE = re.compile(
    r"((?:[A-Z][\w&.,'\-]*\s+){0,6}?"
    r"(?:Sdn\.?\s*Bhd\.?|Berhad|Inc\.?|Incorporated|Ltd\.?|Limited|Corporation|Corp\.?|"
    r"Company|PLC|B\.V\.|S\.A\.|GmbH|Services))(?=\W|$)")
KNOWN_COMPANY_RE = re.compile(
    r"\b(EPMI|Esso(?:\s+Production\s+Malaysia(?:\s+Inc\.?)?)?|Exxon\w*|PETRONAS(?:\s+Carigali)?|Carigali|"
    r"Schlumberger|Flopetrol(?:\s+Johnston)?|Johnston|Halliburton|Baker\s+\w+|Expro|Core\s+Lab\w*|"
    r"Shell|Sarawak\s+Shell|Petroleum\s+Testing\s+Services|Western\s+Atlas|Dresser\s+Atlas|Otis)\b",
    re.IGNORECASE)

FIELD_STOP = {
    "OIL", "GAS", "THE", "THIS", "THAT", "IN", "OF", "OFFSHORE", "ONSHORE", "MAGNETIC", "ELECTRIC",
    "FLOW", "TEST", "NEAR", "ENTIRE", "SAME", "WHOLE", "PRODUCING", "NEW", "MAIN", "ADJACENT", "EACH",
    "WHICH", "FROM", "AND", "FOR", "BY", "ITS", "OUR", "A", "AN", "TO", "AT", "ON", "AS", "IS", "OR",
    "EXPLORATION", "DEVELOPMENT", "PRODUCTION", "PRESSURE", "WELLSITE", "DATA", "GRAVITY", "SAMPLE",
    "VELOCITY", "SONIC", "FREQUENCY", "FREE", "TEXT", "INPUT", "BLANK", "NAME", "WELL", "OILFIELD",
    "GASFIELD", "OTHER", "SURROUNDING", "STRESS", "FAR", "NEARBY", "DISCOVERY", "COMMERCIAL",
}
WELL_PREFIX_STOP = {
    "DST", "TEST", "ZONE", "RUN", "PAGE", "FIGURE", "FIG", "TABLE", "APPENDIX", "SECTION", "ISO",
    "API", "TYPE", "MODEL", "SERIAL", "PART", "ITEM", "SAMPLE", "CHART", "GAUGE", "PLATE", "SHEET",
    "NUMBER", "REV", "REVISION", "STAGE", "PHASE", "STEP", "CASE", "UNIT", "BLOCK", "PERIOD", "FLOW",
    "SHUT", "BUILDUP", "DRAWDOWN", "CHOKE", "BOTTLE", "CYLINDER", "TANK", "SEPARATOR", "LINE",
    "PERFORATION", "INTERVAL", "RESERVOIR", "SAND", "LAYER", "UPPER", "LOWER", "MIDDLE", "ANNEX",
    "VOLUME", "CHAPTER", "PART", "PLOT", "DATE", "TIME", "HOUR", "HOURS", "DAY", "DAYS",
}


def clean(s):
    return re.sub(r"\s+", " ", s).strip(" \t:;,.-|")


def lines_of(page):
    return [clean(l) for l in page.splitlines()]


def letters_ratio(s):
    return sum(c.isalpha() for c in s) / max(len(s), 1)


def upper_ratio(s):
    letters = [c for c in s if c.isalpha()]
    return sum(c.isupper() for c in letters) / max(len(letters), 1)


def page_count(pdf):
    out = subprocess.run(["qpdf", "--show-npages", str(pdf)], capture_output=True, text=True)
    return int(out.stdout.strip()) if out.returncode == 0 and out.stdout.strip().isdigit() else None


def text_layer_pages(pdf, n_pages):
    """Per-page text from the searchable PDF's text layer (existing + OCR), via local Ghostscript."""
    TMP_DIR.mkdir(parents=True, exist_ok=True)
    tmp = TMP_DIR / "page_text.txt"
    pages = []
    for i in range(1, n_pages + 1):
        subprocess.run(["gs", "-q", "-dNOPAUSE", "-dBATCH", "-dSAFER", "-sDEVICE=txtwrite",
                        f"-dFirstPage={i}", f"-dLastPage={i}", f"-sOutputFile={tmp}", str(pdf)],
                       capture_output=True)
        pages.append(tmp.read_text(errors="ignore") if tmp.exists() else "")
    tmp.unlink(missing_ok=True)
    return pages


def load_pages(txt, pdf, n_pages):
    """Combine OCR sidecar and PDF text layer, taking the richer text for each page.

    With --redo-ocr the sidecar can omit a page's pre-existing text, so neither
    source alone is complete.
    """
    side = txt.read_text(errors="ignore").split("\f") if txt.exists() else []
    layer = text_layer_pages(pdf, n_pages) if pdf.exists() and n_pages else []
    n = max(len(side), len(layer))
    side += [""] * (n - len(side))
    layer += [""] * (n - len(layer))
    score = lambda t: len(re.sub(r"\s", "", t))
    return [a if score(a) >= score(b) else b for a, b in zip(side, layer)]


# ---------------------------------------------------------------- extractors

TITLE_KEY = re.compile(r"\b(report|well\s+test|drill\s*stem|D\.?S\.?T\.?|production\s+test|evaluation)\b",
                       re.IGNORECASE)
TITLE_BREAK = re.compile(r"prepared|submitted|presented|confidential|sdn|bhd|inc\b|ltd|page\s+\d",
                         re.IGNORECASE)


def extract_title(pages):
    for pi, page in enumerate(pages[:3]):
        ls = lines_of(page)
        for i, line in enumerate(ls):
            if not (8 <= len(line) <= 120 and letters_ratio(line) >= 0.55 and TITLE_KEY.search(line)):
                continue
            if TITLE_BREAK.search(line) or DATE_RE.search(line):
                continue
            # Grow the block with adjacent title-like lines (same paragraph, mostly upper case).
            start = i
            while start > 0 and ls[start - 1] and len(ls[start - 1]) <= 80 and upper_ratio(ls[start - 1]) > 0.8 \
                    and letters_ratio(ls[start - 1]) >= 0.4 and not TITLE_BREAK.search(ls[start - 1]) \
                    and not DATE_RE.search(ls[start - 1]) and i - start < 3:
                start -= 1
            end = i
            while end + 1 < len(ls) and ls[end + 1] and len(ls[end + 1]) <= 80 and upper_ratio(ls[end + 1]) > 0.8 \
                    and letters_ratio(ls[end + 1]) >= 0.3 and not TITLE_BREAK.search(ls[end + 1]) \
                    and not DATE_RE.search(ls[end + 1]) and end - i < 3:
                end += 1
            title = " ".join(ls[start:end + 1])
            return title, f"page {pi + 1}, lines {start + 1}-{end + 1} (keyword block)", pi
    return NOT_FOUND, "no title-like line with report/test keywords on pages 1-3", None


def company_names_in(text):
    names = []
    for m in COMPANY_SUFFIX_RE.finditer(text):
        n = clean(m.group(1))
        if 3 <= len(n) <= 80 and letters_ratio(n) > 0.6:
            names.append(n)
    for m in KNOWN_COMPANY_RE.finditer(text):
        names.append(clean(m.group(1)))
    return names


def extract_companies(pages):
    roles = OrderedDict()   # (role) -> [names]
    others = OrderedDict()  # name.lower() -> (name, page)
    for pi, page in enumerate(pages[:FRONT_PAGES]):
        ls = lines_of(page)
        for i, line in enumerate(ls):
            m = ROLE_RE.search(line)
            if m:
                role = f"{m.group(1).title()} {m.group(2).title()}"
                rest = m.group(3)
                if not rest:
                    rest = next((l for l in ls[i + 1:i + 3] if l), "")
                found = company_names_in(rest) or ([rest] if rest and letters_ratio(rest) > 0.6 and len(rest) <= 80 else [])
                for n in found:
                    roles.setdefault(role, [])
                    if n.lower() not in [x.lower() for x in roles[role]]:
                        roles[role].append(f"{n}")
        for n in company_names_in(page):
            others.setdefault(n.lower(), (n, pi + 1))

    role_names = {n.lower() for v in roles.values() for n in v}
    # Drop names that are a substring of a longer name already kept (e.g. "Esso" vs "Esso Production Malaysia Inc").
    plain = [v for k, v in others.items() if k not in role_names]
    plain = [v for v in plain if not any(v[0].lower() != o[0].lower() and v[0].lower() in o[0].lower() for o in plain)]

    parts = [f"{role}: {', '.join(names)}" for role, names in roles.items()]
    if plain:
        parts.append(", ".join(n for n, _ in plain))
    if not parts:
        return NOT_FOUND, f"no company name pattern on pages 1-{FRONT_PAGES}"
    note = "; ".join([f"role from explicit '{r}' label" for r in roles] +
                     [f"'{n}' on page {p} (no role stated)" for n, p in plain])
    return "; ".join(parts), note


def extract_date(pages, title_page):
    # 1) Explicitly labelled report/issue date on the front pages.
    for pi, page in enumerate(pages[:FRONT_PAGES]):
        ls = lines_of(page)
        for i, line in enumerate(ls):
            lm = DATE_LABEL_RE.search(line)
            if not lm or DATE_EXCLUDE_RE.search(line):
                continue
            tail = line[lm.end():]
            dm = DATE_RE.search(tail) or (DATE_RE.search(ls[i + 1]) if i + 1 < len(ls) and not tail.strip() else None)
            if dm:
                return clean(dm.group(0)), f"page {pi + 1}: labelled '{clean(lm.group(1))}'"
    # 2) Unlabelled date(s) on the cover / title page (cover = pages 1..title page).
    last = title_page if title_page is not None else 0
    found = OrderedDict()
    for pi in range(last + 1):
        for line in lines_of(pages[pi]):
            if DATE_EXCLUDE_RE.search(line):
                continue
            for m in DATE_RE.finditer(line):
                found.setdefault(clean(m.group(0)), pi + 1)
    where = f"pages 1-{last + 1}" if last else "page 1"
    if len(found) == 1:
        d, pg = next(iter(found.items()))
        return f"{d} ({REVIEW}: unlabelled date on cover/title page)", f"page {pg}: only date on {where}, no label"
    if len(found) > 1:
        return (f"{REVIEW}: multiple unlabelled dates ({', '.join(found)})",
                f"{where}: " + ", ".join(f"p.{p}" for p in found.values()) + "; none labelled as report date")
    return NOT_FOUND, f"no labelled date on pages 1-{FRONT_PAGES}; no date on {where}"


FIELD_RE = re.compile(r"\b([A-Za-z][A-Za-z\-]{2,})\s+(?:(?:oil|gas)\s+)?field\b", re.IGNORECASE)
FIELD_LABEL_RE = re.compile(r"\bfield(?:\s+name)?\s*:\s*([A-Za-z][A-Za-z\-]{2,}(?:\s+[A-Za-z][A-Za-z\-]{2,})?)",
                            re.IGNORECASE)


def extract_fields(pages):
    counts, first_page, labelled = Counter(), {}, set()
    for pi, page in enumerate(pages):
        for m in FIELD_RE.finditer(page):
            w = m.group(1)
            if not w[0].isupper():
                continue
            n = w.upper()
            if n in FIELD_STOP:
                continue
            counts[n] += 1
            first_page.setdefault(n, pi + 1)
        for m in FIELD_LABEL_RE.finditer(page):
            n = clean(m.group(1)).upper()
            if n.split()[0] in FIELD_STOP:
                continue
            labelled.add(n)
            counts[n] += 1
            first_page.setdefault(n, pi + 1)
    keep = [n for n, c in counts.most_common() if c >= 2 or n in labelled]
    if not keep:
        return NOT_FOUND, "no '<Name> Field' or 'Field: <Name>' pattern", []
    note = ", ".join(f"{n} (x{counts[n]}, first p.{first_page[n]})" for n in keep)
    return ", ".join(keep), note, keep


WELL_HYPHEN_RE = re.compile(r"\b([A-Z][A-Za-z]{2,})\s?[-–]\s?(\d{1,2}[A-Z]{0,2}(?:ST\d?)?)\b")
WELL_SPACE_RE = re.compile(r"\b([A-Z][A-Za-z]{2,})\s(\d{1,2}[A-Z]{0,2})\b")
WELL_LABEL_RE = re.compile(r"\bwell(?:\s+(?:name|no\.?))?\s*:\s*([A-Za-z]{3,}\s?[-–]?\s?\d{1,2}[A-Za-z]{0,2})\b",
                           re.IGNORECASE)


def norm_well(prefix, num):
    return f"{prefix.upper()}-{num.upper()}"


def extract_wells(pages, fields):
    field_set = set(fields)
    counts, first_page, labelled = Counter(), {}, set()

    def add(name, pi):
        counts[name] += 1
        first_page.setdefault(name, pi + 1)

    for pi, page in enumerate(pages):
        for m in WELL_HYPHEN_RE.finditer(page):
            p = m.group(1)
            if p.upper() in WELL_PREFIX_STOP or not (p.isupper() or p.upper() in field_set):
                continue
            add(norm_well(p, m.group(2)), pi)
        for m in WELL_SPACE_RE.finditer(page):
            if m.group(1).upper() in field_set:    # space-separated only when prefix is a known field
                add(norm_well(m.group(1), m.group(2)), pi)
        for m in WELL_LABEL_RE.finditer(page):
            mm = re.match(r"([A-Za-z]{3,})\s?[-–]?\s?(\d{1,2}[A-Za-z]{0,2})", m.group(1))
            if mm and mm.group(1).upper() not in WELL_PREFIX_STOP:
                n = norm_well(mm.group(1), mm.group(2))
                labelled.add(n)
                add(n, pi)

    # Keep wells whose prefix is a known field name, explicitly labelled, or repeated (>=3) elsewhere.
    keep = [n for n, c in counts.most_common()
            if n.split("-")[0] in field_set or n in labelled or c >= 3]
    if not keep:
        return NOT_FOUND, "no well-name pattern found"
    keep.sort(key=lambda n: (n.split("-")[0], int(re.match(r"\d+", n.split("-")[1]).group()), n))
    note = ", ".join(f"{n} (x{counts[n]}, first p.{first_page[n]})" for n in keep)
    return ", ".join(keep), note


# ---------------------------------------------------------------- main

COLUMNS = ["File Name", "Report Title", "Company Name", "Report Date",
           "Total Number of Report Pages", "Field Name", "Well Name"]


def main():
    sources = sorted(p for p in SRC_ROOT.rglob("*") if p.is_file() and p.suffix.lower() == ".pdf")
    wb = Workbook()
    ws = wb.active
    ws.title = "Metadata"
    ws.append(COLUMNS)
    notes = wb.create_sheet("Extraction_Notes")
    notes.append(["File Name", "Relative Path", "Field", "Evidence / Rule"])

    for src in sources:
        rel = src.relative_to(SRC_ROOT)
        txt = TXT_DIR / rel.parent / f"{rel.stem}.txt"
        n_pages = page_count(src)
        ocr_pdf = PDF_DIR / rel.parent / f"{rel.stem}_OCR.pdf"
        pages = load_pages(txt, ocr_pdf, n_pages)

        if pages:
            title, t_note, title_page = extract_title(pages)
            company, c_note = extract_companies(pages)
            date, d_note = extract_date(pages, title_page)
            fields, f_note, field_list = extract_fields(pages)
            wells, w_note = extract_wells(pages, field_list)
        else:
            title = company = date = fields = wells = f"{REVIEW}: OCR text missing"
            t_note = c_note = d_note = f_note = w_note = f"missing {txt}"

        row = [src.name, title, company, date, n_pages if n_pages is not None else REVIEW, fields, wells]
        ws.append(row)
        for col, note in zip(COLUMNS[1:], [t_note, c_note, d_note,
                                           f"qpdf page count of original PDF", f_note, w_note]):
            notes.append([src.name, str(rel), col, note])

        status = []
        for col, val in zip(COLUMNS[1:], row[1:]):
            s = str(val)
            status.append(f"{col}: {'NOT FOUND' if s == NOT_FOUND else 'REVIEW' if REVIEW in s else 'found'}")
        tp = "n/a" if not pages or title_page is None else title_page + 1
        print(f"{src.name}  (text pages: {len(pages)}, title page: {tp})\n  " + "\n  ".join(status))

    # Formatting
    for sheet, widths in ((ws, [55, 50, 60, 30, 14, 25, 50]), (notes, [55, 55, 28, 100])):
        for i, w in enumerate(widths):
            sheet.column_dimensions[chr(65 + i)].width = w
        for cell in sheet[1]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill("solid", fgColor="DDEBF7")
        for r in sheet.iter_rows(min_row=1):
            for cell in r:
                cell.alignment = Alignment(wrap_text=True, vertical="top")
        sheet.freeze_panes = "A2"

    wb.save(XLSX)
    print(f"\nSaved: {XLSX}")


if __name__ == "__main__":
    main()
