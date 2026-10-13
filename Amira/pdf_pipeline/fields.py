"""Deterministic, rule-based field extraction from page text.

Only evidence found in the document text is used for the report fields. Embedded
PDF metadata is reported separately (PDF Author / PDF Creation Date) and never used
to fill a document field. Missing values stay empty; nothing is guessed.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date

from dateutil import parser as dateparser

from .config import Rules
from .extract import PageText

FIELD_NAMES = [
    "Report Title",
    "Company on Cover",
    "Created By (Contractor)",
    "Created For (Operator)",
    "Report Date",
    "Field Name(s)",
    "Well Name(s)",
]


@dataclass
class Evidence:
    field: str
    value: str
    pages: list[int]
    source: str  # "text" / "ocr" / "pdf_metadata"
    rule: str
    snippet: str = ""


@dataclass
class Extraction:
    values: dict[str, list[str]] = field(default_factory=lambda: {f: [] for f in FIELD_NAMES})
    evidence: list[Evidence] = field(default_factory=list)

    def add(self, name: str, value: str, page: PageText, rule: str, snippet: str = "") -> None:
        """Add a value, de-duplicating case-insensitively; extra pages are appended to evidence."""
        value = _clean(value)
        if not value:
            return
        key = value.casefold()
        for ev in self.evidence:
            if ev.field == name and ev.value.casefold() == key:
                if page.number not in ev.pages:
                    ev.pages.append(page.number)
                return
        self.values[name].append(value)
        self.evidence.append(Evidence(name, value, [page.number], page.source, rule, _clean(snippet)[:160]))


def _clean(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip().strip(".,;:-–").strip()


def _lines(text: str) -> list[str]:
    return [ln.strip() for ln in text.splitlines() if ln.strip()]


def _snippet(text: str, start: int, end: int, width: int = 50) -> str:
    return text[max(0, start - width): end + width]


# ---------------------------------------------------------------- dates

_MONTHS = {
    "jan": 1, "january": 1, "januari": 1,
    "feb": 2, "february": 2, "februari": 2,
    "mar": 3, "march": 3, "mac": 3,
    "apr": 4, "april": 4,
    "may": 5, "mei": 5,
    "jun": 6, "june": 6,
    "jul": 7, "july": 7, "julai": 7,
    "aug": 8, "august": 8, "ogos": 8,
    "sep": 9, "sept": 9, "september": 9,
    "oct": 10, "october": 10, "oktober": 10,
    "nov": 11, "november": 11,
    "dec": 12, "december": 12, "disember": 12,
}
_MON = "|".join(sorted(_MONTHS, key=len, reverse=True))
_DATE_PATTERNS = [
    ("day month year", re.compile(rf"\b(\d{{1,2}})(?:st|nd|rd|th)?[\s.\-]+({_MON})\.?[\s.\-,]+(\d{{4}})\b", re.I)),
    ("month day year", re.compile(rf"\b({_MON})\.?\s+(\d{{1,2}})(?:st|nd|rd|th)?,?\s+(\d{{4}})\b", re.I)),
    ("iso", re.compile(r"\b(\d{4})-(\d{1,2})-(\d{1,2})\b")),
    ("numeric", re.compile(r"\b(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{4})\b")),
    ("month year", re.compile(rf"\b({_MON})\.?,?\s+(\d{{4}})\b", re.I)),
]


def _valid(y: int, m: int, d: int | None = None) -> bool:
    if not 1900 <= y <= 2100:
        return False
    try:
        date(y, m, d or 1)
        return True
    except ValueError:
        return False


def find_dates(text: str, dayfirst: bool = True) -> list[tuple[str, str, int, int]]:
    """Return (iso_value, rule, start, end) for each date, in text order.

    Month-only dates are returned as YYYY-MM (the day is not invented). Numeric dates
    with only 4-digit years are accepted; dayfirst decides 03/04/1997 ordering.
    """
    found, taken = [], []
    for rule, rx in _DATE_PATTERNS:
        for m in rx.finditer(text):
            if any(m.start() < e and s < m.end() for s, e in taken):
                continue
            g = m.groups()
            iso = None
            if rule == "day month year":
                d, mo, y = int(g[0]), _MONTHS[g[1].lower()], int(g[2])
                iso = f"{y:04d}-{mo:02d}-{d:02d}" if _valid(y, mo, d) else None
            elif rule == "month day year":
                mo, d, y = _MONTHS[g[0].lower()], int(g[1]), int(g[2])
                iso = f"{y:04d}-{mo:02d}-{d:02d}" if _valid(y, mo, d) else None
            elif rule == "iso":
                y, mo, d = map(int, g)
                iso = f"{y:04d}-{mo:02d}-{d:02d}" if _valid(y, mo, d) else None
            elif rule == "numeric":
                try:
                    dt = dateparser.parse(m.group(0), dayfirst=dayfirst)
                    iso = dt.strftime("%Y-%m-%d") if _valid(dt.year, dt.month, dt.day) else None
                except (ValueError, OverflowError):
                    iso = None
                rule = "numeric (day first)" if dayfirst else "numeric (month first)"
            elif rule == "month year":
                mo, y = _MONTHS[g[0].lower()], int(g[1])
                iso = f"{y:04d}-{mo:02d}" if _valid(y, mo) else None
            if iso:
                found.append((iso, rule, m.start(), m.end()))
                taken.append((m.start(), m.end()))
    return sorted(found, key=lambda x: x[2])


# ---------------------------------------------------------------- label helpers

def _label_values(page: PageText, labels: list[re.Pattern], all_labels: list[re.Pattern]):
    """Yield (value, line) for 'Label: value' lines, or the next line when the label stands alone."""
    lines = _lines(page.text)
    for i, line in enumerate(lines):
        for rx in labels:
            m = rx.match(line)
            if not m:
                continue
            value = _clean(m.group("value") or "")
            if not value and i + 1 < len(lines) and not any(r.match(lines[i + 1]) for r in all_labels):
                value = _clean(lines[i + 1])
            if value:
                yield value[:120], line
            break


def _strip_label(line: str, labels: list[re.Pattern]) -> str:
    for rx in labels:
        m = rx.match(line)
        if m:
            return m.group("value") or ""
    return line


# ---------------------------------------------------------------- individual fields

def _is_title_line(line: str, rules: Rules) -> bool:
    return (
        4 <= len(line) <= 120
        and sum(c.isalpha() for c in line) >= 0.6 * len(line.replace(" ", ""))
        and not any(rx.match(line) for rx in rules.all_labels)
        and not find_dates(line)
        and not any(rx.search(line) for rx in rules.company_res)
    )


def find_title(cover: list[PageText], rules: Rules, ex: Extraction) -> None:
    for page in cover:
        lines = _lines(page.text)
        for i, line in enumerate(lines[:40]):
            if not (rules.title_keyword_re.search(line) and _is_title_line(line, rules)):
                continue
            block = [line]
            j = i - 1
            while j >= 0 and len(block) < 4 and lines[j].isupper() and _is_title_line(lines[j], rules):
                block.insert(0, lines[j])
                j -= 1
            j = i + 1
            while j < len(lines) and len(block) < 4 and lines[j].isupper() and _is_title_line(lines[j], rules):
                block.append(lines[j])
                j += 1
            ex.add("Report Title", " ".join(block), page, "cover line with title keyword", " / ".join(block))
            return


def find_companies(cover: list[PageText], rules: Rules, ex: Extraction) -> None:
    page = cover[0] if cover else None
    if page is None:
        return
    found: list[tuple[str, str]] = []
    for line in _lines(page.text):
        text = _strip_label(line, rules.all_labels)
        for rx in rules.company_res:
            for m in rx.finditer(text):
                found.append((_clean(m.group(0)), line))
    # Drop names that are part of a longer name found on the cover (PETRONAS vs PETRONAS Carigali Sdn Bhd).
    names = [n for n, _ in found]
    for name, line in found:
        if name and not any(name.casefold() != o.casefold() and name.casefold() in o.casefold() for o in names):
            ex.add("Company on Cover", name, page, "company dictionary / suffix on cover page", line)


def find_parties(cover: list[PageText], rules: Rules, ex: Extraction) -> None:
    for page in cover:
        for value, line in _label_values(page, rules.contractor_labels, rules.all_labels):
            ex.add("Created By (Contractor)", value, page, "contractor label", line)
        for value, line in _label_values(page, rules.operator_labels, rules.all_labels):
            ex.add("Created For (Operator)", value, page, "operator label", line)


def find_report_date(cover: list[PageText], rules: Rules, ex: Extraction) -> None:
    for page in cover:
        for value, line in _label_values(page, rules.date_labels, rules.all_labels):
            dates = find_dates(value, rules.date_dayfirst)
            if dates:
                ex.add("Report Date", dates[0][0], page, f"date label ({dates[0][1]})", line)
                return
    for page in cover:
        dates = find_dates(page.text, rules.date_dayfirst)
        if dates:
            iso, rule, s, e = dates[0]
            ex.add("Report Date", iso, page, f"first date on cover pages ({rule})", _snippet(page.text, s, e))
            return


_FIELD_SUFFIX_RE = re.compile(r"\b([A-Z][A-Za-z]+)\s+(?:Field|FIELD|Fields|FIELDS)\b")
_FIELD_LABEL_RE = re.compile(r"\bField(?:\s+Name)?\s*:\s*([A-Z][A-Za-z]+(?:[ -][A-Z][A-Za-z]+)?)")


def find_fields(pages: list[PageText], rules: Rules, ex: Extraction) -> None:
    for page in pages:
        if rules.known_field_re:
            for m in rules.known_field_re.finditer(page.text):
                canonical = next(f for f in rules.known_fields if f.casefold() == m.group(1).casefold())
                ex.add("Field Name(s)", canonical, page, "known field dictionary", _snippet(page.text, m.start(), m.end()))
        for rx, rule in ((_FIELD_LABEL_RE, "'Field:' label"), (_FIELD_SUFFIX_RE, "'<Name> Field' pattern")):
            for m in rx.finditer(page.text):
                name = m.group(1)
                if any(w.lower() in rules.field_stopwords for w in re.split(r"[ -]", name)):
                    continue
                ex.add("Field Name(s)", name, page, rule, _snippet(page.text, m.start(), m.end()))


def _trim_well_label(value: str) -> str:
    """'Dulang A-12 Spud date 1998' -> 'Dulang A-12' (stop at the first token with a digit).

    Returns '' when no token has a digit, since well names always carry a number.
    """
    kept = []
    for token in value.split():
        if token[0].islower():
            break
        kept.append(token)
        if any(c.isdigit() for c in token):
            return " ".join(kept)
    return ""


def _is_excluded_well(name: str, rules: Rules) -> bool:
    return (
        not name
        or name.split()[0].lower() in rules.well_stopwords
        or any(rx.search(name) for rx in rules.well_exclude_res)
    )


def find_wells(pages: list[PageText], rules: Rules, ex: Extraction) -> None:
    for page in pages:
        for rx in rules.well_res:
            for m in rx.finditer(page.text):
                name = _clean(m.group(1) if rx.groups else m.group(0))
                if not _is_excluded_well(name, rules):
                    ex.add("Well Name(s)", name, page, "well name pattern", _snippet(page.text, m.start(), m.end()))
        for m in rules.well_label_re.finditer(page.text):
            name = _trim_well_label(_clean(m.group(1)))
            if not _is_excluded_well(name, rules):
                ex.add("Well Name(s)", name, page, "'Well:' label", _snippet(page.text, m.start(), m.end()))


def extract_fields(pages: list[PageText], rules: Rules) -> Extraction:
    ex = Extraction()
    cover = [p for p in pages[: rules.cover_pages] if p.text.strip()]
    find_title(cover, rules, ex)
    find_companies(cover, rules, ex)
    find_parties(cover, rules, ex)
    find_report_date(cover, rules, ex)
    find_fields(pages, rules, ex)
    find_wells(pages, rules, ex)
    return ex
