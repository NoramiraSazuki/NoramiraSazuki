"""Deterministic, rule-based attribute extraction from page text.

Only evidence found in the document text is used for document attributes. Embedded
PDF metadata is reported in evidence.csv only and never used to fill a document
attribute. Missing values stay empty; nothing is guessed.
"""

from __future__ import annotations

import re
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date

from dateutil import parser as dateparser

from .config import Rules
from .extract import PageText
from .schema import CANDIDATE, CONFIG, CONFIRMED, INFERRED, LABELLED, MULTIPLE, PATTERN, SINGLE_VALUED


@dataclass
class Evidence:
    attribute: str
    value: str
    pages: list[int]
    source: str  # "text" / "ocr" / "pdf_metadata" / "filesystem" / "pdf_structure" / "config"
    method: str
    validation: str
    snippet: str = ""


@dataclass
class Depth:
    value: str
    uom: str
    reference: str
    evidence: list[Evidence] = field(default_factory=list, compare=False)


@dataclass
class Extraction:
    values: dict[str, list[str]] = field(default_factory=lambda: defaultdict(list))
    depths: dict[str, list[Depth]] = field(default_factory=lambda: {"Top Depth": [], "Base Depth": []})
    evidence: list[Evidence] = field(default_factory=list)
    review: list[str] = field(default_factory=list)  # notes for EXTRACTION STATUS, e.g. "Title (candidates differ)"

    def add(self, attribute: str, value: str, page: PageText | None, method: str, validation: str,
            snippet: str = "", source: str | None = None, record: bool = True) -> bool:
        """Add a value, de-duplicating case-insensitively; repeat sightings add their page to evidence.

        ``record=False`` writes evidence only (attribute not exported as a column value).
        Returns True when the value is new.
        """
        value = _clean(value)
        if not value:
            return False
        key = value.casefold()
        for ev in self.evidence:
            if record and ev.validation.startswith(CANDIDATE):
                continue  # a reliable sighting is exported even if the value was a candidate earlier
            if ev.attribute == attribute and ev.value.casefold() == key:
                if page is not None and page.number not in ev.pages:
                    ev.pages.append(page.number)
                return False
        if record:
            self.values[attribute].append(value)
        self.evidence.append(Evidence(
            attribute, value, [page.number] if page else [], source or (page.source if page else ""),
            method, validation, _clean(snippet)[:160],
        ))
        return True

    def add_depth(self, attribute: str, depth: Depth, page: PageText, method: str, snippet: str) -> None:
        """Record a depth; the same value+unit seen again only adds its page to the evidence."""
        for existing in self.depths[attribute]:
            if (existing.value, existing.uom) == (depth.value, depth.uom):
                for ev in existing.evidence:
                    if page.number not in ev.pages:
                        ev.pages.append(page.number)
                return
        self.depths[attribute].append(depth)
        for attr, val in ((attribute, depth.value), (f"{attribute} UOM", depth.uom),
                          (f"{attribute} Reference", depth.reference)):
            if val:
                ev = Evidence(attr, val, [page.number], page.source, method, LABELLED, _clean(snippet)[:160])
                depth.evidence.append(ev)
                self.evidence.append(ev)

    def get(self, attribute: str) -> list[str]:
        return self.values.get(attribute, [])

    def candidate(self, attribute: str, value: str, page: PageText | None, reason: str,
                  snippet: str = "", source: str | None = None) -> None:
        """Evidence of an uncertain value that is deliberately NOT exported."""
        self.add(attribute, value, page, reason, CANDIDATE, snippet, source=source, record=False)

    def flag_multiple(self) -> list[str]:
        """Mark single-valued attributes that received several distinct values; return review notes."""
        flagged = []
        for attr in sorted(SINGLE_VALUED):
            count = len(self.depths[attr]) if attr in self.depths else len(self.get(attr))
            if count > 1:
                flagged.append(f"{attr} (multiple values)")
                for ev in self.evidence:
                    if ev.validation.startswith(CANDIDATE):
                        continue
                    if ev.attribute == attr or (attr in self.depths and ev.attribute.startswith(attr)):
                        if MULTIPLE not in ev.validation:
                            ev.validation += f"; {MULTIPLE}"
        return flagged


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


def _label_ok(text: str, m: re.Match, rules: Rules) -> bool:
    """Accept a 'Label: value' match only when it really is a label.

    - the label is written with a capital letter ('Spud Date', 'SPUD DATE'), not mid-sentence 'spud date';
    - generic words in line_start_labels ('Model') must start their line.
    """
    label = m.group("label")
    first = next((c for c in label if c.isalpha()), "")
    if first and first.islower():
        return False
    if label.casefold() in rules.line_start_labels:
        line_start = text.rfind("\n", 0, m.start()) + 1
        if text[line_start: m.start()].strip():
            return False
    return True


def _cut_value(value: str, rules: Rules) -> str:
    """End a captured value at the next known label or a wide gap: 'X   Spud Date: Y' -> 'X'."""
    m = rules.value_stop_re.search(value)
    return _clean(value[: m.start()] if m else value)[:100]


# ---------------------------------------------------------------- report-level attributes

def _is_title_line(line: str, rules: Rules) -> bool:
    return (
        4 <= len(line) <= 120
        and sum(c.isalpha() for c in line) >= 0.6 * len(line.replace(" ", ""))
        and not any(rx.match(line) for rx in rules.all_labels)
        and not find_dates(line)
        and not any(rx.search(line) for rx in rules.company_res)
    )


_INITIALS_RE = re.compile(r"\b[A-Z]\.\s")
_YEAR_RE = re.compile(r"\b(?:19|20)\d{2}\b")
_SMALL_WORDS = {"a", "an", "and", "at", "by", "for", "from", "in", "of", "on", "or", "the", "to", "with", "dan", "di"}


def _continues_title(line: str, upper_mode: bool, rules: Rules) -> bool:
    """Can this neighbouring line belong to the same title block?"""
    if not line or not _is_title_line(line, rules):
        return False
    if upper_mode:
        return line.isupper()
    # Title-case block: mostly capitalised words, no author initials, no year, not a sentence.
    words = [w for w in re.findall(r"[A-Za-z][\w'-]*", line) if w.lower() not in _SMALL_WORDS]
    capitalised = sum(w[0].isupper() for w in words)
    return (bool(words) and capitalised >= 0.6 * len(words) and not _INITIALS_RE.search(line)
            and not _YEAR_RE.search(line) and not line.endswith("."))


def find_title_candidate(cover: list[PageText], rules: Rules) -> tuple[str, PageText] | None:
    """The cover title block: a line with a title keyword plus adjacent lines in the same style."""
    for page in cover:
        raw = [ln.strip() for ln in page.text.splitlines()]
        for i, line in enumerate(raw[:60]):
            if not (line and rules.title_keyword_re.search(line) and _is_title_line(line, rules)):
                continue
            upper = line.isupper()
            block = [line]
            j = i - 1
            while j >= 0 and len(block) < 4 and _continues_title(raw[j], upper, rules):
                block.insert(0, raw[j])
                j -= 1
            j = i + 1
            while j < len(raw) and len(block) < 4 and _continues_title(raw[j], upper, rules):
                block.append(raw[j])
                j += 1
            return " ".join(block), page
    return None


def _norm(text: str) -> str:
    return re.sub(r"[^0-9a-z]+", " ", text.casefold()).strip()


def find_title(cover: list[PageText], rules: Rules, ex: Extraction, metadata_title: str = "") -> None:
    """Choose the Title from the cover block and the PDF metadata title.

    The metadata title is used only when the same text is printed in the document. Disagreeing
    candidates are kept in evidence.csv and the row is flagged for review; nothing is guessed.
    """
    found = find_title_candidate(cover, rules)
    meta = _clean(metadata_title)
    meta_usable = len(meta) >= 8 and not rules.junk_title_re.match(meta)
    meta_page = next((p for p in cover if meta_usable and _norm(meta) in _norm(p.text)), None)

    if found:
        title, page = found
        if meta_page and _norm(title) == _norm(meta):
            ex.add("Title", title, page, "cover title block; matches PDF metadata title", CONFIRMED, title)
        elif meta_page and _norm(title) in _norm(meta):
            ex.add("Title", meta, meta_page, "PDF metadata title, printed in full in the document "
                   "(cover block was a shorter part of it)", CONFIRMED, meta)
            ex.candidate("Title", title, page, "cover title block (incomplete)", title)
        else:
            ex.add("Title", title, page, "cover line with title keyword", INFERRED, title)
            if meta_usable:
                ex.candidate("Title", meta, meta_page, "PDF metadata title differs from cover title",
                             source="pdf_metadata" if not meta_page else None)
                ex.review.append("Title (document and metadata titles differ)")
    elif meta_page:
        ex.add("Title", meta, meta_page, "PDF metadata title, printed in the document", CONFIRMED, meta)
    elif meta_usable:
        ex.candidate("Title", meta, None, "PDF metadata title not found in document text - not exported",
                     source="pdf_metadata")
        ex.review.append("Title (only an unconfirmed metadata title)")


def find_companies(cover: list[PageText], rules: Rules, ex: Extraction) -> None:
    """Company names on the cover: evidence only, to help validate Created By / Created For."""
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
            ex.add("Company on Cover", name, page, "company dictionary / suffix on cover page", PATTERN, line,
                   record=False)


def find_parties(cover: list[PageText], rules: Rules, ex: Extraction) -> None:
    """Created By = contractor/author labels; Created For = client/operator labels."""
    for page in cover:
        for value, line in _label_values(page, rules.contractor_labels, rules.all_labels):
            ex.add("Created By", value, page, "contractor/author label", LABELLED, line)
        for value, line in _label_values(page, rules.operator_labels, rules.all_labels):
            ex.add("Created For", value, page, "client/operator label", LABELLED, line)


def _typed_label_spans(text: str, rules: Rules) -> list[int]:
    return [m.end() for rx in rules.typed_date_res.values() for m in rx.finditer(text)]


def find_publication_dates(pages: list[PageText], rules: Rules, ex: Extraction) -> list[tuple[int, int]]:
    """Publication dates ('Received', 'Published', '©') -> evidence only. Returns their (page, start) spots."""
    spots = []
    for page in pages:
        for m in rules.publication_date_re.finditer(page.text):
            if not _label_ok(page.text, m, rules):
                continue
            segment = page.text[m.end(): m.end() + 40]
            dates = find_dates(segment, rules.date_dayfirst)
            if not dates:
                # '© 1997 Publisher' - a year only; recorded as text, never converted into a date
                year = re.match(r"\s*((?:19|20)\d{2})\b", segment)
                if year:
                    ex.add("Publication Date", year.group(1), page, f"'{m.group('label')}' (year only)",
                           LABELLED, _snippet(page.text, m.start(), m.end() + year.end()), record=False)
                continue
            if dates[0][2] <= 3:
                ex.add("Publication Date", dates[0][0], page,
                       f"'{m.group('label')}' label - publication date, not used for CREATE_DATE",
                       LABELLED, _snippet(page.text, m.start(), m.end() + dates[0][3]), record=False)
                spots.append((page.number, m.end() + dates[0][2]))
    return spots


def find_create_date(cover: list[PageText], rules: Rules, ex: Extraction) -> None:
    """CREATE_DATE = the document's own report/issue date, from a date label on the cover pages.

    Publication dates and the PDF creation date are different concepts and never used here.
    The unlabelled 'first date on the cover' fallback runs only if create_date_fallback is true.
    """
    for page in cover:
        for value, line in _label_values(page, rules.date_labels, rules.all_labels):
            if rules.publication_date_re.search(line):
                continue
            dates = find_dates(value, rules.date_dayfirst)
            if dates:
                ex.add("CREATE_DATE", dates[0][0], page, f"report date label ({dates[0][1]})", LABELLED, line)
                return
    publication = find_publication_dates(cover, rules, ex)
    for page in cover:
        typed_ends = _typed_label_spans(page.text, rules)
        for iso, rule, s, e in find_dates(page.text, rules.date_dayfirst):
            if any(end <= s <= end + 15 for end in typed_ends) or (page.number, s) in publication:
                continue
            snippet = _snippet(page.text, s, e)
            if rules.create_date_fallback:
                ex.add("CREATE_DATE", iso, page, f"first unlabelled date on cover pages ({rule})", INFERRED, snippet)
            else:
                ex.candidate("CREATE_DATE", iso, page,
                             f"unlabelled date on cover ({rule}) - not exported (create_date_fallback is off)",
                             snippet)
            return


def find_typed_dates(pages: list[PageText], rules: Rules, ex: Extraction) -> None:
    """Spud / Completion / Transmittal / Trip dates: only a date written right after its label."""
    for page in pages:
        for attr, rx in rules.typed_date_res.items():
            for m in rx.finditer(page.text):
                if not _label_ok(page.text, m, rules):
                    continue
                segment = page.text[m.end(): m.end() + 40]
                dates = find_dates(segment, rules.date_dayfirst)
                if dates and dates[0][2] <= 3:
                    ex.add(attr, dates[0][0], page, f"'{m.group(0).strip()}' label ({dates[0][1]})", LABELLED,
                           _snippet(page.text, m.start(), m.end() + dates[0][3]))


# ---------------------------------------------------------------- depths

def _unit_and_ref(rules: Rules, *parts: str | None) -> tuple[str, str]:
    """First recognised unit (mapped to m/ft) and depth reference among the given text parts."""
    unit = ref = ""
    for part in parts:
        if not part:
            continue
        if not unit:
            for token in re.findall(r"[A-Za-z]+", part):
                if token.lower() in rules.depth_unit_map:
                    unit = rules.depth_unit_map[token.lower()]
                    break
        if not ref:
            m = rules.depth_ref_re.search(part)
            if m:
                ref = m.group(1).upper()
    return unit, ref


def _number(raw: str) -> str:
    return raw.replace(",", "")  # thousands separators only; digits and decimals unchanged


def _depth_problem(depth: Depth, rules: Rules) -> str:
    """Why a labelled depth is too uncertain to export ('' when it is fine)."""
    if not depth.uom:
        return "no unit written"
    limit = rules.depth_max.get(depth.uom)
    if limit is not None and float(depth.value) > limit:
        return f"above plausible maximum ({limit:g} {depth.uom})"
    return ""


def _record_depth(attr: str, depth: Depth, page: PageText, method: str, snippet: str, rules: Rules,
                  ex: Extraction) -> None:
    problem = _depth_problem(depth, rules)
    if problem:
        shown = f"{depth.value} {depth.uom}".strip()
        ex.candidate(attr, shown, page, f"{method} - not exported: {problem}", snippet)
    else:
        ex.add_depth(attr, depth, page, method, snippet)


def find_depths(pages: list[PageText], rules: Rules, ex: Extraction) -> None:
    """Top/Base depth only from capitalised depth labels with a written unit (m/ft)."""
    for page in pages:
        text = page.text
        for attr, rx in (("Top Depth", rules.top_depth_re), ("Base Depth", rules.base_depth_re)):
            for m in rx.finditer(text):
                if not _label_ok(text, m, rules):
                    continue
                unit, ref = _unit_and_ref(rules, m.group("u"), m.group("r"), m.group("paren"))
                _record_depth(attr, Depth(_number(m.group("v")), unit, ref), page,
                              f"'{m.group('label')}' label", _snippet(text, m.start(), m.end()), rules, ex)
        for m in rules.interval_re.finditer(text):
            if not _label_ok(text, m, rules):
                continue
            paren = m.group("paren")
            # A unit/reference written once applies to both ends: "1200 - 2450 m MDRT"
            u1, r1 = _unit_and_ref(rules, m.group("u1"), m.group("r1"), paren, m.group("u2"), m.group("r2"))
            u2, r2 = _unit_and_ref(rules, m.group("u2"), m.group("r2"), paren, m.group("u1"), m.group("r1"))
            snippet = _snippet(text, m.start(), m.end())
            label = m.group("label")
            _record_depth("Top Depth", Depth(_number(m.group("v1")), u1, r1), page,
                          f"'{label}' label (top)", snippet, rules, ex)
            _record_depth("Base Depth", Depth(_number(m.group("v2")), u2, r2), page,
                          f"'{label}' label (base)", snippet, rules, ex)


# ---------------------------------------------------------------- labelled technical values

def find_labelled_values(pages: list[PageText], rules: Rules, ex: Extraction) -> None:
    for page in pages:
        for attr, rx in rules.labelled_value_res.items():
            for m in rx.finditer(page.text):
                if not _label_ok(page.text, m, rules):
                    continue
                value = _cut_value(m.group("value"), rules)
                if attr == "Unique Wellbore Identifier":
                    value = value.split()[0] if value else ""
                    if not any(c.isdigit() for c in value):
                        continue
                ex.add(attr, value, page, "label", LABELLED, _snippet(page.text, m.start(), m.end()))
        for m in rules.log_run_re.finditer(page.text):
            ex.add("REPORT_LOG_RUN", m.group(1), page, "run number label", LABELLED,
                   _snippet(page.text, m.start(), m.end()))
        for line in _lines(page.text):
            label = re.search(r"\bscale\b", line, re.IGNORECASE)
            if label:
                for m in rules.scale_re.finditer(line, label.end()):
                    ex.add("SCALE_RATIO", re.sub(r"\s+", "", m.group(0)), page, "scale label", LABELLED, line)


# ---------------------------------------------------------------- fields and wells

_FIELD_SUFFIX_RE = re.compile(r"\b([A-Z][A-Za-z'’]+)\s+(?:Field|FIELD|Fields|FIELDS)\b(?:[ \t]+([A-Za-z]+))?")
_FIELD_LABEL_RE = re.compile(r"^[ \t]*(?:Field|FIELD)(?:\s+(?:Name|NAME))?[ \t]*:[ \t]*"
                             r"([A-Z][A-Za-z]+(?:[ -][A-Z][A-Za-z]+)?)", re.MULTILINE)


def find_fields(pages: list[PageText], rules: Rules, ex: Extraction) -> None:
    """FIELD_NAME is exported only from known_fields or an explicit 'Field:' line.

    '<Name> Field' phrases ('KTB Field Laboratory', "Earth's Thermal Field") are never exported;
    plausible ones are listed as CANDIDATE evidence for review.
    """
    for page in pages:
        if rules.known_field_re:
            # A field word inside a well name ("Dulang-1") is not evidence of the field itself.
            well_spans = [m.span() for rx in rules.well_res for m in rx.finditer(page.text)]
            for m in rules.known_field_re.finditer(page.text):
                if any(s <= m.start() < e for s, e in well_spans):
                    continue
                canonical = next(f for f in rules.known_fields if f.casefold() == m.group(1).casefold())
                ex.add("FIELD_NAME", canonical, page, "known field dictionary", PATTERN,
                       _snippet(page.text, m.start(), m.end()))
        for m in _FIELD_LABEL_RE.finditer(page.text):
            name = m.group(1)
            if not any(w.lower() in rules.field_stopwords for w in re.split(r"[ -]", name)):
                ex.add("FIELD_NAME", name, page, "'Field:' line", LABELLED, _snippet(page.text, m.start(), m.end()))
        for m in _FIELD_SUFFIX_RE.finditer(page.text):
            name, following = m.group(1), (m.group(2) or "").lower()
            if name.lower() in rules.field_stopwords or following in rules.field_context_stopwords:
                continue
            if rules.known_field_re and rules.known_field_re.fullmatch(name):
                continue  # already exported via the dictionary
            ex.candidate("FIELD_NAME", name, page, "'<Name> Field' phrase - not exported (add to known_fields "
                         "if it is a field)", _snippet(page.text, m.start(), m.end()))


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
                    ex.add("Well Name", name, page, "well name pattern", PATTERN,
                           _snippet(page.text, m.start(), m.end()))
        for m in rules.well_label_re.finditer(page.text):
            name = _trim_well_label(_clean(m.group(1)))
            if not _is_excluded_well(name, rules):
                ex.add("Well Name", name, page, "'Well:' label", LABELLED, _snippet(page.text, m.start(), m.end()))


# ---------------------------------------------------------------- classification

def classify(file_name: str, cover: list[PageText], rules: Rules, ex: Extraction) -> None:
    """Apply the first matching classification rule from the config (fills empty attributes only)."""
    text = " ".join(ex.get("Title") + [p.text for p in cover])
    for rule in rules.classification_rules:
        if rule.filename_re and not rule.filename_re.search(file_name):
            continue
        if rule.keyword_re and not rule.keyword_re.search(text):
            continue
        for attr, value in rule.values.items():
            if not ex.get(attr):
                ex.add(attr, value, None, f"classification rule '{rule.name}'", CONFIG, source="config")
        return


def extract_fields(pages: list[PageText], rules: Rules, file_name: str = "", metadata_title: str = "") -> Extraction:
    """Run all extractors. ``pages`` should already have reference sections removed."""
    ex = Extraction()
    cover = [p for p in pages[: rules.cover_pages] if p.text.strip()]
    find_title(cover, rules, ex, metadata_title)
    find_companies(cover, rules, ex)
    find_parties(cover, rules, ex)
    find_create_date(cover, rules, ex)
    find_typed_dates(pages, rules, ex)
    find_depths(pages, rules, ex)
    find_labelled_values(pages, rules, ex)
    find_fields(pages, rules, ex)
    find_wells(pages, rules, ex)
    classify(file_name, cover, rules, ex)
    return ex
