"""Hide bibliography/reference sections and citation lines from document-level extraction.

Text inside a references section describes *other* publications (their depths, models,
fields, dates), so it must not fill attributes of the current report. Removed lines are
replaced by empty lines, so page numbers and line structure stay intact.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from .config import Rules
from .extract import PageText

_YEAR = re.compile(r"\b(?:1[89]|20)\d{2}[a-z]?\b")
_AUTHOR_START = [
    re.compile(r"^\s*\[\d{1,3}\]\s+\S"),                                     # [12] Smith ...
    re.compile(r"^\s*[A-Z][A-Za-zÀ-ÿ'’\-]+,\s*(?:[A-Z]\.\s*-?\s*){1,4}"),    # Smith, J. A.
    re.compile(r"^\s*[A-Z][A-Za-zÀ-ÿ'’\-]+\s+(?:[A-Z]\.){1,3}\s*[,(]"),       # Smith J.A., / Smith J. (
]
_DOI = re.compile(r"\bdoi(?:\.org)?\s*[:/]", re.IGNORECASE)
_VENUE = re.compile(
    r"\b(?:J\.|Journal|Geophys|Proc\.|Proceedings|Bull\.|Bulletin|Vol\.|pp\.|Press|Verlag|Eds?\.|In:"
    r"|Trans\.|Rev\.|Lett\.|Conf\.|Symp\.|Publ\.)", re.IGNORECASE)


def is_citation_line(line: str) -> bool:
    """A line that reads like a bibliography entry, e.g. 'Smith, J., 1997. Title. J. Geophys. 12, 1-9.'"""
    if not line.strip():
        return False
    if _DOI.search(line):
        return True
    has_year = bool(_YEAR.search(line))
    if has_year and any(p.search(line) for p in _AUTHOR_START):
        return True
    return has_year and "," in line and bool(_VENUE.search(line))


@dataclass
class ExcludedSection:
    heading: str
    pages: list[int] = field(default_factory=list)
    lines: int = 0


@dataclass
class FilterResult:
    pages: list[PageText]
    sections: list[ExcludedSection]
    citation_lines: dict[int, int]  # page -> citation lines removed outside reference sections


def _followed_by_citations(lines: list[tuple[int, str]], start: int, needed: int = 2, window: int = 12) -> bool:
    """True when the next non-empty lines contain citation entries (rules out table-of-contents hits)."""
    seen = hits = 0
    for _, line in lines[start + 1:]:
        if not line.strip():
            continue
        seen += 1
        hits += is_citation_line(line)
        if hits >= needed:
            return True
        if seen >= window:
            break
    return False


def filter_references(pages: list[PageText], rules: Rules) -> FilterResult:
    flat = [(i, line) for i, p in enumerate(pages) for line in p.text.split("\n")]
    kept: list[list[str]] = [[] for _ in pages]
    sections: list[ExcludedSection] = []
    citations: dict[int, int] = {}
    current: ExcludedSection | None = None

    for pos, (i, line) in enumerate(flat):
        number = pages[i].number
        if current is not None:
            if rules.section_end_re.match(line):
                current = None
                kept[i].append(line)
            else:
                kept[i].append("")
                current.lines += bool(line.strip())
                if number not in current.pages:
                    current.pages.append(number)
            continue
        if rules.reference_heading_re.match(line) and _followed_by_citations(flat, pos):
            current = ExcludedSection(line.strip(), [number])
            sections.append(current)
            kept[i].append("")
            continue
        if is_citation_line(line):
            citations[number] = citations.get(number, 0) + 1
            kept[i].append("")
            continue
        kept[i].append(line)

    filtered = [PageText(p.number, "\n".join(kept[i]), p.source) for i, p in enumerate(pages)]
    return FilterResult(filtered, sections, citations)
