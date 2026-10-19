"""Original page-number selections such as "1-20, 45, 100-102" (1-based, as printed by PDF viewers)."""

from __future__ import annotations

import re

_PART = re.compile(r"^\s*(\d+)\s*(?:-\s*(\d+)\s*)?$")


class PageSpecError(ValueError):
    pass


def parse_page_spec(spec: str, page_count: int | None = None) -> list[int]:
    spec = (spec or "").strip().lower()
    if spec in ("", "all"):
        if page_count is None:
            raise PageSpecError("'all' needs the page count")
        return list(range(1, page_count + 1))
    pages: set[int] = set()
    for part in spec.split(","):
        m = _PART.match(part)
        if not m:
            raise PageSpecError(f"not a page or range: '{part.strip()}'")
        a = int(m.group(1))
        b = int(m.group(2) or a)
        if a < 1 or b < a:
            raise PageSpecError(f"invalid range: '{part.strip()}'")
        if page_count is not None and b > page_count:
            raise PageSpecError(f"page {b} is beyond the last page ({page_count})")
        pages.update(range(a, b + 1))
    return sorted(pages)


def to_spec(pages: list[int]) -> str:
    """[1,2,3,7,9,10] -> '1-3,7,9-10' (same output as the CLI's OCR page spec)."""
    from pdf_pipeline.ocr import _page_spec

    return _page_spec(sorted(set(pages))) if pages else ""


def chunk(pages: list[int], size: int) -> list[list[int]]:
    return [pages[i:i + size] for i in range(0, len(pages), size)]
