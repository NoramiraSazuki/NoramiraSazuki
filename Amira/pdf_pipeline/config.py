"""Load the editable extraction dictionaries and compile them into regexes."""

from __future__ import annotations

import json
import re
from dataclasses import dataclass
from pathlib import Path

from . import WORKSPACE

DEFAULT_CONFIG = WORKSPACE / "config" / "dictionaries.json"


def _alternation(words: list[str]) -> str:
    # Longest first so "PETRONAS Carigali" wins over "PETRONAS".
    return "|".join(re.escape(w) for w in sorted(words, key=len, reverse=True))


def _label_regex(label: str) -> re.Pattern:
    """Match a label at the start of a line, e.g. 'Prepared by: X' or 'Prepared by' alone.

    Single-word labels ('For', 'Client') need a colon/dash or must stand alone on the
    line, so ordinary sentences like 'For example...' are not captured.
    """
    esc = re.escape(label).replace(r"\ ", r"\s+")
    if " " in label:
        pattern = rf"^\s*{esc}(?=\s|[:\-–]|$)\s*[:\-–]?\s*(?P<value>.*)$"
    else:
        pattern = rf"^\s*{esc}\s*(?:[:\-–]\s*(?P<value>.*))?$"
    return re.compile(pattern, re.IGNORECASE)


@dataclass
class Rules:
    cover_pages: int
    date_dayfirst: bool
    title_keyword_re: re.Pattern
    company_res: list[re.Pattern]
    contractor_labels: list[re.Pattern]
    operator_labels: list[re.Pattern]
    date_labels: list[re.Pattern]
    known_fields: list[str]
    known_field_re: re.Pattern | None
    field_stopwords: set[str]
    well_res: list[re.Pattern]
    well_label_re: re.Pattern
    well_stopwords: set[str]
    well_exclude_res: list[re.Pattern]

    @property
    def all_labels(self) -> list[re.Pattern]:
        return self.contractor_labels + self.operator_labels + self.date_labels


def load_rules(path: Path = DEFAULT_CONFIG) -> Rules:
    with open(path, encoding="utf-8") as fh:
        cfg = json.load(fh)

    company_res = [re.compile(cfg["company_suffix_regex"])]
    if cfg.get("known_companies"):
        company_res.insert(
            0, re.compile(rf"\b(?:{_alternation(cfg['known_companies'])})\b", re.IGNORECASE)
        )

    well_res = []
    if cfg.get("well_prefixes"):
        # Prefix is case-insensitive; the rest is case-sensitive so ordinary words
        # ("Dulang is 12") don't match. E.g. Dulang-1, DULANG A-12, DL-B5, Dulang-3ST1
        prefixes = _alternation(cfg["well_prefixes"])
        well_res.append(
            re.compile(rf"\b(?i:{prefixes})(?:[- ]?[A-Z]{{1,2}})?[- ]?\d{{1,3}}[A-Z]?(?:[- ]?ST\d{{0,2}})?\b")
        )
    well_res += [re.compile(p) for p in cfg.get("extra_well_regexes", [])]

    known_fields = cfg.get("known_fields", [])
    return Rules(
        cover_pages=int(cfg.get("cover_pages", 3)),
        date_dayfirst=bool(cfg.get("date_dayfirst", True)),
        title_keyword_re=re.compile(rf"\b(?:{_alternation(cfg['title_keywords'])})\b", re.IGNORECASE),
        company_res=company_res,
        contractor_labels=[_label_regex(x) for x in cfg["contractor_labels"]],
        operator_labels=[_label_regex(x) for x in cfg["operator_labels"]],
        date_labels=[_label_regex(x) for x in cfg["date_labels"]],
        known_fields=known_fields,
        known_field_re=(
            re.compile(rf"\b({_alternation(known_fields)})\b", re.IGNORECASE) if known_fields else None
        ),
        field_stopwords={w.lower() for w in cfg.get("field_stopwords", [])},
        well_res=well_res,
        well_label_re=re.compile(cfg["well_label_regex"]),
        well_stopwords={w.lower() for w in cfg.get("well_stopwords", [])},
        well_exclude_res=[re.compile(p) for p in cfg.get("well_exclude_regexes", [])],
    )
