"""Load the editable extraction dictionaries and compile them into regexes."""

from __future__ import annotations

import json
import re
from dataclasses import dataclass
from pathlib import Path

from . import WORKSPACE
from .schema import CONFIGURABLE_COLUMNS, NEVER_FILLED, RESULT_COLUMNS

PACKAGE_DIR = Path(__file__).resolve().parent

# Looked up in order: inside the package (pdf_pipeline/config/), then beside it (config/).
CONFIG_CANDIDATES = [
    PACKAGE_DIR / "config" / "dictionaries.json",
    WORKSPACE / "config" / "dictionaries.json",
]
DEFAULT_CONFIG = next((p for p in CONFIG_CANDIDATES if p.is_file()), CONFIG_CANDIDATES[0])


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


def _anywhere_label_regex(labels: list[str]) -> re.Pattern:
    """'Label: value' anywhere in a line (log headers often have several labels per line)."""
    return re.compile(rf"\b(?P<label>{_alternation(labels)})\b[ \t]*[:=#][ \t]*(?P<value>[^\n]+)", re.IGNORECASE)


def _heading_regex(headings: list[str], designator: bool = False) -> re.Pattern:
    """A line holding only a heading, optionally numbered: 'References', '7. REFERENCES', 'VI. Bibliography:'.

    With ``designator`` the heading may carry a letter/number and caption: 'Appendix A', 'APPENDIX 2: Core Data'.
    """
    suffix = r"(?:[ \t]+(?:[A-Z]|\d{1,2}|[IVX]{1,4})\b[^\n]{0,80})?" if designator else ""
    return re.compile(
        rf"^[ \t]*(?:\d+(?:\.\d+)*\.?[ \t]+|[IVX]+\.[ \t]+)?(?:{_alternation(headings)}){suffix}[ \t]*:?[ \t]*$",
        re.IGNORECASE | re.MULTILINE,
    )


def _depth_regexes(cfg: dict) -> tuple[re.Pattern, re.Pattern, re.Pattern]:
    units = _alternation([u for names in cfg["depth_units"].values() for u in names])
    refs = _alternation(cfg["depth_references"])
    sp = r"[ \t]*"  # never cross a line break

    def num(n):
        return rf"(?P<{n}>\d{{1,3}}(?:,\d{{3}})+(?:\.\d+)?|\d+(?:\.\d+)?)"

    def unit(n):
        return rf"(?:(?P<{n}>{units})\b\.?)?"

    def ref(n):
        return rf"(?:\(?(?P<{n}>{refs})\)?\b)?"

    paren = r"(?:\((?P<paren>[^)\n]{1,20})\))?"
    single = rf"{sp}{paren}{sp}[:=\-–]?{sp}{num('v')}{sp}{unit('u')}{sp}{ref('r')}"
    interval = (rf"{sp}{paren}{sp}[:=]?{sp}{num('v1')}{sp}{unit('u1')}{sp}{ref('r1')}"
                rf"{sp}(?:-|–|to){sp}{num('v2')}{sp}{unit('u2')}{sp}{ref('r2')}")
    flags = re.IGNORECASE
    return (
        re.compile(rf"\b(?P<label>{_alternation(cfg['top_depth_labels'])})\b{single}", flags),
        re.compile(rf"\b(?P<label>{_alternation(cfg['base_depth_labels'])})\b{single}", flags),
        re.compile(rf"\b(?P<label>{_alternation(cfg['interval_labels'])})\b{interval}", flags),
    )


@dataclass
class ClassificationRule:
    name: str
    filename_re: re.Pattern | None
    keyword_re: re.Pattern | None
    values: dict[str, str]


def _check_columns(columns, where: str) -> None:
    for col in columns:
        if col in NEVER_FILLED:
            raise ValueError(f"{where}: '{col}' must never be filled (secrets are not exported)")
        if col not in CONFIGURABLE_COLUMNS:
            raise ValueError(f"{where}: '{col}' is not a configurable column of the schema")


def _classification_rules(cfg: dict) -> list[ClassificationRule]:
    rules = []
    for i, r in enumerate(cfg.get("classification_rules", []), start=1):
        where = f"classification_rules[{i}]"
        if not r.get("set"):
            raise ValueError(f"{where}: needs a non-empty 'set'")
        if not (r.get("filename_regex") or r.get("text_keywords")):
            raise ValueError(f"{where}: needs filename_regex and/or text_keywords")
        _check_columns(r["set"], where)
        rules.append(ClassificationRule(
            name=r.get("name", where),
            filename_re=re.compile(r["filename_regex"]) if r.get("filename_regex") else None,
            keyword_re=(re.compile(rf"\b(?:{_alternation(r['text_keywords'])})\b", re.IGNORECASE)
                        if r.get("text_keywords") else None),
            values={k: str(v) for k, v in r["set"].items()},
        ))
    return rules


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
    typed_date_res: dict[str, re.Pattern]
    depth_unit_map: dict[str, str]
    top_depth_re: re.Pattern
    base_depth_re: re.Pattern
    interval_re: re.Pattern
    depth_ref_re: re.Pattern
    labelled_value_res: dict[str, re.Pattern]
    value_stop_re: re.Pattern
    log_run_re: re.Pattern
    scale_re: re.Pattern
    required_columns: list[str]
    classification_rules: list[ClassificationRule]
    constant_columns: dict[str, str]
    reference_heading_re: re.Pattern
    section_end_re: re.Pattern
    publication_date_re: re.Pattern
    create_date_fallback: bool
    field_context_stopwords: set[str]
    line_start_labels: set[str]
    depth_max: dict[str, float]
    junk_title_re: re.Pattern

    @property
    def all_labels(self) -> list[re.Pattern]:
        return self.contractor_labels + self.operator_labels + self.date_labels


def load_rules(path: Path = DEFAULT_CONFIG) -> Rules:
    if not Path(path).is_file():
        tried = "\n  ".join(str(p) for p in ([Path(path)] + [c for c in CONFIG_CANDIDATES if c != Path(path)]))
        raise FileNotFoundError(f"Config file not found. Looked in:\n  {tried}\nPass one with --config.")
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
        # ("Dulang is 12") don't match. E.g. Dulang-1, DULANG A-12, DL-B5, Dulang-3ST1.
        # The final lookahead rejects pieces of longer identifiers such as UWI "DL-000-B5-0001".
        prefixes = _alternation(cfg["well_prefixes"])
        well_res.append(re.compile(
            rf"\b(?i:{prefixes})(?:[- ]?[A-Z]{{1,2}})?[- ]?\d{{1,3}}[A-Z]?(?:[- ]?ST\d{{0,2}})?\b(?![-/][A-Za-z0-9])"
        ))
    well_res += [re.compile(p) for p in cfg.get("extra_well_regexes", [])]

    known_fields = cfg.get("known_fields", [])

    typed_dates = cfg.get("typed_date_labels", {})
    labelled = cfg.get("labelled_values", {})
    _check_columns(list(typed_dates) + list(labelled), "typed_date_labels/labelled_values")
    # Any known label ends a captured value: "UWI: X   Spud Date: Y" -> UWI = "X"
    stop_labels = [lb for v in list(typed_dates.values()) + list(labelled.values()) for lb in v]
    stop_labels += cfg["top_depth_labels"] + cfg["base_depth_labels"] + cfg["interval_labels"]
    stop_labels += ["Well Name", "Well", "Field", "Field Name", "Run No", "Run Number", "Scale", "Date"]

    required = cfg.get("required_columns", [])
    unknown = [c for c in required if c not in RESULT_COLUMNS]
    if unknown:
        raise ValueError(f"required_columns: unknown column(s) {unknown}")
    constants = {k: str(v) for k, v in cfg.get("constant_columns", {}).items()}
    _check_columns(constants, "constant_columns")

    top_re, base_re, interval_re = _depth_regexes(cfg)
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
        typed_date_res={
            col: re.compile(
                rf"\b(?P<label>{_alternation(labels)})\b[ \t]*(?:\([^)\n]{{0,15}}\))?[ \t]*(?:[:=\-–]|on\b)?",
                re.IGNORECASE)
            for col, labels in typed_dates.items()
        },
        depth_unit_map={u.lower(): canon for canon, names in cfg["depth_units"].items() for u in names},
        top_depth_re=top_re,
        base_depth_re=base_re,
        interval_re=interval_re,
        depth_ref_re=re.compile(rf"\b({_alternation(cfg['depth_references'])})\b", re.IGNORECASE),
        labelled_value_res={col: _anywhere_label_regex(labels) for col, labels in labelled.items()},
        value_stop_re=re.compile(rf"\s{{3,}}|\b(?:{_alternation(stop_labels)})\b[ \t]*[:=#]", re.IGNORECASE),
        log_run_re=re.compile(cfg["log_run_regex"]),
        scale_re=re.compile(cfg["scale_regex"]),
        required_columns=required,
        classification_rules=_classification_rules(cfg),
        constant_columns=constants,
        reference_heading_re=_heading_regex(cfg.get("reference_headings", ["References", "Bibliography"])),
        section_end_re=_heading_regex(cfg.get("section_end_headings", ["Appendix"]), designator=True),
        # (?<!\w)/(?!\w) instead of \b so non-word labels such as "©" also work.
        publication_date_re=re.compile(
            rf"(?<!\w)(?P<label>{_alternation(cfg.get('publication_date_labels', []) or ['Published'])})(?!\w)"
            r"[ \t]*(?:[:=\-–]|on\b|in\b)?", re.IGNORECASE),
        create_date_fallback=bool(cfg.get("create_date_fallback", False)),
        field_context_stopwords={w.lower() for w in cfg.get("field_context_stopwords", [])},
        line_start_labels={w.casefold() for w in cfg.get("line_start_labels", [])},
        depth_max={k: float(v) for k, v in cfg.get("depth_max", {}).items()},
        junk_title_re=re.compile(cfg.get("junk_metadata_title_regex", r"(?i)^untitled$")),
    )
