"""Per-page routing: measure evidence about a page, then apply configurable rules.

A page is not considered "digital" just because it has a few characters: image coverage and
text coverage are compared too, so a scanned page with a small digital header still goes to OCR.
The decision, the rule that made it, and the evidence are all stored with the page.
"""

from __future__ import annotations

import json
import re
import unicodedata
from dataclasses import dataclass
from pathlib import Path

import pypdfium2.raw as pdfium_c

DEFAULT_RULES = Path(__file__).resolve().parent / "config" / "routing.json"
A4_AREA = 595.28 * 841.89
METRICS = {"chars", "text_area", "image_area", "image_count", "garbage_ratio", "rotation", "size_vs_a4"}
OPS = {"lt": lambda a, b: a < b, "le": lambda a, b: a <= b, "gt": lambda a, b: a > b, "ge": lambda a, b: a >= b}
CID_RE = re.compile(r"\(cid:\d+\)")


@dataclass
class Rule:
    name: str
    conditions: list[tuple[str, str, float]]
    method: str = ""

    def matches(self, ev: dict) -> bool:
        return all(OPS[op](ev[m], v) for m, op, v in self.conditions)


@dataclass
class RoutingRules:
    version: str
    default: str
    rules: list[Rule]
    flags: list[Rule]


def _parse_when(when: dict, where: str) -> list[tuple[str, str, float]]:
    out = []
    for key, value in when.items():
        metric, _, op = key.rpartition("_")
        if metric not in METRICS or op not in OPS:
            raise ValueError(f"{where}: unknown condition '{key}' (metrics: {sorted(METRICS)}; ops: lt/le/gt/ge)")
        out.append((metric, op, float(value)))
    if not out:
        raise ValueError(f"{where}: needs at least one condition")
    return out


def load_rules(path: Path = DEFAULT_RULES, known_methods: set[str] | None = None) -> RoutingRules:
    cfg = json.loads(Path(path).read_text(encoding="utf-8"))
    rules, flags = [], []
    for i, r in enumerate(cfg.get("rules", []), start=1):
        where = f"routing rules[{i}]"
        if known_methods is not None and r.get("method") not in known_methods:
            raise ValueError(f"{where}: unknown method '{r.get('method')}'")
        rules.append(Rule(r.get("name", where), _parse_when(r.get("when", {}), where), r["method"]))
    for i, r in enumerate(cfg.get("flags", []), start=1):
        where = f"routing flags[{i}]"
        flags.append(Rule(r.get("name", where), _parse_when(r.get("when", {}), where)))
    default = cfg.get("default", "textlayer")
    if known_methods is not None and default not in known_methods:
        raise ValueError(f"routing default: unknown method '{default}'")
    return RoutingRules(str(cfg.get("version", "1")), default, rules, flags)


def _garbage_ratio(text: str) -> float:
    visible = [c for c in text if not c.isspace()]
    if not visible:
        return 0.0
    bad = sum(1 for c in visible if c == "�" or unicodedata.category(c) in ("Co", "Cc", "Cs", "Cn"))
    bad += sum(len(m.group()) for m in CID_RE.finditer(text))
    return min(1.0, bad / len(visible))


def page_evidence(page) -> dict:
    """Cheap measurements from the PDF structure of one page (no rendering, no OCR)."""
    w, h = page.get_size()
    area = max(w * h, 1.0)
    tp = page.get_textpage()
    try:
        text = tp.get_text_range()
        n = tp.count_rects()
        text_area = 0.0
        for i in range(n):
            l, b, r, t = tp.get_rect(i)
            text_area += max(0.0, r - l) * max(0.0, t - b)
    finally:
        tp.close()
    image_area, image_count = 0.0, 0
    for obj in page.get_objects(filter=[pdfium_c.FPDF_PAGEOBJ_IMAGE], max_depth=3):
        l, b, r, t = obj.get_bounds()
        image_area += max(0.0, r - l) * max(0.0, t - b)
        image_count += 1
    return {
        "chars": sum(1 for c in text if not c.isspace()),
        "text_area": round(min(1.0, text_area / area), 4),
        "image_area": round(min(1.0, image_area / area), 4),
        "image_count": image_count,
        "garbage_ratio": round(_garbage_ratio(text), 4),
        "rotation": page.get_rotation(),
        "size_vs_a4": round(area / A4_AREA, 3),
        "width_pt": round(w, 1), "height_pt": round(h, 1),
    }


def decide(ev: dict, rules: RoutingRules) -> tuple[str, str, list[str]]:
    """Return (method, reason, flags)."""
    flags = [f.name for f in rules.flags if f.matches(ev)]
    for rule in rules.rules:
        if rule.matches(ev):
            return rule.method, f"rule: {rule.name}", flags
    return rules.default, "default (no rule matched)", flags
