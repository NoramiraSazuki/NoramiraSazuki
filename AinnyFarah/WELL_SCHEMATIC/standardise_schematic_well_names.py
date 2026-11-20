#!/usr/bin/env python3
"""Standardise wellInfo.name in the well schematic JSON files (DULANG-<platform><number><suffix>).

Rule: DULANG- prefix, leading zeros dropped from the well number and numeric suffixes, suffix letters kept
(Dulang_A12ST1 -> DULANG-A12ST1, Dulang_A11ST -> DULANG-A11ST). Owner decisions: 'DULANG A004S01 OP4' -> DULANG-A4S1
(OP tag dropped), 'DLA-01' -> DULANG-A1. Only the name value is edited in the text, so the rest of each file stays
byte-identical. _previous_runs/ folders are left alone. Originals are copied to backup_before_well_name_update/.
"""
import json
import re
import shutil
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parent
BACKUP = BASE / "backup_before_well_name_update"
NAME_RE = re.compile(r"(?:DULANG[_ ]?)?([A-D])(\d+)((?:[A-Z]+\d*)*)(?: OP\d+)?", re.I)


def standard(name):
    if m := re.fullmatch(r"DLA-(\d+)", name):
        return f"DULANG-A{int(m.group(1))}"
    m = NAME_RE.fullmatch(name.strip())
    if not m:
        return None
    sfx = re.sub(r"([A-Z]+)(\d*)", lambda s: s.group(1) + (str(int(s.group(2))) if s.group(2) else ""),
                 m.group(3).upper())
    return f"DULANG-{m.group(1).upper()}{int(m.group(2))}{sfx}"


def main():
    if BACKUP.exists():
        sys.exit(f"STOP: {BACKUP} already exists - not overwritten")
    todo, unresolved = [], []
    for p in sorted(BASE.rglob("*.json")):
        if "_previous_runs" in p.parts or BACKUP in p.parents:
            continue
        try:
            d = json.loads(p.read_text())
            old = d["wellInfo"]["name"]
        except (ValueError, KeyError, TypeError):
            continue
        new = standard(old)
        if new is None:
            unresolved.append((p, old))
        elif new != old:
            todo.append((p, old, new))
    for p, old, new in todo:
        dst = BACKUP / p.relative_to(BASE)
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(p, dst)
    for p, old, new in todo:
        text = p.read_text()
        pat = re.compile(r'("name"\s*:\s*")' + re.escape(json.dumps(old)[1:-1]) + '"')
        assert len(pat.findall(text)) == 1, (p, old)
        out = pat.sub(lambda m: m.group(1) + new + '"', text)
        a, b = json.loads(text), json.loads(out)
        a["wellInfo"]["name"] = new
        assert a == b, p
        tmp = p.with_name(p.name + ".partial")
        tmp.write_text(out)
        tmp.replace(p)
        print(f"{p.relative_to(BASE)}: {old!r} -> {new!r}")
    print(f"updated {len(todo)} files; unresolved {len(unresolved)}: {[(str(p.relative_to(BASE)), o) for p, o in unresolved]}")


if __name__ == "__main__":
    main()
