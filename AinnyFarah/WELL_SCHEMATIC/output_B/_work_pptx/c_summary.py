"""Writes output_C/_completion_summary.md: inputs + hashes, well list, validation, issues, unmapped content, full traceability."""
import datetime as dt
import json
import re


def esc(v):
    return str(v).replace("|", "\\|").replace("\n", " / ")


def write(path, wells, results, before, after, dumps):
    L = []
    a = L.append
    a("# Dulang C — well-schematic extraction: completion summary")
    a("")
    a(f"Run: {dt.datetime.now().isoformat(timespec='seconds')} (local processing only). Script: `output_C/_work/build_schematic_C.py`.")
    a("")
    a("## Inputs (read only) — SHA-256 before / after")
    a("")
    a("| file | before | after | unchanged |")
    a("|---|---|---|---|")
    for k in before:
        a(f"| `{k}` | `{before[k]}` | `{after[k]}` | {'YES' if before[k] == after[k] else 'NO'} |")
    a("")
    a("Template note: `well-schematic-template-json.rtf` is an **example JSON document, not a formal JSON Schema**. Conformity was checked "
      "against its section order, `wellInfo` keys/order and value types; component keys beyond the example (topMdM/bottomMdM, zone, string, "
      "weight, grade, bottomTvd, estimated…, actualType, source) follow the conventions of the earlier Dulang A/B outputs. "
      "`floatEquipment` / `plugBackCement` entries use key names from the Excel template sheets (Float_Equipment, Plug_Back_Cement) because the "
      "JSON example lists are empty.")
    a("")
    a("## Wells identified (one JSON per resolved well)")
    a("")
    a("| well file | original label | slide | schematic source | components | flagged | conflicts | missing | unclear | validation |")
    a("|---|---|---|---|---|---|---|---|---|---|")
    for w, doc in wells:
        kinds = [i[0] for i in w.issues]
        src = "embedded image" if w.slide in (3, 16, 22, 25, 29, 31) else ("title only" if w.slide == 17 else "native slide shapes/text")
        a(f"| `{w.name}.json` | {esc(w.label)} | {w.slide} | {src} | {len(doc['components'])} | {sum(1 for c in doc['components'] if c.get('estimated'))} | "
          f"{kinds.count('Conflict')} | {kinds.count('Missing')} | {kinds.count('Unclear')} | {'PASS' if all(results[w.key].values()) else 'FAIL'} |")
    a("")
    a(f"JSON files written: **{len(wells)}** = resolved wells **{len(wells)}**. Slide 30 ('C30 doesn’t exist') -> no well, no file. "
      "No hidden slides (no `show=\"0\"`); 31 slides read in presentation order; notes pages contain only 'Open' + slide number.")
    a("")
    a("## Validation checks (per JSON)")
    a("")
    names = list(next(iter(results.values())).keys())
    a("| check | " + " | ".join(w.key for w, _ in wells) + " |")
    a("|---|" + "---|" * len(wells))
    for n in names:
        a(f"| {n} | " + " | ".join("✓" if results[w.key][n] else "✗" for w, _ in wells) + " |")
    a("")
    for w, doc in wells:
        a(f"## {w.name} — slide {w.slide}")
        a("")
        a(f"Original label: {esc(w.label)}")
        if getattr(w, "datum", None):
            a(f"Equipment depth datum: {w.datum}")
        a("")
        if w.issues:
            a("### Conflicts / missing / unclear")
            a("")
            a("| kind | where | detail |")
            a("|---|---|---|")
            for k, wh, t in w.issues:
                a(f"| {k} | {esc(wh)} | {esc(t)} |")
            a("")
        if w.unmapped:
            a("### Unmapped content (kept here, not in the JSON)")
            a("")
            a("| where | text | reason |")
            a("|---|---|---|")
            for wh, t, r in w.unmapped:
                a(f"| {esc(wh)} | {esc(t)[:300]} | {esc(r)} |")
            a("")
        a("### Traceability (every populated value)")
        a("")
        a("| JSON path | value | slide / object / location | original text | conversion / interpretation |")
        a("|---|---|---|---|---|")
        for t in w.trace:
            a(f"| `{esc(t['path'])}` | {esc(t['value'])} | {esc(t['where'])} | {esc(t['original'])[:160]} | {esc(t['method'])} |")
        a("")
    path.write_text("\n".join(L) + "\n", encoding="utf-8")
