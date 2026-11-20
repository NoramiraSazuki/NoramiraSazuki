"""Writes output_C/_completion_summary.md: inputs + hashes, well list, validation, issues, unmapped content, full traceability."""
import datetime as dt
import json
import re


def esc(v):
    return str(v).replace("|", "\\|").replace("\n", " / ")


def write(path, wells, results, before, after, covered, nslides, nas, nas_sha, nas_state):
    L = []
    a = L.append
    a("# Dulang D — well-schematic extraction: completion summary")
    a("")
    a(f"Run: {dt.datetime.now().isoformat(timespec='seconds')} (local processing only). Script: `output_D/_work/build_schematic_D.py`.")
    a("")
    a("## Inputs (read only) — SHA-256 before / after")
    a("")
    a("| file | before | after | unchanged |")
    a("|---|---|---|---|")
    for k in before:
        a(f"| `{k}` | `{before[k]}` | `{after[k]}` | {'YES' if before[k] == after[k] else 'NO'} |")
    a("")
    a(f"NAS original `{nas}`: SHA-256 at the start of processing `{nas_sha}` (identical to the local working copy above). "
      + (f"Re-hashed at the end: `{nas_state}` — {'UNCHANGED' if nas_state == nas_sha else 'CHANGED'}." if nas_state else
         "At the end of processing the NAS share was no longer mounted, so the end-of-run hash of the NAS file could not be taken; "
         "the file was only ever read (copied once to `output_D/_work/src/`). Re-check with `shasum -a 256` once the share is mounted."))
    a("")
    a("Template note: `well-schematic-template-json.rtf` is an **example JSON document, not a formal JSON Schema**. Conformity was checked "
      "against its section order, `wellInfo` keys/order and value types; component keys beyond the example (topMdM/bottomMdM, zone, string, "
      "weight, grade, bottomTvd, estimated…, actualType, source) follow the conventions of the earlier Dulang A/B outputs. "
      "`floatEquipment` / `plugBackCement` entries use key names from the Excel template sheets (Float_Equipment, Plug_Back_Cement) because the "
      "JSON example lists are empty.")
    a("")
    a("## Wells identified (one JSON per resolved well)")
    a("")
    a("| well file | original label | slide(s) | schematic source | components | flagged | conflicts | missing | unclear | validation |")
    a("|---|---|---|---|---|---|---|---|---|---|")
    for w, doc in wells:
        kinds = [i[0] for i in w.issues]
        src = w.method.split(" (")[0] + f" — {w.media}"
        a(f"| `{w.name}.json` | {esc(w.label)} | {','.join(map(str, w.S['slides']))} | {esc(src)} | {len(doc['components'])} | {sum(1 for c in doc['components'] if c.get('estimated'))} | "
          f"{kinds.count('Conflict')} | {kinds.count('Missing')} | {kinds.count('Unclear')} | {'PASS' if all(results[w.key].values()) else 'FAIL'} |")
    a("")
    a(f"JSON files written: **{len(wells)}** = resolved wells **{len(wells)}**. Slides read: {nslides} (presentation order, no hidden slides "
      f"— no `show=\"0\"`); slides covered by a well: {covered}. Pictures placed outside the visible slide area (slides 4, 7, 16, 19, 23, 31) "
      "were inspected and are listed per well as unmapped hidden content. Combined slides: D28 = slides 28 (schematic) + 29 (completion report, "
      "different datum — notes only); D32 = slides 33 (schematic) + 34 (completion report — notes only); D30 slide 31 holds a second, off-slide "
      "schematic (notes only).")
    a("")
    a("Extraction method: all Dulang D schematics are pictures. EMF drawings (slides 2, 10, 12, 13, 20) were read from their native text "
      "records; screenshots / scans were transcribed visually from x3-upscaled crops with tesseract OCR as a cross-check. Transcriptions with "
      "per-row locations are in `output_D/_work/specs/*.txt`; low-legibility reads are marked '?' there and flagged in the JSON.")
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
        a(f"Equipment depth datum: {w.S['datum']}  ·  source: {esc(w.media)} ({esc(w.method)})")
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
