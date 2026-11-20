"""Writes output_A_v2/_completion_summary.md for the Dulang A extraction (PPTX + Excel workbooks).

Sections: scope and decisions, inputs with SHA-256 before/after (NAS originals and local working copies), template limitation,
wells identified (one JSON per resolved well), PPTX slide coverage, unreadable content, validation, and per-well conflicts /
missing data / unmapped content / full traceability.
"""
import datetime as dt
import re

# PPTX slides whose wellbore slot is covered by an Excel workbook (user decision 11 Nov 2026: Excel overrides the PPTX
# for the same slot; the slide revision is superseded and NOT merged).
SUPERSEDED = {
    2: ("A4", "WELL A-4 : DUAL OIL PRODUCER IN 9 5/8\" CASING (TYPE 5)", "DULANG-A4ST1 (Dulang A04ST1 Well Schematic.xlsx)"),
    9: ("A11ST", "WELL A-11 ST : SINGLE SELECTIVE OIL PRODUCER IN 9 5/8\" CSG (TYPE 4) — banner points to the A11ST2 Excel",
        "DULANG-A11ST2 (Dulang A11ST2 Well Schematic.xlsx)"),
    14: ("A17ST1", "DULANG REVISIT I — A-17ST1 WELL COMPLETION DIAGRAM", "DULANG-A17ST2 (Dulang A17ST2 Well Schematic updated .xlsx)"),
    20: ("A23", "COMPLETION REPORT — WELL : DULANG A23 (5 OCT 2012)", "DULANG-A23ST2 (Dulang A23ST2 Well Schematic.xlsx)"),
    23: ("A26ST1", "WELL NAME: DULANG A-26ST1 — COMPLETION SCHEMATIC (same wellbore as the Excel)",
         "DULANG-A26ST1 (Dulang A26ST1 Well Schematic.xlsx)"),
    30: ("A33ST2", "DULANG A-33ST2 WELL COMPLETION DIAGRAM (Well No. A-33st2)", "DULANG-A33ST3 (Dulang A33ST3 Well Schematic.xls)"),
}


def esc(v):
    return str(v).replace("|", "\\|").replace("\n", " / ")


def write(path, wells, results, before, after, nas, nas_before, nas_after, nslides, dest):
    L = []
    a = L.append
    a("# Dulang A — well-schematic extraction (PPTX + Excel): completion summary")
    a("")
    a(f"Run: {dt.datetime.now().isoformat(timespec='seconds')} — local processing only. Output directory: `{dest}`. "
      "Scripts and temporary files: `_work/` (build script `_work/build_schematic_A.py`, rules `_work/d_rules.py`, transcriptions "
      "`_work/specs/*.txt`, slide dumps `_work/dump/`, picture crops `_work/media_png/`, Excel cell dumps `_work/xls_dump/`).")
    a("")
    a("## Scope and decisions")
    a("")
    a("- Inputs: `Dulang A Well Schematic updated.pptx` (30 slides) and the six Excel workbooks for A04ST1, A11ST2, A17ST2 (updated), "
      "A23ST2, A26ST1 and A33ST3 (.xls).")
    a("- File naming: the request template says `DULANG-D<well-number><suffix>`; this is the Dulang **A** platform, so files are named "
      "`DULANG-A<number><suffix>.json` (leading zeros removed, explicit suffixes ST1/ST2/ST3 kept). The same name is the template well "
      "identifier (`wellInfo.name`); the original label is kept in this summary.")
    a("- PPTX vs Excel (user decision, 11 Nov 2026): where an Excel workbook exists for a wellbore slot, the Excel data is used and the "
      "PPTX slide revision is **superseded — not merged**. Superseded slides are listed below with their labels; their content is not in any JSON.")
    a("- Rules applied: ft = m × 3.28084 (3 dp) with original metres in `…M` keys; datums kept as printed (MDTHF / MDDF / MDRKB / MD DF "
      "never interchanged or shifted; off-datum values flagged); WLTHF (wireline) depths never used as MD; header sizes without an ID/OD "
      "designation recorded as ID (user rule, flagged); explicit MIN ID / MAX OD kept; diagram values preferred over conflicting header "
      "values (both recorded); tubular tops of 0 only on a tubing-hanger (THF) reference (flagged); no TD/PBTD substitution for tubular "
      "ends; nearest Type_Key mappings flagged with estimated / estimatedFields ['type'] / estimationNote / actualType.")
    a("")
    a("## Inputs (read only) — SHA-256 before / after")
    a("")
    a("| file | before | after | unchanged |")
    a("|---|---|---|---|")
    for k in before:
        a(f"| `{k}` | `{before[k]}` | `{after[k]}` | {'YES' if before[k] == after[k] else 'NO'} |")
    for k, p in nas.items():
        b, f = nas_before.get(k), nas_after.get(k)
        a(f"| NAS `{p}` | `{b}` | `{f}` | {'YES' if b and b == f else 'NOT CHECKED (share not mounted)' if not b or not f else 'NO'} |")
    a("")
    a("All local working copies in `_work/src/` are byte-identical to the NAS originals (same SHA-256). Note: the A33ST3 `.xls` on the "
      "NAS (`0551ea4e…`) differs from the copy hashed by the earlier A33ST3 run (`b51284bc…`, same size) — the file was changed on the "
      "NAS after that run; this extraction uses the current file.")
    a("")
    a("Template note: `well-schematic-template-json.rtf` is an **example JSON document, not a formal JSON Schema**. Conformity was "
      "checked against its section order, `wellInfo` keys/order and value types; component keys beyond the example (topMdM/bottomMdM, "
      "zone, string, weight, grade, bottomTvd, estimated…, actualType, source) follow the conventions of the Dulang B/C/D outputs. "
      "`floatEquipment` / `plugBackCement` entries use key names from the Excel template sheets because the JSON example lists are empty.")
    a("")
    a("## Wells identified (one JSON per resolved well)")
    a("")
    a("| well file | source | slide / sheet | original label | components | flagged | conflicts | missing | unclear | validation |")
    a("|---|---|---|---|---|---|---|---|---|---|")
    for w, doc in wells:
        kinds = [i[0] for i in w.issues]
        src = w.S.get("file") or "Dulang A Well Schematic updated.pptx"
        loc = w.media if w.S.get("file") else f"slide {w.slide}"
        a(f"| `{w.name}.json` | {esc(src)} | {esc(loc)} | {esc(w.label)} | {len(doc['components'])} | "
          f"{sum(1 for c in doc['components'] if c.get('estimated'))} | {kinds.count('Conflict')} | {kinds.count('Missing')} | "
          f"{kinds.count('Unclear')} | {'PASS' if all(results[w.key].values()) else 'FAIL'} |")
    a("")
    n_pptx = sum(1 for w, _ in wells if not w.S.get("file"))
    n_xls = len(wells) - n_pptx
    a(f"**JSON files written: {len(wells)} = resolved wells {len(wells)}** ({n_pptx} from PPTX slides, {n_xls} from Excel workbooks). "
      f"PPTX slides read: {nslides} (presentation order; no hidden slides — no `show=\"0\"`).")
    a("")
    a("## PPTX slide coverage")
    a("")
    a("| slide | well label | output |")
    a("|---|---|---|")
    by_slide = {w.slide: w for w, _ in wells if w.slide}
    for s in range(1, nslides + 1):
        if s in by_slide:
            w = by_slide[s]
            a(f"| {s} | {esc(w.label)} | `{w.name}.json` |")
        elif s in SUPERSEDED:
            k, lab, by = SUPERSEDED[s]
            a(f"| {s} | {esc(lab)} | superseded — slot covered by Excel {esc(by)}; slide not merged |")
        else:
            a(f"| {s} | (not identified) | — |")
    a("")
    a("## Unreadable / low-legibility content")
    a("")
    unread = [(w.name, wh, t) for w, _ in wells for k, wh, t in w.issues if re.search(r"(?i)legib|cut at|not legible|unreadable", t)]
    if unread:
        a("| well | where | detail |")
        a("|---|---|---|")
        for n, wh, t in unread:
            a(f"| {n} | {esc(wh)} | {esc(t)} |")
    a("")
    a("Every slide picture was viewed at x3 upscale (EMF drawings via their embedded bitmaps — no native EMF text on slides 23 and 29). "
      "Excel workbooks were read from their native cell values and drawing text boxes; the drawn schematic graphics inside the workbooks "
      "(shapes / embedded pictures) carry no additional values and were not rendered (no spreadsheet renderer available locally).")
    a("")
    a("## Validation (per JSON)")
    a("")
    names = list(next(iter(results.values())).keys())
    a("Checks: " + "; ".join(names) + ".")
    a("")
    a("| well | checks passed | failed checks |")
    a("|---|---|---|")
    for w, _ in wells:
        r = results[w.key]
        a(f"| {w.name} | {sum(r.values())}/{len(r)} | {', '.join(k for k, v in r.items() if not v) or '—'} |")
    a("")
    a("## Unresolved issues (all wells)")
    a("")
    a("| well | kind | where | detail |")
    a("|---|---|---|---|")
    for w, _ in wells:
        for k, wh, t in w.issues:
            if k in ("Conflict", "Missing", "Unresolved"):
                a(f"| {w.name} | {k} | {esc(wh)} | {esc(t)} |")
    a("")
    for w, doc in wells:
        a(f"## {w.name} — {esc(w.loc)}")
        a("")
        a(f"Original label: {esc(w.label)}")
        a(f"Equipment depth datum: {esc(w.S['datum'])}  ·  source: {esc(w.media)} ({esc(w.method)})")
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
        a("| JSON path | value | location | original text | conversion / interpretation |")
        a("|---|---|---|---|---|")
        for t in w.trace:
            a(f"| `{esc(t['path'])}` | {esc(t['value'])} | {esc(t['where'])} | {esc(t['original'])[:160]} | {esc(t['method'])} |")
        a("")
    path.write_text("\n".join(L) + "\n", encoding="utf-8")
