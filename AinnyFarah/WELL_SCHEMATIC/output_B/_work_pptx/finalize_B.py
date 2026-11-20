#!/usr/bin/env python3
"""Publish the Dulang B PPTX extraction: copy stage_table/ (native-table slides, build_schematic_Bt.py) and stage_spec/
(pictures / irregular layout, build_schematic_Bs.py) JSONs into output_B/, re-validate every file, re-hash the inputs and
write one combined output_B/_completion_summary.md. Never overwrites an existing output."""
import datetime as dt
import hashlib
import json
import os
import re
import shutil
import subprocess
from pathlib import Path

W = Path(__file__).resolve().parent
OUT = W.parent
ROOT = OUT.parent
NAS = Path("/Volumes/homes/public/DulangOSS/20261016/DULANG_WELLS/Well_Schematic/1. Well Schematic/Dulang B/Dulang B Well Schematic updated .pptx")
NAS_SHA = "e390e7e13ec18fc7294cb0a8b7b0ad09caf07893527be518626876bbb9e3fa21"
INPUTS = [NAS, W / "src" / NAS.name, ROOT / "well-schematic-template-json.rtf", ROOT / "well-schematic-sample-template.xlsx"]
SLIDES = {"B1": "1", "B2": "2", "B3": "3", "B4": "4", "B5": "5", "B6": "6", "B7": "7", "B8": "8", "B9": "9", "B10": "10", "B11": "11",
          "B12": "12", "B13": "13", "B14": "14", "B15": "15", "B16ST3": "16", "B17ST3": "17 (+33 obsolete)", "B18": "18", "B19": "19",
          "B20": "20", "B21": "21", "B22": "22", "B23": "23", "B24ST2": "24", "B25ST4": "25", "B26": "26", "B27": "27", "B28ST2": "28",
          "B29": "29", "B30ST1": "30", "B31ST1": "31", "B32": "32"}


def sha(p):
    h = hashlib.sha256()
    fd = os.open(p, os.O_RDONLY)
    while (b := os.read(fd, 1 << 20)):
        h.update(b)
    try:
        os.close(fd)
    except OSError:
        pass
    return h.hexdigest()


def key_of(p):
    return p.stem.replace("Dulang_", "")


def main():
    before = {str(p): sha(p) for p in INPUTS}
    assert before[str(NAS)] == NAS_SHA and before[str(INPUTS[1])] == NAS_SHA, "PPTX hash differs from the hash taken before processing"
    files = sorted(list((W / "stage_table").glob("Dulang_B*.json")) + list((W / "stage_spec").glob("Dulang_B*.json")),
                   key=lambda p: int(re.match(r"B(\d+)", key_of(p)).group(1)))
    keys = [key_of(p) for p in files]
    assert len(keys) == len(set(keys)) == 32 and set(keys) == set(SLIDES), sorted(set(keys) ^ set(SLIDES))
    targets = [OUT / p.name for p in files] + [OUT / "_completion_summary.md"]
    clash = [str(t) for t in targets if t.exists()]
    if clash:
        raise SystemExit("outputs exist — not overwriting without confirmation:\n" + "\n".join(clash))
    tpl = json.loads(subprocess.run(["textutil", "-convert", "txt", "-stdout", str(INPUTS[2])], capture_output=True, text=True, check=True).stdout)
    rows = []
    for p in files:
        doc = json.loads(p.read_text())
        ok = list(doc) == list(tpl) and list(doc["wellInfo"]) == list(tpl["wellInfo"]) and doc["wellInfo"]["name"] == p.stem \
            and len({c["id"] for c in doc["components"]}) == len(doc["components"])
        assert ok, p
        rows.append((p, doc))
        shutil.copy2(p, OUT / p.name)
    after = {str(p): sha(p) for p in INPUTS}
    t_sum = (W / "stage_table" / "_completion_summary.md").read_text()
    s_sum = (W / "stage_spec" / "_completion_summary.md").read_text()

    def section(text, start, end=None):
        i = text.index(start)
        j = text.index(end, i) if end else len(text)
        return text[i:j]

    def wells_table(text):
        sec = section(text, "## Wells identified", "## Validation checks")
        return [l for l in sec.splitlines() if l.startswith("| `Dulang_")]

    def val_table(text):
        sec = section(text, "## Validation checks", "\n## Dulang_")
        return sec.strip().splitlines()[2:]

    L = []
    a = L.append
    a("# Dulang B (PPTX) — well-schematic extraction: completion summary")
    a("")
    a(f"Run: {dt.datetime.now().isoformat(timespec='seconds')} (local processing only). Scripts: `output_B/_work_pptx/build_schematic_Bt.py` "
      "(native table slides), `build_schematic_Bs.py` (pictures / irregular layout, from `specs/*.txt` transcriptions), `finalize_B.py` (publish).")
    a("")
    a("## Inputs (read only) — SHA-256 before / after")
    a("")
    a("| file | before | after | unchanged |")
    a("|---|---|---|---|")
    for k in before:
        a(f"| `{k}` | `{before[k]}` | `{after[k]}` | {'YES' if before[k] == after[k] else 'NO'} |")
    a("")
    a(f"The NAS PPTX was hashed before processing (`{NAS_SHA}`), copied once to `output_B/_work_pptx/src/`, and re-hashed at the end (table above).")
    a("")
    a("Template note: `well-schematic-template-json.rtf` is an **example JSON document, not a formal JSON Schema**. Conformity was checked against "
      "its section order, `wellInfo` keys/order and value types. Extra component keys (topMdM/bottomMdM, zone, string, weight, grade, bottomTvd, "
      "estimated…, actualType, source) follow the earlier Dulang outputs; `floatEquipment` / `plugBackCement` entries use the key names of the "
      "Excel template sheets because the JSON example lists are empty.")
    a("")
    a("File naming: the request text says `Dulang_D<well-number>…`; for this Dulang B deck the files are named `Dulang_B<number><suffix>.json` "
      "(e.g. `Dulang_B24ST2.json`), consistent with the slide labels and the earlier `output_B` naming.")
    a("")
    a("## Wells identified (one JSON per resolved well)")
    a("")
    a(f"33 slides read in presentation order (no hidden slides: no `show=\"0\"`; notes pages hold only 'Open' + slide number plus a few "
      "update notes). 32 wells resolved -> **32 JSON files**.")
    a("")
    a("- Native PowerPoint table layout (STATUS | MIN ID | SHORT STRING | DEPTH ‖ DEPTH | LONG STRING | MIN ID | STATUS), parsed from the "
      "native text with positional row pairing and checked against slide renders: slides 1–15, 17–23, 26, 27, 29, 32 (single-string layout on 8 and 12).")
    a("- Native shapes in a non-table layout: slide 16 (B-16ST3) — multi-line depth/ID/OD text boxes paired with labels visually (see its issues).")
    a("- Pictures (visual transcription of x3 crops + OCR cross-check): slide 24 (B-24ST2, scan), 25 (B25 ST-4), 28 (B-28ST2), "
      "30 (B30 ST-1), 31 (B31 ST-1). Slides 25, 30, 31 are low-resolution pictures: many values are flagged as uncertain, and each slide "
      "states that the latest editable diagram is an Excel file in ECMS (not part of this input).")
    a("- Slide 33 is an older B-17 ST3 schematic labelled 'OBSOLETE SCHEMATIC FOR B17' (Feb 2022): not merged; slide 17 (June 2023) is mapped and the "
      "conflict is recorded in `Dulang_B17ST3`.")
    a("")
    a("| well file | original label | slide(s) | schematic source | components | flagged | conflicts | missing | unclear | validation |")
    a("|---|---|---|---|---|---|---|---|---|---|")
    for line in wells_table(t_sum) + wells_table(s_sum):
        cells = [c.strip() for c in line.strip("|").split("|")]
        key = cells[0].strip("`").replace("Dulang_", "").replace(".json", "")
        cells[2] = SLIDES[key]
        if not cells[3] or cells[3] == "native slide shapes/text":
            cells[3] = "native slide table text"
        a("| " + " | ".join(cells) + " |")
    a("")
    a("## Validation checks (per JSON)")
    a("")
    a("| check | result |")
    a("|---|---|")
    checks = {}
    for line in val_table(t_sum) + val_table(s_sum):
        cells = [c.strip() for c in line.strip("|").split("|")]
        if cells and cells[0]:
            checks.setdefault(cells[0], []).extend(cells[1:])
    for k, v in checks.items():
        a(f"| {k} | {'PASS for every file (' + str(len(v)) + ')' if v and all(x == '✓' for x in v) else 'see per-file table in stage summary'} |")
    a("")
    a("Per-well conflicts, missing / unclear values, unmapped content and full traceability (JSON path, value, slide/object, original text, "
      "conversion) follow.")
    a("")
    body = section(t_sum, "\n## Dulang_") + "\n" + section(s_sum, "\n## Dulang_")
    a(body)
    (OUT / "_completion_summary.md").write_text("\n".join(L) + "\n", encoding="utf-8")
    assert before == after
    print(len(rows), "JSON files published to", OUT)


if __name__ == "__main__":
    main()
