#!/usr/bin/env python3
"""Metre version of Dulang_A2: reads output_A/Dulang_A2.json (read-only, unchanged) and writes output_A/metres/Dulang_A2.json
with topMd/bottomMd in metres exactly as printed on PDF page 3 (no ft conversion). The *M duplicate keys are dropped because
the main depth keys now hold the metres. Diameters stay in inches and weights in ppf, as printed. Everything else is identical.
"""
import hashlib
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "output_A" / "Dulang_A2.json"
OUT = ROOT / "output_A" / "metres" / "Dulang_A2.json"
FT = 3.28084


def main():
    h0 = hashlib.sha256(SRC.read_bytes()).hexdigest()
    doc = json.loads(SRC.read_text())
    for c in doc["components"]:
        for k in ("topMd", "bottomMd"):
            m = c.pop(k + "M")
            assert (m is None and c[k] is None) or c[k] == round(m * FT, 3), (c["id"], k)
            c[k] = m
        c["source"] = c["source"].replace("; CALCULATED: ft = m x 3.28084 (3 dp)", "; depths in metres as printed (no conversion)")
    if OUT.exists():
        raise SystemExit(f"{OUT} already exists — not overwriting")
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text(json.dumps(doc, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")

    back = json.loads(OUT.read_text())
    src = json.loads(SRC.read_text())
    cs = back["components"]
    checks = [
        ("valid JSON", True),
        ("same sections / wellInfo / component count and order as Dulang_A2.json",
         list(back) == list(src) and back["wellInfo"] == src["wellInfo"] and [c["id"] for c in cs] == [c["id"] for c in src["components"]]),
        ("topMd/bottomMd = printed metres (= former *M values)",
         all(c["topMd"] == s["topMdM"] and c["bottomMd"] == s["bottomMdM"] for c, s in zip(cs, src["components"]))),
        ("no *M keys and no ft wording left", all("topMdM" not in c and "ft = m" not in c["source"] for c in cs)),
        ("source Dulang_A2.json unchanged", hashlib.sha256(SRC.read_bytes()).hexdigest() == h0),
    ]
    for a, b in checks:
        print("PASS" if b else "FAIL", "-", a)
    print("written:", OUT.relative_to(ROOT), "|", len(cs), "components")


if __name__ == "__main__":
    main()
