#!/usr/bin/env python3
"""Write output_A/Dulang_A17ST2_2.json = output_A/Dulang_A17ST2.json + the user's listed corrections only (28 Oct 2026).

Decisions confirmed by the user:
  * flow coupling depth 822.67 ft -> upper coupling (U29, flow-coupling-1), topMd only
  * TRSV depth removed although S30 holds 250.738 m (conflict flagged)
  * 13CR-L80 tubing bottom set to missing (requested 5,446 ft = AB15 1660 m, a length by formula, not a depth)
  * liner / tubing sizes without an ID/OD designation (header column V 'SIZE') treated as ID — user-defined assumption
The initial JSON and the Excel source are only read; the output is never overwritten.
"""
import copy
import json
from pathlib import Path

OUT_A = Path(__file__).resolve().parent.parent / "output_A"
SRC, DST = OUT_A / "Dulang_A17ST2.json", OUT_A / "Dulang_A17ST2_2.json"
NOTE = "REVISED 2026-10-28 (user correction)"


def main():
    if DST.exists():
        raise SystemExit(f"{DST} exists — not overwriting without confirmation")
    raw = SRC.read_text(encoding="utf-8")
    orig = json.loads(raw)
    doc = copy.deepcopy(orig)
    c = {x["id"]: x for x in doc["components"]}

    # 1. tubing hanger: bottomMd already missing (null) -> nothing to change
    assert c["tubing-hanger"]["bottomMd"] is None and c["tubing-hanger"]["bottomMdM"] is None

    # 2. TRSV depth removed
    t = c["trsv"]
    assert t["topMd"] == 822.631 and t["topMdM"] == 250.738 and t["bottomMd"] is None
    t["topMd"] = t["topMdM"] = None
    t["source"] += f"; {NOTE}: topMd removed — installation depth treated as not specified (CONFLICT: S30 = 250.738 m-MDDF on the TRSV row)"

    # 3. upper flow coupling (U29) depth 822.67 ft, topMd only
    f = c["flow-coupling-1"]
    assert f["topMd"] is None and f["bottomMd"] is None and "U29" in f["source"]
    f["topMd"] = 822.67
    f["source"] += f"; {NOTE}: topMd 822.67 ft supplied by user (no depth printed in S29); bottomMd left missing; no metre value in source"

    # 4. TRMAXX / TRSV: already one item (U30 'TRMAXX-5E ...'); no separate 'Triple SV' entry exists -> nothing to consolidate
    assert sum("TRMAXX" in x["name"] or "TRIPLE" in x["name"].upper() for x in doc["components"]) == 1

    # 5. 13CR-L80 tubing bottom -> missing
    tl = c["tubing-lower"]
    assert tl["grade"] == "13CR-L80" and tl["bottomMd"] == 6271.457 and tl["estimatedFields"] == ["topMd", "bottomMd"]
    tl["bottomMd"] = tl["bottomMdM"] = None
    tl["estimatedFields"] = ["topMd"]
    tl["estimationNote"] = "DERIVED top = TRSV depth S30; header AB15 1660 is a LENGTH (formula =ROUNDUP(S71-S30,0)), not a depth, and is not used as one"
    tl["source"] += (f"; {NOTE}: bottomMd removed (was 1911.54 m DERIVED from mule shoe S68); requested 5,446 ft equals AB15 1660 m, "
                     "a length by formula, so the bottom is left missing")

    # 6. liner / tubing sizes (header column V8 'SIZE', no ID/OD designation) -> ID (user-defined assumption)
    for cid in ("liner-7", "tubing-upper", "tubing-lower"):
        x = c[cid]
        assert x["outerDiameter"] is not None and x["innerDiameter"] is None
        x["innerDiameter"], x["outerDiameter"] = x["outerDiameter"], None
        x["source"] += f"; {NOTE}: size column V 'SIZE' has no ID/OD designation -> recorded as innerDiameter (user-defined assumption); outerDiameter missing"

    out = json.dumps(doc, indent=2, ensure_ascii=False) + "\n"
    DST.write_text(out, encoding="utf-8")

    # ---------------- validation: only the intended fields differ
    back = json.loads(DST.read_text(encoding="utf-8"))
    assert SRC.read_text(encoding="utf-8") == raw
    diffs = []
    assert [k for k in back] == [k for k in orig]
    for k in orig:
        if k != "components":
            assert back[k] == orig[k], k
    assert [x["id"] for x in back["components"]] == [x["id"] for x in orig["components"]]
    for a, b in zip(orig["components"], back["components"]):
        assert list(a) == list(b), a["id"]
        for key in a:
            if a[key] != b[key]:
                diffs.append((a["id"], key, a[key], b[key]))
    expected = {("trsv", "topMd"), ("trsv", "topMdM"), ("trsv", "source"), ("flow-coupling-1", "topMd"), ("flow-coupling-1", "source"),
                ("tubing-lower", "bottomMd"), ("tubing-lower", "bottomMdM"), ("tubing-lower", "estimatedFields"),
                ("tubing-lower", "estimationNote"), ("tubing-lower", "source")} | \
               {(i, k) for i in ("liner-7", "tubing-upper", "tubing-lower") for k in ("outerDiameter", "innerDiameter", "source")}
    assert {(i, k) for i, k, _, _ in diffs} == expected, {(i, k) for i, k, _, _ in diffs} ^ expected
    assert all(isinstance(v, (int, float)) or v is None for x in back["components"]
               for k, v in x.items() if k in ("topMd", "bottomMd", "outerDiameter", "innerDiameter"))
    print("valid JSON; components", len(back["components"]), "(unchanged count); differences limited to the requested fields:")
    for i, k, a, b in diffs:
        if k not in ("source",):
            print(f"  {i}.{k}: {a!r} -> {b!r}")


if __name__ == "__main__":
    main()
