# Dulang A (PPTX: A6, A8, A19, A30) — well-schematic extraction: completion summary

Run: 2026-11-11T11:01:56 (local processing only). Script: `output_A/_work_pptx/build_schematic_A.py`.

## Inputs (read only) — SHA-256 before / after

| file | before | after | unchanged |
|---|---|---|---|
| `/Users/farahyusof/Documents/GitHub/NoramiraSazuki/AinnyFarah/WELL_SCHEMATIC/output_A/_work_pptx/src/Dulang A Well Schematic updated.pptx` | `a5ca9a3b0076122633de31730fc97b5b8e3185d813c93bdcfdb3ad56bef2861b` | `a5ca9a3b0076122633de31730fc97b5b8e3185d813c93bdcfdb3ad56bef2861b` | YES |
| `/Users/farahyusof/Documents/GitHub/NoramiraSazuki/AinnyFarah/WELL_SCHEMATIC/well-schematic-template-json.rtf` | `c7db15efb1917dfa76fa3102e270fe17115a8561f24eb0a155239870d079e712` | `c7db15efb1917dfa76fa3102e270fe17115a8561f24eb0a155239870d079e712` | YES |
| `/Users/farahyusof/Documents/GitHub/NoramiraSazuki/AinnyFarah/WELL_SCHEMATIC/well-schematic-sample-template.xlsx` | `d377a6310d42fda8c95e85db03f2416cb4c1b6f37f53920052331ce814f6d494` | `d377a6310d42fda8c95e85db03f2416cb4c1b6f37f53920052331ce814f6d494` | YES |

NAS original `/Volumes/homes/public/DulangOSS/20261016/DULANG_WELLS/Well_Schematic/1. Well Schematic/Dulang A/Dulang A Well Schematic updated.pptx`: SHA-256 at the start of processing `a5ca9a3b0076122633de31730fc97b5b8e3185d813c93bdcfdb3ad56bef2861b` (identical to the local working copy above). Re-hashed at the end: `a5ca9a3b0076122633de31730fc97b5b8e3185d813c93bdcfdb3ad56bef2861b` — UNCHANGED.

Template note: `well-schematic-template-json.rtf` is an **example JSON document, not a formal JSON Schema**. Conformity was checked against its section order, `wellInfo` keys/order and value types; component keys beyond the example (topMdM/bottomMdM, zone, string, weight, grade, bottomTvd, estimated…, actualType, source) follow the conventions of the earlier Dulang A/B outputs. `floatEquipment` / `plugBackCement` entries use key names from the Excel template sheets (Float_Equipment, Plug_Back_Cement) because the JSON example lists are empty.

## Wells identified (one JSON per resolved well)

| well file | original label | slide(s) | schematic source | components | flagged | conflicts | missing | unclear | validation |
|---|---|---|---|---|---|---|---|---|---|
| `DULANG-A6.json` | WELL A-6 : SINGLE OIL PRODUCER (TYPE 1) | 5 | visual transcription of the screenshot pictures — image11.png + image12.png | 22 | 10 | 1 | 0 | 3 | PASS |
| `DULANG-A8.json` | WELL A-8 : SINGLE SELECTIVE OIL PRODUCER (TYPE 3) | 7 | visual transcription of the screenshot pictures — image16.png + image17.png | 26 | 11 | 1 | 0 | 2 | PASS |
| `DULANG-A19.json` | WELL A-19 : SINGLE SELECTIVE OIL PRODUCER (TYPE 3) | 16 | visual transcription of the screenshot pictures — image36.png + image37.png | 28 | 12 | 0 | 6 | 1 | PASS |
| `DULANG-A30.json` | WELL A-30: DUAL OIL PRODUCER/WATER INJECTOR WITH SHORT STRING SELECTIVE | 27 | visual transcription of the screenshot pictures — image62.png + image63.png | 46 | 18 | 1 | 4 | 3 | PASS |

JSON files written: **4** = resolved wells **4**. Slides read: 30 (presentation order, no hidden slides — no `show="0"`); slides covered by a well: [5, 7, 16, 27]. Pictures placed outside the visible slide area (slides 4, 7, 16, 19, 23, 31) were inspected and are listed per well as unmapped hidden content. Combined slides: D28 = slides 28 (schematic) + 29 (completion report, different datum — notes only); D32 = slides 33 (schematic) + 34 (completion report — notes only); D30 slide 31 holds a second, off-slide schematic (notes only).

Extraction method: the A schematics are screenshot pictures (upper + lower part per slide); records; screenshots / scans were transcribed visually from x3-upscaled crops with tesseract OCR as a cross-check. Transcriptions with per-row locations are in `output_A/_work_pptx/specs/*.txt`; low-legibility reads are marked '?' there and flagged in the JSON.

## Validation checks (per JSON)

| check | A6 | A8 | A19 | A30 |
|---|---|---|---|---|
| sections and order as template | ✓ | ✓ | ✓ | ✓ |
| wellInfo keys and order as template | ✓ | ✓ | ✓ | ✓ |
| wellInfo types (number or null; string) | ✓ | ✓ | ✓ | ✓ |
| unsupported sections empty lists | ✓ | ✓ | ✓ | ✓ |
| approved Type_Keys only | ✓ | ✓ | ✓ | ✓ |
| unique component ids | ✓ | ✓ | ✓ | ✓ |
| numeric fields numeric or null | ✓ | ✓ | ✓ | ✓ |
| ft = m x 3.28084 (3 dp) everywhere | ✓ | ✓ | ✓ | ✓ |
| nearest-type mappings flagged | ✓ | ✓ | ✓ | ✓ |
| flags complete (estimated -> fields + note) | ✓ | ✓ | ✓ | ✓ |
| no template sample values | ✓ | ✓ | ✓ | ✓ |
| well identifier = file name | ✓ | ✓ | ✓ | ✓ |
| every component traced to the source | ✓ | ✓ | ✓ | ✓ |
| every transcribed row mapped or listed | ✓ | ✓ | ✓ | ✓ |

## DULANG-A6 — slide 5

Original label: WELL A-6 : SINGLE OIL PRODUCER (TYPE 1)
Equipment depth datum: m MDTHF (column headers 'DEPTH m MDTHF')  ·  source: image11.png + image12.png (visual transcription of the screenshot pictures (x3 upscale, upper part image11 / lower part image12))

### Conflicts / missing / unclear

| kind | where | detail |
|---|---|---|
| Conflict | pbr-1 | 'SEAL BORE EXT' at 1459.2 m is shallower than the item listed above it (packer-1 at 1459.9 m) — kept as printed |
| Unclear | picture p2 | SEAL BORE EXT printed at 1459.2 m, shallower than the GT packer (1459.9 m) listed above it — kept as printed |
| Unclear | slide 5 obj #10 | status 'GLOV 20/64" (Mar 15)' is a slide text box; matched to SPM 1357.6 m by vertical position (later than the printed diagram) |
| Unclear | header | PBTD 1800 m MDDF (header) and TD 1819.9 m MDTHF (diagram) are on different datums — not shifted; 7" shoe 1834 m MDDF - RTE 14.13 m = 1819.87 m, consistent with the TD |

### Unmapped content (kept here, not in the JSON)

| where | text | reason |
|---|---|---|
| slide 5 image11.png header | RTE / drill floor to tubing hanger : 14.13 m | offset, not an elevation above datum -> kbElevation left empty (no datum shift) |
| picture p1 | 1/4" CONTROL LINE (to the SCSSV) | note (status / history / survey / other revision) |
| picture p2 | 7" 23 PPF CASING label on the diagram (no depth) | note (status / history / survey / other revision) |
| picture p2 | DOGLEG (DEPTH): 135.0 - 185.0, 414.8 - 703.1 | note (status / history / survey / other revision) |
| slide 5 obj #35 | PB valve set at TRSV profile. Jammed Open \|\| TR flapper lock open in May 2021 | note (status / history / survey / other revision) |
| slide 5 obj #7 | 28 Apr 2026 \|\| RIH 2.25" GR tagged D No Go at 1464m-MDTHF \|\| RIH 2.00" LIB tagged HUD at 1529m-MDTHF | note (status / history / survey / other revision) |
| slide 5 obj #6 | Damaged 2" OD gun left in hole (July 2002). Top of fish unknown | note (status / history / survey / other revision) |
| slide 5 obj #18 | FB – SE1 | slide text box over/beside the picture (status / history / note) |
| slide 5 obj #9 | CLOSED | slide text box over/beside the picture (status / history / note) |
| slide 5 obj #10 | GLOV 20/64" (Mar 15) | slide text box over/beside the picture (status / history / note) |
| slide 5 obj #35 | PB valve set at TRSV profile. Jammed Open TR flapper lock open in May 2021 | slide text box over/beside the picture (status / history / note) |
| slide 5 obj #7 | 28 Apr 2026 RIH 2.25” GR tagged D No Go at 1464m-MDTHF RIH 2.00” LIB tagged HUD at 1529m-MDTHF | slide text box over/beside the picture (status / history / note) |
| slide 5 obj #6 | Damaged 2” OD gun left in hole (July 2002). Top of fish unknown | slide text box over/beside the picture (status / history / note) |
| slide 5 obj #11 | mMDTHF 1482.9 – 1486.7 1487.4 – 1496.9 | slide text box over/beside the picture (status / history / note) |

### Traceability (every populated value)

| JSON path | value | slide / object / location | original text | conversion / interpretation |
|---|---|---|---|---|
| `components[christmas-tree-1].name` | INGRAM CACTUS, SINGLE, 3M | slide 5 image11.png header | INGRAM CACTUS, SINGLE, 3M | visual transcription of the screenshot pictures (x3 upscale, upper part image11 / lower part image12) |
| `components[christmas-tree-1].type` | christmas-tree | slide 5 image11.png header | INGRAM CACTUS, SINGLE, 3M | EXACT TYPE (Component_Types) |
| `components[flow-coupling-1].name` | FLOW COUPLING | slide 5 image11.png (upper picture) | FLOW COUPLING | visual transcription of the screenshot pictures (x3 upscale, upper part image11 / lower part image12) |
| `components[flow-coupling-1].type` | flow-coupling | slide 5 image11.png (upper picture) | FLOW COUPLING | EXACT TYPE (Component_Types) |
| `components[flow-coupling-1].topMd` | 465.223 | slide 5 image11.png (upper picture) | 141.8 | 141.8 m x 3.28084 = 465.223 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[flow-coupling-1].topMdM` | 141.8 | slide 5 image11.png (upper picture) | 141.8 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[flow-coupling-1].innerDiameter` | 2.441 | slide 5 image11.png (upper picture) | 2.441 | column 'MIN ID' = explicit ID |
| `components[scssv-1].name` | TRDP-4A SCSSV | slide 5 image11.png (upper picture) | TRDP-4A SCSSV | visual transcription of the screenshot pictures (x3 upscale, upper part image11 / lower part image12) |
| `components[scssv-1].type` | scssv | slide 5 image11.png (upper picture) | TRDP-4A SCSSV | EXACT TYPE (Component_Types) |
| `components[scssv-1].topMd` | 467.848 | slide 5 image11.png (upper picture) | 142.6 | 142.6 m x 3.28084 = 467.848 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[scssv-1].topMdM` | 142.6 | slide 5 image11.png (upper picture) | 142.6 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[scssv-1].innerDiameter` | 2.312 | slide 5 image11.png (upper picture) | 2.312 | column 'MIN ID' = explicit ID |
| `components[landing-nipple-1].name` | WP-1 NIPPLE | slide 5 image11.png (upper picture) | WP-1 NIPPLE | visual transcription of the screenshot pictures (x3 upscale, upper part image11 / lower part image12) |
| `components[landing-nipple-1].type` | landing-nipple | slide 5 image11.png (upper picture) | WP-1 NIPPLE | NEAREST TYPE (flagged) |
| `components[landing-nipple-1].topMd` | 969.816 | slide 5 image11.png (upper picture) | 295.6 | 295.6 m x 3.28084 = 969.816 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[landing-nipple-1].topMdM` | 295.6 | slide 5 image11.png (upper picture) | 295.6 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[landing-nipple-1].innerDiameter` | 2.312 | slide 5 image11.png (upper picture) | 2.312 | column 'MIN ID' = explicit ID |
| `components[spm-1].name` | KBMG-M-SPM | slide 5 image11.png (upper picture) | KBMG-M-SPM | visual transcription of the screenshot pictures (x3 upscale, upper part image11 / lower part image12) |
| `components[spm-1].type` | spm | slide 5 image11.png (upper picture) | KBMG-M-SPM | EXACT TYPE (Component_Types) |
| `components[spm-1].topMd` | 1781.824 | slide 5 image11.png (upper picture) | 543.1 | 543.1 m x 3.28084 = 1781.824 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[spm-1].topMdM` | 543.1 | slide 5 image11.png (upper picture) | 543.1 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[spm-1].innerDiameter` | 2.371 | slide 5 image11.png (upper picture) | 2.371 | column 'MIN ID' = explicit ID |
| `components[spm-2].name` | 8/64" GLOV | slide 5 image11.png (upper picture) | 8/64" GLOV | visual transcription of the screenshot pictures (x3 upscale, upper part image11 / lower part image12) |
| `components[spm-2].type` | spm | slide 5 image11.png (upper picture) | 8/64" GLOV | NEAREST TYPE (flagged) |
| `components[spm-2].topMd` | 3180.118 | slide 5 image11.png (upper picture) | 969.3 | 969.3 m x 3.28084 = 3180.118 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[spm-2].topMdM` | 969.3 | slide 5 image11.png (upper picture) | 969.3 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[spm-2].innerDiameter` | 2.371 | slide 5 image11.png (upper picture) | 2.371 | column 'MIN ID' = explicit ID |
| `components[spm-3].name` | KBMG-M-SPM | slide 5 image11.png (upper picture) | KBMG-M-SPM | visual transcription of the screenshot pictures (x3 upscale, upper part image11 / lower part image12) |
| `components[spm-3].type` | spm | slide 5 image11.png (upper picture) | KBMG-M-SPM | EXACT TYPE (Component_Types) |
| `components[spm-3].topMd` | 4052.494 | slide 5 image11.png (upper picture) | 1235.2 | 1235.2 m x 3.28084 = 4052.494 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[spm-3].topMdM` | 1235.2 | slide 5 image11.png (upper picture) | 1235.2 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[spm-3].innerDiameter` | 2.371 | slide 5 image11.png (upper picture) | 2.371 | column 'MIN ID' = explicit ID |
| `components[spm-4].name` | KBMG-M-SPM | slide 5 image11.png (upper picture) | KBMG-M-SPM | visual transcription of the screenshot pictures (x3 upscale, upper part image11 / lower part image12) |
| `components[spm-4].type` | spm | slide 5 image11.png (upper picture) | KBMG-M-SPM | EXACT TYPE (Component_Types) |
| `components[spm-4].topMd` | 4454.068 | slide 5 image11.png (upper picture) | 1357.6 | 1357.6 m x 3.28084 = 4454.068 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[spm-4].topMdM` | 1357.6 | slide 5 image11.png (upper picture) | 1357.6 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[spm-4].innerDiameter` | 2.371 | slide 5 image11.png (upper picture) | 2.371 | column 'MIN ID' = explicit ID |
| `components[sliding-sleeve-1].name` | WP-2DE SSD | slide 5 image12.png (lower picture) | WP-2DE SSD | visual transcription of the screenshot pictures (x3 upscale, upper part image11 / lower part image12) |
| `components[sliding-sleeve-1].type` | sliding-sleeve | slide 5 image12.png (lower picture) | WP-2DE SSD | EXACT TYPE (Component_Types) |
| `components[sliding-sleeve-1].topMd` | 4759.515 | slide 5 image12.png (lower picture) | 1450.7 | 1450.7 m x 3.28084 = 4759.515 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[sliding-sleeve-1].topMdM` | 1450.7 | slide 5 image12.png (lower picture) | 1450.7 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[sliding-sleeve-1].innerDiameter` | 2.312 | slide 5 image12.png (lower picture) | 2.312 | column 'MIN ID' = explicit ID |
| `components[pup-joint-1].name` | CENTRALIZER | slide 5 image12.png (lower picture) | CENTRALIZER | visual transcription of the screenshot pictures (x3 upscale, upper part image11 / lower part image12) |
| `components[pup-joint-1].type` | pup-joint | slide 5 image12.png (lower picture) | CENTRALIZER | UNRESOLVED / NEAREST TYPE (flagged) |
| `components[pup-joint-1].topMd` | 4785.433 | slide 5 image12.png (lower picture) | 1458.6 | 1458.6 m x 3.28084 = 4785.433 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[pup-joint-1].topMdM` | 1458.6 | slide 5 image12.png (lower picture) | 1458.6 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[pup-joint-1].innerDiameter` | 2.441 | slide 5 image12.png (lower picture) | 2.441 | column 'MIN ID' = explicit ID |
| `components[seal-assembly-1].name` | LOC. SEAL ASSY. | slide 5 image12.png (lower picture) | LOC. SEAL ASSY. | visual transcription of the screenshot pictures (x3 upscale, upper part image11 / lower part image12) |
| `components[seal-assembly-1].type` | seal-assembly | slide 5 image12.png (lower picture) | LOC. SEAL ASSY. | EXACT TYPE (Component_Types) |
| `components[seal-assembly-1].topMd` | 4787.074 | slide 5 image12.png (lower picture) | 1459.1 | 1459.1 m x 3.28084 = 4787.074 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[seal-assembly-1].topMdM` | 1459.1 | slide 5 image12.png (lower picture) | 1459.1 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[seal-assembly-1].innerDiameter` | 2.441 | slide 5 image12.png (lower picture) | 2.441 | column 'MIN ID' = explicit ID |
| `components[packer-1].name` | GUIB MAGNUM 'GT' PACKER | slide 5 image12.png (lower picture) | GUIB MAGNUM 'GT' PACKER | visual transcription of the screenshot pictures (x3 upscale, upper part image11 / lower part image12) |
| `components[packer-1].type` | packer | slide 5 image12.png (lower picture) | GUIB MAGNUM 'GT' PACKER | EXACT TYPE (Component_Types) |
| `components[packer-1].topMd` | 4789.698 | slide 5 image12.png (lower picture) | 1459.9 | 1459.9 m x 3.28084 = 4789.698 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[packer-1].topMdM` | 1459.9 | slide 5 image12.png (lower picture) | 1459.9 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[packer-1].innerDiameter` | 4.0 | slide 5 image12.png (lower picture) | 4.000 | column 'MIN ID' = explicit ID |
| `components[pbr-1].name` | SEAL BORE EXT | slide 5 image12.png (lower picture) | SEAL BORE EXT | visual transcription of the screenshot pictures (x3 upscale, upper part image11 / lower part image12) |
| `components[pbr-1].type` | pbr | slide 5 image12.png (lower picture) | SEAL BORE EXT | NEAREST TYPE (flagged) |
| `components[pbr-1].topMd` | 4787.402 | slide 5 image12.png (lower picture) | 1459.2 | 1459.2 m x 3.28084 = 4787.402 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[pbr-1].topMdM` | 1459.2 | slide 5 image12.png (lower picture) | 1459.2 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[pbr-1].innerDiameter` | 4.0 | slide 5 image12.png (lower picture) | 4.000 | column 'MIN ID' = explicit ID |
| `components[nipple-no-go-1].name` | 'D' NO-GO NIPPLE | slide 5 image12.png (lower picture) | 'D' NO-GO NIPPLE | visual transcription of the screenshot pictures (x3 upscale, upper part image11 / lower part image12) |
| `components[nipple-no-go-1].type` | nipple (no-go) | slide 5 image12.png (lower picture) | 'D' NO-GO NIPPLE | EXACT TYPE (Component_Types) |
| `components[nipple-no-go-1].topMd` | 4805.446 | slide 5 image12.png (lower picture) | 1464.7 | 1464.7 m x 3.28084 = 4805.446 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[nipple-no-go-1].topMdM` | 1464.7 | slide 5 image12.png (lower picture) | 1464.7 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[nipple-no-go-1].innerDiameter` | 2.25 | slide 5 image12.png (lower picture) | 2.250 | column 'MIN ID' = explicit ID |
| `components[mule-shoe-1].name` | WEG | slide 5 image12.png (lower picture) | WEG | visual transcription of the screenshot pictures (x3 upscale, upper part image11 / lower part image12) |
| `components[mule-shoe-1].type` | mule-shoe | slide 5 image12.png (lower picture) | WEG | NEAREST TYPE (flagged) |
| `components[mule-shoe-1].topMd` | 4805.774 | slide 5 image12.png (lower picture) | 1464.8 | 1464.8 m x 3.28084 = 4805.774 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[mule-shoe-1].topMdM` | 1464.8 | slide 5 image12.png (lower picture) | 1464.8 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[mule-shoe-1].innerDiameter` | 2.441 | slide 5 image12.png (lower picture) | 2.441 | column 'MIN ID' = explicit ID |
| `components[conductor-1].name` | 13 3/8" COND.# @ 374 mMDDF | slide 5 image11.png header | 13 3/8" COND.# @ 374 mMDDF | visual transcription of the screenshot pictures (x3 upscale, upper part image11 / lower part image12) |
| `components[conductor-1].type` | casing | slide 5 image11.png header | 13 3/8" COND.# @ 374 mMDDF | EXACT TYPE (Component_Types) |
| `components[conductor-1].bottomMd` | 1227.034 | slide 5 image11.png header | 13 3/8" COND.# @ 374 mMDDF | 374.0 m x 3.28084 = 1227.034 ft (3 dp); m MDDF as printed |
| `components[conductor-1].bottomMdM` | 374.0 | slide 5 image11.png header | 13 3/8" COND.# @ 374 mMDDF | original metres; m MDDF as printed |
| `components[conductor-1].innerDiameter` | 13.375 | slide 5 image11.png header | 13.375 | USER-DEFINED RULE: header size without ID/OD designation -> innerDiameter |
| `components[production-casing-1].name` | 9 5/8" N-80, 40 ppf @ 808 m MDDF | slide 5 image11.png header | 9 5/8" N-80, 40 ppf @ 808 m MDDF | visual transcription of the screenshot pictures (x3 upscale, upper part image11 / lower part image12) |
| `components[production-casing-1].type` | casing | slide 5 image11.png header | 9 5/8" N-80, 40 ppf @ 808 m MDDF | EXACT TYPE (Component_Types) |
| `components[production-casing-1].bottomMd` | 2650.919 | slide 5 image11.png header | 9 5/8" N-80, 40 ppf @ 808 m MDDF | 808.0 m x 3.28084 = 2650.919 ft (3 dp); m MDDF as printed |
| `components[production-casing-1].bottomMdM` | 808.0 | slide 5 image11.png header | 9 5/8" N-80, 40 ppf @ 808 m MDDF | original metres; m MDDF as printed |
| `components[production-casing-1].innerDiameter` | 9.625 | slide 5 image11.png header | 9.625 | USER-DEFINED RULE: header size without ID/OD designation -> innerDiameter |
| `components[production-casing-1].weight` | 40.0 | slide 5 image11.png header | 9 5/8" N-80, 40 ppf @ 808 m MDDF | parsed from the label (ppf = lb/ft) |
| `components[production-casing-1].grade` | N-80 | slide 5 image11.png header | 9 5/8" N-80, 40 ppf @ 808 m MDDF | parsed from the label |
| `components[production-casing-2].name` | 7" 23 ppf, K-55 @ 1834 m MDDF | slide 5 image11.png header | 7" 23 ppf, K-55 @ 1834 m MDDF | visual transcription of the screenshot pictures (x3 upscale, upper part image11 / lower part image12) |
| `components[production-casing-2].type` | casing | slide 5 image11.png header | 7" 23 ppf, K-55 @ 1834 m MDDF | EXACT TYPE (Component_Types) |
| `components[production-casing-2].bottomMd` | 6017.061 | slide 5 image11.png header | 7" 23 ppf, K-55 @ 1834 m MDDF | 1834.0 m x 3.28084 = 6017.061 ft (3 dp); m MDDF as printed |
| `components[production-casing-2].bottomMdM` | 1834.0 | slide 5 image11.png header | 7" 23 ppf, K-55 @ 1834 m MDDF | original metres; m MDDF as printed |
| `components[production-casing-2].innerDiameter` | 7.0 | slide 5 image11.png header | 7 | USER-DEFINED RULE: header size without ID/OD designation -> innerDiameter |
| `components[production-casing-2].weight` | 23.0 | slide 5 image11.png header | 7" 23 ppf, K-55 @ 1834 m MDDF | parsed from the label (ppf = lb/ft) |
| `components[production-casing-2].grade` | K-55 | slide 5 image11.png header | 7" 23 ppf, K-55 @ 1834 m MDDF | parsed from the label |
| `components[tubing-1].name` | 2 7/8" N-80 NK3SB | slide 5 image11.png header | 2 7/8" N-80 NK3SB | visual transcription of the screenshot pictures (x3 upscale, upper part image11 / lower part image12) |
| `components[tubing-1].type` | tubing | slide 5 image11.png header | 2 7/8" N-80 NK3SB | EXACT TYPE (Component_Types) |
| `components[tubing-1].topMd` | 0.0 | slide 5 image11.png header | depth reference m MDTHF / THF | 0.0 m x 3.28084 = 0.0 ft (3 dp); ASSUMED 0 m (tubing hangs from the tubing hanger flange) |
| `components[tubing-1].topMdM` | 0.0 | slide 5 image11.png header | depth reference m MDTHF / THF | original metres; ASSUMED 0 m (tubing hangs from the tubing hanger flange) |
| `components[tubing-1].bottomMd` | 4805.774 | slide 5 image11.png header | WEG | 1464.8 m x 3.28084 = 4805.774 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF'); depth of 'WEG' (mule-shoe-1) |
| `components[tubing-1].bottomMdM` | 1464.8 | slide 5 image11.png header | WEG | original metres; m MDTHF (column headers 'DEPTH m MDTHF'); depth of 'WEG' (mule-shoe-1) |
| `components[tubing-1].innerDiameter` | 2.875 | slide 5 image11.png header | 2 7/8" N-80 NK3SB | USER-DEFINED RULE: header size without ID/OD designation -> innerDiameter |
| `components[tubing-1].grade` | N-80 | slide 5 image11.png header | 2 7/8" N-80 NK3SB | parsed from the label |
| `components[zone-e-14-1].name` | E-14 1482.9-1486.7 m MDTHF | slide 5 image11.png slide obj #11 | E-14 1482.9-1486.7 m MDTHF | visual transcription of the screenshot pictures (x3 upscale, upper part image11 / lower part image12) |
| `components[zone-e-14-1].type` | perforation | slide 5 image11.png slide obj #11 | E-14 1482.9-1486.7 m MDTHF | EXACT TYPE (Component_Types) |
| `components[zone-e-14-1].topMd` | 4865.158 | slide 5 image11.png slide obj #11 | slide text box obj #11 'mMDTHF 1482.9 – 1486.7' beside SAND: E-14 | 1482.9 m x 3.28084 = 4865.158 ft (3 dp); m MDTHF as printed |
| `components[zone-e-14-1].topMdM` | 1482.9 | slide 5 image11.png slide obj #11 | slide text box obj #11 'mMDTHF 1482.9 – 1486.7' beside SAND: E-14 | original metres; m MDTHF as printed |
| `components[zone-e-14-1].bottomMd` | 4877.625 | slide 5 image11.png slide obj #11 | slide text box obj #11 'mMDTHF 1482.9 – 1486.7' beside SAND: E-14 | 1486.7 m x 3.28084 = 4877.625 ft (3 dp); m MDTHF as printed |
| `components[zone-e-14-1].bottomMdM` | 1486.7 | slide 5 image11.png slide obj #11 | slide text box obj #11 'mMDTHF 1482.9 – 1486.7' beside SAND: E-14 | original metres; m MDTHF as printed |
| `components[zone-e-14-1].zone` | E-14 | slide 5 image11.png slide obj #11 | E-14 | zone label as printed |
| `components[zone-e-14-2].name` | E-14 1487.4-1496.9 m MDTHF | slide 5 image11.png slide obj #11 | E-14 1487.4-1496.9 m MDTHF | visual transcription of the screenshot pictures (x3 upscale, upper part image11 / lower part image12) |
| `components[zone-e-14-2].type` | perforation | slide 5 image11.png slide obj #11 | E-14 1487.4-1496.9 m MDTHF | EXACT TYPE (Component_Types) |
| `components[zone-e-14-2].topMd` | 4879.921 | slide 5 image11.png slide obj #11 | slide text box obj #11 'mMDTHF 1487.4 – 1496.9' beside SAND: E-14 | 1487.4 m x 3.28084 = 4879.921 ft (3 dp); m MDTHF as printed |
| `components[zone-e-14-2].topMdM` | 1487.4 | slide 5 image11.png slide obj #11 | slide text box obj #11 'mMDTHF 1487.4 – 1496.9' beside SAND: E-14 | original metres; m MDTHF as printed |
| `components[zone-e-14-2].bottomMd` | 4911.089 | slide 5 image11.png slide obj #11 | slide text box obj #11 'mMDTHF 1487.4 – 1496.9' beside SAND: E-14 | 1496.9 m x 3.28084 = 4911.089 ft (3 dp); m MDTHF as printed |
| `components[zone-e-14-2].bottomMdM` | 1496.9 | slide 5 image11.png slide obj #11 | slide text box obj #11 'mMDTHF 1487.4 – 1496.9' beside SAND: E-14 | original metres; m MDTHF as printed |
| `components[zone-e-14-2].zone` | E-14 | slide 5 image11.png slide obj #11 | E-14 | zone label as printed |
| `components[pbtd-1].name` | PBTD : 1800 mMDDF (header) | slide 5 image11.png diagram | PBTD : 1800 mMDDF (header) | visual transcription of the screenshot pictures (x3 upscale, upper part image11 / lower part image12) |
| `components[pbtd-1].type` | pbtd | slide 5 image11.png diagram | PBTD : 1800 mMDDF (header) | EXACT TYPE (Component_Types) |
| `components[pbtd-1].topMd` | 5905.512 | slide 5 image11.png diagram | PBTD : 1800 mMDDF (header) | 1800.0 m x 3.28084 = 5905.512 ft (3 dp); m MDDF |
| `components[pbtd-1].topMdM` | 1800.0 | slide 5 image11.png diagram | PBTD : 1800 mMDDF (header) | original metres; m MDDF |
| `wellInfo.completionDate` | 1991-03-12 | slide 5 image11.png header | 12 MARCH 1991 | ISO date () |
| `wellInfo.completionFluid` | - |  |  | not supported by the source -> template empty value ('-' text / null number) |
| `wellInfo.completionType` | SINGLE OIL PRODUCER (TYPE 1) | slide 5 image11.png header | SINGLE OIL PRODUCER (TYPE 1) | as printed |
| `wellInfo.groundElevation` | None |  |  | not supported by the source -> template empty value ('-' text / null number) |
| `wellInfo.kbElevation` | None |  |  | not supported by the source -> template empty value ('-' text / null number) |
| `wellInfo.maxDeviation` | 48.0 DEG @ 812.80 MDDF | slide 5 image11.png header | 48.0 DEG @ 812.80 MDDF | as printed |
| `wellInfo.name` | DULANG-A6 | slide 5 | WELL A-6 : SINGLE OIL PRODUCER (TYPE 1) | identifier rule DULANG-A<no><suffix>; original label kept in the summary |
| `wellInfo.packerFluid` | 9.6 PPG 3% KCl + NaCl + 0.05% NALCO 3900 BIOCIDE & CORR. INHIBITOR | slide 5 image11.png header | 9.6 PPG 3% KCl + NaCl + 0.05% NALCO 3900 BIOCIDE & CORR. INHIBITOR | as printed |
| `wellInfo.rig` | TIOMAN 1 | slide 5 image11.png header | TIOMAN 1 | as printed |
| `wellInfo.surfaceCasing` | - |  |  | not supported by the source -> template empty value ('-' text / null number) |
| `wellInfo.totalDepth` | 5970.801 | slide 5 image11.png header/diagram | TD @ 1819.9 m MDTHF | 1819.9 m x 3.28084 = 5970.801 ft; datum MDTHF as printed (not shifted) |
| `wellInfo.trajectoryType` | - |  |  | not supported by the source -> template empty value ('-' text / null number) |
| `wellInfo.tubing` | 2 7/8" N-80 NK3SB | slide 5 image11.png header | 2 7/8" N-80 NK3SB | as printed |
| `wellInfo.wellhead` | - |  |  | not supported by the source -> template empty value ('-' text / null number) |

## DULANG-A8 — slide 7

Original label: WELL A-8 : SINGLE SELECTIVE OIL PRODUCER (TYPE 3)
Equipment depth datum: m MDTHF (column headers 'DEPTH m MDTHF')  ·  source: image16.png + image17.png (visual transcription of the screenshot pictures (x3 upscale, upper part image16 / lower part image17))

### Conflicts / missing / unclear

| kind | where | detail |
|---|---|---|
| Conflict | header / diagram | PBTD 1937 m MDDF (header) vs PBTD @ 2282 m (diagram, datum not printed) — diagram value used, header recorded; 7" shoe in the header (1965.2 m MDDF) is shallower than TD 2329 m MDTHF — header casing depths may be outdated |
| Unclear | slide 7 | zone intervals LD1 1698-1706 and E7 1804-1818 printed without unit/datum in slide text boxes; read as m MDTHF like the depth columns |
| Unclear | slide 7 | statuses DV (Mar/07), GLV (Jan/07), GLOV 16/64" (Jan/07), OPEN/CLOSED JAN 2009 and 'CA' Plug (Jan 2009) are slide text boxes matched to rows by vertical position |

### Unmapped content (kept here, not in the JSON)

| where | text | reason |
|---|---|---|
| slide 7 image16.png header | RTE / drill floor to tubing hanger : 14.13 m | offset, not an elevation above datum -> kbElevation left empty (no datum shift) |
| picture p1 | 1/4" CONTROL LINE (to the SCSSV) | note (status / history / survey / other revision) |
| picture p2 | 7" 23 PPF CASING label on the diagram (no depth) | note (status / history / survey / other revision) |
| picture p1 | SAND: LOWER D1 (label beside the blast joint) | note (status / history / survey / other revision) |
| picture p2 | SAND: E-7 (HIGH GOR) E10/11 Add perf | note (status / history / survey / other revision) |
| picture p1 | red note: 10.01.07 GLV SET AT 143?m, SPM #3; Type: TR1 18.06; Port size: 8/16"; Ptro: 300psi; Model: RIPO | note (status / history / survey / other revision) |
| slide 7 obj #13 | retrieve 2.31" CSA @141m-MDTHF (retrieved — not installed equipment) | note (status / history / survey / other revision) |
| slide 7 obj #24 | 'B' Tbg stop + SC (195psi) set @ TRSV profile. TR flapper lock open in April 2015 | note (status / history / survey / other revision) |
| slide 7 obj #12 | 2 suspected leak points @ 16m & 26m from MIT 2021 | note (status / history / survey / other revision) |
| slide 7 obj #9 | Mar 2024 : Fluid level @ 1432 WLTHF (wireline depth — not used as MD) | note (status / history / survey / other revision) |
| slide 7 obj #6 | HUD@1726mMDTHF (Apr 2023). Suspect fine sand | note (status / history / survey / other revision) |
| slide 7 obj #10 | 07/Aug/2001 | note (status / history / survey / other revision) |
| slide 7 obj #31 | Updated by Pravin Nair Feb 2023 | note (status / history / survey / other revision) |
| slide 7 obj #18 | FB – N3C | slide text box over/beside the picture (status / history / note) |
| slide 7 obj #7 | OPEN JAN 2009 | slide text box over/beside the picture (status / history / note) |
| slide 7 obj #8 | CLOSED JAN 2009 | slide text box over/beside the picture (status / history / note) |
| slide 7 obj #10 | 07/Aug/2001 | slide text box over/beside the picture (status / history / note) |
| slide 7 obj #13 | retrieve 2.31" CSA @141m-MDTHF | slide text box over/beside the picture (status / history / note) |
| slide 7 obj #15 | DV (Mar/07) | slide text box over/beside the picture (status / history / note) |
| slide 7 obj #16 | GLV (Jan/07) | slide text box over/beside the picture (status / history / note) |
| slide 7 obj #19 | GLOV 16/64” (Jan/07) | slide text box over/beside the picture (status / history / note) |
| slide 7 obj #20 | DV (Mar/07) | slide text box over/beside the picture (status / history / note) |
| slide 7 obj #21 | ‘CA’ Plug (Jan 2009) | slide text box over/beside the picture (status / history / note) |
| slide 7 obj #25 | LD1 1698 - 1706 | slide text box over/beside the picture (status / history / note) |
| slide 7 obj #26 | E7 – 1804– 1818 | slide text box over/beside the picture (status / history / note) |
| slide 7 obj #6 | HUD@1726mMDTHF (Apr 2023). Suspect fine sand | slide text box over/beside the picture (status / history / note) |
| slide 7 obj #31 | Updated by Pravin Nair Feb 2023 | slide text box over/beside the picture (status / history / note) |
| slide 7 obj #24 | ‘B’ Tbg stop + SC (195psi) set @ TRSV profile. TR flapper lock open in April 2015 | slide text box over/beside the picture (status / history / note) |
| slide 7 obj #12 | 2 suspected leak points @ 16m & 26m from MIT 2021 | slide text box over/beside the picture (status / history / note) |
| slide 7 obj #9 | Mar 2024 : Fluid level @ 1432 WLTHF | slide text box over/beside the picture (status / history / note) |

### Traceability (every populated value)

| JSON path | value | slide / object / location | original text | conversion / interpretation |
|---|---|---|---|---|
| `components[christmas-tree-1].name` | SINGLE INGRAM CACTUS, 3M | slide 7 image16.png header | SINGLE INGRAM CACTUS, 3M | visual transcription of the screenshot pictures (x3 upscale, upper part image16 / lower part image17) |
| `components[christmas-tree-1].type` | christmas-tree | slide 7 image16.png header | SINGLE INGRAM CACTUS, 3M | EXACT TYPE (Component_Types) |
| `components[flow-coupling-1].name` | FLOW COUPLING | slide 7 image16.png (upper picture) | FLOW COUPLING | visual transcription of the screenshot pictures (x3 upscale, upper part image16 / lower part image17) |
| `components[flow-coupling-1].type` | flow-coupling | slide 7 image16.png (upper picture) | FLOW COUPLING | EXACT TYPE (Component_Types) |
| `components[flow-coupling-1].topMd` | 459.646 | slide 7 image16.png (upper picture) | 140.1 | 140.1 m x 3.28084 = 459.646 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[flow-coupling-1].topMdM` | 140.1 | slide 7 image16.png (upper picture) | 140.1 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[flow-coupling-1].innerDiameter` | 2.441 | slide 7 image16.png (upper picture) | 2.441 | column 'MIN ID' = explicit ID |
| `components[scssv-1].name` | TRDP-4A SCSSV | slide 7 image16.png (upper picture) | TRDP-4A SCSSV | visual transcription of the screenshot pictures (x3 upscale, upper part image16 / lower part image17) |
| `components[scssv-1].type` | scssv | slide 7 image16.png (upper picture) | TRDP-4A SCSSV | EXACT TYPE (Component_Types) |
| `components[scssv-1].topMd` | 461.942 | slide 7 image16.png (upper picture) | 140.8 | 140.8 m x 3.28084 = 461.942 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[scssv-1].topMdM` | 140.8 | slide 7 image16.png (upper picture) | 140.8 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[scssv-1].innerDiameter` | 2.312 | slide 7 image16.png (upper picture) | 2.312 | column 'MIN ID' = explicit ID |
| `components[landing-nipple-1].name` | WP-1 NIPPLE | slide 7 image16.png (upper picture) | WP-1 NIPPLE | visual transcription of the screenshot pictures (x3 upscale, upper part image16 / lower part image17) |
| `components[landing-nipple-1].type` | landing-nipple | slide 7 image16.png (upper picture) | WP-1 NIPPLE | NEAREST TYPE (flagged) |
| `components[landing-nipple-1].topMd` | 969.16 | slide 7 image16.png (upper picture) | 295.4 | 295.4 m x 3.28084 = 969.16 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[landing-nipple-1].topMdM` | 295.4 | slide 7 image16.png (upper picture) | 295.4 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[landing-nipple-1].innerDiameter` | 2.312 | slide 7 image16.png (upper picture) | 2.312 | column 'MIN ID' = explicit ID |
| `components[spm-1].name` | KBMG-M-SPM | slide 7 image16.png (upper picture) | KBMG-M-SPM | visual transcription of the screenshot pictures (x3 upscale, upper part image16 / lower part image17) |
| `components[spm-1].type` | spm | slide 7 image16.png (upper picture) | KBMG-M-SPM | EXACT TYPE (Component_Types) |
| `components[spm-1].topMd` | 1840.223 | slide 7 image16.png (upper picture) | 560.9 | 560.9 m x 3.28084 = 1840.223 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[spm-1].topMdM` | 560.9 | slide 7 image16.png (upper picture) | 560.9 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[spm-1].innerDiameter` | 2.371 | slide 7 image16.png (upper picture) | 2.371 | column 'MIN ID' = explicit ID |
| `components[spm-2].name` | KBMG-M-SPM | slide 7 image16.png (upper picture) | KBMG-M-SPM | visual transcription of the screenshot pictures (x3 upscale, upper part image16 / lower part image17) |
| `components[spm-2].type` | spm | slide 7 image16.png (upper picture) | KBMG-M-SPM | EXACT TYPE (Component_Types) |
| `components[spm-2].topMd` | 3687.008 | slide 7 image16.png (upper picture) | 1123.8 | 1123.8 m x 3.28084 = 3687.008 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[spm-2].topMdM` | 1123.8 | slide 7 image16.png (upper picture) | 1123.8 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[spm-2].innerDiameter` | 2.371 | slide 7 image16.png (upper picture) | 2.371 | column 'MIN ID' = explicit ID |
| `components[spm-3].name` | KBMG-M-SPM | slide 7 image16.png (upper picture) | KBMG-M-SPM | visual transcription of the screenshot pictures (x3 upscale, upper part image16 / lower part image17) |
| `components[spm-3].type` | spm | slide 7 image16.png (upper picture) | KBMG-M-SPM | EXACT TYPE (Component_Types) |
| `components[spm-3].topMd` | 4683.727 | slide 7 image16.png (upper picture) | 1427.6 | 1427.6 m x 3.28084 = 4683.727 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[spm-3].topMdM` | 1427.6 | slide 7 image16.png (upper picture) | 1427.6 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[spm-3].innerDiameter` | 2.371 | slide 7 image16.png (upper picture) | 2.371 | column 'MIN ID' = explicit ID |
| `components[spm-4].name` | KBMG-M-SPM | slide 7 image16.png (upper picture) | KBMG-M-SPM | visual transcription of the screenshot pictures (x3 upscale, upper part image16 / lower part image17) |
| `components[spm-4].type` | spm | slide 7 image16.png (upper picture) | KBMG-M-SPM | EXACT TYPE (Component_Types) |
| `components[spm-4].topMd` | 5149.606 | slide 7 image16.png (upper picture) | 1569.6 | 1569.6 m x 3.28084 = 5149.606 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[spm-4].topMdM` | 1569.6 | slide 7 image16.png (upper picture) | 1569.6 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[spm-4].innerDiameter` | 2.371 | slide 7 image16.png (upper picture) | 2.371 | column 'MIN ID' = explicit ID |
| `components[sliding-sleeve-1].name` | WP-2DE SSD | slide 7 image16.png (upper picture) | WP-2DE SSD | visual transcription of the screenshot pictures (x3 upscale, upper part image16 / lower part image17) |
| `components[sliding-sleeve-1].type` | sliding-sleeve | slide 7 image16.png (upper picture) | WP-2DE SSD | EXACT TYPE (Component_Types) |
| `components[sliding-sleeve-1].topMd` | 5490.486 | slide 7 image16.png (upper picture) | 1673.5 | 1673.5 m x 3.28084 = 5490.486 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[sliding-sleeve-1].topMdM` | 1673.5 | slide 7 image16.png (upper picture) | 1673.5 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[sliding-sleeve-1].innerDiameter` | 2.312 | slide 7 image16.png (upper picture) | 2.312 | column 'MIN ID' = explicit ID |
| `components[packer-1].name` | SINGLE HYD. PKR GUIB UNI-VII | slide 7 image16.png (upper picture) | SINGLE HYD. PKR GUIB UNI-VII | visual transcription of the screenshot pictures (x3 upscale, upper part image16 / lower part image17) |
| `components[packer-1].type` | packer | slide 7 image16.png (upper picture) | SINGLE HYD. PKR GUIB UNI-VII | EXACT TYPE (Component_Types) |
| `components[packer-1].topMd` | 5517.061 | slide 7 image16.png (upper picture) | 1681.6 | 1681.6 m x 3.28084 = 5517.061 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[packer-1].topMdM` | 1681.6 | slide 7 image16.png (upper picture) | 1681.6 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[packer-1].innerDiameter` | 2.441 | slide 7 image16.png (upper picture) | 2.441 | column 'MIN ID' = explicit ID |
| `components[sliding-sleeve-2].name` | WP-2DE SSD | slide 7 image16.png (upper picture) | WP-2DE SSD | visual transcription of the screenshot pictures (x3 upscale, upper part image16 / lower part image17) |
| `components[sliding-sleeve-2].type` | sliding-sleeve | slide 7 image16.png (upper picture) | WP-2DE SSD | EXACT TYPE (Component_Types) |
| `components[sliding-sleeve-2].topMd` | 5541.011 | slide 7 image16.png (upper picture) | 1688.9 | 1688.9 m x 3.28084 = 5541.011 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[sliding-sleeve-2].topMdM` | 1688.9 | slide 7 image16.png (upper picture) | 1688.9 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[sliding-sleeve-2].innerDiameter` | 2.312 | slide 7 image16.png (upper picture) | 2.312 | column 'MIN ID' = explicit ID |
| `components[blast-joint-1].name` | BLAST JOINT | slide 7 image16.png (upper picture) | BLAST JOINT | visual transcription of the screenshot pictures (x3 upscale, upper part image16 / lower part image17) |
| `components[blast-joint-1].type` | blast-joint | slide 7 image16.png (upper picture) | BLAST JOINT | EXACT TYPE (Component_Types) |
| `components[blast-joint-1].topMd` | 5553.806 | slide 7 image16.png (upper picture) | 1692.8 | 1692.8 m x 3.28084 = 5553.806 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[blast-joint-1].topMdM` | 1692.8 | slide 7 image16.png (upper picture) | 1692.8 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[blast-joint-1].bottomMd` | 5611.877 | slide 7 image16.png (upper picture) | 1710.5 | 1710.5 m x 3.28084 = 5611.877 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[blast-joint-1].bottomMdM` | 1710.5 | slide 7 image16.png (upper picture) | 1710.5 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[blast-joint-1].innerDiameter` | 2.441 | slide 7 image16.png (upper picture) | 2.441 | column 'MIN ID' = explicit ID |
| `components[landing-nipple-2].name` | WP-1 NIPPLE | slide 7 image17.png (lower picture) | WP-1 NIPPLE | visual transcription of the screenshot pictures (x3 upscale, upper part image16 / lower part image17) |
| `components[landing-nipple-2].type` | landing-nipple | slide 7 image17.png (lower picture) | WP-1 NIPPLE | NEAREST TYPE (flagged) |
| `components[landing-nipple-2].topMd` | 5660.105 | slide 7 image17.png (lower picture) | 1725.2 | 1725.2 m x 3.28084 = 5660.105 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[landing-nipple-2].topMdM` | 1725.2 | slide 7 image17.png (lower picture) | 1725.2 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[landing-nipple-2].innerDiameter` | 2.312 | slide 7 image17.png (lower picture) | 2.312 | column 'MIN ID' = explicit ID |
| `components[pup-joint-1].name` | CENTRALIZER | slide 7 image17.png (lower picture) | CENTRALIZER | visual transcription of the screenshot pictures (x3 upscale, upper part image16 / lower part image17) |
| `components[pup-joint-1].type` | pup-joint | slide 7 image17.png (lower picture) | CENTRALIZER | UNRESOLVED / NEAREST TYPE (flagged) |
| `components[pup-joint-1].topMd` | 5677.166 | slide 7 image17.png (lower picture) | 1730.4 | 1730.4 m x 3.28084 = 5677.166 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[pup-joint-1].topMdM` | 1730.4 | slide 7 image17.png (lower picture) | 1730.4 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[pup-joint-1].innerDiameter` | 2.441 | slide 7 image17.png (lower picture) | 2.441 | column 'MIN ID' = explicit ID |
| `components[seal-assembly-1].name` | LOC. SEAL ASSY. | slide 7 image17.png (lower picture) | LOC. SEAL ASSY. | visual transcription of the screenshot pictures (x3 upscale, upper part image16 / lower part image17) |
| `components[seal-assembly-1].type` | seal-assembly | slide 7 image17.png (lower picture) | LOC. SEAL ASSY. | EXACT TYPE (Component_Types) |
| `components[seal-assembly-1].topMd` | 5682.415 | slide 7 image17.png (lower picture) | 1732.0 | 1732.0 m x 3.28084 = 5682.415 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[seal-assembly-1].topMdM` | 1732.0 | slide 7 image17.png (lower picture) | 1732.0 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[seal-assembly-1].innerDiameter` | 2.441 | slide 7 image17.png (lower picture) | 2.441 | column 'MIN ID' = explicit ID |
| `components[packer-2].name` | GUIB MAGNUM 'GT' PACKER | slide 7 image17.png (lower picture) | GUIB MAGNUM 'GT' PACKER | visual transcription of the screenshot pictures (x3 upscale, upper part image16 / lower part image17) |
| `components[packer-2].type` | packer | slide 7 image17.png (lower picture) | GUIB MAGNUM 'GT' PACKER | EXACT TYPE (Component_Types) |
| `components[packer-2].topMd` | 5685.368 | slide 7 image17.png (lower picture) | 1732.9 | 1732.9 m x 3.28084 = 5685.368 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[packer-2].topMdM` | 1732.9 | slide 7 image17.png (lower picture) | 1732.9 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[packer-2].innerDiameter` | 4.0 | slide 7 image17.png (lower picture) | 4.000 | column 'MIN ID' = explicit ID |
| `components[pbr-1].name` | SEAL BORE EXT | slide 7 image17.png (lower picture) | SEAL BORE EXT | visual transcription of the screenshot pictures (x3 upscale, upper part image16 / lower part image17) |
| `components[pbr-1].type` | pbr | slide 7 image17.png (lower picture) | SEAL BORE EXT | NEAREST TYPE (flagged) |
| `components[pbr-1].topMd` | 5687.336 | slide 7 image17.png (lower picture) | 1733.5 | 1733.5 m x 3.28084 = 5687.336 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[pbr-1].topMdM` | 1733.5 | slide 7 image17.png (lower picture) | 1733.5 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[pbr-1].innerDiameter` | 4.0 | slide 7 image17.png (lower picture) | 4.000 | column 'MIN ID' = explicit ID |
| `components[nipple-no-go-1].name` | 'D' NO-GO NIPPLE | slide 7 image17.png (lower picture) | 'D' NO-GO NIPPLE | visual transcription of the screenshot pictures (x3 upscale, upper part image16 / lower part image17) |
| `components[nipple-no-go-1].type` | nipple (no-go) | slide 7 image17.png (lower picture) | 'D' NO-GO NIPPLE | EXACT TYPE (Component_Types) |
| `components[nipple-no-go-1].topMd` | 5708.99 | slide 7 image17.png (lower picture) | 1740.1 | 1740.1 m x 3.28084 = 5708.99 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[nipple-no-go-1].topMdM` | 1740.1 | slide 7 image17.png (lower picture) | 1740.1 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[nipple-no-go-1].innerDiameter` | 2.25 | slide 7 image17.png (lower picture) | 2.250 | column 'MIN ID' = explicit ID |
| `components[mule-shoe-1].name` | WEG | slide 7 image17.png (lower picture) | WEG | visual transcription of the screenshot pictures (x3 upscale, upper part image16 / lower part image17) |
| `components[mule-shoe-1].type` | mule-shoe | slide 7 image17.png (lower picture) | WEG | NEAREST TYPE (flagged) |
| `components[mule-shoe-1].topMd` | 5709.974 | slide 7 image17.png (lower picture) | 1740.4 | 1740.4 m x 3.28084 = 5709.974 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[mule-shoe-1].topMdM` | 1740.4 | slide 7 image17.png (lower picture) | 1740.4 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[mule-shoe-1].innerDiameter` | 2.441 | slide 7 image17.png (lower picture) | 2.441 | column 'MIN ID' = explicit ID |
| `components[production-casing-1].name` | 13 3/8" K-55 @ 370 m MDDF | slide 7 image16.png header | 13 3/8" K-55 @ 370 m MDDF | visual transcription of the screenshot pictures (x3 upscale, upper part image16 / lower part image17) |
| `components[production-casing-1].type` | casing | slide 7 image16.png header | 13 3/8" K-55 @ 370 m MDDF | EXACT TYPE (Component_Types) |
| `components[production-casing-1].bottomMd` | 1213.911 | slide 7 image16.png header | 13 3/8" K-55 @ 370 m MDDF | 370.0 m x 3.28084 = 1213.911 ft (3 dp); m MDDF as printed |
| `components[production-casing-1].bottomMdM` | 370.0 | slide 7 image16.png header | 13 3/8" K-55 @ 370 m MDDF | original metres; m MDDF as printed |
| `components[production-casing-1].innerDiameter` | 13.375 | slide 7 image16.png header | 13.375 | USER-DEFINED RULE: header size without ID/OD designation -> innerDiameter |
| `components[production-casing-1].grade` | K-55 | slide 7 image16.png header | 13 3/8" K-55 @ 370 m MDDF | parsed from the label |
| `components[production-casing-2].name` | 9 5/8" N-80 @ 945.0 m MDDF | slide 7 image16.png header | 9 5/8" N-80 @ 945.0 m MDDF | visual transcription of the screenshot pictures (x3 upscale, upper part image16 / lower part image17) |
| `components[production-casing-2].type` | casing | slide 7 image16.png header | 9 5/8" N-80 @ 945.0 m MDDF | EXACT TYPE (Component_Types) |
| `components[production-casing-2].bottomMd` | 3100.394 | slide 7 image16.png header | 9 5/8" N-80 @ 945.0 m MDDF | 945.0 m x 3.28084 = 3100.394 ft (3 dp); m MDDF as printed |
| `components[production-casing-2].bottomMdM` | 945.0 | slide 7 image16.png header | 9 5/8" N-80 @ 945.0 m MDDF | original metres; m MDDF as printed |
| `components[production-casing-2].innerDiameter` | 9.625 | slide 7 image16.png header | 9.625 | USER-DEFINED RULE: header size without ID/OD designation -> innerDiameter |
| `components[production-casing-2].grade` | N-80 | slide 7 image16.png header | 9 5/8" N-80 @ 945.0 m MDDF | parsed from the label |
| `components[production-casing-3].name` | 7" 23 ppf, K-55 @ 1965.2 m MDDF | slide 7 image16.png header | 7" 23 ppf, K-55 @ 1965.2 m MDDF | visual transcription of the screenshot pictures (x3 upscale, upper part image16 / lower part image17) |
| `components[production-casing-3].type` | casing | slide 7 image16.png header | 7" 23 ppf, K-55 @ 1965.2 m MDDF | EXACT TYPE (Component_Types) |
| `components[production-casing-3].bottomMd` | 6447.507 | slide 7 image16.png header | 7" 23 ppf, K-55 @ 1965.2 m MDDF | 1965.2 m x 3.28084 = 6447.507 ft (3 dp); m MDDF as printed |
| `components[production-casing-3].bottomMdM` | 1965.2 | slide 7 image16.png header | 7" 23 ppf, K-55 @ 1965.2 m MDDF | original metres; m MDDF as printed |
| `components[production-casing-3].innerDiameter` | 7.0 | slide 7 image16.png header | 7 | USER-DEFINED RULE: header size without ID/OD designation -> innerDiameter |
| `components[production-casing-3].weight` | 23.0 | slide 7 image16.png header | 7" 23 ppf, K-55 @ 1965.2 m MDDF | parsed from the label (ppf = lb/ft) |
| `components[production-casing-3].grade` | K-55 | slide 7 image16.png header | 7" 23 ppf, K-55 @ 1965.2 m MDDF | parsed from the label |
| `components[tubing-1].name` | 2 7/8" 6.5 PPF N-80 NK3SB | slide 7 image16.png header | 2 7/8" 6.5 PPF N-80 NK3SB | visual transcription of the screenshot pictures (x3 upscale, upper part image16 / lower part image17) |
| `components[tubing-1].type` | tubing | slide 7 image16.png header | 2 7/8" 6.5 PPF N-80 NK3SB | EXACT TYPE (Component_Types) |
| `components[tubing-1].topMd` | 0.0 | slide 7 image16.png header | depth reference m MDTHF / THF | 0.0 m x 3.28084 = 0.0 ft (3 dp); ASSUMED 0 m (tubing hangs from the tubing hanger flange) |
| `components[tubing-1].topMdM` | 0.0 | slide 7 image16.png header | depth reference m MDTHF / THF | original metres; ASSUMED 0 m (tubing hangs from the tubing hanger flange) |
| `components[tubing-1].bottomMd` | 5709.974 | slide 7 image16.png header | WEG | 1740.4 m x 3.28084 = 5709.974 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF'); depth of 'WEG' (mule-shoe-1) |
| `components[tubing-1].bottomMdM` | 1740.4 | slide 7 image16.png header | WEG | original metres; m MDTHF (column headers 'DEPTH m MDTHF'); depth of 'WEG' (mule-shoe-1) |
| `components[tubing-1].innerDiameter` | 2.875 | slide 7 image16.png header | 2 7/8" 6.5 PPF N-80 NK3SB | USER-DEFINED RULE: header size without ID/OD designation -> innerDiameter |
| `components[tubing-1].weight` | 6.5 | slide 7 image16.png header | 2 7/8" 6.5 PPF N-80 NK3SB | parsed from the label (ppf = lb/ft) |
| `components[tubing-1].grade` | N-80 | slide 7 image16.png header | 2 7/8" 6.5 PPF N-80 NK3SB | parsed from the label |
| `components[zone-ld1-1].name` | LD1 1698-1706 m MDTHF | slide 7 image16.png slide obj #25 | LD1 1698-1706 m MDTHF | visual transcription of the screenshot pictures (x3 upscale, upper part image16 / lower part image17) |
| `components[zone-ld1-1].type` | perforation | slide 7 image16.png slide obj #25 | LD1 1698-1706 m MDTHF | EXACT TYPE (Component_Types) |
| `components[zone-ld1-1].topMd` | 5570.866 | slide 7 image16.png slide obj #25 | slide text box obj #25 'LD1 1698 - 1706' beside SAND: LOWER D1 (no unit/datum printed) | 1698.0 m x 3.28084 = 5570.866 ft (3 dp); m MDTHF as printed |
| `components[zone-ld1-1].topMdM` | 1698.0 | slide 7 image16.png slide obj #25 | slide text box obj #25 'LD1 1698 - 1706' beside SAND: LOWER D1 (no unit/datum printed) | original metres; m MDTHF as printed |
| `components[zone-ld1-1].bottomMd` | 5597.113 | slide 7 image16.png slide obj #25 | slide text box obj #25 'LD1 1698 - 1706' beside SAND: LOWER D1 (no unit/datum printed) | 1706.0 m x 3.28084 = 5597.113 ft (3 dp); m MDTHF as printed |
| `components[zone-ld1-1].bottomMdM` | 1706.0 | slide 7 image16.png slide obj #25 | slide text box obj #25 'LD1 1698 - 1706' beside SAND: LOWER D1 (no unit/datum printed) | original metres; m MDTHF as printed |
| `components[zone-ld1-1].zone` | LD1 | slide 7 image16.png slide obj #25 | LD1 | zone label as printed |
| `components[zone-e7-1].name` | E7 1804-1818 m MDTHF | slide 7 image16.png slide obj #26 | E7 1804-1818 m MDTHF | visual transcription of the screenshot pictures (x3 upscale, upper part image16 / lower part image17) |
| `components[zone-e7-1].type` | perforation | slide 7 image16.png slide obj #26 | E7 1804-1818 m MDTHF | EXACT TYPE (Component_Types) |
| `components[zone-e7-1].topMd` | 5918.635 | slide 7 image16.png slide obj #26 | slide text box obj #26 'E7 – 1804– 1818' beside SAND: E-7 (HIGH GOR) E10/11 Add perf (no unit/datum printed) | 1804.0 m x 3.28084 = 5918.635 ft (3 dp); m MDTHF as printed |
| `components[zone-e7-1].topMdM` | 1804.0 | slide 7 image16.png slide obj #26 | slide text box obj #26 'E7 – 1804– 1818' beside SAND: E-7 (HIGH GOR) E10/11 Add perf (no unit/datum printed) | original metres; m MDTHF as printed |
| `components[zone-e7-1].bottomMd` | 5964.567 | slide 7 image16.png slide obj #26 | slide text box obj #26 'E7 – 1804– 1818' beside SAND: E-7 (HIGH GOR) E10/11 Add perf (no unit/datum printed) | 1818.0 m x 3.28084 = 5964.567 ft (3 dp); m MDTHF as printed |
| `components[zone-e7-1].bottomMdM` | 1818.0 | slide 7 image16.png slide obj #26 | slide text box obj #26 'E7 – 1804– 1818' beside SAND: E-7 (HIGH GOR) E10/11 Add perf (no unit/datum printed) | original metres; m MDTHF as printed |
| `components[zone-e7-1].zone` | E7 | slide 7 image16.png slide obj #26 | E7 | zone label as printed |
| `components[pbtd-1].name` | PBTD @ 2282 m (diagram; header: PBTD : 1937 mMDDF) | slide 7 image16.png diagram | PBTD @ 2282 m (diagram; header: PBTD : 1937 mMDDF) | visual transcription of the screenshot pictures (x3 upscale, upper part image16 / lower part image17) |
| `components[pbtd-1].type` | pbtd | slide 7 image16.png diagram | PBTD @ 2282 m (diagram; header: PBTD : 1937 mMDDF) | EXACT TYPE (Component_Types) |
| `components[pbtd-1].topMd` | 7486.877 | slide 7 image16.png diagram | PBTD @ 2282 m (diagram; header: PBTD : 1937 mMDDF) | 2282.0 m x 3.28084 = 7486.877 ft (3 dp); m datum not printed |
| `components[pbtd-1].topMdM` | 2282.0 | slide 7 image16.png diagram | PBTD @ 2282 m (diagram; header: PBTD : 1937 mMDDF) | original metres; m datum not printed |
| `wellInfo.completionDate` | 1991-05-07 | slide 7 image16.png header | 7 MAY 1991 | ISO date () |
| `wellInfo.completionFluid` | - |  |  | not supported by the source -> template empty value ('-' text / null number) |
| `wellInfo.completionType` | SINGLE SELECTIVE OIL PRODUCER (TYPE 3) | slide 7 image16.png header | SINGLE SELECTIVE OIL PRODUCER (TYPE 3) | as printed |
| `wellInfo.groundElevation` | None |  |  | not supported by the source -> template empty value ('-' text / null number) |
| `wellInfo.kbElevation` | None |  |  | not supported by the source -> template empty value ('-' text / null number) |
| `wellInfo.maxDeviation` | 58.8 DEG @ 762.10 MDDF | slide 7 image16.png header | 58.8 DEG @ 762.10 MDDF | as printed |
| `wellInfo.name` | DULANG-A8 | slide 7 | WELL A-8 : SINGLE SELECTIVE OIL PRODUCER (TYPE 3) | identifier rule DULANG-A<no><suffix>; original label kept in the summary |
| `wellInfo.packerFluid` | 9.6 PPG 3% KCl + NaCl + 0.05% NALCO 3900 BIOCIDE & CORR. INHIBITOR | slide 7 image16.png header | 9.6 PPG 3% KCl + NaCl + 0.05% NALCO 3900 BIOCIDE & CORR. INHIBITOR | as printed |
| `wellInfo.rig` | TIOMAN 1 | slide 7 image16.png header | TIOMAN 1 | as printed |
| `wellInfo.surfaceCasing` | - |  |  | not supported by the source -> template empty value ('-' text / null number) |
| `wellInfo.totalDepth` | 7641.076 | slide 7 image16.png header/diagram | TD @ 2329 m MDTHF | 2329.0 m x 3.28084 = 7641.076 ft; datum MDTHF as printed (not shifted) |
| `wellInfo.trajectoryType` | - |  |  | not supported by the source -> template empty value ('-' text / null number) |
| `wellInfo.tubing` | 2 7/8" 6.5 PPF N-80 NK3SB | slide 7 image16.png header | 2 7/8" 6.5 PPF N-80 NK3SB | as printed |
| `wellInfo.wellhead` | - |  |  | not supported by the source -> template empty value ('-' text / null number) |

## DULANG-A19 — slide 16

Original label: WELL A-19 : SINGLE SELECTIVE OIL PRODUCER (TYPE 3)
Equipment depth datum: m MDTHF (column headers 'DEPTH m MDTHF')  ·  source: image36.png + image37.png (visual transcription of the screenshot pictures (x3 upscale, upper part image36 / lower part image37))

### Conflicts / missing / unclear

| kind | where | detail |
|---|---|---|
| Missing | slide 16 image36.png header | '20" 94.0 ppf X-56 @ mMDDF': setting depth not printed |
| Missing | slide 16 image36.png header | '13 3/8" K-55 @ mMDDF': setting depth not printed |
| Missing | slide 16 image36.png header | '7" 23 ppf, K-55 @ m MDDF': setting depth not printed |
| Missing | slide 16 | PBTD not printed |
| Missing | header | casing setting depths and PBTD left blank on the schematic ('@ mMDDF', 'PBTD : m MDDF') |
| Missing | picture p2 | GUIB MAGNUM 'GT' PACKER and SEAL BORE EXT have no printed depth (between LOC. SEAL ASSY. 2354 and 'D' NO-GO NIPPLE 2359) |
| Unclear | slide 16 | later statuses (CLOSED – STIFF Aug 2015, OPEN OCT 2023 x2, OPEN Feb 2015) are slide text boxes matched to rows by vertical position |

### Unmapped content (kept here, not in the JSON)

| where | text | reason |
|---|---|---|
| slide 16 image36.png header | RTE / drill floor to tubing hanger : 14.52 m | offset, not an elevation above datum -> kbElevation left empty (no datum shift) |
| picture p1 | 1/4" CONTROL LINE (to the SCSSV) | note (status / history / survey / other revision) |
| picture p2 | 7" 23 PPF CASING label on the diagram (no depth) | note (status / history / survey / other revision) |
| picture p1/p2 | SAND: LOWER D-1 (beside blast joint 2196-2211), SAND: LD-3 (beside blast joint 2257-2266) — intervals not printed | note (status / history / survey / other revision) |
| slide 16 obj #12 | Tubing punch at ~2174m (July 2016) for communication with PCP (Previously installed GLV Straddle Assembly) | note (status / history / survey / other revision) |
| slide 16 obj #14 | Latest Update: Nov 2023 by Pravin Nair | note (status / history / survey / other revision) |
| slide 16 obj #18 | FB – N3C | slide text box over/beside the picture (status / history / note) |
| slide 16 obj #14 | Latest Update: Nov 2023 by Pravin Nair | slide text box over/beside the picture (status / history / note) |
| slide 16 obj #11 | OPEN (Feb 2015) | slide text box over/beside the picture (status / history / note) |
| slide 16 obj #23 | CLOSED – STIFF (unable to open in Aug 2015) | slide text box over/beside the picture (status / history / note) |
| slide 16 obj #12 | Tubing punch at ~2174m (July 2016) for communication with PCP (Previously installed GLV Straddle Assembly) | slide text box over/beside the picture (status / history / note) |
| slide 16 obj #10 | 2386 – 2402m MDTHF | slide text box over/beside the picture (status / history / note) |
| slide 16 obj #7 | OPEN (OCT 2023) | slide text box over/beside the picture (status / history / note) |
| slide 16 obj #9 | OPEN (OCT 2023) | slide text box over/beside the picture (status / history / note) |

### Traceability (every populated value)

| JSON path | value | slide / object / location | original text | conversion / interpretation |
|---|---|---|---|---|
| `components[christmas-tree-1].name` | SINGLE INGRAM CACTUS, 3M | slide 16 image36.png header | SINGLE INGRAM CACTUS, 3M | visual transcription of the screenshot pictures (x3 upscale, upper part image36 / lower part image37) |
| `components[christmas-tree-1].type` | christmas-tree | slide 16 image36.png header | SINGLE INGRAM CACTUS, 3M | EXACT TYPE (Component_Types) |
| `components[flow-coupling-1].name` | FLOW COUPLING | slide 16 image36.png (upper picture) | FLOW COUPLING | visual transcription of the screenshot pictures (x3 upscale, upper part image36 / lower part image37) |
| `components[flow-coupling-1].type` | flow-coupling | slide 16 image36.png (upper picture) | FLOW COUPLING | EXACT TYPE (Component_Types) |
| `components[flow-coupling-1].topMd` | 462.598 | slide 16 image36.png (upper picture) | 141 | 141.0 m x 3.28084 = 462.598 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[flow-coupling-1].topMdM` | 141.0 | slide 16 image36.png (upper picture) | 141 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[flow-coupling-1].innerDiameter` | 2.441 | slide 16 image36.png (upper picture) | 2.441 | column 'MIN ID' = explicit ID |
| `components[scssv-1].name` | TRDP-4A SCSSV | slide 16 image36.png (upper picture) | TRDP-4A SCSSV | visual transcription of the screenshot pictures (x3 upscale, upper part image36 / lower part image37) |
| `components[scssv-1].type` | scssv | slide 16 image36.png (upper picture) | TRDP-4A SCSSV | EXACT TYPE (Component_Types) |
| `components[scssv-1].topMd` | 465.879 | slide 16 image36.png (upper picture) | 142 | 142.0 m x 3.28084 = 465.879 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[scssv-1].topMdM` | 142.0 | slide 16 image36.png (upper picture) | 142 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[scssv-1].innerDiameter` | 2.312 | slide 16 image36.png (upper picture) | 2.312 | column 'MIN ID' = explicit ID |
| `components[landing-nipple-1].name` | WP-1 NIPPLE | slide 16 image36.png (upper picture) | WP-1 NIPPLE | visual transcription of the screenshot pictures (x3 upscale, upper part image36 / lower part image37) |
| `components[landing-nipple-1].type` | landing-nipple | slide 16 image36.png (upper picture) | WP-1 NIPPLE | NEAREST TYPE (flagged) |
| `components[landing-nipple-1].topMd` | 974.409 | slide 16 image36.png (upper picture) | 297 | 297.0 m x 3.28084 = 974.409 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[landing-nipple-1].topMdM` | 297.0 | slide 16 image36.png (upper picture) | 297 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[landing-nipple-1].innerDiameter` | 2.312 | slide 16 image36.png (upper picture) | 2.312 | column 'MIN ID' = explicit ID |
| `components[spm-1].name` | KBMG-M-SPM | slide 16 image36.png (upper picture) | KBMG-M-SPM | visual transcription of the screenshot pictures (x3 upscale, upper part image36 / lower part image37) |
| `components[spm-1].type` | spm | slide 16 image36.png (upper picture) | KBMG-M-SPM | EXACT TYPE (Component_Types) |
| `components[spm-1].topMd` | 2391.732 | slide 16 image36.png (upper picture) | 729 | 729.0 m x 3.28084 = 2391.732 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[spm-1].topMdM` | 729.0 | slide 16 image36.png (upper picture) | 729 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[spm-1].innerDiameter` | 2.371 | slide 16 image36.png (upper picture) | 2.371 | column 'MIN ID' = explicit ID |
| `components[spm-2].name` | KBMG-M-SPM | slide 16 image36.png (upper picture) | KBMG-M-SPM | visual transcription of the screenshot pictures (x3 upscale, upper part image36 / lower part image37) |
| `components[spm-2].type` | spm | slide 16 image36.png (upper picture) | KBMG-M-SPM | EXACT TYPE (Component_Types) |
| `components[spm-2].topMd` | 4668.635 | slide 16 image36.png (upper picture) | 1423 | 1423.0 m x 3.28084 = 4668.635 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[spm-2].topMdM` | 1423.0 | slide 16 image36.png (upper picture) | 1423 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[spm-2].innerDiameter` | 2.371 | slide 16 image36.png (upper picture) | 2.371 | column 'MIN ID' = explicit ID |
| `components[spm-3].name` | KBMG-M-SPM | slide 16 image36.png (upper picture) | KBMG-M-SPM | visual transcription of the screenshot pictures (x3 upscale, upper part image36 / lower part image37) |
| `components[spm-3].type` | spm | slide 16 image36.png (upper picture) | KBMG-M-SPM | EXACT TYPE (Component_Types) |
| `components[spm-3].topMd` | 5961.286 | slide 16 image36.png (upper picture) | 1817 | 1817.0 m x 3.28084 = 5961.286 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[spm-3].topMdM` | 1817.0 | slide 16 image36.png (upper picture) | 1817 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[spm-3].innerDiameter` | 2.371 | slide 16 image36.png (upper picture) | 2.371 | column 'MIN ID' = explicit ID |
| `components[spm-4].name` | KBMG-M-SPM | slide 16 image36.png (upper picture) | KBMG-M-SPM | visual transcription of the screenshot pictures (x3 upscale, upper part image36 / lower part image37) |
| `components[spm-4].type` | spm | slide 16 image36.png (upper picture) | KBMG-M-SPM | EXACT TYPE (Component_Types) |
| `components[spm-4].topMd` | 6594.488 | slide 16 image36.png (upper picture) | 2010 | 2010.0 m x 3.28084 = 6594.488 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[spm-4].topMdM` | 2010.0 | slide 16 image36.png (upper picture) | 2010 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[spm-4].innerDiameter` | 2.371 | slide 16 image36.png (upper picture) | 2.371 | column 'MIN ID' = explicit ID |
| `components[sliding-sleeve-1].name` | WP-2DE SSD | slide 16 image36.png (upper picture) | WP-2DE SSD | visual transcription of the screenshot pictures (x3 upscale, upper part image36 / lower part image37) |
| `components[sliding-sleeve-1].type` | sliding-sleeve | slide 16 image36.png (upper picture) | WP-2DE SSD | EXACT TYPE (Component_Types) |
| `components[sliding-sleeve-1].topMd` | 7132.546 | slide 16 image36.png (upper picture) | 2174 | 2174.0 m x 3.28084 = 7132.546 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[sliding-sleeve-1].topMdM` | 2174.0 | slide 16 image36.png (upper picture) | 2174 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[sliding-sleeve-1].innerDiameter` | 2.312 | slide 16 image36.png (upper picture) | 2.312 | column 'MIN ID' = explicit ID |
| `components[packer-1].name` | SINGLE HYD. PKR GUIB UNI-VII | slide 16 image36.png (upper picture) | SINGLE HYD. PKR GUIB UNI-VII | visual transcription of the screenshot pictures (x3 upscale, upper part image36 / lower part image37) |
| `components[packer-1].type` | packer | slide 16 image36.png (upper picture) | SINGLE HYD. PKR GUIB UNI-VII | EXACT TYPE (Component_Types) |
| `components[packer-1].topMd` | 7155.512 | slide 16 image36.png (upper picture) | 2181 | 2181.0 m x 3.28084 = 7155.512 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[packer-1].topMdM` | 2181.0 | slide 16 image36.png (upper picture) | 2181 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[packer-1].innerDiameter` | 2.441 | slide 16 image36.png (upper picture) | 2.441 | column 'MIN ID' = explicit ID |
| `components[sliding-sleeve-2].name` | WP-2DE SSD | slide 16 image36.png (upper picture) | WP-2DE SSD | visual transcription of the screenshot pictures (x3 upscale, upper part image36 / lower part image37) |
| `components[sliding-sleeve-2].type` | sliding-sleeve | slide 16 image36.png (upper picture) | WP-2DE SSD | EXACT TYPE (Component_Types) |
| `components[sliding-sleeve-2].topMd` | 7191.601 | slide 16 image36.png (upper picture) | 2192 | 2192.0 m x 3.28084 = 7191.601 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[sliding-sleeve-2].topMdM` | 2192.0 | slide 16 image36.png (upper picture) | 2192 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[sliding-sleeve-2].innerDiameter` | 2.312 | slide 16 image36.png (upper picture) | 2.312 | column 'MIN ID' = explicit ID |
| `components[blast-joint-1].name` | BLAST JOINT | slide 16 image36.png (upper picture) | BLAST JOINT | visual transcription of the screenshot pictures (x3 upscale, upper part image36 / lower part image37) |
| `components[blast-joint-1].type` | blast-joint | slide 16 image36.png (upper picture) | BLAST JOINT | EXACT TYPE (Component_Types) |
| `components[blast-joint-1].topMd` | 7204.725 | slide 16 image36.png (upper picture) | 2196 | 2196.0 m x 3.28084 = 7204.725 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[blast-joint-1].topMdM` | 2196.0 | slide 16 image36.png (upper picture) | 2196 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[blast-joint-1].bottomMd` | 7253.937 | slide 16 image36.png (upper picture) | 2211 | 2211.0 m x 3.28084 = 7253.937 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[blast-joint-1].bottomMdM` | 2211.0 | slide 16 image36.png (upper picture) | 2211 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[blast-joint-1].innerDiameter` | 2.441 | slide 16 image36.png (upper picture) | 2.441 | column 'MIN ID' = explicit ID |
| `components[landing-nipple-2].name` | WP-1 NIPPLE | slide 16 image36.png (upper picture) | WP-1 NIPPLE | visual transcription of the screenshot pictures (x3 upscale, upper part image36 / lower part image37) |
| `components[landing-nipple-2].type` | landing-nipple | slide 16 image36.png (upper picture) | WP-1 NIPPLE | NEAREST TYPE (flagged) |
| `components[landing-nipple-2].topMd` | 7345.801 | slide 16 image36.png (upper picture) | 2239 | 2239.0 m x 3.28084 = 7345.801 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[landing-nipple-2].topMdM` | 2239.0 | slide 16 image36.png (upper picture) | 2239 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[landing-nipple-2].innerDiameter` | 2.312 | slide 16 image36.png (upper picture) | 2.312 | column 'MIN ID' = explicit ID |
| `components[packer-2].name` | SINGLE HYD. PKR GUIB UNI-VII | slide 16 image37.png (lower picture) | SINGLE HYD. PKR GUIB UNI-VII | visual transcription of the screenshot pictures (x3 upscale, upper part image36 / lower part image37) |
| `components[packer-2].type` | packer | slide 16 image37.png (lower picture) | SINGLE HYD. PKR GUIB UNI-VII | EXACT TYPE (Component_Types) |
| `components[packer-2].topMd` | 7365.486 | slide 16 image37.png (lower picture) | 2245 | 2245.0 m x 3.28084 = 7365.486 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[packer-2].topMdM` | 2245.0 | slide 16 image37.png (lower picture) | 2245 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[packer-2].innerDiameter` | 2.441 | slide 16 image37.png (lower picture) | 2.441 | column 'MIN ID' = explicit ID |
| `components[sliding-sleeve-3].name` | WP-2DE SSD | slide 16 image37.png (lower picture) | WP-2DE SSD | visual transcription of the screenshot pictures (x3 upscale, upper part image36 / lower part image37) |
| `components[sliding-sleeve-3].type` | sliding-sleeve | slide 16 image37.png (lower picture) | WP-2DE SSD | EXACT TYPE (Component_Types) |
| `components[sliding-sleeve-3].topMd` | 7395.013 | slide 16 image37.png (lower picture) | 2254 | 2254.0 m x 3.28084 = 7395.013 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[sliding-sleeve-3].topMdM` | 2254.0 | slide 16 image37.png (lower picture) | 2254 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[sliding-sleeve-3].innerDiameter` | 2.312 | slide 16 image37.png (lower picture) | 2.312 | column 'MIN ID' = explicit ID |
| `components[blast-joint-2].name` | BLAST JOINT | slide 16 image37.png (lower picture) | BLAST JOINT | visual transcription of the screenshot pictures (x3 upscale, upper part image36 / lower part image37) |
| `components[blast-joint-2].type` | blast-joint | slide 16 image37.png (lower picture) | BLAST JOINT | EXACT TYPE (Component_Types) |
| `components[blast-joint-2].topMd` | 7404.856 | slide 16 image37.png (lower picture) | 2257 | 2257.0 m x 3.28084 = 7404.856 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[blast-joint-2].topMdM` | 2257.0 | slide 16 image37.png (lower picture) | 2257 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[blast-joint-2].bottomMd` | 7434.383 | slide 16 image37.png (lower picture) | 2266 | 2266.0 m x 3.28084 = 7434.383 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[blast-joint-2].bottomMdM` | 2266.0 | slide 16 image37.png (lower picture) | 2266 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[blast-joint-2].innerDiameter` | 2.441 | slide 16 image37.png (lower picture) | 2.441 | column 'MIN ID' = explicit ID |
| `components[landing-nipple-3].name` | WP-1 NIPPLE | slide 16 image37.png (lower picture) | WP-1 NIPPLE | visual transcription of the screenshot pictures (x3 upscale, upper part image36 / lower part image37) |
| `components[landing-nipple-3].type` | landing-nipple | slide 16 image37.png (lower picture) | WP-1 NIPPLE | NEAREST TYPE (flagged) |
| `components[landing-nipple-3].topMd` | 7634.515 | slide 16 image37.png (lower picture) | 2327 | 2327.0 m x 3.28084 = 7634.515 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[landing-nipple-3].topMdM` | 2327.0 | slide 16 image37.png (lower picture) | 2327 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[landing-nipple-3].innerDiameter` | 2.312 | slide 16 image37.png (lower picture) | 2.312 | column 'MIN ID' = explicit ID |
| `components[pup-joint-1].name` | CENTRALIZER | slide 16 image37.png (lower picture) | CENTRALIZER | visual transcription of the screenshot pictures (x3 upscale, upper part image36 / lower part image37) |
| `components[pup-joint-1].type` | pup-joint | slide 16 image37.png (lower picture) | CENTRALIZER | UNRESOLVED / NEAREST TYPE (flagged) |
| `components[pup-joint-1].topMd` | 7719.817 | slide 16 image37.png (lower picture) | 2353 | 2353.0 m x 3.28084 = 7719.817 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[pup-joint-1].topMdM` | 2353.0 | slide 16 image37.png (lower picture) | 2353 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[pup-joint-1].innerDiameter` | 2.441 | slide 16 image37.png (lower picture) | 2.441 | column 'MIN ID' = explicit ID |
| `components[seal-assembly-1].name` | LOC. SEAL ASSY. | slide 16 image37.png (lower picture) | LOC. SEAL ASSY. | visual transcription of the screenshot pictures (x3 upscale, upper part image36 / lower part image37) |
| `components[seal-assembly-1].type` | seal-assembly | slide 16 image37.png (lower picture) | LOC. SEAL ASSY. | EXACT TYPE (Component_Types) |
| `components[seal-assembly-1].topMd` | 7723.097 | slide 16 image37.png (lower picture) | 2354 | 2354.0 m x 3.28084 = 7723.097 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[seal-assembly-1].topMdM` | 2354.0 | slide 16 image37.png (lower picture) | 2354 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[seal-assembly-1].innerDiameter` | 2.441 | slide 16 image37.png (lower picture) | 2.441 | column 'MIN ID' = explicit ID |
| `components[packer-3].name` | GUIB MAGNUM 'GT' PACKER | slide 16 image37.png (lower picture) | GUIB MAGNUM 'GT' PACKER | visual transcription of the screenshot pictures (x3 upscale, upper part image36 / lower part image37) |
| `components[packer-3].type` | packer | slide 16 image37.png (lower picture) | GUIB MAGNUM 'GT' PACKER | EXACT TYPE (Component_Types) |
| `components[packer-3].innerDiameter` | 4.0 | slide 16 image37.png (lower picture) | 4.000 | column 'MIN ID' = explicit ID |
| `components[pbr-1].name` | SEAL BORE EXT | slide 16 image37.png (lower picture) | SEAL BORE EXT | visual transcription of the screenshot pictures (x3 upscale, upper part image36 / lower part image37) |
| `components[pbr-1].type` | pbr | slide 16 image37.png (lower picture) | SEAL BORE EXT | NEAREST TYPE (flagged) |
| `components[pbr-1].innerDiameter` | 4.0 | slide 16 image37.png (lower picture) | 4.000 | column 'MIN ID' = explicit ID |
| `components[nipple-no-go-1].name` | 'D' NO-GO NIPPLE | slide 16 image37.png (lower picture) | 'D' NO-GO NIPPLE | visual transcription of the screenshot pictures (x3 upscale, upper part image36 / lower part image37) |
| `components[nipple-no-go-1].type` | nipple (no-go) | slide 16 image37.png (lower picture) | 'D' NO-GO NIPPLE | EXACT TYPE (Component_Types) |
| `components[nipple-no-go-1].topMd` | 7739.502 | slide 16 image37.png (lower picture) | 2359 | 2359.0 m x 3.28084 = 7739.502 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[nipple-no-go-1].topMdM` | 2359.0 | slide 16 image37.png (lower picture) | 2359 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[nipple-no-go-1].innerDiameter` | 2.25 | slide 16 image37.png (lower picture) | 2.250 | column 'MIN ID' = explicit ID |
| `components[mule-shoe-1].name` | WEG | slide 16 image37.png (lower picture) | WEG | visual transcription of the screenshot pictures (x3 upscale, upper part image36 / lower part image37) |
| `components[mule-shoe-1].type` | mule-shoe | slide 16 image37.png (lower picture) | WEG | NEAREST TYPE (flagged) |
| `components[mule-shoe-1].topMd` | 7742.782 | slide 16 image37.png (lower picture) | 2360 | 2360.0 m x 3.28084 = 7742.782 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[mule-shoe-1].topMdM` | 2360.0 | slide 16 image37.png (lower picture) | 2360 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[mule-shoe-1].innerDiameter` | 2.441 | slide 16 image37.png (lower picture) | 2.441 | column 'MIN ID' = explicit ID |
| `components[production-casing-1].name` | 20" 94.0 ppf X-56 @ mMDDF | slide 16 image36.png header | 20" 94.0 ppf X-56 @ mMDDF | visual transcription of the screenshot pictures (x3 upscale, upper part image36 / lower part image37) |
| `components[production-casing-1].type` | casing | slide 16 image36.png header | 20" 94.0 ppf X-56 @ mMDDF | EXACT TYPE (Component_Types) |
| `components[production-casing-1].innerDiameter` | 20.0 | slide 16 image36.png header | 20 | USER-DEFINED RULE: header size without ID/OD designation -> innerDiameter |
| `components[production-casing-1].weight` | 94.0 | slide 16 image36.png header | 20" 94.0 ppf X-56 @ mMDDF | parsed from the label (ppf = lb/ft) |
| `components[production-casing-1].grade` | X-56 | slide 16 image36.png header | 20" 94.0 ppf X-56 @ mMDDF | parsed from the label |
| `components[production-casing-2].name` | 13 3/8" K-55 @ mMDDF | slide 16 image36.png header | 13 3/8" K-55 @ mMDDF | visual transcription of the screenshot pictures (x3 upscale, upper part image36 / lower part image37) |
| `components[production-casing-2].type` | casing | slide 16 image36.png header | 13 3/8" K-55 @ mMDDF | EXACT TYPE (Component_Types) |
| `components[production-casing-2].innerDiameter` | 13.375 | slide 16 image36.png header | 13.375 | USER-DEFINED RULE: header size without ID/OD designation -> innerDiameter |
| `components[production-casing-2].grade` | K-55 | slide 16 image36.png header | 13 3/8" K-55 @ mMDDF | parsed from the label |
| `components[production-casing-3].name` | 7" 23 ppf, K-55 @ m MDDF | slide 16 image36.png header | 7" 23 ppf, K-55 @ m MDDF | visual transcription of the screenshot pictures (x3 upscale, upper part image36 / lower part image37) |
| `components[production-casing-3].type` | casing | slide 16 image36.png header | 7" 23 ppf, K-55 @ m MDDF | EXACT TYPE (Component_Types) |
| `components[production-casing-3].innerDiameter` | 7.0 | slide 16 image36.png header | 7 | USER-DEFINED RULE: header size without ID/OD designation -> innerDiameter |
| `components[production-casing-3].weight` | 23.0 | slide 16 image36.png header | 7" 23 ppf, K-55 @ m MDDF | parsed from the label (ppf = lb/ft) |
| `components[production-casing-3].grade` | K-55 | slide 16 image36.png header | 7" 23 ppf, K-55 @ m MDDF | parsed from the label |
| `components[tubing-1].name` | 2 7/8" 6.5 PPF N-80 NK3SB | slide 16 image36.png header | 2 7/8" 6.5 PPF N-80 NK3SB | visual transcription of the screenshot pictures (x3 upscale, upper part image36 / lower part image37) |
| `components[tubing-1].type` | tubing | slide 16 image36.png header | 2 7/8" 6.5 PPF N-80 NK3SB | EXACT TYPE (Component_Types) |
| `components[tubing-1].topMd` | 0.0 | slide 16 image36.png header | depth reference m MDTHF / THF | 0.0 m x 3.28084 = 0.0 ft (3 dp); ASSUMED 0 m (tubing hangs from the tubing hanger flange) |
| `components[tubing-1].topMdM` | 0.0 | slide 16 image36.png header | depth reference m MDTHF / THF | original metres; ASSUMED 0 m (tubing hangs from the tubing hanger flange) |
| `components[tubing-1].bottomMd` | 7742.782 | slide 16 image36.png header | WEG | 2360.0 m x 3.28084 = 7742.782 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF'); depth of 'WEG' (mule-shoe-1) |
| `components[tubing-1].bottomMdM` | 2360.0 | slide 16 image36.png header | WEG | original metres; m MDTHF (column headers 'DEPTH m MDTHF'); depth of 'WEG' (mule-shoe-1) |
| `components[tubing-1].innerDiameter` | 2.875 | slide 16 image36.png header | 2 7/8" 6.5 PPF N-80 NK3SB | USER-DEFINED RULE: header size without ID/OD designation -> innerDiameter |
| `components[tubing-1].weight` | 6.5 | slide 16 image36.png header | 2 7/8" 6.5 PPF N-80 NK3SB | parsed from the label (ppf = lb/ft) |
| `components[tubing-1].grade` | N-80 | slide 16 image36.png header | 2 7/8" 6.5 PPF N-80 NK3SB | parsed from the label |
| `components[zone-e-7-1].name` | E-7 2386-2402 m MDTHF | slide 16 image36.png slide obj #10 | E-7 2386-2402 m MDTHF | visual transcription of the screenshot pictures (x3 upscale, upper part image36 / lower part image37) |
| `components[zone-e-7-1].type` | perforation | slide 16 image36.png slide obj #10 | E-7 2386-2402 m MDTHF | EXACT TYPE (Component_Types) |
| `components[zone-e-7-1].topMd` | 7828.084 | slide 16 image36.png slide obj #10 | slide text box obj #10 '2386 – 2402m MDTHF' beside SAND: E-7 | 2386.0 m x 3.28084 = 7828.084 ft (3 dp); m MDTHF as printed |
| `components[zone-e-7-1].topMdM` | 2386.0 | slide 16 image36.png slide obj #10 | slide text box obj #10 '2386 – 2402m MDTHF' beside SAND: E-7 | original metres; m MDTHF as printed |
| `components[zone-e-7-1].bottomMd` | 7880.578 | slide 16 image36.png slide obj #10 | slide text box obj #10 '2386 – 2402m MDTHF' beside SAND: E-7 | 2402.0 m x 3.28084 = 7880.578 ft (3 dp); m MDTHF as printed |
| `components[zone-e-7-1].bottomMdM` | 2402.0 | slide 16 image36.png slide obj #10 | slide text box obj #10 '2386 – 2402m MDTHF' beside SAND: E-7 | original metres; m MDTHF as printed |
| `components[zone-e-7-1].zone` | E-7 | slide 16 image36.png slide obj #10 | E-7 | zone label as printed |
| `wellInfo.completionDate` | 1992-03-01 | slide 16 image36.png header | 1 MARCH 1992 | ISO date () |
| `wellInfo.completionFluid` | - |  |  | not supported by the source -> template empty value ('-' text / null number) |
| `wellInfo.completionType` | SINGLE SELECTIVE OIL PRODUCER (TYPE 3) | slide 16 image36.png header | SINGLE SELECTIVE OIL PRODUCER (TYPE 3) | as printed |
| `wellInfo.groundElevation` | None |  |  | not supported by the source -> template empty value ('-' text / null number) |
| `wellInfo.kbElevation` | None |  |  | not supported by the source -> template empty value ('-' text / null number) |
| `wellInfo.maxDeviation` | 63.8 DEG | slide 16 image36.png header | 63.8 DEG | as printed |
| `wellInfo.name` | DULANG-A19 | slide 16 | WELL A-19 : SINGLE SELECTIVE OIL PRODUCER (TYPE 3) | identifier rule DULANG-A<no><suffix>; original label kept in the summary |
| `wellInfo.packerFluid` | 3% KCl + NaCl + 0.05% NALCO 3900 BIOCIDE & CORR. INHIBITOR | slide 16 image36.png header | 3% KCl + NaCl + 0.05% NALCO 3900 BIOCIDE & CORR. INHIBITOR | as printed |
| `wellInfo.rig` | TIOMAN 1 | slide 16 image36.png header | TIOMAN 1 | as printed |
| `wellInfo.surfaceCasing` | - |  |  | not supported by the source -> template empty value ('-' text / null number) |
| `wellInfo.totalDepth` | 8435.04 | slide 16 image36.png header/diagram | TD @ 2571 m MDTHF | 2571.0 m x 3.28084 = 8435.04 ft; datum MDTHF as printed (not shifted) |
| `wellInfo.trajectoryType` | - |  |  | not supported by the source -> template empty value ('-' text / null number) |
| `wellInfo.tubing` | 2 7/8" 6.5 PPF N-80 NK3SB | slide 16 image36.png header | 2 7/8" 6.5 PPF N-80 NK3SB | as printed |
| `wellInfo.wellhead` | - |  |  | not supported by the source -> template empty value ('-' text / null number) |

## DULANG-A30 — slide 27

Original label: WELL A-30: DUAL OIL PRODUCER/WATER INJECTOR WITH SHORT STRING SELECTIVE
Equipment depth datum: m MDTHF (column headers 'DEPTH m MDTHF')  ·  source: image62.png + image63.png (visual transcription of the screenshot pictures (x3 upscale, upper part image62 / lower part image63))

### Conflicts / missing / unclear

| kind | where | detail |
|---|---|---|
| Conflict | l-flow-coupling-1 | 'FLOW COUPLING' at 135.0 m is shallower than the item listed above it (packer-2 at 1763.9 m) — kept as printed |
| Missing | slide 27 image62.png header | '18 5/8" X-56 COND. 87.5 ppf @ m MDDF': setting depth not printed |
| Missing | slide 27 image62.png header | '13 3/8" K-55 SURF., 54.5 ppf @ mMDDF': setting depth not printed |
| Missing | slide 27 image62.png header | '9 5/8" 40 ppf, K-55, BTC @ mMDDF': setting depth not printed |
| Missing | header | casing setting depths left blank on the schematic ('@ m MDDF') |
| Unclear | picture p1/p2 | perforation intervals printed in 'm' without datum; read as m MDTHF like the depth columns |
| Unclear | slide 27 | later statuses OPEN / CLOSED 14/AUG/2009 are slide text boxes matched to the short-string SSD 1678.8 and 'D' NO-GO 1766.6 by vertical position |
| Unclear | picture p1 | long-string FLOW COUPLING rows (135.0, 139.2) have no MIN ID printed; packers and lower-completion items have no MIN ID printed |

### Unmapped content (kept here, not in the JSON)

| where | text | reason |
|---|---|---|
| slide 27 image62.png header | RTE / drill floor to tubing hanger : 14.52 m | offset, not an elevation above datum -> kbElevation left empty (no datum shift) |
| slide 27 image62.png diagram | 9 5/8" CASING SHOE | casing-shoe label without its own depth (9 5/8" CASING SHOE (label without depth)) |
| picture p1 | 1/4" CONTROL LINE on both strings (to the SCSSVs) | note (status / history / survey / other revision) |
| slide 27 obj #9 | OPEN 21/MAR/1994 (slide text box near the long-string lower completion; not matched to a row) | note (status / history / survey / other revision) |
| slide 27 obj #11 | Sand HUD @ 1676m MDTHF (Apr 2017) – 2.29" Gauge Ring | note (status / history / survey / other revision) |
| slide 27 obj #13353 | 44ft Fish (125 wire + 20' Toolstring Fish LIH (MA Plug + W1 RT + SSJ + 2'NRS + KJ + 5'NRS + KJ + 5'NRS + SJ + RS + Flo Petrol Cutter + Go-devil + Flo petrol cutter + 4' drop bar (Top of Fish at 1616m WLTHF) – Mar 2019 (partly outside the slide) | note (status / history / survey / other revision) |
| slide 27 obj #14 | 12 leak spots @ 645, 653.5, 670, 673.5, 820, 1027.7, 1036, 1026, 1545, 1611, 1708 & 1726 mMDTHF (MPLT Jan 2011) | note (status / history / survey / other revision) |
| slide 27 obj #12 | Latest Update: DEC 2023 BY PRAVIN NAIR | note (status / history / survey / other revision) |
| slide 27 obj #3 | Jan/2013 | note (status / history / survey / other revision) |
| slide 27 obj #18 | FB – SE1 | slide text box over/beside the picture (status / history / note) |
| slide 27 obj #3 | Jan/2013 | slide text box over/beside the picture (status / history / note) |
| slide 27 obj #7 | OPEN 14/AUG/2009 | slide text box over/beside the picture (status / history / note) |
| slide 27 obj #8 | CLOSED 14/AUG/2009 | slide text box over/beside the picture (status / history / note) |
| slide 27 obj #9 | OPEN 21/MAR/1994 | slide text box over/beside the picture (status / history / note) |
| slide 27 obj #12 | Latest Update: DEC 2023 BY PRAVIN NAIR | slide text box over/beside the picture (status / history / note) |
| slide 27 obj #11 | Sand HUD @ 1676m MDTHF (Apr 2017) – 2.29” Gauge Ring | slide text box over/beside the picture (status / history / note) |
| slide 27 obj #13353 | 44ft Fish (125 wire + 20’ Toolstring Fish LIH (MA Plug + W1 RT + SSJ + 2’NRS + KJ + 5’NRS + KJ + 5’NRS + SJ + RS + Flo Petrol Cutter + Go-devil + Flo petrol cutter + 4’ drop bar (Top of Fish at 1616m WLTHF) – Mar 2019 | slide text box over/beside the picture (status / history / note) |
| slide 27 obj #14 | 12 leak spots @ 645, 653.5, 670, 673.5, 820, 1027.7, 1036, 1026, 1545, 1611, 1708 & 1726 mMDTHF (MPLT Jan 2011) | slide text box over/beside the picture (status / history / note) |

### Traceability (every populated value)

| JSON path | value | slide / object / location | original text | conversion / interpretation |
|---|---|---|---|---|
| `components[christmas-tree-1].name` | INGRAM CACTUS, DUAL 3 1/8" X 3 1/8" X 11", 3M | slide 27 image62.png header | INGRAM CACTUS, DUAL 3 1/8" X 3 1/8" X 11", 3M | visual transcription of the screenshot pictures (x3 upscale, upper part image62 / lower part image63) |
| `components[christmas-tree-1].type` | christmas-tree | slide 27 image62.png header | INGRAM CACTUS, DUAL 3 1/8" X 3 1/8" X 11", 3M | EXACT TYPE (Component_Types) |
| `components[s-flow-coupling-1].name` | FLOW COUPLING | slide 27 image62.png (upper picture) | FLOW COUPLING | visual transcription of the screenshot pictures (x3 upscale, upper part image62 / lower part image63) |
| `components[s-flow-coupling-1].type` | flow-coupling | slide 27 image62.png (upper picture) | FLOW COUPLING | EXACT TYPE (Component_Types) |
| `components[s-flow-coupling-1].topMd` | 414.042 | slide 27 image62.png (upper picture) | 126.2 | 126.2 m x 3.28084 = 414.042 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[s-flow-coupling-1].topMdM` | 126.2 | slide 27 image62.png (upper picture) | 126.2 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[s-flow-coupling-1].innerDiameter` | 2.812 | slide 27 image62.png (upper picture) | 2.812 | column 'MIN ID' = explicit ID |
| `components[s-flow-coupling-1].string` | Short | slide 27 image62.png (upper picture) |  | string column of the schematic (short / long / dual packer) as printed |
| `components[s-scssv-1].name` | TRDP-4A SCSSV | slide 27 image62.png (upper picture) | TRDP-4A SCSSV | visual transcription of the screenshot pictures (x3 upscale, upper part image62 / lower part image63) |
| `components[s-scssv-1].type` | scssv | slide 27 image62.png (upper picture) | TRDP-4A SCSSV | EXACT TYPE (Component_Types) |
| `components[s-scssv-1].topMd` | 416.339 | slide 27 image62.png (upper picture) | 126.9 | 126.9 m x 3.28084 = 416.339 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[s-scssv-1].topMdM` | 126.9 | slide 27 image62.png (upper picture) | 126.9 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[s-scssv-1].innerDiameter` | 2.812 | slide 27 image62.png (upper picture) | 2.812 | column 'MIN ID' = explicit ID |
| `components[s-scssv-1].string` | Short | slide 27 image62.png (upper picture) |  | string column of the schematic (short / long / dual packer) as printed |
| `components[s-landing-nipple-1].name` | WP-1 NIPPLE | slide 27 image62.png (upper picture) | WP-1 NIPPLE | visual transcription of the screenshot pictures (x3 upscale, upper part image62 / lower part image63) |
| `components[s-landing-nipple-1].type` | landing-nipple | slide 27 image62.png (upper picture) | WP-1 NIPPLE | NEAREST TYPE (flagged) |
| `components[s-landing-nipple-1].topMd` | 947.507 | slide 27 image62.png (upper picture) | 288.8 | 288.8 m x 3.28084 = 947.507 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[s-landing-nipple-1].topMdM` | 288.8 | slide 27 image62.png (upper picture) | 288.8 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[s-landing-nipple-1].innerDiameter` | 2.812 | slide 27 image62.png (upper picture) | 2.812 | column 'MIN ID' = explicit ID |
| `components[s-landing-nipple-1].string` | Short | slide 27 image62.png (upper picture) |  | string column of the schematic (short / long / dual packer) as printed |
| `components[s-spm-1].name` | KBMG-M SPM | slide 27 image62.png (upper picture) | KBMG-M SPM | visual transcription of the screenshot pictures (x3 upscale, upper part image62 / lower part image63) |
| `components[s-spm-1].type` | spm | slide 27 image62.png (upper picture) | KBMG-M SPM | EXACT TYPE (Component_Types) |
| `components[s-spm-1].topMd` | 2240.486 | slide 27 image62.png (upper picture) | 682.9 | 682.9 m x 3.28084 = 2240.486 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[s-spm-1].topMdM` | 682.9 | slide 27 image62.png (upper picture) | 682.9 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[s-spm-1].innerDiameter` | 2.92 | slide 27 image62.png (upper picture) | 2.920 | column 'MIN ID' = explicit ID |
| `components[s-spm-1].string` | Short | slide 27 image62.png (upper picture) |  | string column of the schematic (short / long / dual packer) as printed |
| `components[s-spm-2].name` | KBMG-M SPM | slide 27 image62.png (upper picture) | KBMG-M SPM | visual transcription of the screenshot pictures (x3 upscale, upper part image62 / lower part image63) |
| `components[s-spm-2].type` | spm | slide 27 image62.png (upper picture) | KBMG-M SPM | EXACT TYPE (Component_Types) |
| `components[s-spm-2].topMd` | 3448.163 | slide 27 image62.png (upper picture) | 1051.0 | 1051.0 m x 3.28084 = 3448.163 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[s-spm-2].topMdM` | 1051.0 | slide 27 image62.png (upper picture) | 1051.0 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[s-spm-2].innerDiameter` | 2.92 | slide 27 image62.png (upper picture) | 2.920 | column 'MIN ID' = explicit ID |
| `components[s-spm-2].string` | Short | slide 27 image62.png (upper picture) |  | string column of the schematic (short / long / dual packer) as printed |
| `components[s-spm-3].name` | KBMG-M SPM | slide 27 image62.png (upper picture) | KBMG-M SPM | visual transcription of the screenshot pictures (x3 upscale, upper part image62 / lower part image63) |
| `components[s-spm-3].type` | spm | slide 27 image62.png (upper picture) | KBMG-M SPM | EXACT TYPE (Component_Types) |
| `components[s-spm-3].topMd` | 4407.48 | slide 27 image62.png (upper picture) | 1343.4 | 1343.4 m x 3.28084 = 4407.48 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[s-spm-3].topMdM` | 1343.4 | slide 27 image62.png (upper picture) | 1343.4 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[s-spm-3].innerDiameter` | 2.92 | slide 27 image62.png (upper picture) | 2.920 | column 'MIN ID' = explicit ID |
| `components[s-spm-3].string` | Short | slide 27 image62.png (upper picture) |  | string column of the schematic (short / long / dual packer) as printed |
| `components[s-spm-4].name` | KBMG-M SPM | slide 27 image62.png (upper picture) | KBMG-M SPM | visual transcription of the screenshot pictures (x3 upscale, upper part image62 / lower part image63) |
| `components[s-spm-4].type` | spm | slide 27 image62.png (upper picture) | KBMG-M SPM | EXACT TYPE (Component_Types) |
| `components[s-spm-4].topMd` | 5115.814 | slide 27 image62.png (upper picture) | 1559.3 | 1559.3 m x 3.28084 = 5115.814 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[s-spm-4].topMdM` | 1559.3 | slide 27 image62.png (upper picture) | 1559.3 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[s-spm-4].innerDiameter` | 2.92 | slide 27 image62.png (upper picture) | 2.920 | column 'MIN ID' = explicit ID |
| `components[s-spm-4].string` | Short | slide 27 image62.png (upper picture) |  | string column of the schematic (short / long / dual packer) as printed |
| `components[s-spm-5].name` | KBMG-M SPM | slide 27 image62.png (upper picture) | KBMG-M SPM | visual transcription of the screenshot pictures (x3 upscale, upper part image62 / lower part image63) |
| `components[s-spm-5].type` | spm | slide 27 image62.png (upper picture) | KBMG-M SPM | EXACT TYPE (Component_Types) |
| `components[s-spm-5].topMd` | 5425.197 | slide 27 image62.png (upper picture) | 1653.6 | 1653.6 m x 3.28084 = 5425.197 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[s-spm-5].topMdM` | 1653.6 | slide 27 image62.png (upper picture) | 1653.6 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[s-spm-5].innerDiameter` | 2.92 | slide 27 image62.png (upper picture) | 2.920 | column 'MIN ID' = explicit ID |
| `components[s-spm-5].string` | Short | slide 27 image62.png (upper picture) |  | string column of the schematic (short / long / dual packer) as printed |
| `components[s-sliding-sleeve-1].name` | CSW-10-T | slide 27 image62.png (upper picture) | CSW-10-T | visual transcription of the screenshot pictures (x3 upscale, upper part image62 / lower part image63) |
| `components[s-sliding-sleeve-1].type` | sliding-sleeve | slide 27 image62.png (upper picture) | CSW-10-T | EXACT TYPE (Component_Types) |
| `components[s-sliding-sleeve-1].topMd` | 5443.898 | slide 27 image62.png (upper picture) | 1659.3 | 1659.3 m x 3.28084 = 5443.898 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[s-sliding-sleeve-1].topMdM` | 1659.3 | slide 27 image62.png (upper picture) | 1659.3 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[s-sliding-sleeve-1].innerDiameter` | 2.812 | slide 27 image62.png (upper picture) | 2.812 | column 'MIN ID' = explicit ID |
| `components[s-sliding-sleeve-1].string` | Short | slide 27 image62.png (upper picture) |  | string column of the schematic (short / long / dual packer) as printed |
| `components[packer-1].name` | OTIS RDH DUAL HYD. PKR (PKR-3) | slide 27 image62.png (upper picture) | OTIS RDH DUAL HYD. PKR (PKR-3) | visual transcription of the screenshot pictures (x3 upscale, upper part image62 / lower part image63) |
| `components[packer-1].type` | packer | slide 27 image62.png (upper picture) | OTIS RDH DUAL HYD. PKR (PKR-3) | EXACT TYPE (Component_Types) |
| `components[packer-1].topMd` | 5478.675 | slide 27 image62.png (upper picture) | 1669.9 | 1669.9 m x 3.28084 = 5478.675 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[packer-1].topMdM` | 1669.9 | slide 27 image62.png (upper picture) | 1669.9 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[packer-1].string` | Short & Long | slide 27 image62.png (upper picture) |  | string column of the schematic (short / long / dual packer) as printed |
| `components[s-expansion-joint-1].name` | TELESCOPIC JOINT | slide 27 image62.png (upper picture) | TELESCOPIC JOINT | visual transcription of the screenshot pictures (x3 upscale, upper part image62 / lower part image63) |
| `components[s-expansion-joint-1].type` | expansion-joint | slide 27 image62.png (upper picture) | TELESCOPIC JOINT | NEAREST TYPE (flagged) |
| `components[s-expansion-joint-1].topMd` | 5483.268 | slide 27 image62.png (upper picture) | 1671.3 | 1671.3 m x 3.28084 = 5483.268 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[s-expansion-joint-1].topMdM` | 1671.3 | slide 27 image62.png (upper picture) | 1671.3 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[s-expansion-joint-1].string` | Short | slide 27 image62.png (upper picture) |  | string column of the schematic (short / long / dual packer) as printed |
| `components[s-sliding-sleeve-2].name` | 'CSW-1D'-T-SSD | slide 27 image62.png (upper picture) | 'CSW-1D'-T-SSD | visual transcription of the screenshot pictures (x3 upscale, upper part image62 / lower part image63) |
| `components[s-sliding-sleeve-2].type` | sliding-sleeve | slide 27 image62.png (upper picture) | 'CSW-1D'-T-SSD | EXACT TYPE (Component_Types) |
| `components[s-sliding-sleeve-2].topMd` | 5507.874 | slide 27 image62.png (upper picture) | 1678.8 | 1678.8 m x 3.28084 = 5507.874 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[s-sliding-sleeve-2].topMdM` | 1678.8 | slide 27 image62.png (upper picture) | 1678.8 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[s-sliding-sleeve-2].innerDiameter` | 2.812 | slide 27 image62.png (upper picture) | 2.812 | column 'MIN ID' = explicit ID |
| `components[s-sliding-sleeve-2].string` | Short | slide 27 image62.png (upper picture) |  | string column of the schematic (short / long / dual packer) as printed |
| `components[s-landing-nipple-2].name` | WP-1 NIPPLE | slide 27 image62.png (upper picture) | WP-1 NIPPLE | visual transcription of the screenshot pictures (x3 upscale, upper part image62 / lower part image63) |
| `components[s-landing-nipple-2].type` | landing-nipple | slide 27 image62.png (upper picture) | WP-1 NIPPLE | NEAREST TYPE (flagged) |
| `components[s-landing-nipple-2].topMd` | 5759.187 | slide 27 image62.png (upper picture) | 1755.4 | 1755.4 m x 3.28084 = 5759.187 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[s-landing-nipple-2].topMdM` | 1755.4 | slide 27 image62.png (upper picture) | 1755.4 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[s-landing-nipple-2].innerDiameter` | 2.812 | slide 27 image62.png (upper picture) | 2.812 | column 'MIN ID' = explicit ID |
| `components[s-landing-nipple-2].string` | Short | slide 27 image62.png (upper picture) |  | string column of the schematic (short / long / dual packer) as printed |
| `components[packer-2].name` | OTIS RDH DUAL HYD. PKR (PKR-2) | slide 27 image62.png (upper picture) | OTIS RDH DUAL HYD. PKR (PKR-2) | visual transcription of the screenshot pictures (x3 upscale, upper part image62 / lower part image63) |
| `components[packer-2].type` | packer | slide 27 image62.png (upper picture) | OTIS RDH DUAL HYD. PKR (PKR-2) | EXACT TYPE (Component_Types) |
| `components[packer-2].topMd` | 5787.074 | slide 27 image62.png (upper picture) | 1763.9 | 1763.9 m x 3.28084 = 5787.074 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[packer-2].topMdM` | 1763.9 | slide 27 image62.png (upper picture) | 1763.9 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[packer-2].string` | Short & Long | slide 27 image62.png (upper picture) |  | string column of the schematic (short / long / dual packer) as printed |
| `components[s-nipple-no-go-1].name` | 'D' NO-GO NIPPLE | slide 27 image63.png (lower picture) | 'D' NO-GO NIPPLE | visual transcription of the screenshot pictures (x3 upscale, upper part image62 / lower part image63) |
| `components[s-nipple-no-go-1].type` | nipple (no-go) | slide 27 image63.png (lower picture) | 'D' NO-GO NIPPLE | EXACT TYPE (Component_Types) |
| `components[s-nipple-no-go-1].topMd` | 5795.932 | slide 27 image63.png (lower picture) | 1766.6 | 1766.6 m x 3.28084 = 5795.932 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[s-nipple-no-go-1].topMdM` | 1766.6 | slide 27 image63.png (lower picture) | 1766.6 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[s-nipple-no-go-1].innerDiameter` | 2.75 | slide 27 image63.png (lower picture) | 2.750 | column 'MIN ID' = explicit ID |
| `components[s-nipple-no-go-1].string` | Short | slide 27 image63.png (lower picture) |  | string column of the schematic (short / long / dual packer) as printed |
| `components[s-mule-shoe-1].name` | W/L ENTRY GUIDE | slide 27 image63.png (lower picture) | W/L ENTRY GUIDE | visual transcription of the screenshot pictures (x3 upscale, upper part image62 / lower part image63) |
| `components[s-mule-shoe-1].type` | mule-shoe | slide 27 image63.png (lower picture) | W/L ENTRY GUIDE | NEAREST TYPE (flagged) |
| `components[s-mule-shoe-1].topMd` | 5797.244 | slide 27 image63.png (lower picture) | 1767.0 | 1767.0 m x 3.28084 = 5797.244 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[s-mule-shoe-1].topMdM` | 1767.0 | slide 27 image63.png (lower picture) | 1767.0 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[s-mule-shoe-1].innerDiameter` | 2.97 | slide 27 image63.png (lower picture) | 2.970 | column 'MIN ID' = explicit ID |
| `components[s-mule-shoe-1].string` | Short | slide 27 image63.png (lower picture) |  | string column of the schematic (short / long / dual packer) as printed |
| `components[l-flow-coupling-1].name` | FLOW COUPLING | slide 27 image62.png (upper picture) | FLOW COUPLING | visual transcription of the screenshot pictures (x3 upscale, upper part image62 / lower part image63) |
| `components[l-flow-coupling-1].type` | flow-coupling | slide 27 image62.png (upper picture) | FLOW COUPLING | EXACT TYPE (Component_Types) |
| `components[l-flow-coupling-1].topMd` | 442.913 | slide 27 image62.png (upper picture) | 135.0 | 135.0 m x 3.28084 = 442.913 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[l-flow-coupling-1].topMdM` | 135.0 | slide 27 image62.png (upper picture) | 135.0 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[l-flow-coupling-1].string` | Long | slide 27 image62.png (upper picture) |  | string column of the schematic (short / long / dual packer) as printed |
| `components[l-scssv-1].name` | TRDP-4A SCSSV | slide 27 image62.png (upper picture) | TRDP-4A SCSSV | visual transcription of the screenshot pictures (x3 upscale, upper part image62 / lower part image63) |
| `components[l-scssv-1].type` | scssv | slide 27 image62.png (upper picture) | TRDP-4A SCSSV | EXACT TYPE (Component_Types) |
| `components[l-scssv-1].topMd` | 445.866 | slide 27 image62.png (upper picture) | 135.9 | 135.9 m x 3.28084 = 445.866 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[l-scssv-1].topMdM` | 135.9 | slide 27 image62.png (upper picture) | 135.9 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[l-scssv-1].innerDiameter` | 2.812 | slide 27 image62.png (upper picture) | 2.812 | column 'MIN ID' = explicit ID |
| `components[l-scssv-1].string` | Long | slide 27 image62.png (upper picture) |  | string column of the schematic (short / long / dual packer) as printed |
| `components[l-flow-coupling-2].name` | FLOW COUPLING | slide 27 image62.png (upper picture) | FLOW COUPLING | visual transcription of the screenshot pictures (x3 upscale, upper part image62 / lower part image63) |
| `components[l-flow-coupling-2].type` | flow-coupling | slide 27 image62.png (upper picture) | FLOW COUPLING | EXACT TYPE (Component_Types) |
| `components[l-flow-coupling-2].topMd` | 456.693 | slide 27 image62.png (upper picture) | 139.2 | 139.2 m x 3.28084 = 456.693 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[l-flow-coupling-2].topMdM` | 139.2 | slide 27 image62.png (upper picture) | 139.2 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[l-flow-coupling-2].string` | Long | slide 27 image62.png (upper picture) |  | string column of the schematic (short / long / dual packer) as printed |
| `components[l-landing-nipple-1].name` | WP-1 NIPPLE | slide 27 image62.png (upper picture) | WP-1 NIPPLE | visual transcription of the screenshot pictures (x3 upscale, upper part image62 / lower part image63) |
| `components[l-landing-nipple-1].type` | landing-nipple | slide 27 image62.png (upper picture) | WP-1 NIPPLE | NEAREST TYPE (flagged) |
| `components[l-landing-nipple-1].topMd` | 979.331 | slide 27 image62.png (upper picture) | 298.5 | 298.5 m x 3.28084 = 979.331 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[l-landing-nipple-1].topMdM` | 298.5 | slide 27 image62.png (upper picture) | 298.5 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[l-landing-nipple-1].innerDiameter` | 2.812 | slide 27 image62.png (upper picture) | 2.812 | column 'MIN ID' = explicit ID |
| `components[l-landing-nipple-1].string` | Long | slide 27 image62.png (upper picture) |  | string column of the schematic (short / long / dual packer) as printed |
| `components[l-spm-1].name` | KBMG-M SPM | slide 27 image62.png (upper picture) | KBMG-M SPM | visual transcription of the screenshot pictures (x3 upscale, upper part image62 / lower part image63) |
| `components[l-spm-1].type` | spm | slide 27 image62.png (upper picture) | KBMG-M SPM | EXACT TYPE (Component_Types) |
| `components[l-spm-1].topMd` | 2270.013 | slide 27 image62.png (upper picture) | 691.9 | 691.9 m x 3.28084 = 2270.013 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[l-spm-1].topMdM` | 691.9 | slide 27 image62.png (upper picture) | 691.9 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[l-spm-1].innerDiameter` | 2.92 | slide 27 image62.png (upper picture) | 2.920 | column 'MIN ID' = explicit ID |
| `components[l-spm-1].string` | Long | slide 27 image62.png (upper picture) |  | string column of the schematic (short / long / dual packer) as printed |
| `components[l-sliding-sleeve-1].name` | CWS-1D-T-SSD | slide 27 image62.png (upper picture) | CWS-1D-T-SSD | visual transcription of the screenshot pictures (x3 upscale, upper part image62 / lower part image63) |
| `components[l-sliding-sleeve-1].type` | sliding-sleeve | slide 27 image62.png (upper picture) | CWS-1D-T-SSD | EXACT TYPE (Component_Types) |
| `components[l-sliding-sleeve-1].topMd` | 5451.772 | slide 27 image62.png (upper picture) | 1661.7 | 1661.7 m x 3.28084 = 5451.772 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[l-sliding-sleeve-1].topMdM` | 1661.7 | slide 27 image62.png (upper picture) | 1661.7 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[l-sliding-sleeve-1].innerDiameter` | 2.812 | slide 27 image62.png (upper picture) | 2.812 | column 'MIN ID' = explicit ID |
| `components[l-sliding-sleeve-1].string` | Long | slide 27 image62.png (upper picture) |  | string column of the schematic (short / long / dual packer) as printed |
| `components[l-sliding-sleeve-2].name` | CWS-1D-T-SSD | slide 27 image62.png (upper picture) | CWS-1D-T-SSD | visual transcription of the screenshot pictures (x3 upscale, upper part image62 / lower part image63) |
| `components[l-sliding-sleeve-2].type` | sliding-sleeve | slide 27 image62.png (upper picture) | CWS-1D-T-SSD | EXACT TYPE (Component_Types) |
| `components[l-sliding-sleeve-2].topMd` | 5516.732 | slide 27 image62.png (upper picture) | 1681.5 | 1681.5 m x 3.28084 = 5516.732 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[l-sliding-sleeve-2].topMdM` | 1681.5 | slide 27 image62.png (upper picture) | 1681.5 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[l-sliding-sleeve-2].innerDiameter` | 2.812 | slide 27 image62.png (upper picture) | 2.812 | column 'MIN ID' = explicit ID |
| `components[l-sliding-sleeve-2].string` | Long | slide 27 image62.png (upper picture) |  | string column of the schematic (short / long / dual packer) as printed |
| `components[l-blast-joint-1].name` | BLAST JOINT | slide 27 image62.png (upper picture) | BLAST JOINT | visual transcription of the screenshot pictures (x3 upscale, upper part image62 / lower part image63) |
| `components[l-blast-joint-1].type` | blast-joint | slide 27 image62.png (upper picture) | BLAST JOINT | EXACT TYPE (Component_Types) |
| `components[l-blast-joint-1].topMd` | 5529.856 | slide 27 image62.png (upper picture) | 1685.5 | 1685.5 m x 3.28084 = 5529.856 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[l-blast-joint-1].topMdM` | 1685.5 | slide 27 image62.png (upper picture) | 1685.5 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[l-blast-joint-1].bottomMd` | 5760.827 | slide 27 image62.png (upper picture) | 1755.9 | 1755.9 m x 3.28084 = 5760.827 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[l-blast-joint-1].bottomMdM` | 1755.9 | slide 27 image62.png (upper picture) | 1755.9 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[l-blast-joint-1].innerDiameter` | 2.992 | slide 27 image62.png (upper picture) | 2.992 | column 'MIN ID' = explicit ID |
| `components[l-blast-joint-1].string` | Long | slide 27 image62.png (upper picture) |  | string column of the schematic (short / long / dual packer) as printed |
| `components[l-landing-nipple-2].name` | WP-1 NIPPLE | slide 27 image62.png (upper picture) | WP-1 NIPPLE | visual transcription of the screenshot pictures (x3 upscale, upper part image62 / lower part image63) |
| `components[l-landing-nipple-2].type` | landing-nipple | slide 27 image62.png (upper picture) | WP-1 NIPPLE | NEAREST TYPE (flagged) |
| `components[l-landing-nipple-2].topMd` | 5770.998 | slide 27 image62.png (upper picture) | 1759.0 | 1759.0 m x 3.28084 = 5770.998 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[l-landing-nipple-2].topMdM` | 1759.0 | slide 27 image62.png (upper picture) | 1759.0 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[l-landing-nipple-2].innerDiameter` | 2.812 | slide 27 image62.png (upper picture) | 2.812 | column 'MIN ID' = explicit ID |
| `components[l-landing-nipple-2].string` | Long | slide 27 image62.png (upper picture) |  | string column of the schematic (short / long / dual packer) as printed |
| `components[l-sliding-sleeve-3].name` | CWS-1D-T-SSD | slide 27 image63.png (lower picture) | CWS-1D-T-SSD | visual transcription of the screenshot pictures (x3 upscale, upper part image62 / lower part image63) |
| `components[l-sliding-sleeve-3].type` | sliding-sleeve | slide 27 image63.png (lower picture) | CWS-1D-T-SSD | EXACT TYPE (Component_Types) |
| `components[l-sliding-sleeve-3].topMd` | 5806.103 | slide 27 image63.png (lower picture) | 1769.7 | 1769.7 m x 3.28084 = 5806.103 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[l-sliding-sleeve-3].topMdM` | 1769.7 | slide 27 image63.png (lower picture) | 1769.7 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[l-sliding-sleeve-3].innerDiameter` | 2.812 | slide 27 image63.png (lower picture) | 2.812 | column 'MIN ID' = explicit ID |
| `components[l-sliding-sleeve-3].string` | Long | slide 27 image63.png (lower picture) |  | string column of the schematic (short / long / dual packer) as printed |
| `components[l-blast-joint-2].name` | BLAST JOINT | slide 27 image63.png (lower picture) | BLAST JOINT | visual transcription of the screenshot pictures (x3 upscale, upper part image62 / lower part image63) |
| `components[l-blast-joint-2].type` | blast-joint | slide 27 image63.png (lower picture) | BLAST JOINT | EXACT TYPE (Component_Types) |
| `components[l-blast-joint-2].topMd` | 5819.226 | slide 27 image63.png (lower picture) | 1773.7 | 1773.7 m x 3.28084 = 5819.226 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[l-blast-joint-2].topMdM` | 1773.7 | slide 27 image63.png (lower picture) | 1773.7 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[l-blast-joint-2].bottomMd` | 5897.966 | slide 27 image63.png (lower picture) | 1797.7 | 1797.7 m x 3.28084 = 5897.966 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[l-blast-joint-2].bottomMdM` | 1797.7 | slide 27 image63.png (lower picture) | 1797.7 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[l-blast-joint-2].innerDiameter` | 2.992 | slide 27 image63.png (lower picture) | 2.992 | column 'MIN ID' = explicit ID |
| `components[l-blast-joint-2].string` | Long | slide 27 image63.png (lower picture) |  | string column of the schematic (short / long / dual packer) as printed |
| `components[l-landing-nipple-3].name` | WP-1 NIPPLE | slide 27 image63.png (lower picture) | WP-1 NIPPLE | visual transcription of the screenshot pictures (x3 upscale, upper part image62 / lower part image63) |
| `components[l-landing-nipple-3].type` | landing-nipple | slide 27 image63.png (lower picture) | WP-1 NIPPLE | NEAREST TYPE (flagged) |
| `components[l-landing-nipple-3].topMd` | 6249.672 | slide 27 image63.png (lower picture) | 1904.9 | 1904.9 m x 3.28084 = 6249.672 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[l-landing-nipple-3].topMdM` | 1904.9 | slide 27 image63.png (lower picture) | 1904.9 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[l-landing-nipple-3].innerDiameter` | 2.812 | slide 27 image63.png (lower picture) | 2.812 | column 'MIN ID' = explicit ID |
| `components[l-landing-nipple-3].string` | Long | slide 27 image63.png (lower picture) |  | string column of the schematic (short / long / dual packer) as printed |
| `components[l-pup-joint-1].name` | CENTRALIZER | slide 27 image63.png (lower picture) | CENTRALIZER | visual transcription of the screenshot pictures (x3 upscale, upper part image62 / lower part image63) |
| `components[l-pup-joint-1].type` | pup-joint | slide 27 image63.png (lower picture) | CENTRALIZER | UNRESOLVED / NEAREST TYPE (flagged) |
| `components[l-pup-joint-1].topMd` | 6268.373 | slide 27 image63.png (lower picture) | 1910.6 | 1910.6 m x 3.28084 = 6268.373 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[l-pup-joint-1].topMdM` | 1910.6 | slide 27 image63.png (lower picture) | 1910.6 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[l-pup-joint-1].string` | Long | slide 27 image63.png (lower picture) |  | string column of the schematic (short / long / dual packer) as printed |
| `components[l-seal-assembly-1].name` | BAKER G-H4 SEAL ASSY. | slide 27 image63.png (lower picture) | BAKER G-H4 SEAL ASSY. | visual transcription of the screenshot pictures (x3 upscale, upper part image62 / lower part image63) |
| `components[l-seal-assembly-1].type` | seal-assembly | slide 27 image63.png (lower picture) | BAKER G-H4 SEAL ASSY. | EXACT TYPE (Component_Types) |
| `components[l-seal-assembly-1].topMd` | 6273.622 | slide 27 image63.png (lower picture) | 1912.2 | 1912.2 m x 3.28084 = 6273.622 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[l-seal-assembly-1].topMdM` | 1912.2 | slide 27 image63.png (lower picture) | 1912.2 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[l-seal-assembly-1].string` | Long | slide 27 image63.png (lower picture) |  | string column of the schematic (short / long / dual packer) as printed |
| `components[l-packer-1].name` | BAKER D-1 PERM. PKR (PKR-1) | slide 27 image63.png (lower picture) | BAKER D-1 PERM. PKR (PKR-1) | visual transcription of the screenshot pictures (x3 upscale, upper part image62 / lower part image63) |
| `components[l-packer-1].type` | packer | slide 27 image63.png (lower picture) | BAKER D-1 PERM. PKR (PKR-1) | EXACT TYPE (Component_Types) |
| `components[l-packer-1].topMd` | 6277.887 | slide 27 image63.png (lower picture) | 1913.5 | 1913.5 m x 3.28084 = 6277.887 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[l-packer-1].topMdM` | 1913.5 | slide 27 image63.png (lower picture) | 1913.5 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[l-packer-1].string` | Long | slide 27 image63.png (lower picture) |  | string column of the schematic (short / long / dual packer) as printed |
| `components[l-pbr-1].name` | SEAL BORE EXT. | slide 27 image63.png (lower picture) | SEAL BORE EXT. | visual transcription of the screenshot pictures (x3 upscale, upper part image62 / lower part image63) |
| `components[l-pbr-1].type` | pbr | slide 27 image63.png (lower picture) | SEAL BORE EXT. | NEAREST TYPE (flagged) |
| `components[l-pbr-1].topMd` | 6278.872 | slide 27 image63.png (lower picture) | 1913.8 | 1913.8 m x 3.28084 = 6278.872 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[l-pbr-1].topMdM` | 1913.8 | slide 27 image63.png (lower picture) | 1913.8 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[l-pbr-1].string` | Long | slide 27 image63.png (lower picture) |  | string column of the schematic (short / long / dual packer) as printed |
| `components[l-nipple-no-go-1].name` | 'D' NO-GO NIPPLE | slide 27 image63.png (lower picture) | 'D' NO-GO NIPPLE | visual transcription of the screenshot pictures (x3 upscale, upper part image62 / lower part image63) |
| `components[l-nipple-no-go-1].type` | nipple (no-go) | slide 27 image63.png (lower picture) | 'D' NO-GO NIPPLE | EXACT TYPE (Component_Types) |
| `components[l-nipple-no-go-1].topMd` | 6295.604 | slide 27 image63.png (lower picture) | 1918.9 | 1918.9 m x 3.28084 = 6295.604 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[l-nipple-no-go-1].topMdM` | 1918.9 | slide 27 image63.png (lower picture) | 1918.9 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[l-nipple-no-go-1].innerDiameter` | 2.75 | slide 27 image63.png (lower picture) | 2.750 | column 'MIN ID' = explicit ID |
| `components[l-nipple-no-go-1].string` | Long | slide 27 image63.png (lower picture) |  | string column of the schematic (short / long / dual packer) as printed |
| `components[l-mule-shoe-1].name` | WEG | slide 27 image63.png (lower picture) | WEG | visual transcription of the screenshot pictures (x3 upscale, upper part image62 / lower part image63) |
| `components[l-mule-shoe-1].type` | mule-shoe | slide 27 image63.png (lower picture) | WEG | NEAREST TYPE (flagged) |
| `components[l-mule-shoe-1].topMd` | 6296.588 | slide 27 image63.png (lower picture) | 1919.2 | 1919.2 m x 3.28084 = 6296.588 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[l-mule-shoe-1].topMdM` | 1919.2 | slide 27 image63.png (lower picture) | 1919.2 | original metres; m MDTHF (column headers 'DEPTH m MDTHF') (equipment depth column as printed) |
| `components[l-mule-shoe-1].innerDiameter` | 2.97 | slide 27 image63.png (lower picture) | 2.970 | column 'MIN ID' = explicit ID |
| `components[l-mule-shoe-1].string` | Long | slide 27 image63.png (lower picture) |  | string column of the schematic (short / long / dual packer) as printed |
| `components[conductor-1].name` | 18 5/8" X-56 COND. 87.5 ppf @ m MDDF | slide 27 image62.png header | 18 5/8" X-56 COND. 87.5 ppf @ m MDDF | visual transcription of the screenshot pictures (x3 upscale, upper part image62 / lower part image63) |
| `components[conductor-1].type` | casing | slide 27 image62.png header | 18 5/8" X-56 COND. 87.5 ppf @ m MDDF | EXACT TYPE (Component_Types) |
| `components[conductor-1].innerDiameter` | 18.625 | slide 27 image62.png header | 18.625 | USER-DEFINED RULE: header size without ID/OD designation -> innerDiameter |
| `components[conductor-1].weight` | 87.5 | slide 27 image62.png header | 18 5/8" X-56 COND. 87.5 ppf @ m MDDF | parsed from the label (ppf = lb/ft) |
| `components[conductor-1].grade` | X-56 | slide 27 image62.png header | 18 5/8" X-56 COND. 87.5 ppf @ m MDDF | parsed from the label |
| `components[surface-casing-1].name` | 13 3/8" K-55 SURF., 54.5 ppf @ mMDDF | slide 27 image62.png header | 13 3/8" K-55 SURF., 54.5 ppf @ mMDDF | visual transcription of the screenshot pictures (x3 upscale, upper part image62 / lower part image63) |
| `components[surface-casing-1].type` | casing | slide 27 image62.png header | 13 3/8" K-55 SURF., 54.5 ppf @ mMDDF | EXACT TYPE (Component_Types) |
| `components[surface-casing-1].innerDiameter` | 13.375 | slide 27 image62.png header | 13.375 | USER-DEFINED RULE: header size without ID/OD designation -> innerDiameter |
| `components[surface-casing-1].weight` | 54.5 | slide 27 image62.png header | 13 3/8" K-55 SURF., 54.5 ppf @ mMDDF | parsed from the label (ppf = lb/ft) |
| `components[surface-casing-1].grade` | K-55 | slide 27 image62.png header | 13 3/8" K-55 SURF., 54.5 ppf @ mMDDF | parsed from the label |
| `components[production-casing-1].name` | 9 5/8" 40 ppf, K-55, BTC @ mMDDF | slide 27 image62.png header | 9 5/8" 40 ppf, K-55, BTC @ mMDDF | visual transcription of the screenshot pictures (x3 upscale, upper part image62 / lower part image63) |
| `components[production-casing-1].type` | casing | slide 27 image62.png header | 9 5/8" 40 ppf, K-55, BTC @ mMDDF | EXACT TYPE (Component_Types) |
| `components[production-casing-1].innerDiameter` | 9.625 | slide 27 image62.png header | 9.625 | USER-DEFINED RULE: header size without ID/OD designation -> innerDiameter |
| `components[production-casing-1].weight` | 40.0 | slide 27 image62.png header | 9 5/8" 40 ppf, K-55, BTC @ mMDDF | parsed from the label (ppf = lb/ft) |
| `components[production-casing-1].grade` | K-55 | slide 27 image62.png header | 9 5/8" 40 ppf, K-55, BTC @ mMDDF | parsed from the label |
| `components[s-tubing-1].name` | 3 1/2" 9.2 PPF N-80 TKC 40/40 (short string) | slide 27 image62.png header | 3 1/2" 9.2 PPF N-80 TKC 40/40 (short string) | visual transcription of the screenshot pictures (x3 upscale, upper part image62 / lower part image63) |
| `components[s-tubing-1].type` | tubing | slide 27 image62.png header | 3 1/2" 9.2 PPF N-80 TKC 40/40 (short string) | EXACT TYPE (Component_Types) |
| `components[s-tubing-1].topMd` | 0.0 | slide 27 image62.png header | depth reference m MDTHF / THF | 0.0 m x 3.28084 = 0.0 ft (3 dp); ASSUMED 0 m (tubing hangs from the tubing hanger flange) |
| `components[s-tubing-1].topMdM` | 0.0 | slide 27 image62.png header | depth reference m MDTHF / THF | original metres; ASSUMED 0 m (tubing hangs from the tubing hanger flange) |
| `components[s-tubing-1].bottomMd` | 5797.244 | slide 27 image62.png header | W/L ENTRY GUIDE | 1767.0 m x 3.28084 = 5797.244 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF'); depth of 'W/L ENTRY GUIDE' (s-mule-shoe-1) |
| `components[s-tubing-1].bottomMdM` | 1767.0 | slide 27 image62.png header | W/L ENTRY GUIDE | original metres; m MDTHF (column headers 'DEPTH m MDTHF'); depth of 'W/L ENTRY GUIDE' (s-mule-shoe-1) |
| `components[s-tubing-1].innerDiameter` | 3.5 | slide 27 image62.png header | 3 1/2" 9.2 PPF N-80 TKC 40/40 (short string) | USER-DEFINED RULE: header size without ID/OD designation -> innerDiameter |
| `components[s-tubing-1].weight` | 9.2 | slide 27 image62.png header | 3 1/2" 9.2 PPF N-80 TKC 40/40 (short string) | parsed from the label (ppf = lb/ft) |
| `components[s-tubing-1].grade` | N-80 | slide 27 image62.png header | 3 1/2" 9.2 PPF N-80 TKC 40/40 (short string) | parsed from the label |
| `components[s-tubing-1].string` | Short | slide 27 image62.png header |  | string column of the schematic (short / long / dual packer) as printed |
| `components[l-tubing-1].name` | 3 1/2" 9.2 PPF N-80 TKC 40/40 (long string) | slide 27 image62.png header | 3 1/2" 9.2 PPF N-80 TKC 40/40 (long string) | visual transcription of the screenshot pictures (x3 upscale, upper part image62 / lower part image63) |
| `components[l-tubing-1].type` | tubing | slide 27 image62.png header | 3 1/2" 9.2 PPF N-80 TKC 40/40 (long string) | EXACT TYPE (Component_Types) |
| `components[l-tubing-1].topMd` | 0.0 | slide 27 image62.png header | depth reference m MDTHF / THF | 0.0 m x 3.28084 = 0.0 ft (3 dp); ASSUMED 0 m (tubing hangs from the tubing hanger flange) |
| `components[l-tubing-1].topMdM` | 0.0 | slide 27 image62.png header | depth reference m MDTHF / THF | original metres; ASSUMED 0 m (tubing hangs from the tubing hanger flange) |
| `components[l-tubing-1].bottomMd` | 6296.588 | slide 27 image62.png header | WEG | 1919.2 m x 3.28084 = 6296.588 ft (3 dp); m MDTHF (column headers 'DEPTH m MDTHF'); depth of 'WEG' (l-mule-shoe-1) |
| `components[l-tubing-1].bottomMdM` | 1919.2 | slide 27 image62.png header | WEG | original metres; m MDTHF (column headers 'DEPTH m MDTHF'); depth of 'WEG' (l-mule-shoe-1) |
| `components[l-tubing-1].innerDiameter` | 3.5 | slide 27 image62.png header | 3 1/2" 9.2 PPF N-80 TKC 40/40 (long string) | USER-DEFINED RULE: header size without ID/OD designation -> innerDiameter |
| `components[l-tubing-1].weight` | 9.2 | slide 27 image62.png header | 3 1/2" 9.2 PPF N-80 TKC 40/40 (long string) | parsed from the label (ppf = lb/ft) |
| `components[l-tubing-1].grade` | N-80 | slide 27 image62.png header | 3 1/2" 9.2 PPF N-80 TKC 40/40 (long string) | parsed from the label |
| `components[l-tubing-1].string` | Long | slide 27 image62.png header |  | string column of the schematic (short / long / dual packer) as printed |
| `components[zone-ld-3-1].name` | LD-3 1688.5-1692.0 m MDTHF | slide 27 image62.png (upper picture) | LD-3 1688.5-1692.0 m MDTHF | visual transcription of the screenshot pictures (x3 upscale, upper part image62 / lower part image63) |
| `components[zone-ld-3-1].type` | perforation | slide 27 image62.png (upper picture) | LD-3 1688.5-1692.0 m MDTHF | EXACT TYPE (Component_Types) |
| `components[zone-ld-3-1].topMd` | 5539.698 | slide 27 image62.png (upper picture) | SAND : LD-3 & E-2 PERFT: LD3 1688.5 - 1692.0 m (no datum printed) | 1688.5 m x 3.28084 = 5539.698 ft (3 dp); m MDTHF as printed |
| `components[zone-ld-3-1].topMdM` | 1688.5 | slide 27 image62.png (upper picture) | SAND : LD-3 & E-2 PERFT: LD3 1688.5 - 1692.0 m (no datum printed) | original metres; m MDTHF as printed |
| `components[zone-ld-3-1].bottomMd` | 5551.181 | slide 27 image62.png (upper picture) | SAND : LD-3 & E-2 PERFT: LD3 1688.5 - 1692.0 m (no datum printed) | 1692.0 m x 3.28084 = 5551.181 ft (3 dp); m MDTHF as printed |
| `components[zone-ld-3-1].bottomMdM` | 1692.0 | slide 27 image62.png (upper picture) | SAND : LD-3 & E-2 PERFT: LD3 1688.5 - 1692.0 m (no datum printed) | original metres; m MDTHF as printed |
| `components[zone-ld-3-1].zone` | LD-3 | slide 27 image62.png (upper picture) | LD-3 | zone label as printed |
| `components[zone-e-2-1].name` | E-2 1740.0-1744.0 m MDTHF | slide 27 image62.png (upper picture) | E-2 1740.0-1744.0 m MDTHF | visual transcription of the screenshot pictures (x3 upscale, upper part image62 / lower part image63) |
| `components[zone-e-2-1].type` | perforation | slide 27 image62.png (upper picture) | E-2 1740.0-1744.0 m MDTHF | EXACT TYPE (Component_Types) |
| `components[zone-e-2-1].topMd` | 5708.662 | slide 27 image62.png (upper picture) | : E-2 1740.0 - 1744.0 m (no datum printed) | 1740.0 m x 3.28084 = 5708.662 ft (3 dp); m MDTHF as printed |
| `components[zone-e-2-1].topMdM` | 1740.0 | slide 27 image62.png (upper picture) | : E-2 1740.0 - 1744.0 m (no datum printed) | original metres; m MDTHF as printed |
| `components[zone-e-2-1].bottomMd` | 5721.785 | slide 27 image62.png (upper picture) | : E-2 1740.0 - 1744.0 m (no datum printed) | 1744.0 m x 3.28084 = 5721.785 ft (3 dp); m MDTHF as printed |
| `components[zone-e-2-1].bottomMdM` | 1744.0 | slide 27 image62.png (upper picture) | : E-2 1740.0 - 1744.0 m (no datum printed) | original metres; m MDTHF as printed |
| `components[zone-e-2-1].zone` | E-2 | slide 27 image62.png (upper picture) | E-2 | zone label as printed |
| `components[zone-e-2-2].name` | E-2 1746.0-1749.0 m MDTHF | slide 27 image62.png (upper picture) | E-2 1746.0-1749.0 m MDTHF | visual transcription of the screenshot pictures (x3 upscale, upper part image62 / lower part image63) |
| `components[zone-e-2-2].type` | perforation | slide 27 image62.png (upper picture) | E-2 1746.0-1749.0 m MDTHF | EXACT TYPE (Component_Types) |
| `components[zone-e-2-2].topMd` | 5728.347 | slide 27 image62.png (upper picture) | 1746.0 - 1749.0 m (no datum printed) | 1746.0 m x 3.28084 = 5728.347 ft (3 dp); m MDTHF as printed |
| `components[zone-e-2-2].topMdM` | 1746.0 | slide 27 image62.png (upper picture) | 1746.0 - 1749.0 m (no datum printed) | original metres; m MDTHF as printed |
| `components[zone-e-2-2].bottomMd` | 5738.189 | slide 27 image62.png (upper picture) | 1746.0 - 1749.0 m (no datum printed) | 1749.0 m x 3.28084 = 5738.189 ft (3 dp); m MDTHF as printed |
| `components[zone-e-2-2].bottomMdM` | 1749.0 | slide 27 image62.png (upper picture) | 1746.0 - 1749.0 m (no datum printed) | original metres; m MDTHF as printed |
| `components[zone-e-2-2].zone` | E-2 | slide 27 image62.png (upper picture) | E-2 | zone label as printed |
| `components[zone-e-7-1].name` | E-7 1775.5-1786.5 m MDTHF | slide 27 image63.png (lower picture) | E-7 1775.5-1786.5 m MDTHF | visual transcription of the screenshot pictures (x3 upscale, upper part image62 / lower part image63) |
| `components[zone-e-7-1].type` | perforation | slide 27 image63.png (lower picture) | E-7 1775.5-1786.5 m MDTHF | EXACT TYPE (Component_Types) |
| `components[zone-e-7-1].topMd` | 5825.131 | slide 27 image63.png (lower picture) | SAND : E-7 PERFT: 1775.5 - 1786.5 m (no datum printed) | 1775.5 m x 3.28084 = 5825.131 ft (3 dp); m MDTHF as printed |
| `components[zone-e-7-1].topMdM` | 1775.5 | slide 27 image63.png (lower picture) | SAND : E-7 PERFT: 1775.5 - 1786.5 m (no datum printed) | original metres; m MDTHF as printed |
| `components[zone-e-7-1].bottomMd` | 5861.221 | slide 27 image63.png (lower picture) | SAND : E-7 PERFT: 1775.5 - 1786.5 m (no datum printed) | 1786.5 m x 3.28084 = 5861.221 ft (3 dp); m MDTHF as printed |
| `components[zone-e-7-1].bottomMdM` | 1786.5 | slide 27 image63.png (lower picture) | SAND : E-7 PERFT: 1775.5 - 1786.5 m (no datum printed) | original metres; m MDTHF as printed |
| `components[zone-e-7-1].zone` | E-7 | slide 27 image63.png (lower picture) | E-7 | zone label as printed |
| `components[zone-e-14-1].name` | E-14 1928.0-1943.5 m MDTHF | slide 27 image63.png (lower picture) | E-14 1928.0-1943.5 m MDTHF | visual transcription of the screenshot pictures (x3 upscale, upper part image62 / lower part image63) |
| `components[zone-e-14-1].type` | perforation | slide 27 image63.png (lower picture) | E-14 1928.0-1943.5 m MDTHF | EXACT TYPE (Component_Types) |
| `components[zone-e-14-1].topMd` | 6325.46 | slide 27 image63.png (lower picture) | SAND : E-14 (WATER INJECTOR) PERFT: 1928.0 - 1943.5 m (no datum printed) | 1928.0 m x 3.28084 = 6325.46 ft (3 dp); m MDTHF as printed |
| `components[zone-e-14-1].topMdM` | 1928.0 | slide 27 image63.png (lower picture) | SAND : E-14 (WATER INJECTOR) PERFT: 1928.0 - 1943.5 m (no datum printed) | original metres; m MDTHF as printed |
| `components[zone-e-14-1].bottomMd` | 6376.313 | slide 27 image63.png (lower picture) | SAND : E-14 (WATER INJECTOR) PERFT: 1928.0 - 1943.5 m (no datum printed) | 1943.5 m x 3.28084 = 6376.313 ft (3 dp); m MDTHF as printed |
| `components[zone-e-14-1].bottomMdM` | 1943.5 | slide 27 image63.png (lower picture) | SAND : E-14 (WATER INJECTOR) PERFT: 1928.0 - 1943.5 m (no datum printed) | original metres; m MDTHF as printed |
| `components[zone-e-14-1].zone` | E-14 | slide 27 image63.png (lower picture) | E-14 | zone label as printed |
| `components[pbtd-1].name` | PBTD : 2017.5 m MDDF (header) | slide 27 image62.png diagram | PBTD : 2017.5 m MDDF (header) | visual transcription of the screenshot pictures (x3 upscale, upper part image62 / lower part image63) |
| `components[pbtd-1].type` | pbtd | slide 27 image62.png diagram | PBTD : 2017.5 m MDDF (header) | EXACT TYPE (Component_Types) |
| `components[pbtd-1].topMd` | 6619.095 | slide 27 image62.png diagram | PBTD : 2017.5 m MDDF (header) | 2017.5 m x 3.28084 = 6619.095 ft (3 dp); m MDDF |
| `components[pbtd-1].topMdM` | 2017.5 | slide 27 image62.png diagram | PBTD : 2017.5 m MDDF (header) | original metres; m MDDF |
| `wellInfo.completionDate` | 1994-03-15 | slide 27 image62.png header | 15 MARCH 1994 | ISO date () |
| `wellInfo.completionFluid` | - |  |  | not supported by the source -> template empty value ('-' text / null number) |
| `wellInfo.completionType` | DUAL OIL PRODUCER/WATER INJECTOR WITH SHORT STRING SELECTIVE | slide 27 image62.png header | DUAL OIL PRODUCER/WATER INJECTOR WITH SHORT STRING SELECTIVE | as printed |
| `wellInfo.groundElevation` | None |  |  | not supported by the source -> template empty value ('-' text / null number) |
| `wellInfo.kbElevation` | None |  |  | not supported by the source -> template empty value ('-' text / null number) |
| `wellInfo.maxDeviation` | 55.3 DEG | slide 27 image62.png header | 55.3 DEG | as printed |
| `wellInfo.name` | DULANG-A30 | slide 27 | WELL A-30: DUAL OIL PRODUCER/WATER INJECTOR WITH SHORT STRING SELECTIVE | identifier rule DULANG-A<no><suffix>; original label kept in the summary |
| `wellInfo.packerFluid` | 8.8 PPG 3% KCl + NaCl + 0.05% NALCO 3900 BIOCIDE & CORR. INHIBITOR | slide 27 image62.png header | 8.8 PPG 3% KCl + NaCl + 0.05% NALCO 3900 BIOCIDE & CORR. INHIBITOR | as printed |
| `wellInfo.rig` | TEKNIK BERKAT | slide 27 image62.png header | TEKNIK BERKAT | as printed |
| `wellInfo.surfaceCasing` | 13 3/8" K-55 SURF., 54.5 ppf @ mMDDF | slide 27 image62.png header | 13 3/8" K-55 SURF., 54.5 ppf @ mMDDF | SURF / INT. CASING header line as printed |
| `wellInfo.totalDepth` | 6724.082 | slide 27 image62.png header/diagram | TD @ 2049.5 m MDTHF | 2049.5 m x 3.28084 = 6724.082 ft; datum MDTHF as printed (not shifted) |
| `wellInfo.trajectoryType` | - |  |  | not supported by the source -> template empty value ('-' text / null number) |
| `wellInfo.tubing` | 3 1/2" 9.2 PPF N-80 TKC 40/40 | slide 27 image62.png header | 3 1/2" 9.2 PPF N-80 TKC 40/40 | as printed |
| `wellInfo.wellhead` | - |  |  | not supported by the source -> template empty value ('-' text / null number) |

