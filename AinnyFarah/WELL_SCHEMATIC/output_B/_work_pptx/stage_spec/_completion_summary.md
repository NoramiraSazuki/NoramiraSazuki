# Dulang D — well-schematic extraction: completion summary

Run: 2026-10-30T15:54:35 (local processing only). Script: `output_D/_work/build_schematic_D.py`.

## Inputs (read only) — SHA-256 before / after

| file | before | after | unchanged |
|---|---|---|---|
| `/Users/farahyusof/Documents/GitHub/NoramiraSazuki/AinnyFarah/WELL_SCHEMATIC/output_B/_work_pptx/src/Dulang B Well Schematic updated .pptx` | `e390e7e13ec18fc7294cb0a8b7b0ad09caf07893527be518626876bbb9e3fa21` | `e390e7e13ec18fc7294cb0a8b7b0ad09caf07893527be518626876bbb9e3fa21` | YES |
| `/Users/farahyusof/Documents/GitHub/NoramiraSazuki/AinnyFarah/WELL_SCHEMATIC/well-schematic-template-json.rtf` | `c7db15efb1917dfa76fa3102e270fe17115a8561f24eb0a155239870d079e712` | `c7db15efb1917dfa76fa3102e270fe17115a8561f24eb0a155239870d079e712` | YES |
| `/Users/farahyusof/Documents/GitHub/NoramiraSazuki/AinnyFarah/WELL_SCHEMATIC/well-schematic-sample-template.xlsx` | `d377a6310d42fda8c95e85db03f2416cb4c1b6f37f53920052331ce814f6d494` | `d377a6310d42fda8c95e85db03f2416cb4c1b6f37f53920052331ce814f6d494` | YES |

NAS original `/Volumes/homes/public/DulangOSS/20261016/DULANG_WELLS/Well_Schematic/1. Well Schematic/Dulang B/Dulang B Well Schematic updated .pptx`: SHA-256 at the start of processing `e390e7e13ec18fc7294cb0a8b7b0ad09caf07893527be518626876bbb9e3fa21` (identical to the local working copy above). Re-hashed at the end: `e390e7e13ec18fc7294cb0a8b7b0ad09caf07893527be518626876bbb9e3fa21` — UNCHANGED.

Template note: `well-schematic-template-json.rtf` is an **example JSON document, not a formal JSON Schema**. Conformity was checked against its section order, `wellInfo` keys/order and value types; component keys beyond the example (topMdM/bottomMdM, zone, string, weight, grade, bottomTvd, estimated…, actualType, source) follow the conventions of the earlier Dulang A/B outputs. `floatEquipment` / `plugBackCement` entries use key names from the Excel template sheets (Float_Equipment, Plug_Back_Cement) because the JSON example lists are empty.

## Wells identified (one JSON per resolved well)

| well file | original label | slide(s) | schematic source | components | flagged | conflicts | missing | unclear | validation |
|---|---|---|---|---|---|---|---|---|---|
| `Dulang_B16ST3.json` | B-16ST3 WELL COMPLETION DIAGRAM (DULANG UNIT REVISIT I); Well No. : B-16ST3 | 16 | native PowerPoint text boxes; depth / MIN ID / MAX OD columns are single multi-line text boxes, rows paired with labels by vertical position on the render — native slide text | 27 | 15 | 1 | 3 | 1 | PASS |
| `Dulang_B24ST2.json` | WELL NAME: DULANG B-24ST2 (LOCATION : DULANG-B) | 24 | visual transcription of the scanned picture — image24.png | 50 | 30 | 1 | 1 | 4 | PASS |
| `Dulang_B25ST4.json` | WELL NAME: B25 ST-4 (LOCATION / SLOT: DULANG B PLATFORM /SLOT 25) | 25 | visual transcription of a low-resolution picture — image25.png | 25 | 16 | 0 | 0 | 3 | PASS |
| `Dulang_B28ST2.json` | WELL NAME: DULANG B-28ST2 (LOCATION : DULANG-B) | 28 | visual transcription of the picture — image28.png | 50 | 20 | 3 | 5 | 4 | PASS |
| `Dulang_B30ST1.json` | WELL NAME: B30 ST-1 (LOCATION / SLOT: DULANG B PLATFORM /SLOT 26) | 30 | visual transcription of a low-resolution picture — image30.png | 49 | 48 | 0 | 0 | 3 | PASS |
| `Dulang_B31ST1.json` | WELL NAME: B31 ST-1 (LOCATION / SLOT: DULANG B PLATFORM /SLOT 14) | 31 | visual transcription of a low-resolution picture — image31.png | 29 | 27 | 0 | 0 | 3 | PASS |

JSON files written: **6** = resolved wells **6**. Slides read: 33 (presentation order, no hidden slides — no `show="0"`); slides covered by a well: [16, 24, 25, 28, 30, 31]. Pictures placed outside the visible slide area (slides 4, 7, 16, 19, 23, 31) were inspected and are listed per well as unmapped hidden content. Combined slides: D28 = slides 28 (schematic) + 29 (completion report, different datum — notes only); D32 = slides 33 (schematic) + 34 (completion report — notes only); D30 slide 31 holds a second, off-slide schematic (notes only).

Extraction method: all Dulang D schematics are pictures. EMF drawings (slides 2, 10, 12, 13, 20) were read from their native text records; screenshots / scans were transcribed visually from x3-upscaled crops with tesseract OCR as a cross-check. Transcriptions with per-row locations are in `output_D/_work/specs/*.txt`; low-legibility reads are marked '?' there and flagged in the JSON.

## Validation checks (per JSON)

| check | B16ST3 | B24ST2 | B25ST4 | B28ST2 | B30ST1 | B31ST1 |
|---|---|---|---|---|---|---|
| sections and order as template | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |
| wellInfo keys and order as template | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |
| wellInfo types (number or null; string) | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |
| unsupported sections empty lists | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |
| approved Type_Keys only | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |
| unique component ids | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |
| numeric fields numeric or null | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |
| ft = m x 3.28084 (3 dp) everywhere | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |
| nearest-type mappings flagged | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |
| flags complete (estimated -> fields + note) | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |
| no template sample values | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |
| well identifier = file name | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |
| every component traced to the source | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |
| every transcribed row mapped or listed | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |

## Dulang_B16ST3 — slide 16

Original label: B-16ST3 WELL COMPLETION DIAGRAM (DULANG UNIT REVISIT I); Well No. : B-16ST3
Equipment depth datum: m MDDF ('All Depth in : m-MDDF')  ·  source: native slide text (native PowerPoint text boxes; depth / MIN ID / MAX OD columns are single multi-line text boxes, rows paired with labels by vertical position on the render (visual check))

### Conflicts / missing / unclear

| kind | where | detail |
|---|---|---|
| Unresolved | slide 16 native slide text #175 | 'TR-SC-SSSV': no Type_Key |
| Unresolved | slide 16 native slide text #265 | '4-1/2" L-80 LCT B-P Blank': no Type_Key |
| Unresolved | slide 16 native slide text #266 | '4" P110 NU B-P WWS 8 ga': no Type_Key |
| Unresolved | slide 16 native slide text #268 | '4-1/2" L-80 LCT B-P Blank': no Type_Key |
| Unresolved | slide 16 native slide text #269 | '4" J-55 NU B-P WWS 8 ga': no Type_Key |
| Missing | slide 16 native slide text header | '4-1/2" TOL / 4-1/2" Liner Shoe': setting depth not printed |
| Missing | slide 16 | TD not printed |
| Unclear | slide 16 obj #297/#163/#296 | depth, MIN ID and MAX OD columns are single multi-line text boxes whose lines do not align with the equipment labels; only rows with a clear vertical match were given depths, uncertain pairings are marked '?' and flagged; 9 equipment labels have no depth |
| Conflict | slide 16 obj #318 | tubing was fully cut at 2857 m MDDF in May 2023 — the diagram below the cut (no-go nipple, EOT) is kept as printed with this flag |
| Missing | slide 16 | TD not printed with a label (3121 / 3189 / 3223 printed below the frame without labels); 4-1/2" liner shoe depth not assignable |

### Unmapped content (kept here, not in the JSON)

| where | text | reason |
|---|---|---|
| slide 16 native slide text #175 | TR-SC-SSSV | no Type_Key could be assigned |
| slide 16 native slide text #265 | 4-1/2" L-80 LCT B-P Blank | no Type_Key could be assigned |
| slide 16 native slide text #266 | 4" P110 NU B-P WWS 8 ga | no Type_Key could be assigned |
| slide 16 native slide text #268 | 4-1/2" L-80 LCT B-P Blank | no Type_Key could be assigned |
| slide 16 native slide text #269 | 4" J-55 NU B-P WWS 8 ga | no Type_Key could be assigned |
| slide 16 obj #297 | unassigned depths in the depth column: 1629 (SPM # 1 / SSD#3), 2860 and 2889 (next to the 7" liner shoe / PBTD labels, with ID/OD 2.250/3.175 and 2.441/3.175), 3121, 3189, 3223 (below the diagram frame; possibly 4-1/2" liner shoe / TD) — not mapped | note (status / history / survey / other revision) |
| slide 16 obj #160 | RKB : 37.3m (to sea level); RKB-THF : 21.19m | note (status / history / survey / other revision) |
| slide 16 obj #318 | Tubing full cut at 2857m MDDF (May 2023) due to unable to retrieve stiff PXN plug body at no-go nipple | note (status / history / survey / other revision) |
| slide 16 | zone status boxes: E 12/13 CLOSED; E 14 OPEN; E 23 OPEN; E 32 OPEN; LAST UPDATE: MAY 2023 BY PRAVIN NAIR | note (status / history / survey / other revision) |
| slide 16 obj #299 | LAST UPDATE: MAY 2023 BY PRAVIN NAIR | slide text box over/beside the picture (status / history / note) |
| slide 16 obj #161 | Well No. : B-16ST3 Location : BB Wellhead : Cooper cameron, 3-1/8",3000 psi Tubing : 2 7/8", 6.4ppf, L-80, Vam Ace | slide text box over/beside the picture (status / history / note) |
| slide 16 obj #162 | DULANG UNIT REVISIT I B-16ST3 WELL COMPLETION DIAGRAM | slide text box over/beside the picture (status / history / note) |
| slide 16 obj #170 | MIN I.D (IN) | slide text box over/beside the picture (status / history / note) |
| slide 16 obj #171 | MAX O.D (IN) | slide text box over/beside the picture (status / history / note) |
| slide 16 obj #172 | SINGLE STRING | slide text box over/beside the picture (status / history / note) |
| slide 16 obj #178 | 7" Liner Shoe | slide text box over/beside the picture (status / history / note) |
| slide 16 obj #182 | DEPTH (m) MDDF | slide text box over/beside the picture (status / history / note) |
| slide 16 obj #219 | 9-5/8" Casing Shoe | slide text box over/beside the picture (status / history / note) |
| slide 16 obj #280 | 4-1/2" Liner Shoe | slide text box over/beside the picture (status / history / note) |
| slide 16 obj #289 | 7" TOL | slide text box over/beside the picture (status / history / note) |
| slide 16 obj #290 | 4-1/2" TOL | slide text box over/beside the picture (status / history / note) |
| slide 16 obj #291 | PBTD | slide text box over/beside the picture (status / history / note) |
| slide 16 obj #328 | CLOSED | slide text box over/beside the picture (status / history / note) |
| slide 16 obj #330 | OPEN | slide text box over/beside the picture (status / history / note) |
| slide 16 obj #332 | OPEN | slide text box over/beside the picture (status / history / note) |
| slide 16 obj #333 | OPEN | slide text box over/beside the picture (status / history / note) |

### Traceability (every populated value)

| JSON path | value | slide / object / location | original text | conversion / interpretation |
|---|---|---|---|---|
| `components[flow-coupling-1].name` | FLOW COUPLING | slide 16 native slide text #174/#297 | FLOW COUPLING | native PowerPoint text boxes; depth / MIN ID / MAX OD columns are single multi-line text boxes, rows paired with labels by vertical position on the render (visual check) |
| `components[flow-coupling-1].type` | flow-coupling | slide 16 native slide text #174/#297 | FLOW COUPLING | EXACT TYPE (Component_Types) |
| `components[flow-coupling-1].topMd` | 508.53 | slide 16 native slide text #174/#297 | 155 | 155.0 m x 3.28084 = 508.53 ft (3 dp); m MDDF ('All Depth in : m-MDDF') (equipment depth column as printed) |
| `components[flow-coupling-1].topMdM` | 155.0 | slide 16 native slide text #174/#297 | 155 | original metres; m MDDF ('All Depth in : m-MDDF') (equipment depth column as printed) |
| `components[flow-coupling-1].innerDiameter` | 2.313 | slide 16 native slide text #174/#297 | 2.313 | column 'MIN ID' = explicit ID |
| `components[flow-coupling-1].outerDiameter` | 4.65 | slide 16 native slide text #174/#297 | 4.650 | column 'MAX OD' = explicit OD |
| `components[flow-coupling-2].name` | FLOW COUPLING | slide 16 native slide text #207 | FLOW COUPLING | native PowerPoint text boxes; depth / MIN ID / MAX OD columns are single multi-line text boxes, rows paired with labels by vertical position on the render (visual check) |
| `components[flow-coupling-2].type` | flow-coupling | slide 16 native slide text #207 | FLOW COUPLING | EXACT TYPE (Component_Types) |
| `components[nipple-x-1].name` | X NIPPLE | slide 16 native slide text #205/#297 | X NIPPLE | native PowerPoint text boxes; depth / MIN ID / MAX OD columns are single multi-line text boxes, rows paired with labels by vertical position on the render (visual check) |
| `components[nipple-x-1].type` | nipple (x) | slide 16 native slide text #205/#297 | X NIPPLE | EXACT TYPE (Component_Types) |
| `components[nipple-x-1].topMd` | 948.163 | slide 16 native slide text #205/#297 | 289 | 289.0 m x 3.28084 = 948.163 ft (3 dp); m MDDF ('All Depth in : m-MDDF') (equipment depth column as printed) |
| `components[nipple-x-1].topMdM` | 289.0 | slide 16 native slide text #205/#297 | 289 | original metres; m MDDF ('All Depth in : m-MDDF') (equipment depth column as printed) |
| `components[nipple-x-1].innerDiameter` | 2.313 | slide 16 native slide text #205/#297 | 2.313 | column 'MIN ID' = explicit ID |
| `components[nipple-x-1].outerDiameter` | 4.65 | slide 16 native slide text #205/#297 | 4.650 | column 'MAX OD' = explicit OD |
| `components[spm-1].name` | SPM # 3 | slide 16 native slide text #176/#297 | SPM # 3 | native PowerPoint text boxes; depth / MIN ID / MAX OD columns are single multi-line text boxes, rows paired with labels by vertical position on the render (visual check) |
| `components[spm-1].type` | spm | slide 16 native slide text #176/#297 | SPM # 3 | EXACT TYPE (Component_Types) |
| `components[spm-1].topMd` | 2145.669 | slide 16 native slide text #176/#297 | 654 | 654.0 m x 3.28084 = 2145.669 ft (3 dp); m MDDF ('All Depth in : m-MDDF') (equipment depth column as printed) |
| `components[spm-1].topMdM` | 654.0 | slide 16 native slide text #176/#297 | 654 | original metres; m MDDF ('All Depth in : m-MDDF') (equipment depth column as printed) |
| `components[spm-1].innerDiameter` | 2.441 | slide 16 native slide text #176/#297 | 2.441 | column 'MIN ID' = explicit ID |
| `components[spm-1].outerDiameter` | 4.375 | slide 16 native slide text #176/#297 | 4.375 | column 'MAX OD' = explicit OD |
| `components[spm-2].name` | SPM # 2 | slide 16 native slide text #194/#297 | SPM # 2 | native PowerPoint text boxes; depth / MIN ID / MAX OD columns are single multi-line text boxes, rows paired with labels by vertical position on the render (visual check) |
| `components[spm-2].type` | spm | slide 16 native slide text #194/#297 | SPM # 2 | EXACT TYPE (Component_Types) |
| `components[spm-2].topMd` | 3697.507 | slide 16 native slide text #194/#297 | 1127 | 1127.0 m x 3.28084 = 3697.507 ft (3 dp); m MDDF ('All Depth in : m-MDDF') (equipment depth column as printed) |
| `components[spm-2].topMdM` | 1127.0 | slide 16 native slide text #194/#297 | 1127 | original metres; m MDDF ('All Depth in : m-MDDF') (equipment depth column as printed) |
| `components[spm-2].innerDiameter` | 2.441 | slide 16 native slide text #194/#297 | 2.441 | column 'MIN ID' = explicit ID |
| `components[spm-2].outerDiameter` | 4.375 | slide 16 native slide text #194/#297 | 4.375 | column 'MAX OD' = explicit OD |
| `components[spm-3].name` | SPM # 1 | slide 16 native slide text #195 | SPM # 1 | native PowerPoint text boxes; depth / MIN ID / MAX OD columns are single multi-line text boxes, rows paired with labels by vertical position on the render (visual check) |
| `components[spm-3].type` | spm | slide 16 native slide text #195 | SPM # 1 | EXACT TYPE (Component_Types) |
| `components[sliding-sleeve-1].name` | SSD#3 | slide 16 native slide text #286 | SSD#3 | native PowerPoint text boxes; depth / MIN ID / MAX OD columns are single multi-line text boxes, rows paired with labels by vertical position on the render (visual check) |
| `components[sliding-sleeve-1].type` | sliding-sleeve | slide 16 native slide text #286 | SSD#3 | EXACT TYPE (Component_Types) |
| `components[sliding-sleeve-1].innerDiameter` | 2.441 | slide 16 native slide text #286 | 2.441 | column 'MIN ID' = explicit ID |
| `components[sliding-sleeve-1].outerDiameter` | 4.375 | slide 16 native slide text #286 | 4.375 | column 'MAX OD' = explicit OD |
| `components[packer-1].name` | 9-5/8" BHR HYD PKR | slide 16 native slide text #287/#297 | 9-5/8" BHR HYD PKR | native PowerPoint text boxes; depth / MIN ID / MAX OD columns are single multi-line text boxes, rows paired with labels by vertical position on the render (visual check) |
| `components[packer-1].type` | packer | slide 16 native slide text #287/#297 | 9-5/8" BHR HYD PKR | NEAREST TYPE (flagged) |
| `components[packer-1].topMd` | 5370.735 | slide 16 native slide text #287/#297 | 1637 | 1637.0 m x 3.28084 = 5370.735 ft (3 dp); m MDDF ('All Depth in : m-MDDF') (equipment depth column as printed) |
| `components[packer-1].topMdM` | 1637.0 | slide 16 native slide text #287/#297 | 1637 | original metres; m MDDF ('All Depth in : m-MDDF') (equipment depth column as printed) |
| `components[packer-1].innerDiameter` | 2.441 | slide 16 native slide text #287/#297 | 2.441 | column 'MIN ID' = explicit ID |
| `components[packer-1].outerDiameter` | 4.375 | slide 16 native slide text #287/#297 | 4.375 | column 'MAX OD' = explicit OD |
| `components[sliding-sleeve-2].name` | SSD#2 | slide 16 native slide text #288 | SSD#2 | native PowerPoint text boxes; depth / MIN ID / MAX OD columns are single multi-line text boxes, rows paired with labels by vertical position on the render (visual check) |
| `components[sliding-sleeve-2].type` | sliding-sleeve | slide 16 native slide text #288 | SSD#2 | EXACT TYPE (Component_Types) |
| `components[lc-packer-1].name` | 7" BWD PKR 4.5" x 3.88" Bore | slide 16 native slide text #264/#297 | 7" BWD PKR 4.5" x 3.88" Bore | native PowerPoint text boxes; depth / MIN ID / MAX OD columns are single multi-line text boxes, rows paired with labels by vertical position on the render (visual check) |
| `components[lc-packer-1].type` | packer | slide 16 native slide text #264/#297 | 7" BWD PKR 4.5" x 3.88" Bore | EXACT TYPE (Component_Types) |
| `components[lc-packer-1].topMd` | 6092.52 | slide 16 native slide text #264/#297 | 1857 | 1857.0 m x 3.28084 = 6092.52 ft (3 dp); m MDDF ('All Depth in : m-MDDF') (equipment depth column as printed) |
| `components[lc-packer-1].topMdM` | 1857.0 | slide 16 native slide text #264/#297 | 1857 | original metres; m MDDF ('All Depth in : m-MDDF') (equipment depth column as printed) |
| `components[lc-pbr-1].name` | 3.25"  Seal Bore Assembly | slide 16 native slide text #293 | 3.25"  Seal Bore Assembly | native PowerPoint text boxes; depth / MIN ID / MAX OD columns are single multi-line text boxes, rows paired with labels by vertical position on the render (visual check) |
| `components[lc-pbr-1].type` | pbr | slide 16 native slide text #293 | 3.25"  Seal Bore Assembly | NEAREST TYPE (flagged) |
| `components[lc-pbr-1].innerDiameter` | 3.25 | slide 16 native slide text #293 | 3.25"  Seal Bore Assembly | USER-DEFINED RULE: label size without ID/OD designation -> innerDiameter |
| `components[lc-packer-2].name` | 7" BWD PKR 4.5" x 3.88" Bore | slide 16 native slide text #267 | 7" BWD PKR 4.5" x 3.88" Bore | native PowerPoint text boxes; depth / MIN ID / MAX OD columns are single multi-line text boxes, rows paired with labels by vertical position on the render (visual check) |
| `components[lc-packer-2].type` | packer | slide 16 native slide text #267 | 7" BWD PKR 4.5" x 3.88" Bore | EXACT TYPE (Component_Types) |
| `components[sliding-sleeve-3].name` | SLIMLINE SSD#1 | slide 16 native slide text #177 | SLIMLINE SSD#1 | native PowerPoint text boxes; depth / MIN ID / MAX OD columns are single multi-line text boxes, rows paired with labels by vertical position on the render (visual check) |
| `components[sliding-sleeve-3].type` | sliding-sleeve | slide 16 native slide text #177 | SLIMLINE SSD#1 | EXACT TYPE (Component_Types) |
| `components[lc-pbr-2].name` | 3.25"  Seal Bore Assembly | slide 16 native slide text #292 | 3.25"  Seal Bore Assembly | native PowerPoint text boxes; depth / MIN ID / MAX OD columns are single multi-line text boxes, rows paired with labels by vertical position on the render (visual check) |
| `components[lc-pbr-2].type` | pbr | slide 16 native slide text #292 | 3.25"  Seal Bore Assembly | NEAREST TYPE (flagged) |
| `components[lc-pbr-2].innerDiameter` | 3.25 | slide 16 native slide text #292 | 3.25"  Seal Bore Assembly | USER-DEFINED RULE: label size without ID/OD designation -> innerDiameter |
| `components[lc-packer-3].name` | 7" BWD PKR 4.5" x 3.88" Bore | slide 16 native slide text #295 | 7" BWD PKR 4.5" x 3.88" Bore | native PowerPoint text boxes; depth / MIN ID / MAX OD columns are single multi-line text boxes, rows paired with labels by vertical position on the render (visual check) |
| `components[lc-packer-3].type` | packer | slide 16 native slide text #295 | 7" BWD PKR 4.5" x 3.88" Bore | EXACT TYPE (Component_Types) |
| `components[nipple-no-go-1].name` | XN NO GO NIPPLE | slide 16 native slide text #206/#297 | XN NO GO NIPPLE | native PowerPoint text boxes; depth / MIN ID / MAX OD columns are single multi-line text boxes, rows paired with labels by vertical position on the render (visual check) |
| `components[nipple-no-go-1].type` | nipple (no-go) | slide 16 native slide text #206/#297 | XN NO GO NIPPLE | EXACT TYPE (Component_Types) |
| `components[nipple-no-go-1].topMd` | 9153.544 | slide 16 native slide text #206/#297 | 2790 | 2790.0 m x 3.28084 = 9153.544 ft (3 dp); m MDDF ('All Depth in : m-MDDF') (equipment depth column as printed) |
| `components[nipple-no-go-1].topMdM` | 2790.0 | slide 16 native slide text #206/#297 | 2790 | original metres; m MDDF ('All Depth in : m-MDDF') (equipment depth column as printed) |
| `components[nipple-no-go-1].innerDiameter` | 3.25 | slide 16 native slide text #206/#297 | 3.250 | column 'MIN ID' = explicit ID |
| `components[nipple-no-go-1].outerDiameter` | 5.875 | slide 16 native slide text #206/#297 | 5.875 | column 'MAX OD' = explicit OD |
| `components[eot-1].name` | EOT | slide 16 native slide text #298/#297 | EOT | native PowerPoint text boxes; depth / MIN ID / MAX OD columns are single multi-line text boxes, rows paired with labels by vertical position on the render (visual check) |
| `components[eot-1].type` | eot | slide 16 native slide text #298/#297 | EOT | EXACT TYPE (Component_Types) |
| `components[eot-1].topMd` | 9169.948 | slide 16 native slide text #298/#297 | 2795 | 2795.0 m x 3.28084 = 9169.948 ft (3 dp); m MDDF ('All Depth in : m-MDDF') (equipment depth column as printed) |
| `components[eot-1].topMdM` | 2795.0 | slide 16 native slide text #298/#297 | 2795 | original metres; m MDDF ('All Depth in : m-MDDF') (equipment depth column as printed) |
| `components[eot-1].innerDiameter` | 2.313 | slide 16 native slide text #298/#297 | 2.313 | column 'MIN ID' = explicit ID |
| `components[eot-1].outerDiameter` | 3.21 | slide 16 native slide text #298/#297 | 3.210 | column 'MAX OD' = explicit OD |
| `components[liner-1].name` | 7" TOL / 7" Liner Shoe | slide 16 native slide text header | 7" TOL / 7" Liner Shoe | native PowerPoint text boxes; depth / MIN ID / MAX OD columns are single multi-line text boxes, rows paired with labels by vertical position on the render (visual check) |
| `components[liner-1].type` | liner | slide 16 native slide text header | 7" TOL / 7" Liner Shoe | EXACT TYPE (Component_Types) |
| `components[liner-1].topMd` | 5396.982 | slide 16 native slide text header | 7" TOL (label, depth 1645 printed beside it) | 1645.0 m x 3.28084 = 5396.982 ft (3 dp); TOL from the diagram label (m MDTHF) |
| `components[liner-1].topMdM` | 1645.0 | slide 16 native slide text header | 7" TOL (label, depth 1645 printed beside it) | original metres; TOL from the diagram label (m MDTHF) |
| `components[liner-1].bottomMd` | 9356.956 | slide 16 native slide text header | 7" TOL / 7" Liner Shoe | 2852.0 m x 3.28084 = 9356.956 ft (3 dp); m MDDF as printed |
| `components[liner-1].bottomMdM` | 2852.0 | slide 16 native slide text header | 7" TOL / 7" Liner Shoe | original metres; m MDDF as printed |
| `components[liner-1].innerDiameter` | 7.0 | slide 16 native slide text header | 7 | USER-DEFINED RULE: header size without ID/OD designation -> innerDiameter |
| `components[liner-2].name` | 4-1/2" TOL / 4-1/2" Liner Shoe | slide 16 native slide text header | 4-1/2" TOL / 4-1/2" Liner Shoe | native PowerPoint text boxes; depth / MIN ID / MAX OD columns are single multi-line text boxes, rows paired with labels by vertical position on the render (visual check) |
| `components[liner-2].type` | liner | slide 16 native slide text header | 4-1/2" TOL / 4-1/2" Liner Shoe | EXACT TYPE (Component_Types) |
| `components[liner-2].innerDiameter` | 4.5 | slide 16 native slide text header | 4.5 | USER-DEFINED RULE: header size without ID/OD designation -> innerDiameter |
| `components[casing-shoe-1].name` | 9-5/8" Casing Shoe | slide 16 native slide text diagram | 9-5/8" Casing Shoe | native PowerPoint text boxes; depth / MIN ID / MAX OD columns are single multi-line text boxes, rows paired with labels by vertical position on the render (visual check) |
| `components[casing-shoe-1].type` | casing | slide 16 native slide text diagram | 9-5/8" Casing Shoe | EXACT TYPE (Component_Types) |
| `components[casing-shoe-1].bottomMd` | 5436.352 | slide 16 native slide text diagram | 9-5/8" Casing Shoe (depth 1657 on the line below the label) | 1657.0 m x 3.28084 = 5436.352 ft (3 dp); m MDDF as printed |
| `components[casing-shoe-1].bottomMdM` | 1657.0 | slide 16 native slide text diagram | 9-5/8" Casing Shoe (depth 1657 on the line below the label) | original metres; m MDDF as printed |
| `components[casing-shoe-1].innerDiameter` | 9.625 | slide 16 native slide text diagram | 9-5/8" Casing Shoe | USER-DEFINED RULE: size without ID/OD designation -> innerDiameter |
| `components[tubing-1].name` | 2 7/8", 6.4ppf, L-80,  Vam Ace | slide 16 native slide text header | 2 7/8", 6.4ppf, L-80,  Vam Ace | native PowerPoint text boxes; depth / MIN ID / MAX OD columns are single multi-line text boxes, rows paired with labels by vertical position on the render (visual check) |
| `components[tubing-1].type` | tubing | slide 16 native slide text header | 2 7/8", 6.4ppf, L-80,  Vam Ace | EXACT TYPE (Component_Types) |
| `components[tubing-1].bottomMd` | 9169.948 | slide 16 native slide text header | EOT | 2795.0 m x 3.28084 = 9169.948 ft (3 dp); m MDDF ('All Depth in : m-MDDF'); depth of 'EOT' (eot-1) |
| `components[tubing-1].bottomMdM` | 2795.0 | slide 16 native slide text header | EOT | original metres; m MDDF ('All Depth in : m-MDDF'); depth of 'EOT' (eot-1) |
| `components[tubing-1].innerDiameter` | 2.875 | slide 16 native slide text header | 2 7/8", 6.4ppf, L-80,  Vam Ace | USER-DEFINED RULE: header size without ID/OD designation -> innerDiameter |
| `components[tubing-1].weight` | 6.4 | slide 16 native slide text header | 2 7/8", 6.4ppf, L-80,  Vam Ace | parsed from the label (ppf = lb/ft) |
| `components[tubing-1].grade` | L-80 | slide 16 native slide text header | 2 7/8", 6.4ppf, L-80,  Vam Ace | parsed from the label |
| `components[zone-e-12-13-1].name` | E 12/13 2760-2784 m MDDF | slide 16 native slide text #197 | E 12/13 2760-2784 m MDDF | native PowerPoint text boxes; depth / MIN ID / MAX OD columns are single multi-line text boxes, rows paired with labels by vertical position on the render (visual check) |
| `components[zone-e-12-13-1].type` | perforation | slide 16 native slide text #197 | E 12/13 2760-2784 m MDDF | NEAREST TYPE (flagged) |
| `components[zone-e-12-13-1].topMd` | 9055.118 | slide 16 native slide text #197 | E 12/13 (2760-2784) (FB S3; WIR) | 2760.0 m x 3.28084 = 9055.118 ft (3 dp); m MDDF as printed |
| `components[zone-e-12-13-1].topMdM` | 2760.0 | slide 16 native slide text #197 | E 12/13 (2760-2784) (FB S3; WIR) | original metres; m MDDF as printed |
| `components[zone-e-12-13-1].bottomMd` | 9133.859 | slide 16 native slide text #197 | E 12/13 (2760-2784) (FB S3; WIR) | 2784.0 m x 3.28084 = 9133.859 ft (3 dp); m MDDF as printed |
| `components[zone-e-12-13-1].bottomMdM` | 2784.0 | slide 16 native slide text #197 | E 12/13 (2760-2784) (FB S3; WIR) | original metres; m MDDF as printed |
| `components[zone-e-12-13-1].zone` | E 12/13 | slide 16 native slide text #197 | E 12/13 | zone label as printed |
| `components[zone-e-14-1].name` | E 14 2806-2846 m MDDF | slide 16 native slide text #199 | E 14 2806-2846 m MDDF | native PowerPoint text boxes; depth / MIN ID / MAX OD columns are single multi-line text boxes, rows paired with labels by vertical position on the render (visual check) |
| `components[zone-e-14-1].type` | perforation | slide 16 native slide text #199 | E 14 2806-2846 m MDDF | NEAREST TYPE (flagged) |
| `components[zone-e-14-1].topMd` | 9206.037 | slide 16 native slide text #199 | E 14  (2806-2846) (FB S3; WIR) | 2806.0 m x 3.28084 = 9206.037 ft (3 dp); m MDDF as printed |
| `components[zone-e-14-1].topMdM` | 2806.0 | slide 16 native slide text #199 | E 14  (2806-2846) (FB S3; WIR) | original metres; m MDDF as printed |
| `components[zone-e-14-1].bottomMd` | 9337.271 | slide 16 native slide text #199 | E 14  (2806-2846) (FB S3; WIR) | 2846.0 m x 3.28084 = 9337.271 ft (3 dp); m MDDF as printed |
| `components[zone-e-14-1].bottomMdM` | 2846.0 | slide 16 native slide text #199 | E 14  (2806-2846) (FB S3; WIR) | original metres; m MDDF as printed |
| `components[zone-e-14-1].zone` | E 14 | slide 16 native slide text #199 | E 14 | zone label as printed |
| `components[zone-e-23-1].name` | E 23 3060-3070 m MDDF | slide 16 native slide text #283 | E 23 3060-3070 m MDDF | native PowerPoint text boxes; depth / MIN ID / MAX OD columns are single multi-line text boxes, rows paired with labels by vertical position on the render (visual check) |
| `components[zone-e-23-1].type` | perforation | slide 16 native slide text #283 | E 23 3060-3070 m MDDF | NEAREST TYPE (flagged) |
| `components[zone-e-23-1].topMd` | 10039.37 | slide 16 native slide text #283 | E 23 (FB S3; NDR) (3060-3070) | 3060.0 m x 3.28084 = 10039.37 ft (3 dp); m MDDF as printed |
| `components[zone-e-23-1].topMdM` | 3060.0 | slide 16 native slide text #283 | E 23 (FB S3; NDR) (3060-3070) | original metres; m MDDF as printed |
| `components[zone-e-23-1].bottomMd` | 10072.179 | slide 16 native slide text #283 | E 23 (FB S3; NDR) (3060-3070) | 3070.0 m x 3.28084 = 10072.179 ft (3 dp); m MDDF as printed |
| `components[zone-e-23-1].bottomMdM` | 3070.0 | slide 16 native slide text #283 | E 23 (FB S3; NDR) (3060-3070) | original metres; m MDDF as printed |
| `components[zone-e-23-1].zone` | E 23 | slide 16 native slide text #283 | E 23 | zone label as printed |
| `components[zone-e-23-2].name` | E 23 3100-3110 m MDDF | slide 16 native slide text #283 | E 23 3100-3110 m MDDF | native PowerPoint text boxes; depth / MIN ID / MAX OD columns are single multi-line text boxes, rows paired with labels by vertical position on the render (visual check) |
| `components[zone-e-23-2].type` | perforation | slide 16 native slide text #283 | E 23 3100-3110 m MDDF | NEAREST TYPE (flagged) |
| `components[zone-e-23-2].topMd` | 10170.604 | slide 16 native slide text #283 | (3100-3110) | 3100.0 m x 3.28084 = 10170.604 ft (3 dp); m MDDF as printed |
| `components[zone-e-23-2].topMdM` | 3100.0 | slide 16 native slide text #283 | (3100-3110) | original metres; m MDDF as printed |
| `components[zone-e-23-2].bottomMd` | 10203.412 | slide 16 native slide text #283 | (3100-3110) | 3110.0 m x 3.28084 = 10203.412 ft (3 dp); m MDDF as printed |
| `components[zone-e-23-2].bottomMdM` | 3110.0 | slide 16 native slide text #283 | (3100-3110) | original metres; m MDDF as printed |
| `components[zone-e-23-2].zone` | E 23 | slide 16 native slide text #283 | E 23 | zone label as printed |
| `components[zone-e-32-1].name` | E 32 3160-3170 m MDDF | slide 16 native slide text #204 | E 32 3160-3170 m MDDF | native PowerPoint text boxes; depth / MIN ID / MAX OD columns are single multi-line text boxes, rows paired with labels by vertical position on the render (visual check) |
| `components[zone-e-32-1].type` | perforation | slide 16 native slide text #204 | E 32 3160-3170 m MDDF | NEAREST TYPE (flagged) |
| `components[zone-e-32-1].topMd` | 10367.454 | slide 16 native slide text #204 | E 32 (FB S3; NDR) (3160-3170) | 3160.0 m x 3.28084 = 10367.454 ft (3 dp); m MDDF as printed |
| `components[zone-e-32-1].topMdM` | 3160.0 | slide 16 native slide text #204 | E 32 (FB S3; NDR) (3160-3170) | original metres; m MDDF as printed |
| `components[zone-e-32-1].bottomMd` | 10400.263 | slide 16 native slide text #204 | E 32 (FB S3; NDR) (3160-3170) | 3170.0 m x 3.28084 = 10400.263 ft (3 dp); m MDDF as printed |
| `components[zone-e-32-1].bottomMdM` | 3170.0 | slide 16 native slide text #204 | E 32 (FB S3; NDR) (3160-3170) | original metres; m MDDF as printed |
| `components[zone-e-32-1].zone` | E 32 | slide 16 native slide text #204 | E 32 | zone label as printed |
| `components[pbtd-1].name` | PBTD (label; 3001 printed on its line) | slide 16 native slide text diagram | PBTD (label; 3001 printed on its line) | native PowerPoint text boxes; depth / MIN ID / MAX OD columns are single multi-line text boxes, rows paired with labels by vertical position on the render (visual check) |
| `components[pbtd-1].type` | pbtd | slide 16 native slide text diagram | PBTD (label; 3001 printed on its line) | EXACT TYPE (Component_Types) |
| `components[pbtd-1].topMd` | 9845.801 | slide 16 native slide text diagram | PBTD (label; 3001 printed on its line) | 3001.0 m x 3.28084 = 9845.801 ft (3 dp); m MDDF |
| `components[pbtd-1].topMdM` | 3001.0 | slide 16 native slide text diagram | PBTD (label; 3001 printed on its line) | original metres; m MDDF |
| `wellInfo.completionDate` | 2002-01-15 | slide 16 native slide text header | 15 January 2002 | ISO date () |
| `wellInfo.completionFluid` | - |  |  | not supported by the source -> template empty value ('-' text / null number) |
| `wellInfo.completionType` | SINGLE STRING (B-16ST3 WELL COMPLETION DIAGRAM) | slide 16 native slide text header | SINGLE STRING (B-16ST3 WELL COMPLETION DIAGRAM) | as printed |
| `wellInfo.groundElevation` | None |  |  | not supported by the source -> template empty value ('-' text / null number) |
| `wellInfo.kbElevation` | None |  |  | not supported by the source -> template empty value ('-' text / null number) |
| `wellInfo.maxDeviation` | 80.18 @ 1626m-MDDF | slide 16 native slide text header | 80.18 @ 1626m-MDDF | as printed |
| `wellInfo.name` | Dulang_B16ST3 | slide 16 | B-16ST3 WELL COMPLETION DIAGRAM (DULANG UNIT REVISIT I); Well No. : B-16ST3 | identifier rule Dulang_B<no><suffix>; original label kept in the summary |
| `wellInfo.packerFluid` | - |  |  | not supported by the source -> template empty value ('-' text / null number) |
| `wellInfo.rig` | - |  |  | not supported by the source -> template empty value ('-' text / null number) |
| `wellInfo.surfaceCasing` | - |  |  | not supported by the source -> template empty value ('-' text / null number) |
| `wellInfo.totalDepth` | None |  |  | not supported by the source -> template empty value ('-' text / null number) |
| `wellInfo.trajectoryType` | - |  |  | not supported by the source -> template empty value ('-' text / null number) |
| `wellInfo.tubing` | 2 7/8", 6.4ppf, L-80,  Vam Ace | slide 16 native slide text header | 2 7/8", 6.4ppf, L-80,  Vam Ace | as printed |
| `wellInfo.wellhead` | Cooper cameron, 3-1/8",3000 psi | slide 16 native slide text header | Cooper cameron, 3-1/8",3000 psi | as printed |

## Dulang_B24ST2 — slide 24

Original label: WELL NAME: DULANG B-24ST2 (LOCATION : DULANG-B)
Equipment depth datum: m MDDF (column headers 'DEPTH m MDDF')  ·  source: image24.png (visual transcription of the scanned picture (x3 upscale), OCR cross-check)

### Conflicts / missing / unclear

| kind | where | detail |
|---|---|---|
| Missing | slide 24 image24.png header | tubing '4-1/2" 12.6 PPF L80 13Cr JFE-FOX Outer String' (outer): no documented end-of-string depth -> bottomMd empty |
| Unclear | slide 24 image24.png diagram | 'Double Valve Float Shoe 1480.96' mapped to floatEquipment (no Type_Key for a float shoe); key names from the template sheet Float_Equipment; casingId null (shoe on the screen/inner-string assembly) |
| Unclear | slide 24 obj #14 | 'Oct 2013 PXN Plug set at 1455m-MDTHF': depth 1455.0 m MDTHF not mapped (datum differs) |
| Unclear | slide 24 obj #9 | '8 Mar 2026 Set Paragon Plug at 1418m-MDTHF': depth 1418.0 m MDTHF not mapped (datum differs) |
| Unclear | picture | PROPOSED COMPLETION SCHEMATIC (scanned) — treated as the well's schematic; MIN ID / MAX OD columns empty; zone TVD values not mapped |
| Conflict | slide 24 | slide overlays use m-MDTHF (DF-THF 17.9 m) while the picture is m MDDF; plug depths in MDTHF are not mapped to the MDDF depth axis (no datum shift) |

### Unmapped content (kept here, not in the JSON)

| where | text | reason |
|---|---|---|
| picture header | WATER DEPTH : 78.90 Meter; zone thickness / spacing scale on the right (4 m, 9 m, 5 m, 11 m, 7 m, 45 m, 4 m, 28 m, 14 m, 25 m, 11 m); prepared by ADAM HAREEZI B A AZIZ / AFIQ B A HAMID | note (status / history / survey / other revision) |
| slide 24 | overlay boxes: SSD#1 CLOSED, SSD#2 CLOSED, SSD#3 OPEN, SSD#4 OPEN, SSD#5 OPEN, SSD#6 CLOSED, SSD#7 CLOSED; overlay depth column 'DEPTH m-MDTHF' (231, 1268, 1303, 1429); 'DF-THF : 17.9m'; '7 Mar 2026 RIH 2.29” GR, encounter HUD (sand) at 1432 m- MDTHF (above SSD#6)'; LATEST UPDATE : Sep 2022 | note (status / history / survey / other revision) |
| slide 24 speaker notes | Updated by Pravin on 11th July 2022 Open 24 | slide notes |
| slide 24 obj #52228 | LATEST UPDATE : Sep 2022 | slide text box over/beside the picture (status / history / note) |
| slide 24 obj #52231 | SSD#3 OPEN | slide text box over/beside the picture (status / history / note) |
| slide 24 obj #52233 | SSD#6 CLOSED | slide text box over/beside the picture (status / history / note) |
| slide 24 obj #52234 | SSD#7 CLOSED | slide text box over/beside the picture (status / history / note) |
| slide 24 obj #52236 | SSD#2 CLOSED | slide text box over/beside the picture (status / history / note) |
| slide 24 obj #13 | 7 Mar 2026 RIH 2.29” GR, encounter HUD (sand) at 1432 m- MDTHF (above SSD#6) | slide text box over/beside the picture (status / history / note) |
| slide 24 obj #62 | DEPTH m-MDTHF | slide text box over/beside the picture (status / history / note) |
| slide 24 obj #52280 | 1268 | slide text box over/beside the picture (status / history / note) |
| slide 24 obj #125 | 231 | slide text box over/beside the picture (status / history / note) |
| slide 24 obj #127 | 1429 | slide text box over/beside the picture (status / history / note) |
| slide 24 obj #131 | 1303 | slide text box over/beside the picture (status / history / note) |

### Traceability (every populated value)

| JSON path | value | slide / object / location | original text | conversion / interpretation |
|---|---|---|---|---|
| `components[christmas-tree-1].name` | Conventional Solar Alert 5K Wellhead Single F-F Trim Christmas Tree | slide 24 image24.png header | Conventional Solar Alert 5K Wellhead Single F-F Trim Christmas Tree | visual transcription of the scanned picture (x3 upscale), OCR cross-check |
| `components[christmas-tree-1].type` | christmas-tree | slide 24 image24.png header | Conventional Solar Alert 5K Wellhead Single F-F Trim Christmas Tree | EXACT TYPE (Component_Types) |
| `components[flow-coupling-1].name` | FLOW COUPLING | slide 24 image24.png p1 | FLOW COUPLING | visual transcription of the scanned picture (x3 upscale), OCR cross-check |
| `components[flow-coupling-1].type` | flow-coupling | slide 24 image24.png p1 | FLOW COUPLING | EXACT TYPE (Component_Types) |
| `components[flow-coupling-1].topMd` | 814.042 | slide 24 image24.png p1 | 248.12 | 248.12 m x 3.28084 = 814.042 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[flow-coupling-1].topMdM` | 248.12 | slide 24 image24.png p1 | 248.12 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[trsv-1].name` | TRSSSV | slide 24 image24.png p1 | TRSSSV | visual transcription of the scanned picture (x3 upscale), OCR cross-check |
| `components[trsv-1].type` | trsv | slide 24 image24.png p1 | TRSSSV | EXACT TYPE (Component_Types) |
| `components[trsv-1].topMd` | 816.765 | slide 24 image24.png p1 | 248.95 | 248.95 m x 3.28084 = 816.765 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[trsv-1].topMdM` | 248.95 | slide 24 image24.png p1 | 248.95 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[flow-coupling-2].name` | FLOW COUPLING | slide 24 image24.png p1 | FLOW COUPLING | visual transcription of the scanned picture (x3 upscale), OCR cross-check |
| `components[flow-coupling-2].type` | flow-coupling | slide 24 image24.png p1 | FLOW COUPLING | EXACT TYPE (Component_Types) |
| `components[flow-coupling-2].topMd` | 821.654 | slide 24 image24.png p1 | 250.44 | 250.44 m x 3.28084 = 821.654 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[flow-coupling-2].topMdM` | 250.44 | slide 24 image24.png p1 | 250.44 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[spm-1].name` | 3-1/2" SPM | slide 24 image24.png p1 | 3-1/2" SPM | visual transcription of the scanned picture (x3 upscale), OCR cross-check |
| `components[spm-1].type` | spm | slide 24 image24.png p1 | 3-1/2" SPM | EXACT TYPE (Component_Types) |
| `components[spm-1].topMd` | 1065.617 | slide 24 image24.png p1 | 324.80 | 324.8 m x 3.28084 = 1065.617 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[spm-1].topMdM` | 324.8 | slide 24 image24.png p1 | 324.80 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[spm-2].name` | 3-1/2" SPM | slide 24 image24.png p1 | 3-1/2" SPM | visual transcription of the scanned picture (x3 upscale), OCR cross-check |
| `components[spm-2].type` | spm | slide 24 image24.png p1 | 3-1/2" SPM | EXACT TYPE (Component_Types) |
| `components[spm-2].topMd` | 2054.823 | slide 24 image24.png p1 | 626.31 | 626.31 m x 3.28084 = 2054.823 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[spm-2].topMdM` | 626.31 | slide 24 image24.png p1 | 626.31 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[spm-3].name` | 3-1/2" SPM | slide 24 image24.png p1 | 3-1/2" SPM | visual transcription of the scanned picture (x3 upscale), OCR cross-check |
| `components[spm-3].type` | spm | slide 24 image24.png p1 | 3-1/2" SPM | EXACT TYPE (Component_Types) |
| `components[spm-3].topMd` | 2941.7 | slide 24 image24.png p1 | 896.63 | 896.63 m x 3.28084 = 2941.7 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[spm-3].topMdM` | 896.63 | slide 24 image24.png p1 | 896.63 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[spm-4].name` | 3-1/2" SPM | slide 24 image24.png p1 | 3-1/2" SPM | visual transcription of the scanned picture (x3 upscale), OCR cross-check |
| `components[spm-4].type` | spm | slide 24 image24.png p1 | 3-1/2" SPM | EXACT TYPE (Component_Types) |
| `components[spm-4].topMd` | 3643.406 | slide 24 image24.png p1 | 1110.51 | 1110.51 m x 3.28084 = 3643.406 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[spm-4].topMdM` | 1110.51 | slide 24 image24.png p1 | 1110.51 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[spm-5].name` | 3-1/2" ORIFICE | slide 24 image24.png p1 | 3-1/2" ORIFICE | visual transcription of the scanned picture (x3 upscale), OCR cross-check |
| `components[spm-5].type` | spm | slide 24 image24.png p1 | 3-1/2" ORIFICE | NEAREST TYPE (flagged) |
| `components[spm-5].topMd` | 4128.019 | slide 24 image24.png p1 | 1258.22 | 1258.22 m x 3.28084 = 4128.019 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[spm-5].topMdM` | 1258.22 | slide 24 image24.png p1 | 1258.22 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[pup-joint-1].name` | PDG | slide 24 image24.png p1 | PDG | visual transcription of the scanned picture (x3 upscale), OCR cross-check |
| `components[pup-joint-1].type` | pup-joint | slide 24 image24.png p1 | PDG | UNRESOLVED / NEAREST TYPE (flagged) |
| `components[pup-joint-1].topMd` | 4153.576 | slide 24 image24.png p1 | 1266.01 | 1266.01 m x 3.28084 = 4153.576 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[pup-joint-1].topMdM` | 1266.01 | slide 24 image24.png p1 | 1266.01 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[sliding-sleeve-1].name` | 3-1/2" SSD | slide 24 image24.png p2 | 3-1/2" SSD | visual transcription of the scanned picture (x3 upscale), OCR cross-check |
| `components[sliding-sleeve-1].type` | sliding-sleeve | slide 24 image24.png p2 | 3-1/2" SSD | EXACT TYPE (Component_Types) |
| `components[sliding-sleeve-1].topMd` | 4176.083 | slide 24 image24.png p2 | 1272.87 | 1272.87 m x 3.28084 = 4176.083 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[sliding-sleeve-1].topMdM` | 1272.87 | slide 24 image24.png p2 | 1272.87 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[packer-1].name` | 7" x 3-1/2" AHC PACKER | slide 24 image24.png p2 | 7" x 3-1/2" AHC PACKER | visual transcription of the scanned picture (x3 upscale), OCR cross-check |
| `components[packer-1].type` | packer | slide 24 image24.png p2 | 7" x 3-1/2" AHC PACKER | EXACT TYPE (Component_Types) |
| `components[packer-1].topMd` | 4196.522 | slide 24 image24.png p2 | 1279.10 | 1279.1 m x 3.28084 = 4196.522 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[packer-1].topMdM` | 1279.1 | slide 24 image24.png p2 | 1279.10 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[nipple-no-go-1].name` | 3-1/2" XN-NIPPLE | slide 24 image24.png p2 | 3-1/2" XN-NIPPLE | visual transcription of the scanned picture (x3 upscale), OCR cross-check |
| `components[nipple-no-go-1].type` | nipple (no-go) | slide 24 image24.png p2 | 3-1/2" XN-NIPPLE | EXACT TYPE (Component_Types) |
| `components[nipple-no-go-1].topMd` | 4210.105 | slide 24 image24.png p2 | 1283.24 | 1283.24 m x 3.28084 = 4210.105 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[nipple-no-go-1].topMdM` | 1283.24 | slide 24 image24.png p2 | 1283.24 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[st-slot-locator-1].name` | 3-1/2" NO-GO LOCATOR | slide 24 image24.png p2 | 3-1/2" NO-GO LOCATOR | visual transcription of the scanned picture (x3 upscale), OCR cross-check |
| `components[st-slot-locator-1].type` | st-slot-locator | slide 24 image24.png p2 | 3-1/2" NO-GO LOCATOR | NEAREST TYPE (flagged) |
| `components[st-slot-locator-1].topMd` | 4269.095 | slide 24 image24.png p2 | 1301.22 | 1301.22 m x 3.28084 = 4269.095 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[st-slot-locator-1].topMdM` | 1301.22 | slide 24 image24.png p2 | 1301.22 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[mule-shoe-1].name` | 3-1/2" SELF ALIGNING MULE SHOE | slide 24 image24.png p2 | 3-1/2" SELF ALIGNING MULE SHOE | visual transcription of the scanned picture (x3 upscale), OCR cross-check |
| `components[mule-shoe-1].type` | mule-shoe | slide 24 image24.png p2 | 3-1/2" SELF ALIGNING MULE SHOE | EXACT TYPE (Component_Types) |
| `components[mule-shoe-1].topMd` | 4269.652 | slide 24 image24.png p2 | 1301.39 | 1301.39 m x 3.28084 = 4269.652 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[mule-shoe-1].topMdM` | 1301.39 | slide 24 image24.png p2 | 1301.39 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[lc-packer-1].name` | 7" X 4-1/2" VTA PACKER | slide 24 image24.png p2 | 7" X 4-1/2" VTA PACKER | visual transcription of the scanned picture (x3 upscale), OCR cross-check |
| `components[lc-packer-1].type` | packer | slide 24 image24.png p2 | 7" X 4-1/2" VTA PACKER | EXACT TYPE (Component_Types) |
| `components[lc-packer-1].topMd` | 4267.717 | slide 24 image24.png p2 | 1300.80 | 1300.8 m x 3.28084 = 4267.717 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[lc-packer-1].topMdM` | 1300.8 | slide 24 image24.png p2 | 1300.80 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[lc-crossover-1].name` | 3 Way Sub : 4-1/2" x 4-1/2" x 2-7/8" | slide 24 image24.png p2 | 3 Way Sub : 4-1/2" x 4-1/2" x 2-7/8" | visual transcription of the scanned picture (x3 upscale), OCR cross-check |
| `components[lc-crossover-1].type` | crossover | slide 24 image24.png p2 | 3 Way Sub : 4-1/2" x 4-1/2" x 2-7/8" | NEAREST TYPE (flagged) |
| `components[lc-crossover-1].topMd` | 4284.777 | slide 24 image24.png p2 | 1306.00 | 1306.0 m x 3.28084 = 4284.777 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[lc-crossover-1].topMdM` | 1306.0 | slide 24 image24.png p2 | 1306.00 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[lc-packer-2].name` | 3 Meter Oil Swell Packer | slide 24 image24.png p2 | 3 Meter Oil Swell Packer | visual transcription of the scanned picture (x3 upscale), OCR cross-check |
| `components[lc-packer-2].type` | packer | slide 24 image24.png p2 | 3 Meter Oil Swell Packer | EXACT TYPE (Component_Types) |
| `components[lc-packer-2].topMd` | 4295.998 | slide 24 image24.png p2 | 1309.42 | 1309.42 m x 3.28084 = 4295.998 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[lc-packer-2].topMdM` | 1309.42 | slide 24 image24.png p2 | 1309.42 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[sliding-sleeve-2].name` | 3.22" SLIM SSD Inner String | slide 24 image24.png p2 | 3.22" SLIM SSD Inner String | visual transcription of the scanned picture (x3 upscale), OCR cross-check |
| `components[sliding-sleeve-2].type` | sliding-sleeve | slide 24 image24.png p2 | 3.22" SLIM SSD Inner String | EXACT TYPE (Component_Types) |
| `components[sliding-sleeve-2].topMd` | 4324.016 | slide 24 image24.png p2 | 1317.96 | 1317.96 m x 3.28084 = 4324.016 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[sliding-sleeve-2].topMdM` | 1317.96 | slide 24 image24.png p2 | 1317.96 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[sliding-sleeve-2].innerDiameter` | 3.22 | slide 24 image24.png p2 | 3.22" SLIM SSD Inner String | USER-DEFINED RULE: label size without ID/OD designation -> innerDiameter |
| `components[st-slot-locator-2].name` | No-Go Locator Inner String | slide 24 image24.png p2 | No-Go Locator Inner String | visual transcription of the scanned picture (x3 upscale), OCR cross-check |
| `components[st-slot-locator-2].type` | st-slot-locator | slide 24 image24.png p2 | No-Go Locator Inner String | NEAREST TYPE (flagged) |
| `components[st-slot-locator-2].topMd` | 4343.471 | slide 24 image24.png p2 | 1323.89 | 1323.89 m x 3.28084 = 4343.471 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[st-slot-locator-2].topMdM` | 1323.89 | slide 24 image24.png p2 | 1323.89 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[pbr-1].name` | 3.25" Seal Bore Ext | slide 24 image24.png p2 | 3.25" Seal Bore Ext | visual transcription of the scanned picture (x3 upscale), OCR cross-check |
| `components[pbr-1].type` | pbr | slide 24 image24.png p2 | 3.25" Seal Bore Ext | NEAREST TYPE (flagged) |
| `components[pbr-1].topMd` | 4345.571 | slide 24 image24.png p2 | 1324.53 | 1324.53 m x 3.28084 = 4345.571 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[pbr-1].topMdM` | 1324.53 | slide 24 image24.png p2 | 1324.53 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[pbr-1].innerDiameter` | 3.25 | slide 24 image24.png p2 | 3.25" Seal Bore Ext | USER-DEFINED RULE: label size without ID/OD designation -> innerDiameter |
| `components[lc-packer-3].name` | 3 Meter Oil Swell Packer | slide 24 image24.png p2 | 3 Meter Oil Swell Packer | visual transcription of the scanned picture (x3 upscale), OCR cross-check |
| `components[lc-packer-3].type` | packer | slide 24 image24.png p2 | 3 Meter Oil Swell Packer | EXACT TYPE (Component_Types) |
| `components[lc-packer-3].topMd` | 4362.5 | slide 24 image24.png p2 | 1329.69 | 1329.69 m x 3.28084 = 4362.5 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[lc-packer-3].topMdM` | 1329.69 | slide 24 image24.png p2 | 1329.69 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[sliding-sleeve-3].name` | 3.22" SLIM SSD Inner String | slide 24 image24.png p2 | 3.22" SLIM SSD Inner String | visual transcription of the scanned picture (x3 upscale), OCR cross-check |
| `components[sliding-sleeve-3].type` | sliding-sleeve | slide 24 image24.png p2 | 3.22" SLIM SSD Inner String | EXACT TYPE (Component_Types) |
| `components[sliding-sleeve-3].topMd` | 4369.488 | slide 24 image24.png p2 | 1331.82 | 1331.82 m x 3.28084 = 4369.488 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[sliding-sleeve-3].topMdM` | 1331.82 | slide 24 image24.png p2 | 1331.82 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[sliding-sleeve-3].innerDiameter` | 3.22 | slide 24 image24.png p2 | 3.22" SLIM SSD Inner String | USER-DEFINED RULE: label size without ID/OD designation -> innerDiameter |
| `components[pbr-2].name` | 3.25" Seal Bore Ext | slide 24 image24.png p2 | 3.25" Seal Bore Ext | visual transcription of the scanned picture (x3 upscale), OCR cross-check |
| `components[pbr-2].type` | pbr | slide 24 image24.png p2 | 3.25" Seal Bore Ext | NEAREST TYPE (flagged) |
| `components[pbr-2].topMd` | 4396.293 | slide 24 image24.png p2 | 1339.99 | 1339.99 m x 3.28084 = 4396.293 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[pbr-2].topMdM` | 1339.99 | slide 24 image24.png p2 | 1339.99 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[pbr-2].innerDiameter` | 3.25 | slide 24 image24.png p2 | 3.25" Seal Bore Ext | USER-DEFINED RULE: label size without ID/OD designation -> innerDiameter |
| `components[lc-packer-4].name` | 3 Meter Oil Swell Packer | slide 24 image24.png p2 | 3 Meter Oil Swell Packer | visual transcription of the scanned picture (x3 upscale), OCR cross-check |
| `components[lc-packer-4].type` | packer | slide 24 image24.png p2 | 3 Meter Oil Swell Packer | EXACT TYPE (Component_Types) |
| `components[lc-packer-4].topMd` | 4405.414 | slide 24 image24.png p2 | 1342.77 | 1342.77 m x 3.28084 = 4405.414 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[lc-packer-4].topMdM` | 1342.77 | slide 24 image24.png p2 | 1342.77 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[sliding-sleeve-4].name` | 3.22" SLIM SSD Inner String | slide 24 image24.png p3 | 3.22" SLIM SSD Inner String | visual transcription of the scanned picture (x3 upscale), OCR cross-check |
| `components[sliding-sleeve-4].type` | sliding-sleeve | slide 24 image24.png p3 | 3.22" SLIM SSD Inner String | EXACT TYPE (Component_Types) |
| `components[sliding-sleeve-4].topMd` | 4464.469 | slide 24 image24.png p3 | 1360.77 | 1360.77 m x 3.28084 = 4464.469 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[sliding-sleeve-4].topMdM` | 1360.77 | slide 24 image24.png p3 | 1360.77 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[sliding-sleeve-4].innerDiameter` | 3.22 | slide 24 image24.png p3 | 3.22" SLIM SSD Inner String | USER-DEFINED RULE: label size without ID/OD designation -> innerDiameter |
| `components[pbr-3].name` | 3.25" Seal Bore Ext | slide 24 image24.png p3 | 3.25" Seal Bore Ext | visual transcription of the scanned picture (x3 upscale), OCR cross-check |
| `components[pbr-3].type` | pbr | slide 24 image24.png p3 | 3.25" Seal Bore Ext | NEAREST TYPE (flagged) |
| `components[pbr-3].topMd` | 4486.45 | slide 24 image24.png p3 | 1367.47 | 1367.47 m x 3.28084 = 4486.45 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[pbr-3].topMdM` | 1367.47 | slide 24 image24.png p3 | 1367.47 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[pbr-3].innerDiameter` | 3.25 | slide 24 image24.png p3 | 3.25" Seal Bore Ext | USER-DEFINED RULE: label size without ID/OD designation -> innerDiameter |
| `components[lc-packer-5].name` | 3 Meter Oil Swell Packer | slide 24 image24.png p3 | 3 Meter Oil Swell Packer | visual transcription of the scanned picture (x3 upscale), OCR cross-check |
| `components[lc-packer-5].type` | packer | slide 24 image24.png p3 | 3 Meter Oil Swell Packer | EXACT TYPE (Component_Types) |
| `components[lc-packer-5].topMd` | 4513.55 | slide 24 image24.png p3 | 1375.73 | 1375.73 m x 3.28084 = 4513.55 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[lc-packer-5].topMdM` | 1375.73 | slide 24 image24.png p3 | 1375.73 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[sliding-sleeve-5].name` | 3.22" SLIM SSD Inner String | slide 24 image24.png p3 | 3.22" SLIM SSD Inner String | visual transcription of the scanned picture (x3 upscale), OCR cross-check |
| `components[sliding-sleeve-5].type` | sliding-sleeve | slide 24 image24.png p3 | 3.22" SLIM SSD Inner String | EXACT TYPE (Component_Types) |
| `components[sliding-sleeve-5].topMd` | 4605.545 | slide 24 image24.png p3 | 1403.77 | 1403.77 m x 3.28084 = 4605.545 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[sliding-sleeve-5].topMdM` | 1403.77 | slide 24 image24.png p3 | 1403.77 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[sliding-sleeve-5].innerDiameter` | 3.22 | slide 24 image24.png p3 | 3.22" SLIM SSD Inner String | USER-DEFINED RULE: label size without ID/OD designation -> innerDiameter |
| `components[pbr-4].name` | 3.25" Seal Bore Ext | slide 24 image24.png p3 | 3.25" Seal Bore Ext | visual transcription of the scanned picture (x3 upscale), OCR cross-check |
| `components[pbr-4].type` | pbr | slide 24 image24.png p3 | 3.25" Seal Bore Ext | NEAREST TYPE (flagged) |
| `components[pbr-4].topMd` | 4628.248 | slide 24 image24.png p3 | 1410.69 | 1410.69 m x 3.28084 = 4628.248 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[pbr-4].topMdM` | 1410.69 | slide 24 image24.png p3 | 1410.69 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[pbr-4].innerDiameter` | 3.25 | slide 24 image24.png p3 | 3.25" Seal Bore Ext | USER-DEFINED RULE: label size without ID/OD designation -> innerDiameter |
| `components[lc-packer-6].name` | 3 Meter Oil Swell Packer | slide 24 image24.png p3 | 3 Meter Oil Swell Packer | visual transcription of the scanned picture (x3 upscale), OCR cross-check |
| `components[lc-packer-6].type` | packer | slide 24 image24.png p3 | 3 Meter Oil Swell Packer | EXACT TYPE (Component_Types) |
| `components[lc-packer-6].topMd` | 4660.236 | slide 24 image24.png p3 | 1420.44 | 1420.44 m x 3.28084 = 4660.236 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[lc-packer-6].topMdM` | 1420.44 | slide 24 image24.png p3 | 1420.44 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[sliding-sleeve-6].name` | 3.22" SLIM SSD Inner String | slide 24 image24.png p3 | 3.22" SLIM SSD Inner String | visual transcription of the scanned picture (x3 upscale), OCR cross-check |
| `components[sliding-sleeve-6].type` | sliding-sleeve | slide 24 image24.png p3 | 3.22" SLIM SSD Inner String | EXACT TYPE (Component_Types) |
| `components[sliding-sleeve-6].topMd` | 4743.078 | slide 24 image24.png p3 | 1445.69 | 1445.69 m x 3.28084 = 4743.078 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[sliding-sleeve-6].topMdM` | 1445.69 | slide 24 image24.png p3 | 1445.69 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[sliding-sleeve-6].innerDiameter` | 3.22 | slide 24 image24.png p3 | 3.22" SLIM SSD Inner String | USER-DEFINED RULE: label size without ID/OD designation -> innerDiameter |
| `components[pbr-5].name` | 3.25" Seal Bore Ext | slide 24 image24.png p3 | 3.25" Seal Bore Ext | visual transcription of the scanned picture (x3 upscale), OCR cross-check |
| `components[pbr-5].type` | pbr | slide 24 image24.png p3 | 3.25" Seal Bore Ext | NEAREST TYPE (flagged) |
| `components[pbr-5].topMd` | 4763.156 | slide 24 image24.png p3 | 1451.81 | 1451.81 m x 3.28084 = 4763.156 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[pbr-5].topMdM` | 1451.81 | slide 24 image24.png p3 | 1451.81 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[pbr-5].innerDiameter` | 3.25 | slide 24 image24.png p3 | 3.25" Seal Bore Ext | USER-DEFINED RULE: label size without ID/OD designation -> innerDiameter |
| `components[lc-packer-7].name` | 3 Meter Oil Swell Packer | slide 24 image24.png p3 | 3 Meter Oil Swell Packer | visual transcription of the scanned picture (x3 upscale), OCR cross-check |
| `components[lc-packer-7].type` | packer | slide 24 image24.png p3 | 3 Meter Oil Swell Packer | EXACT TYPE (Component_Types) |
| `components[lc-packer-7].topMd` | 4788.255 | slide 24 image24.png p3 | 1459.46 | 1459.46 m x 3.28084 = 4788.255 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[lc-packer-7].topMdM` | 1459.46 | slide 24 image24.png p3 | 1459.46 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[sliding-sleeve-7].name` | 3.22" SLIM SSD Inner String | slide 24 image24.png p3 | 3.22" SLIM SSD Inner String | visual transcription of the scanned picture (x3 upscale), OCR cross-check |
| `components[sliding-sleeve-7].type` | sliding-sleeve | slide 24 image24.png p3 | 3.22" SLIM SSD Inner String | EXACT TYPE (Component_Types) |
| `components[sliding-sleeve-7].topMd` | 4780.512 | slide 24 image24.png p3 | 1457.10 | 1457.1 m x 3.28084 = 4780.512 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[sliding-sleeve-7].topMdM` | 1457.1 | slide 24 image24.png p3 | 1457.10 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[sliding-sleeve-7].innerDiameter` | 3.22 | slide 24 image24.png p3 | 3.22" SLIM SSD Inner String | USER-DEFINED RULE: label size without ID/OD designation -> innerDiameter |
| `components[nipple-no-go-2].name` | 3.22" SLIM XN-Nipple Inner String | slide 24 image24.png p3 | 3.22" SLIM XN-Nipple Inner String | visual transcription of the scanned picture (x3 upscale), OCR cross-check |
| `components[nipple-no-go-2].type` | nipple (no-go) | slide 24 image24.png p3 | 3.22" SLIM XN-Nipple Inner String | EXACT TYPE (Component_Types) |
| `components[nipple-no-go-2].topMd` | 4825.952 | slide 24 image24.png p3 | 1470.95 | 1470.95 m x 3.28084 = 4825.952 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[nipple-no-go-2].topMdM` | 1470.95 | slide 24 image24.png p3 | 1470.95 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[nipple-no-go-2].innerDiameter` | 3.22 | slide 24 image24.png p3 | 3.22" SLIM XN-Nipple Inner String | USER-DEFINED RULE: label size without ID/OD designation -> innerDiameter |
| `components[mule-shoe-2].name` | 3.22" Self Aligning Muleshoe | slide 24 image24.png p3 | 3.22" Self Aligning Muleshoe | visual transcription of the scanned picture (x3 upscale), OCR cross-check |
| `components[mule-shoe-2].type` | mule-shoe | slide 24 image24.png p3 | 3.22" Self Aligning Muleshoe | EXACT TYPE (Component_Types) |
| `components[mule-shoe-2].topMd` | 4827.625 | slide 24 image24.png p3 | 1471.46 | 1471.46 m x 3.28084 = 4827.625 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[mule-shoe-2].topMdM` | 1471.46 | slide 24 image24.png p3 | 1471.46 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[mule-shoe-2].innerDiameter` | 3.22 | slide 24 image24.png p3 | 3.22" Self Aligning Muleshoe | USER-DEFINED RULE: label size without ID/OD designation -> innerDiameter |
| `components[casing-shoe-1].name` | 7" CASING SHOE | slide 24 image24.png diagram | 7" CASING SHOE | visual transcription of the scanned picture (x3 upscale), OCR cross-check |
| `components[casing-shoe-1].type` | casing | slide 24 image24.png diagram | 7" CASING SHOE | EXACT TYPE (Component_Types) |
| `components[casing-shoe-1].bottomMd` | 4964.928 | slide 24 image24.png diagram | 7" CASING SHOE 1513.31 (highlighted row) | 1513.31 m x 3.28084 = 4964.928 ft (3 dp); m MDDF as printed |
| `components[casing-shoe-1].bottomMdM` | 1513.31 | slide 24 image24.png diagram | 7" CASING SHOE 1513.31 (highlighted row) | original metres; m MDDF as printed |
| `components[casing-shoe-1].innerDiameter` | 7.0 | slide 24 image24.png diagram | 7" CASING SHOE | USER-DEFINED RULE: size without ID/OD designation -> innerDiameter |
| `components[tubing-1].name` | 3-1/2" 9.2 PPF L80 13Cr JFE-FOX Upper Completion | slide 24 image24.png header | 3-1/2" 9.2 PPF L80 13Cr JFE-FOX Upper Completion | visual transcription of the scanned picture (x3 upscale), OCR cross-check |
| `components[tubing-1].type` | tubing | slide 24 image24.png header | 3-1/2" 9.2 PPF L80 13Cr JFE-FOX Upper Completion | EXACT TYPE (Component_Types) |
| `components[tubing-1].bottomMd` | 4269.652 | slide 24 image24.png header | 3-1/2" SELF ALIGNING MULE SHOE | 1301.39 m x 3.28084 = 4269.652 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF'); depth of '3-1/2" SELF ALIGNING MULE SHOE' (mule-shoe-1) |
| `components[tubing-1].bottomMdM` | 1301.39 | slide 24 image24.png header | 3-1/2" SELF ALIGNING MULE SHOE | original metres; m MDDF (column headers 'DEPTH m MDDF'); depth of '3-1/2" SELF ALIGNING MULE SHOE' (mule-shoe-1) |
| `components[tubing-1].innerDiameter` | 3.5 | slide 24 image24.png header | 3-1/2" 9.2 PPF L80 13Cr JFE-FOX Upper Completion | USER-DEFINED RULE: header size without ID/OD designation -> innerDiameter |
| `components[tubing-1].weight` | 9.2 | slide 24 image24.png header | 3-1/2" 9.2 PPF L80 13Cr JFE-FOX Upper Completion | parsed from the label (ppf = lb/ft) |
| `components[tubing-1].grade` | L80 | slide 24 image24.png header | 3-1/2" 9.2 PPF L80 13Cr JFE-FOX Upper Completion | parsed from the label |
| `components[inner-string-tubing-1].name` | 2-7/8" 6.4 PPF L80 13Cr JFE-FOX Inner String | slide 24 image24.png header | 2-7/8" 6.4 PPF L80 13Cr JFE-FOX Inner String | visual transcription of the scanned picture (x3 upscale), OCR cross-check |
| `components[inner-string-tubing-1].type` | tubing | slide 24 image24.png header | 2-7/8" 6.4 PPF L80 13Cr JFE-FOX Inner String | EXACT TYPE (Component_Types) |
| `components[inner-string-tubing-1].bottomMd` | 4827.625 | slide 24 image24.png header | 3.22" Self Aligning Muleshoe | 1471.46 m x 3.28084 = 4827.625 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF'); depth of '3.22" Self Aligning Muleshoe' (mule-shoe-2) |
| `components[inner-string-tubing-1].bottomMdM` | 1471.46 | slide 24 image24.png header | 3.22" Self Aligning Muleshoe | original metres; m MDDF (column headers 'DEPTH m MDDF'); depth of '3.22" Self Aligning Muleshoe' (mule-shoe-2) |
| `components[inner-string-tubing-1].innerDiameter` | 2.875 | slide 24 image24.png header | 2-7/8" 6.4 PPF L80 13Cr JFE-FOX Inner String | USER-DEFINED RULE: header size without ID/OD designation -> innerDiameter |
| `components[inner-string-tubing-1].weight` | 6.4 | slide 24 image24.png header | 2-7/8" 6.4 PPF L80 13Cr JFE-FOX Inner String | parsed from the label (ppf = lb/ft) |
| `components[inner-string-tubing-1].grade` | L80 | slide 24 image24.png header | 2-7/8" 6.4 PPF L80 13Cr JFE-FOX Inner String | parsed from the label |
| `components[tubing-2].name` | 4-1/2" 12.6 PPF L80 13Cr JFE-FOX Outer String | slide 24 image24.png header | 4-1/2" 12.6 PPF L80 13Cr JFE-FOX Outer String | visual transcription of the scanned picture (x3 upscale), OCR cross-check |
| `components[tubing-2].type` | tubing | slide 24 image24.png header | 4-1/2" 12.6 PPF L80 13Cr JFE-FOX Outer String | EXACT TYPE (Component_Types) |
| `components[tubing-2].innerDiameter` | 4.5 | slide 24 image24.png header | 4-1/2" 12.6 PPF L80 13Cr JFE-FOX Outer String | USER-DEFINED RULE: header size without ID/OD designation -> innerDiameter |
| `components[tubing-2].weight` | 12.6 | slide 24 image24.png header | 4-1/2" 12.6 PPF L80 13Cr JFE-FOX Outer String | parsed from the label (ppf = lb/ft) |
| `components[tubing-2].grade` | L80 | slide 24 image24.png header | 4-1/2" 12.6 PPF L80 13Cr JFE-FOX Outer String | parsed from the label |
| `components[zone-e10-11-1].name` | E10-11 1318-1322 m MDDF | slide 24 image24.png p2 | E10-11 1318-1322 m MDDF | visual transcription of the scanned picture (x3 upscale), OCR cross-check |
| `components[zone-e10-11-1].type` | perforation | slide 24 image24.png p2 | E10-11 1318-1322 m MDDF | NEAREST TYPE (flagged) |
| `components[zone-e10-11-1].topMd` | 4324.147 | slide 24 image24.png p2 | E10-11 1318 - 1322 m MDDF 1209.55 - 1213.46 m TVDDF 1214 psi, 193 degF | 1318.0 m x 3.28084 = 4324.147 ft (3 dp); m MDDF as printed |
| `components[zone-e10-11-1].topMdM` | 1318.0 | slide 24 image24.png p2 | E10-11 1318 - 1322 m MDDF 1209.55 - 1213.46 m TVDDF 1214 psi, 193 degF | original metres; m MDDF as printed |
| `components[zone-e10-11-1].bottomMd` | 4337.27 | slide 24 image24.png p2 | E10-11 1318 - 1322 m MDDF 1209.55 - 1213.46 m TVDDF 1214 psi, 193 degF | 1322.0 m x 3.28084 = 4337.27 ft (3 dp); m MDDF as printed |
| `components[zone-e10-11-1].bottomMdM` | 1322.0 | slide 24 image24.png p2 | E10-11 1318 - 1322 m MDDF 1209.55 - 1213.46 m TVDDF 1214 psi, 193 degF | original metres; m MDDF as printed |
| `components[zone-e10-11-1].zone` | E10-11 | slide 24 image24.png p2 | E10-11 | zone label as printed |
| `components[zone-e12-13-1].name` | E12-13 1331-1336 m MDDF | slide 24 image24.png p2 | E12-13 1331-1336 m MDDF | visual transcription of the scanned picture (x3 upscale), OCR cross-check |
| `components[zone-e12-13-1].type` | perforation | slide 24 image24.png p2 | E12-13 1331-1336 m MDDF | NEAREST TYPE (flagged) |
| `components[zone-e12-13-1].topMd` | 4366.798 | slide 24 image24.png p2 | E12-13 1331 - 1336 m MDDF 1222.27 - 1227.17 m TVDDF 1243 psi, 194 degF | 1331.0 m x 3.28084 = 4366.798 ft (3 dp); m MDDF as printed |
| `components[zone-e12-13-1].topMdM` | 1331.0 | slide 24 image24.png p2 | E12-13 1331 - 1336 m MDDF 1222.27 - 1227.17 m TVDDF 1243 psi, 194 degF | original metres; m MDDF as printed |
| `components[zone-e12-13-1].bottomMd` | 4383.202 | slide 24 image24.png p2 | E12-13 1331 - 1336 m MDDF 1222.27 - 1227.17 m TVDDF 1243 psi, 194 degF | 1336.0 m x 3.28084 = 4383.202 ft (3 dp); m MDDF as printed |
| `components[zone-e12-13-1].bottomMdM` | 1336.0 | slide 24 image24.png p2 | E12-13 1331 - 1336 m MDDF 1222.27 - 1227.17 m TVDDF 1243 psi, 194 degF | original metres; m MDDF as printed |
| `components[zone-e12-13-1].zone` | E12-13 | slide 24 image24.png p2 | E12-13 | zone label as printed |
| `components[zone-e14-1].name` | E14 1347-1354 m MDDF | slide 24 image24.png p2 | E14 1347-1354 m MDDF | visual transcription of the scanned picture (x3 upscale), OCR cross-check |
| `components[zone-e14-1].type` | perforation | slide 24 image24.png p2 | E14 1347-1354 m MDDF | NEAREST TYPE (flagged) |
| `components[zone-e14-1].topMd` | 4419.291 | slide 24 image24.png p2 | E14 1347 - 1354 m MDDF 1237.96 - 1244.85 m TVDDF 1452 psi, 196 degF | 1347.0 m x 3.28084 = 4419.291 ft (3 dp); m MDDF as printed |
| `components[zone-e14-1].topMdM` | 1347.0 | slide 24 image24.png p2 | E14 1347 - 1354 m MDDF 1237.96 - 1244.85 m TVDDF 1452 psi, 196 degF | original metres; m MDDF as printed |
| `components[zone-e14-1].bottomMd` | 4442.257 | slide 24 image24.png p2 | E14 1347 - 1354 m MDDF 1237.96 - 1244.85 m TVDDF 1452 psi, 196 degF | 1354.0 m x 3.28084 = 4442.257 ft (3 dp); m MDDF as printed |
| `components[zone-e14-1].bottomMdM` | 1354.0 | slide 24 image24.png p2 | E14 1347 - 1354 m MDDF 1237.96 - 1244.85 m TVDDF 1452 psi, 196 degF | original metres; m MDDF as printed |
| `components[zone-e14-1].zone` | E14 | slide 24 image24.png p2 | E14 | zone label as printed |
| `components[zone-e23-1].name` | E23 1399-1403 m MDDF | slide 24 image24.png p3 | E23 1399-1403 m MDDF | visual transcription of the scanned picture (x3 upscale), OCR cross-check |
| `components[zone-e23-1].type` | perforation | slide 24 image24.png p3 | E23 1399-1403 m MDDF | NEAREST TYPE (flagged) |
| `components[zone-e23-1].topMd` | 4589.895 | slide 24 image24.png p3 | E23 1399 - 1403 m MDDF 1289.13 -1293.07 m TVDDF 1742 psi, 200 degF | 1399.0 m x 3.28084 = 4589.895 ft (3 dp); m MDDF as printed |
| `components[zone-e23-1].topMdM` | 1399.0 | slide 24 image24.png p3 | E23 1399 - 1403 m MDDF 1289.13 -1293.07 m TVDDF 1742 psi, 200 degF | original metres; m MDDF as printed |
| `components[zone-e23-1].bottomMd` | 4603.019 | slide 24 image24.png p3 | E23 1399 - 1403 m MDDF 1289.13 -1293.07 m TVDDF 1742 psi, 200 degF | 1403.0 m x 3.28084 = 4603.019 ft (3 dp); m MDDF as printed |
| `components[zone-e23-1].bottomMdM` | 1403.0 | slide 24 image24.png p3 | E23 1399 - 1403 m MDDF 1289.13 -1293.07 m TVDDF 1742 psi, 200 degF | original metres; m MDDF as printed |
| `components[zone-e23-1].zone` | E23 | slide 24 image24.png p3 | E23 | zone label as printed |
| `components[zone-e34-1].name` | E34 1431-1445 m MDDF | slide 24 image24.png p3 | E34 1431-1445 m MDDF | visual transcription of the scanned picture (x3 upscale), OCR cross-check |
| `components[zone-e34-1].type` | perforation | slide 24 image24.png p3 | E34 1431-1445 m MDDF | NEAREST TYPE (flagged) |
| `components[zone-e34-1].topMd` | 4694.882 | slide 24 image24.png p3 | E34 1431 - 1445 m MDDF 1328.12 - 1334.43 m TVDDF 1526 psi, 205 degF | 1431.0 m x 3.28084 = 4694.882 ft (3 dp); m MDDF as printed |
| `components[zone-e34-1].topMdM` | 1431.0 | slide 24 image24.png p3 | E34 1431 - 1445 m MDDF 1328.12 - 1334.43 m TVDDF 1526 psi, 205 degF | original metres; m MDDF as printed |
| `components[zone-e34-1].bottomMd` | 4740.814 | slide 24 image24.png p3 | E34 1431 - 1445 m MDDF 1328.12 - 1334.43 m TVDDF 1526 psi, 205 degF | 1445.0 m x 3.28084 = 4740.814 ft (3 dp); m MDDF as printed |
| `components[zone-e34-1].bottomMdM` | 1445.0 | slide 24 image24.png p3 | E34 1431 - 1445 m MDDF 1328.12 - 1334.43 m TVDDF 1526 psi, 205 degF | original metres; m MDDF as printed |
| `components[zone-e34-1].zone` | E34 | slide 24 image24.png p3 | E34 | zone label as printed |
| `components[zone-e40-1].name` | E40 1470-1481 m MDDF | slide 24 image24.png p3 | E40 1470-1481 m MDDF | visual transcription of the scanned picture (x3 upscale), OCR cross-check |
| `components[zone-e40-1].type` | perforation | slide 24 image24.png p3 | E40 1470-1481 m MDDF | NEAREST TYPE (flagged) |
| `components[zone-e40-1].topMd` | 4822.835 | slide 24 image24.png p3 | E40 1470 - 1481 m MDDF 1359.04 - 1369.87 m TVDDF 1503 psi, 209 degF | 1470.0 m x 3.28084 = 4822.835 ft (3 dp); m MDDF as printed |
| `components[zone-e40-1].topMdM` | 1470.0 | slide 24 image24.png p3 | E40 1470 - 1481 m MDDF 1359.04 - 1369.87 m TVDDF 1503 psi, 209 degF | original metres; m MDDF as printed |
| `components[zone-e40-1].bottomMd` | 4858.924 | slide 24 image24.png p3 | E40 1470 - 1481 m MDDF 1359.04 - 1369.87 m TVDDF 1503 psi, 209 degF | 1481.0 m x 3.28084 = 4858.924 ft (3 dp); m MDDF as printed |
| `components[zone-e40-1].bottomMdM` | 1481.0 | slide 24 image24.png p3 | E40 1470 - 1481 m MDDF 1359.04 - 1369.87 m TVDDF 1503 psi, 209 degF | original metres; m MDDF as printed |
| `components[zone-e40-1].zone` | E40 | slide 24 image24.png p3 | E40 | zone label as printed |
| `components[pbtd-1].name` | PBTD 1486.84 | slide 24 image24.png diagram | PBTD 1486.84 | visual transcription of the scanned picture (x3 upscale), OCR cross-check |
| `components[pbtd-1].type` | pbtd | slide 24 image24.png diagram | PBTD 1486.84 | EXACT TYPE (Component_Types) |
| `components[pbtd-1].topMd` | 4878.084 | slide 24 image24.png diagram | PBTD 1486.84 | 1486.84 m x 3.28084 = 4878.084 ft (3 dp); m MDDF |
| `components[pbtd-1].topMdM` | 1486.84 | slide 24 image24.png diagram | PBTD 1486.84 | original metres; m MDDF |
| `floatEquipment[0].md` | 4858.793 | slide 24 image24.png diagram | Double Valve Float Shoe 1480.96 | 1480.96 m x 3.28084 (MDDF) |
| `components[int-mechanical-plug-1].name` | Oct 2013 PXN Plug set at 1455m-MDTHF | slide 24 obj #14 | Oct 2013 PXN Plug set at 1455m-MDTHF | visual transcription of the scanned picture (x3 upscale), OCR cross-check |
| `components[int-mechanical-plug-1].type` | mechanical-plug | slide 24 obj #14 | Oct 2013 PXN Plug set at 1455m-MDTHF | NEAREST TYPE (flagged) |
| `components[int-mechanical-plug-2].name` | 8 Mar 2026 Set Paragon Plug at 1418m-MDTHF | slide 24 obj #9 | 8 Mar 2026 Set Paragon Plug at 1418m-MDTHF | visual transcription of the scanned picture (x3 upscale), OCR cross-check |
| `components[int-mechanical-plug-2].type` | mechanical-plug | slide 24 obj #9 | 8 Mar 2026 Set Paragon Plug at 1418m-MDTHF | NEAREST TYPE (flagged) |
| `wellInfo.completionDate` | - |  |  | not supported by the source -> template empty value ('-' text / null number) |
| `wellInfo.completionFluid` | 8.5 ppg Drill Water + 3% KCL (COMPLETION BRINE) | slide 24 image24.png header | 8.5 ppg Drill Water + 3% KCL (COMPLETION BRINE) | as printed |
| `wellInfo.completionType` | PROPOSED COMPLETION SCHEMATIC — Tubing String (OIL PRODUCER); single string | slide 24 image24.png header | PROPOSED COMPLETION SCHEMATIC — Tubing String (OIL PRODUCER); single string | as printed |
| `wellInfo.groundElevation` | None |  |  | not supported by the source -> template empty value ('-' text / null number) |
| `wellInfo.kbElevation` | None |  |  | not supported by the source -> template empty value ('-' text / null number) |
| `wellInfo.maxDeviation` | 41.50 deg @ 922.47 m MDDF /  855.73 m TVDDF | slide 24 image24.png header | 41.50 deg @ 922.47 m MDDF /  855.73 m TVDDF | as printed |
| `wellInfo.name` | Dulang_B24ST2 | slide 24 | WELL NAME: DULANG B-24ST2 (LOCATION : DULANG-B) | identifier rule Dulang_B<no><suffix>; original label kept in the summary |
| `wellInfo.packerFluid` | 8.5 ppg Drill water + KCl based fluid + 0.1% vol. Bacteride + 2% vol. CONQOR 303A Corrosion Inhibitor + 1% vol. O2 scavenger | slide 24 image24.png header | 8.5 ppg Drill water + KCl based fluid + 0.1% vol. Bacteride + 2% vol. CONQOR 303A Corrosion Inhibitor + 1% vol. O2 scavenger | as printed |
| `wellInfo.rig` | KM-1 | slide 24 image24.png header | KM-1 | as printed |
| `wellInfo.surfaceCasing` | - |  |  | not supported by the source -> template empty value ('-' text / null number) |
| `wellInfo.totalDepth` | 4967.192 | slide 24 image24.png header/diagram | TD : 1514.00 m MDDF / 1402.36 m TVDDF | 1514.0 m x 3.28084 = 4967.192 ft; datum MDDF as printed (not shifted) |
| `wellInfo.trajectoryType` | - |  |  | not supported by the source -> template empty value ('-' text / null number) |
| `wellInfo.tubing` | 3-1/2" 9.2 PPF L80 13Cr JFE-FOX Upper Completion; 2-7/8" 6.4 PPF L80 13Cr JFE-FOX Inner String; 4-1/2" 12.6 PPF L80 13Cr JFE-FOX Outer String | slide 24 image24.png header | 3-1/2" 9.2 PPF L80 13Cr JFE-FOX Upper Completion; 2-7/8" 6.4 PPF L80 13Cr JFE-FOX Inner String; 4-1/2" 12.6 PPF L80 13Cr JFE-FOX Outer String | as printed |
| `wellInfo.wellhead` | - |  |  | not supported by the source -> template empty value ('-' text / null number) |

## Dulang_B25ST4 — slide 25

Original label: WELL NAME: B25 ST-4 (LOCATION / SLOT: DULANG B PLATFORM /SLOT 25)
Equipment depth datum: m MDDF (column 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment')  ·  source: image25.png (visual transcription of a low-resolution picture (726x788 px, x3 upscale), OCR cross-check)

### Conflicts / missing / unclear

| kind | where | detail |
|---|---|---|
| Unclear | slide 25 image25.png header/diagram | TD 2155.0 m printed in MDDF; equipment datum m MDDF (column 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') — no datum shift applied |
| Unclear | picture | low-resolution picture: equipment depths and maximum deviation read with low confidence (marked '?'); ID/OD columns readable — a higher-resolution source (Excel in ECMS, per the slide note) should be checked |
| Unclear | slide 25 | slide notes say the latest editable diagram is an Excel file in ECMS (not part of this input) — this picture may not be the latest revision |

### Unmapped content (kept here, not in the JSON)

| where | text | reason |
|---|---|---|
| left table | 13-3/8" 54.5# K55, BTC CSG / 9-5/8" CSG SIDE TRACK WINDOW 515 m-MDDF / 481 m-TVDDF (INC 40.0); 7" CASING PIP TAG 1893 m-MDDF / 1147 m-TVDDF; 7" 29# 13CR-80MY JFE BEAR R3 CSG row at 1651 m-MDDF / 1064 m-TVDDF (INC 73.0; position label on the casing); 'Estimated top of dropped gun' 2099.0 m-MDDF; 'Rat | note (status / history / survey / other revision) |
| picture header | WATER DEPTH : 76.7 m; DFE : 37.1 m-AMSL; DFE-THF : 15.3 m; drift diameters: intermediate 8.679, production casings 6.059, tubing 2.867; Designed by: Diana Madi (Senior Completion Engineer), Daeng Dhamiry (Completion Engineer), Mohd Hisham (Production Technologist); Rev: Final; Date: 11-Aug-22 | note (status / history / survey / other revision) |
| slide 25 | 'FOR B25 LATEST EDITABLE DIAGRAM, PLEASE REFER TO EXCEL FILE IN ECMS - PRAVIN NAIR 2023-'; slide notes 'Updated by Pravin on 11th July 2022' | note (status / history / survey / other revision) |
| slide 25 speaker notes | Updated by Pravin on 11th July 2022 Open 25 | slide notes |
| slide 25 obj #84 | FOR B25 LATEST EDITABLE DIAGRAM, PLEASE REFER TO EXCEL FILE IN ECMS - PRAVIN NAIR 2023- | slide text box over/beside the picture (status / history / note) |

### Traceability (every populated value)

| JSON path | value | slide / object / location | original text | conversion / interpretation |
|---|---|---|---|---|
| `components[christmas-tree-1].name` | Solar Alert 3-1/8" 5K U/HH Trim | slide 25 image25.png header | Solar Alert 3-1/8" 5K U/HH Trim | visual transcription of a low-resolution picture (726x788 px, x3 upscale), OCR cross-check |
| `components[christmas-tree-1].type` | christmas-tree | slide 25 image25.png header | Solar Alert 3-1/8" 5K U/HH Trim | EXACT TYPE (Component_Types) |
| `components[tubing-hanger-1].name` | TUBING HANGER, 13-5/8" x 5"-3SA-2G LH x 3-1/2" 9.2# JFEBEAR, w/ 3" HBPV PROFILE | slide 25 image25.png right table | TUBING HANGER, 13-5/8" x 5"-3SA-2G LH x 3-1/2" 9.2# JFEBEAR, w/ 3" HBPV PROFILE | visual transcription of a low-resolution picture (726x788 px, x3 upscale), OCR cross-check |
| `components[tubing-hanger-1].type` | tubing-hanger | slide 25 image25.png right table | TUBING HANGER, 13-5/8" x 5"-3SA-2G LH x 3-1/2" 9.2# JFEBEAR, w/ 3" HBPV PROFILE | EXACT TYPE (Component_Types) |
| `components[tubing-hanger-1].topMd` | 50.197 | slide 25 image25.png right table | 15.3 | 15.3 m x 3.28084 = 50.197 ft (3 dp); m MDDF (column 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[tubing-hanger-1].topMdM` | 15.3 | slide 25 image25.png right table | 15.3 | original metres; m MDDF (column 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[tubing-hanger-1].innerDiameter` | 2.95 | slide 25 image25.png right table | 2.950 | column 'MIN ID' = explicit ID |
| `components[tubing-hanger-1].outerDiameter` | 13.47 | slide 25 image25.png right table | 13.470 | column 'MAX OD' = explicit OD |
| `components[flow-coupling-1].name` | FLOW COUPLING, 3-1/2", S13CR, JFEBEAR | slide 25 image25.png right table | FLOW COUPLING, 3-1/2", S13CR, JFEBEAR | visual transcription of a low-resolution picture (726x788 px, x3 upscale), OCR cross-check |
| `components[flow-coupling-1].type` | flow-coupling | slide 25 image25.png right table | FLOW COUPLING, 3-1/2", S13CR, JFEBEAR | EXACT TYPE (Component_Types) |
| `components[flow-coupling-1].topMd` | 818.241 | slide 25 image25.png right table | 249.4 | 249.4 m x 3.28084 = 818.241 ft (3 dp); m MDDF (column 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[flow-coupling-1].topMdM` | 249.4 | slide 25 image25.png right table | 249.4 | original metres; m MDDF (column 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[flow-coupling-1].innerDiameter` | 2.992 | slide 25 image25.png right table | 2.992 | column 'MIN ID' = explicit ID |
| `components[flow-coupling-1].outerDiameter` | 3.883 | slide 25 image25.png right table | 3.883 | column 'MAX OD' = explicit OD |
| `components[scssv-1].name` | SCSSV, SELF-EQ, 3-1/2", S13CR, JFEBEAR | slide 25 image25.png right table | SCSSV, SELF-EQ, 3-1/2", S13CR, JFEBEAR | visual transcription of a low-resolution picture (726x788 px, x3 upscale), OCR cross-check |
| `components[scssv-1].type` | scssv | slide 25 image25.png right table | SCSSV, SELF-EQ, 3-1/2", S13CR, JFEBEAR | EXACT TYPE (Component_Types) |
| `components[scssv-1].innerDiameter` | 2.813 | slide 25 image25.png right table | 2.813 | column 'MIN ID' = explicit ID |
| `components[scssv-1].outerDiameter` | 5.0 | slide 25 image25.png right table | 5.000 | column 'MAX OD' = explicit OD |
| `components[flow-coupling-2].name` | FLOW COUPLING, 3-1/2", 13CR, JFEBEAR | slide 25 image25.png right table | FLOW COUPLING, 3-1/2", 13CR, JFEBEAR | visual transcription of a low-resolution picture (726x788 px, x3 upscale), OCR cross-check |
| `components[flow-coupling-2].type` | flow-coupling | slide 25 image25.png right table | FLOW COUPLING, 3-1/2", 13CR, JFEBEAR | EXACT TYPE (Component_Types) |
| `components[flow-coupling-2].innerDiameter` | 2.992 | slide 25 image25.png right table | 2.992 | column 'MIN ID' = explicit ID |
| `components[flow-coupling-2].outerDiameter` | 3.883 | slide 25 image25.png right table | 3.883 | column 'MAX OD' = explicit OD |
| `components[lc-packer-1].name` | Swell Packer #2 | slide 25 image25.png right table | Swell Packer #2 | visual transcription of a low-resolution picture (726x788 px, x3 upscale), OCR cross-check |
| `components[lc-packer-1].type` | packer | slide 25 image25.png right table | Swell Packer #2 | EXACT TYPE (Component_Types) |
| `components[lc-packer-1].topMd` | 1604.331 | slide 25 image25.png right table | 489.0 | 489.0 m x 3.28084 = 1604.331 ft (3 dp); m MDDF (column 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[lc-packer-1].topMdM` | 489.0 | slide 25 image25.png right table | 489.0 | original metres; m MDDF (column 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[lc-packer-2].name` | Swell Packer #1 | slide 25 image25.png right table | Swell Packer #1 | visual transcription of a low-resolution picture (726x788 px, x3 upscale), OCR cross-check |
| `components[lc-packer-2].type` | packer | slide 25 image25.png right table | Swell Packer #1 | EXACT TYPE (Component_Types) |
| `components[lc-packer-2].topMd` | 2244.095 | slide 25 image25.png right table | 684.0 | 684.0 m x 3.28084 = 2244.095 ft (3 dp); m MDDF (column 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[lc-packer-2].topMdM` | 684.0 | slide 25 image25.png right table | 684.0 | original metres; m MDDF (column 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[spm-1].name` | SPM, 3-1/2" 13CR JFEBEAR | slide 25 image25.png right table | SPM, 3-1/2" 13CR JFEBEAR | visual transcription of a low-resolution picture (726x788 px, x3 upscale), OCR cross-check |
| `components[spm-1].type` | spm | slide 25 image25.png right table | SPM, 3-1/2" 13CR JFEBEAR | EXACT TYPE (Component_Types) |
| `components[spm-1].topMd` | 2847.769 | slide 25 image25.png right table | 868.0 | 868.0 m x 3.28084 = 2847.769 ft (3 dp); m MDDF (column 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[spm-1].topMdM` | 868.0 | slide 25 image25.png right table | 868.0 | original metres; m MDDF (column 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[spm-1].innerDiameter` | 2.992 | slide 25 image25.png right table | 2.992 | column 'MIN ID' = explicit ID |
| `components[spm-1].outerDiameter` | 5.313 | slide 25 image25.png right table | 5.313 | column 'MAX OD' = explicit OD |
| `components[spm-2].name` | SPM, 3-1/2", 13CR JFEBEAR | slide 25 image25.png right table | SPM, 3-1/2", 13CR JFEBEAR | visual transcription of a low-resolution picture (726x788 px, x3 upscale), OCR cross-check |
| `components[spm-2].type` | spm | slide 25 image25.png right table | SPM, 3-1/2", 13CR JFEBEAR | EXACT TYPE (Component_Types) |
| `components[spm-2].topMd` | 4853.019 | slide 25 image25.png right table | 1479.2 | 1479.2 m x 3.28084 = 4853.019 ft (3 dp); m MDDF (column 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[spm-2].topMdM` | 1479.2 | slide 25 image25.png right table | 1479.2 | original metres; m MDDF (column 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[spm-2].innerDiameter` | 2.992 | slide 25 image25.png right table | 2.992 | column 'MIN ID' = explicit ID |
| `components[spm-2].outerDiameter` | 5.313 | slide 25 image25.png right table | 5.313 | column 'MAX OD' = explicit OD |
| `components[spm-3].name` | SPM, 3-1/2", 13CR JFEBEAR | slide 25 image25.png right table | SPM, 3-1/2", 13CR JFEBEAR | visual transcription of a low-resolution picture (726x788 px, x3 upscale), OCR cross-check |
| `components[spm-3].type` | spm | slide 25 image25.png right table | SPM, 3-1/2", 13CR JFEBEAR | EXACT TYPE (Component_Types) |
| `components[spm-3].topMd` | 6151.575 | slide 25 image25.png right table | 1875.0 | 1875.0 m x 3.28084 = 6151.575 ft (3 dp); m MDDF (column 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[spm-3].topMdM` | 1875.0 | slide 25 image25.png right table | 1875.0 | original metres; m MDDF (column 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[spm-3].innerDiameter` | 2.992 | slide 25 image25.png right table | 2.992 | column 'MIN ID' = explicit ID |
| `components[spm-3].outerDiameter` | 5.313 | slide 25 image25.png right table | 5.313 | column 'MAX OD' = explicit OD |
| `components[sliding-sleeve-1].name` | SLIMLINE SSD, 3-1/2", 13CR JFEBEAR | slide 25 image25.png right table | SLIMLINE SSD, 3-1/2", 13CR JFEBEAR | visual transcription of a low-resolution picture (726x788 px, x3 upscale), OCR cross-check |
| `components[sliding-sleeve-1].type` | sliding-sleeve | slide 25 image25.png right table | SLIMLINE SSD, 3-1/2", 13CR JFEBEAR | EXACT TYPE (Component_Types) |
| `components[sliding-sleeve-1].topMd` | 6443.898 | slide 25 image25.png right table | 1964.1 | 1964.1 m x 3.28084 = 6443.898 ft (3 dp); m MDDF (column 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[sliding-sleeve-1].topMdM` | 1964.1 | slide 25 image25.png right table | 1964.1 | original metres; m MDDF (column 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[sliding-sleeve-1].innerDiameter` | 2.813 | slide 25 image25.png right table | 2.813 | column 'MIN ID' = explicit ID |
| `components[sliding-sleeve-1].outerDiameter` | 4.2 | slide 25 image25.png right table | 4.200 | column 'MAX OD' = explicit OD |
| `components[packer-1].name` | 7" X 3-1/2" BLUEPACK PACKER, V3 RATING, 13CR JFEBEAR HYDRAULIC SET, PULL-TO-RELEASE | slide 25 image25.png right table | 7" X 3-1/2" BLUEPACK PACKER, V3 RATING, 13CR JFEBEAR HYDRAULIC SET, PULL-TO-RELEASE | visual transcription of a low-resolution picture (726x788 px, x3 upscale), OCR cross-check |
| `components[packer-1].type` | packer | slide 25 image25.png right table | 7" X 3-1/2" BLUEPACK PACKER, V3 RATING, 13CR JFEBEAR HYDRAULIC SET, PULL-TO-RELEASE | EXACT TYPE (Component_Types) |
| `components[packer-1].topMd` | 6569.554 | slide 25 image25.png right table | 2002.4 | 2002.4 m x 3.28084 = 6569.554 ft (3 dp); m MDDF (column 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[packer-1].topMdM` | 2002.4 | slide 25 image25.png right table | 2002.4 | original metres; m MDDF (column 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[packer-1].innerDiameter` | 2.919 | slide 25 image25.png right table | 2.919 | column 'MIN ID' = explicit ID |
| `components[packer-1].outerDiameter` | 5.994 | slide 25 image25.png right table | 5.994 | column 'MAX OD' = explicit OD |
| `components[landing-nipple-1].name` | 2.813" X- LANDING NIPPLE | slide 25 image25.png right table | 2.813" X- LANDING NIPPLE | visual transcription of a low-resolution picture (726x788 px, x3 upscale), OCR cross-check |
| `components[landing-nipple-1].type` | landing-nipple | slide 25 image25.png right table | 2.813" X- LANDING NIPPLE | NEAREST TYPE (flagged) |
| `components[landing-nipple-1].topMd` | 6593.832 | slide 25 image25.png right table | 2009.8 | 2009.8 m x 3.28084 = 6593.832 ft (3 dp); m MDDF (column 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[landing-nipple-1].topMdM` | 2009.8 | slide 25 image25.png right table | 2009.8 | original metres; m MDDF (column 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[landing-nipple-1].innerDiameter` | 2.813 | slide 25 image25.png right table | 2.813 | column 'MIN ID' = explicit ID |
| `components[landing-nipple-1].outerDiameter` | 4.04 | slide 25 image25.png right table | 4.040 | column 'MAX OD' = explicit OD |
| `components[circulating-valve-1].name` | SELF-FILLING SUB, 3-1/2", 13CR JFE BEAR | slide 25 image25.png right table | SELF-FILLING SUB, 3-1/2", 13CR JFE BEAR | visual transcription of a low-resolution picture (726x788 px, x3 upscale), OCR cross-check |
| `components[circulating-valve-1].type` | circulating-valve | slide 25 image25.png right table | SELF-FILLING SUB, 3-1/2", 13CR JFE BEAR | NEAREST TYPE (flagged) |
| `components[circulating-valve-1].topMd` | 6644.685 | slide 25 image25.png right table | 2025.3 | 2025.3 m x 3.28084 = 6644.685 ft (3 dp); m MDDF (column 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[circulating-valve-1].topMdM` | 2025.3 | slide 25 image25.png right table | 2025.3 | original metres; m MDDF (column 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[circulating-valve-1].innerDiameter` | 2.88 | slide 25 image25.png right table | 2.880 | column 'MIN ID' = explicit ID |
| `components[circulating-valve-1].outerDiameter` | 4.93 | slide 25 image25.png right table | 4.930 | column 'MAX OD' = explicit OD |
| `components[mechanical-plug-1].name` | i-MAGNUM DISC, 3-1/2", 13CR JFE BEAR | slide 25 image25.png right table | i-MAGNUM DISC, 3-1/2", 13CR JFE BEAR | visual transcription of a low-resolution picture (726x788 px, x3 upscale), OCR cross-check |
| `components[mechanical-plug-1].type` | mechanical-plug | slide 25 image25.png right table | i-MAGNUM DISC, 3-1/2", 13CR JFE BEAR | NEAREST TYPE (flagged) |
| `components[mechanical-plug-1].topMd` | 6662.402 | slide 25 image25.png right table | 2030.7 | 2030.7 m x 3.28084 = 6662.402 ft (3 dp); m MDDF (column 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[mechanical-plug-1].topMdM` | 2030.7 | slide 25 image25.png right table | 2030.7 | original metres; m MDDF (column 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[mechanical-plug-1].innerDiameter` | 2.92 | slide 25 image25.png right table | 2.920 | column 'MIN ID' = explicit ID |
| `components[mechanical-plug-1].outerDiameter` | 5.0 | slide 25 image25.png right table | 5.000 | column 'MAX OD' = explicit OD |
| `components[nipple-no-go-1].name` | 2.750" XN NIPPLE, 3-1/2", 13CR JFE BEAR | slide 25 image25.png right table | 2.750" XN NIPPLE, 3-1/2", 13CR JFE BEAR | visual transcription of a low-resolution picture (726x788 px, x3 upscale), OCR cross-check |
| `components[nipple-no-go-1].type` | nipple (no-go) | slide 25 image25.png right table | 2.750" XN NIPPLE, 3-1/2", 13CR JFE BEAR | EXACT TYPE (Component_Types) |
| `components[nipple-no-go-1].topMd` | 6720.801 | slide 25 image25.png right table | 2048.5 | 2048.5 m x 3.28084 = 6720.801 ft (3 dp); m MDDF (column 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[nipple-no-go-1].topMdM` | 2048.5 | slide 25 image25.png right table | 2048.5 | original metres; m MDDF (column 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[nipple-no-go-1].innerDiameter` | 2.635 | slide 25 image25.png right table | 2.635 | column 'MIN ID' = explicit ID |
| `components[nipple-no-go-1].outerDiameter` | 4.0 | slide 25 image25.png right table | 4.000 | column 'MAX OD' = explicit OD |
| `components[mule-shoe-1].name` | SELF ALIGNING MULE SHOE, 3-1/2", 13CR JFE BEAR | slide 25 image25.png right table | SELF ALIGNING MULE SHOE, 3-1/2", 13CR JFE BEAR | visual transcription of a low-resolution picture (726x788 px, x3 upscale), OCR cross-check |
| `components[mule-shoe-1].type` | mule-shoe | slide 25 image25.png right table | SELF ALIGNING MULE SHOE, 3-1/2", 13CR JFE BEAR | EXACT TYPE (Component_Types) |
| `components[mule-shoe-1].topMd` | 6730.643 | slide 25 image25.png right table | 2051.5 | 2051.5 m x 3.28084 = 6730.643 ft (3 dp); m MDDF (column 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[mule-shoe-1].topMdM` | 2051.5 | slide 25 image25.png right table | 2051.5 | original metres; m MDDF (column 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[mule-shoe-1].innerDiameter` | 3.015 | slide 25 image25.png right table | 3.015 | column 'MIN ID' = explicit ID |
| `components[mule-shoe-1].outerDiameter` | 3.96 | slide 25 image25.png right table | 3.960 | column 'MAX OD' = explicit OD |
| `components[lc-packer-3].name` | ANNULAR BARRIER PACKER | slide 25 image25.png left table | ANNULAR BARRIER PACKER | visual transcription of a low-resolution picture (726x788 px, x3 upscale), OCR cross-check |
| `components[lc-packer-3].type` | packer | slide 25 image25.png left table | ANNULAR BARRIER PACKER | EXACT TYPE (Component_Types) |
| `components[lc-packer-3].topMd` | 6847.769 | slide 25 image25.png left table | 2087.2 | 2087.2 m x 3.28084 = 6847.769 ft (3 dp); m MDDF (column 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[lc-packer-3].topMdM` | 2087.2 | slide 25 image25.png left table | 2087.2 | original metres; m MDDF (column 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[lc-packer-3].outerDiameter` | 8.25 | slide 25 image25.png left table | 8.250 | column 'MAX OD' = explicit OD |
| `components[surface-casing-1].name` | INTERMEDIATE CASING 9-5/8 40 ppf N-80 BTC | slide 25 image25.png tubular table | INTERMEDIATE CASING 9-5/8 40 ppf N-80 BTC | visual transcription of a low-resolution picture (726x788 px, x3 upscale), OCR cross-check |
| `components[surface-casing-1].type` | casing | slide 25 image25.png tubular table | INTERMEDIATE CASING 9-5/8 40 ppf N-80 BTC | EXACT TYPE (Component_Types) |
| `components[surface-casing-1].bottomMd` | 3280.84 | slide 25 image25.png tubular table | 1000.00 | 1000.0 m x 3.28084 = 3280.84 ft (3 dp); m MDDF (tubular table DEPTH column) |
| `components[surface-casing-1].bottomMdM` | 1000.0 | slide 25 image25.png tubular table | 1000.00 | original metres; m MDDF (tubular table DEPTH column) |
| `components[surface-casing-1].innerDiameter` | 8.835 | slide 25 image25.png tubular table | 8.835 | tubular table 'ID (inch)' = explicit ID |
| `components[surface-casing-1].outerDiameter` | 9.625 | slide 25 image25.png tubular table | 9.625 | tubular table 'OD (inch)' = explicit OD |
| `components[surface-casing-1].weight` | 40.0 | slide 25 image25.png tubular table | INTERMEDIATE CASING 9-5/8 40 ppf N-80 BTC | parsed from the label (ppf = lb/ft) |
| `components[surface-casing-1].grade` | N-80 | slide 25 image25.png tubular table | INTERMEDIATE CASING 9-5/8 40 ppf N-80 BTC | parsed from the label |
| `components[production-casing-1].name` | PRODUCTION CASING #1 7 29 ppf L80 VAMTOP HT (7" 29# L80 VAMTOP HT R3 CSG) | slide 25 image25.png tubular table | PRODUCTION CASING #1 7 29 ppf L80 VAMTOP HT (7" 29# L80 VAMTOP HT R3 CSG) | visual transcription of a low-resolution picture (726x788 px, x3 upscale), OCR cross-check |
| `components[production-casing-1].type` | casing | slide 25 image25.png tubular table | PRODUCTION CASING #1 7 29 ppf L80 VAMTOP HT (7" 29# L80 VAMTOP HT R3 CSG) | EXACT TYPE (Component_Types) |
| `components[production-casing-1].bottomMd` | 1640.42 | slide 25 image25.png tubular table | 500.00 | 500.0 m x 3.28084 = 1640.42 ft (3 dp); m MDDF (tubular table DEPTH column) |
| `components[production-casing-1].bottomMdM` | 500.0 | slide 25 image25.png tubular table | 500.00 | original metres; m MDDF (tubular table DEPTH column) |
| `components[production-casing-1].innerDiameter` | 6.184 | slide 25 image25.png tubular table | 6.184 | tubular table 'ID (inch)' = explicit ID |
| `components[production-casing-1].outerDiameter` | 7.0 | slide 25 image25.png tubular table | 7.000 | tubular table 'OD (inch)' = explicit OD |
| `components[production-casing-1].weight` | 29.0 | slide 25 image25.png tubular table | PRODUCTION CASING #1 7 29 ppf L80 VAMTOP HT (7" 29# L80 VAMTOP HT R3 CSG) | parsed from the label (ppf = lb/ft) |
| `components[production-casing-1].grade` | L80 | slide 25 image25.png tubular table | PRODUCTION CASING #1 7 29 ppf L80 VAMTOP HT (7" 29# L80 VAMTOP HT R3 CSG) | parsed from the label |
| `components[production-casing-2].name` | PRODUCTION CASING #2 7 29 ppf 13Cr-L80 JFE BEAR (7" 29# 13CR-80MY JFE BEAR R3 CSG) | slide 25 image25.png tubular table | PRODUCTION CASING #2 7 29 ppf 13Cr-L80 JFE BEAR (7" 29# 13CR-80MY JFE BEAR R3 CSG) | visual transcription of a low-resolution picture (726x788 px, x3 upscale), OCR cross-check |
| `components[production-casing-2].type` | casing | slide 25 image25.png tubular table | PRODUCTION CASING #2 7 29 ppf 13Cr-L80 JFE BEAR (7" 29# 13CR-80MY JFE BEAR R3 CSG) | EXACT TYPE (Component_Types) |
| `components[production-casing-2].bottomMd` | 7070.21 | slide 25 image25.png tubular table | 2155.00 | 2155.0 m x 3.28084 = 7070.21 ft (3 dp); m MDDF (tubular table DEPTH column) |
| `components[production-casing-2].bottomMdM` | 2155.0 | slide 25 image25.png tubular table | 2155.00 | original metres; m MDDF (tubular table DEPTH column) |
| `components[production-casing-2].innerDiameter` | 6.184 | slide 25 image25.png tubular table | 6.184 | tubular table 'ID (inch)' = explicit ID |
| `components[production-casing-2].outerDiameter` | 7.0 | slide 25 image25.png tubular table | 7.000 | tubular table 'OD (inch)' = explicit OD |
| `components[production-casing-2].weight` | 29.0 | slide 25 image25.png tubular table | PRODUCTION CASING #2 7 29 ppf 13Cr-L80 JFE BEAR (7" 29# 13CR-80MY JFE BEAR R3 CSG) | parsed from the label (ppf = lb/ft) |
| `components[production-casing-2].grade` | 13Cr-L80 | slide 25 image25.png tubular table | PRODUCTION CASING #2 7 29 ppf 13Cr-L80 JFE BEAR (7" 29# 13CR-80MY JFE BEAR R3 CSG) | parsed from the label |
| `components[tubing-1].name` | PRODUCTION TUBING #1 3-1/2 9.2 ppf HP2-13Cr-95 JFE BEAR | slide 25 image25.png tubular table | PRODUCTION TUBING #1 3-1/2 9.2 ppf HP2-13Cr-95 JFE BEAR | visual transcription of a low-resolution picture (726x788 px, x3 upscale), OCR cross-check |
| `components[tubing-1].type` | tubing | slide 25 image25.png tubular table | PRODUCTION TUBING #1 3-1/2 9.2 ppf HP2-13Cr-95 JFE BEAR | EXACT TYPE (Component_Types) |
| `components[tubing-1].bottomMd` | 818.241 | slide 25 image25.png tubular table | PRODUCTION TUBING #1 3-1/2 9.2 ppf HP2-13Cr-95 JFE BEAR 249.40 MDDF | 249.4 m x 3.28084 = 818.241 ft (3 dp); m MDDF (tubular table depth = string end) |
| `components[tubing-1].bottomMdM` | 249.4 | slide 25 image25.png tubular table | PRODUCTION TUBING #1 3-1/2 9.2 ppf HP2-13Cr-95 JFE BEAR 249.40 MDDF | original metres; m MDDF (tubular table depth = string end) |
| `components[tubing-1].innerDiameter` | 2.992 | slide 25 image25.png tubular table | 2.992 | tubular table 'ID (inch)' = explicit ID |
| `components[tubing-1].outerDiameter` | 3.5 | slide 25 image25.png tubular table | 3.500 | tubular table 'OD (inch)' = explicit OD |
| `components[tubing-1].weight` | 9.2 | slide 25 image25.png tubular table | PRODUCTION TUBING #1 3-1/2 9.2 ppf HP2-13Cr-95 JFE BEAR | parsed from the label (ppf = lb/ft) |
| `components[tubing-1].grade` | HP2-13Cr-95 | slide 25 image25.png tubular table | PRODUCTION TUBING #1 3-1/2 9.2 ppf HP2-13Cr-95 JFE BEAR | parsed from the label |
| `components[tubing-2].name` | PRODUCTION TUBING #2 3-1/2 9.2 ppf 13Cr-L80 JFE BEAR | slide 25 image25.png tubular table | PRODUCTION TUBING #2 3-1/2 9.2 ppf 13Cr-L80 JFE BEAR | visual transcription of a low-resolution picture (726x788 px, x3 upscale), OCR cross-check |
| `components[tubing-2].type` | tubing | slide 25 image25.png tubular table | PRODUCTION TUBING #2 3-1/2 9.2 ppf 13Cr-L80 JFE BEAR | EXACT TYPE (Component_Types) |
| `components[tubing-2].bottomMd` | 6730.643 | slide 25 image25.png tubular table | PRODUCTION TUBING #2 3-1/2 9.2 ppf 13Cr-L80 JFE BEAR 2051.50 MDDF | 2051.5 m x 3.28084 = 6730.643 ft (3 dp); m MDDF (tubular table depth = string end) |
| `components[tubing-2].bottomMdM` | 2051.5 | slide 25 image25.png tubular table | PRODUCTION TUBING #2 3-1/2 9.2 ppf 13Cr-L80 JFE BEAR 2051.50 MDDF | original metres; m MDDF (tubular table depth = string end) |
| `components[tubing-2].innerDiameter` | 2.992 | slide 25 image25.png tubular table | 2.992 | tubular table 'ID (inch)' = explicit ID |
| `components[tubing-2].outerDiameter` | 3.5 | slide 25 image25.png tubular table | 3.500 | tubular table 'OD (inch)' = explicit OD |
| `components[tubing-2].weight` | 9.2 | slide 25 image25.png tubular table | PRODUCTION TUBING #2 3-1/2 9.2 ppf 13Cr-L80 JFE BEAR | parsed from the label (ppf = lb/ft) |
| `components[tubing-2].grade` | 13Cr-L80 | slide 25 image25.png tubular table | PRODUCTION TUBING #2 3-1/2 9.2 ppf 13Cr-L80 JFE BEAR | parsed from the label |
| `components[zone-dl1-sand-1].name` | DL1 SAND 2055-2057 m MDDF | slide 25 image25.png left table | DL1 SAND 2055-2057 m MDDF | visual transcription of a low-resolution picture (726x788 px, x3 upscale), OCR cross-check |
| `components[zone-dl1-sand-1].type` | perforation | slide 25 image25.png left table | DL1 SAND 2055-2057 m MDDF | EXACT TYPE (Component_Types) |
| `components[zone-dl1-sand-1].topMd` | 6742.126 | slide 25 image25.png left table | DL1 SAND Top Sand: 2033.5m-MDDF / 1199.5 m-TVDDF; Top Perf: 2055 m-MDDF/ 1207 m-TVDDF; Bottom Perf: 2057 m-MDDF/ 1208 m-TVDDF; Bottom Sand: 2084.1 m-MDDF /1217. | 2055.0 m x 3.28084 = 6742.126 ft (3 dp); m MDDF as printed |
| `components[zone-dl1-sand-1].topMdM` | 2055.0 | slide 25 image25.png left table | DL1 SAND Top Sand: 2033.5m-MDDF / 1199.5 m-TVDDF; Top Perf: 2055 m-MDDF/ 1207 m-TVDDF; Bottom Perf: 2057 m-MDDF/ 1208 m-TVDDF; Bottom Sand: 2084.1 m-MDDF /1217. | original metres; m MDDF as printed |
| `components[zone-dl1-sand-1].bottomMd` | 6748.688 | slide 25 image25.png left table | DL1 SAND Top Sand: 2033.5m-MDDF / 1199.5 m-TVDDF; Top Perf: 2055 m-MDDF/ 1207 m-TVDDF; Bottom Perf: 2057 m-MDDF/ 1208 m-TVDDF; Bottom Sand: 2084.1 m-MDDF /1217. | 2057.0 m x 3.28084 = 6748.688 ft (3 dp); m MDDF as printed |
| `components[zone-dl1-sand-1].bottomMdM` | 2057.0 | slide 25 image25.png left table | DL1 SAND Top Sand: 2033.5m-MDDF / 1199.5 m-TVDDF; Top Perf: 2055 m-MDDF/ 1207 m-TVDDF; Bottom Perf: 2057 m-MDDF/ 1208 m-TVDDF; Bottom Sand: 2084.1 m-MDDF /1217. | original metres; m MDDF as printed |
| `components[zone-dl1-sand-1].zone` | DL1 SAND | slide 25 image25.png left table | DL1 SAND | zone label as printed |
| `components[pbtd-1].name` | PBTD 2113.0 (TVDDF 1228.6; header PBTD : 2113 m-MDDF) | slide 25 image25.png diagram | PBTD 2113.0 (TVDDF 1228.6; header PBTD : 2113 m-MDDF) | visual transcription of a low-resolution picture (726x788 px, x3 upscale), OCR cross-check |
| `components[pbtd-1].type` | pbtd | slide 25 image25.png diagram | PBTD 2113.0 (TVDDF 1228.6; header PBTD : 2113 m-MDDF) | EXACT TYPE (Component_Types) |
| `components[pbtd-1].topMd` | 6932.415 | slide 25 image25.png diagram | PBTD 2113.0 (TVDDF 1228.6; header PBTD : 2113 m-MDDF) | 2113.0 m x 3.28084 = 6932.415 ft (3 dp); m MDDF |
| `components[pbtd-1].topMdM` | 2113.0 | slide 25 image25.png diagram | PBTD 2113.0 (TVDDF 1228.6; header PBTD : 2113 m-MDDF) | original metres; m MDDF |
| `wellInfo.completionDate` | 2022-08-11 | slide 25 image25.png header | 11-Aug-22 | ISO date (two-digit year expanded) |
| `wellInfo.completionFluid` | 8.5ppg Brine (8.33 ppg DW + 3% KCl) filtered to 30 NTU | slide 25 image25.png header | 8.5ppg Brine (8.33 ppg DW + 3% KCl) filtered to 30 NTU | as printed |
| `wellInfo.completionType` | FINAL WELL COMPLETION SCHEMATIC (OIL PRODUCER #1) — Rev: Final, 11-Aug-22 | slide 25 image25.png header | FINAL WELL COMPLETION SCHEMATIC (OIL PRODUCER #1) — Rev: Final, 11-Aug-22 | as printed |
| `wellInfo.groundElevation` | None |  |  | not supported by the source -> template empty value ('-' text / null number) |
| `wellInfo.kbElevation` | None |  |  | not supported by the source -> template empty value ('-' text / null number) |
| `wellInfo.maxDeviation` | ?61.96 deg (blurred) | slide 25 image25.png header | ?61.96 deg (blurred) | as printed |
| `wellInfo.name` | Dulang_B25ST4 | slide 25 | WELL NAME: B25 ST-4 (LOCATION / SLOT: DULANG B PLATFORM /SLOT 25) | identifier rule Dulang_B<no><suffix>; original label kept in the summary |
| `wellInfo.packerFluid` | 8.5ppg Brine (8.33 ppg DW + 3% KCl) + 0.2% Biocide + 2% Oxygen Scavenger + 2% Corrosion Inhibitor | slide 25 image25.png header | 8.5ppg Brine (8.33 ppg DW + 3% KCl) + 0.2% Biocide + 2% Oxygen Scavenger + 2% Corrosion Inhibitor | as printed |
| `wellInfo.rig` | Sapura T9 (TADR) | slide 25 image25.png header | Sapura T9 (TADR) | as printed |
| `wellInfo.surfaceCasing` | - |  |  | not supported by the source -> template empty value ('-' text / null number) |
| `wellInfo.totalDepth` | 7070.21 | slide 25 image25.png header/diagram | TOTAL DEPTH : 2155 m-MDDF | 2155.0 m x 3.28084 = 7070.21 ft; datum MDDF as printed (not shifted) |
| `wellInfo.trajectoryType` | - |  |  | not supported by the source -> template empty value ('-' text / null number) |
| `wellInfo.tubing` | PRODUCTION TUBING #1 3-1/2 9.2 ppf HP2-13Cr-95 JFE BEAR 249.40 m-MDDF; PRODUCTION TUBING #2 3-1/2 9.2 ppf 13Cr-L80 JFE BEAR 2051.50 m-MDDF | slide 25 image25.png header | PRODUCTION TUBING #1 3-1/2 9.2 ppf HP2-13Cr-95 JFE BEAR 249.40 m-MDDF; PRODUCTION TUBING #2 3-1/2 9.2 ppf 13Cr-L80 JFE BEAR 2051.50 m-MDDF | as printed |
| `wellInfo.wellhead` | Solar Alert | slide 25 image25.png header | Solar Alert | as printed |

## Dulang_B28ST2 — slide 28

Original label: WELL NAME: DULANG B-28ST2 (LOCATION : DULANG-B)
Equipment depth datum: m MDDF (column headers 'DEPTH m MDDF')  ·  source: image28.png (visual transcription of the picture (x3 upscale), OCR cross-check)

### Conflicts / missing / unclear

| kind | where | detail |
|---|---|---|
| Conflict | l-flow-coupling-1 | '3-1/2" FLOW COUPLING' at 243.73 m is shallower than the item listed above it (packer-1 at 1222.68 m) — kept as printed |
| Missing | slide 28 image28.png header | 'PRODUCTION CASING 9-5/8" 47 ppf L-80 CS VAM TOP': setting depth not printed |
| Missing | slide 28 image28.png header | tubing '3-1/2" L80 13Cr 9.2 PPF JFE-FOX (Upper Completion)' (Short): no documented end-of-string depth -> bottomMd empty |
| Missing | slide 28 image28.png header | tubing '3-1/2" L80 13Cr 9.2 PPF JFE-FOX (Upper Completion)' (Long): no documented end-of-string depth -> bottomMd empty |
| Missing | slide 28 image28.png header | tubing '2-7/8" SC L80 13Cr 6.4 PPF JFE-FOX (Inner String)' (inner): no documented end-of-string depth -> bottomMd empty |
| Missing | slide 28 | TD not printed |
| Unclear | slide 28 obj #6 | '2.31” PXN Plug @ 1470m MDTHF (Jul 2017)': depth 1470.0 m MDTHF not mapped (datum differs) |
| Unclear | slide 28 obj #7 | '2.2” Slickplug @ 1455m MDTHF (Dec 2017)': depth 1455.0 m MDTHF not mapped (datum differs) |
| Conflict | picture | PBTD on the diagram 'PBTD / BRIDGE PLUG 1550.00' vs header 'PBTD : 1570 m MDDF' — diagram value used, header recorded |
| Conflict | picture p2 | ADJUSTABLE UNION: MIN ID 4.520 > MAX OD 2.950 (columns swapped in the source) — diameters not mapped, both values recorded here |
| Unclear | picture | PROPOSED COMPLETION SCHEMATIC with 'DATE OF COMPLETION : TBA' and 'TD : TBA' — treated as the well's schematic; slide overlays (MDTHF) are later status |
| Unclear | slide 28 | plug depths in the overlays are m MDTHF while the picture is m MDDF — not mapped to the MDDF axis (no datum shift) |

### Unmapped content (kept here, not in the JSON)

| where | text | reason |
|---|---|---|
| picture header | TD : TBA; WATER DEPTH : 78.90 Meter; TUBING HANGER : 20 Meter; 4" L80 13Cr 9.5 PPF VAM FJL (Outer String); prepared by ADAM HAREEZI B A AZIZ / AFIQ B A HAMID | note (status / history / survey / other revision) |
| slide 28 | overlay boxes: 'SSD#1 CLOSED (2024)', CLOSED / OPENED / CLOSED / OPENED status boxes, 'GLOV (10/64”) Feb 2021', '2.3” drift & 2.25” LIB HUD at 1440m MDTHF (fine sand)', '2.3” GR HUD at 1459m MDTHF fine sand (Dec 2017)', LATEST UPDATE: APRIL 2024 BY PRAVIN NAIR; image29.jpeg (small photo, no data) | note (status / history / survey / other revision) |
| slide 28 obj #60418 | LATEST UPDATE: APRIL 2024 BY PRAVIN NAIR | slide text box over/beside the picture (status / history / note) |
| slide 28 obj #60421 | CLOSED | slide text box over/beside the picture (status / history / note) |
| slide 28 obj #60422 | OPENED | slide text box over/beside the picture (status / history / note) |
| slide 28 obj #60423 | CLOSED | slide text box over/beside the picture (status / history / note) |
| slide 28 obj #60424 | OPENED | slide text box over/beside the picture (status / history / note) |
| slide 28 obj #49 | GLOV (10/64”) Feb 2021 | slide text box over/beside the picture (status / history / note) |
| slide 28 obj #60 | 2.3” drift & 2.25” LIB HUD at 1440m MDTHF (fine sand) | slide text box over/beside the picture (status / history / note) |
| slide 28 obj #9 | 2.3” GR HUD at 1459m MDTHF fine sand (Dec 2017) | slide text box over/beside the picture (status / history / note) |
| slide 28 obj #4 | SSD#1 CLOSED (2024) | slide text box over/beside the picture (status / history / note) |
| slide 28 obj #10 | OPENED | slide text box over/beside the picture (status / history / note) |

### Traceability (every populated value)

| JSON path | value | slide / object / location | original text | conversion / interpretation |
|---|---|---|---|---|
| `components[christmas-tree-1].name` | Conventional Solar Alert 5K Wellhead Dual F-F Trim Christmas Tree | slide 28 image28.png header | Conventional Solar Alert 5K Wellhead Dual F-F Trim Christmas Tree | visual transcription of the picture (x3 upscale), OCR cross-check |
| `components[christmas-tree-1].type` | christmas-tree | slide 28 image28.png header | Conventional Solar Alert 5K Wellhead Dual F-F Trim Christmas Tree | EXACT TYPE (Component_Types) |
| `components[s-flow-coupling-1].name` | 3-1/2" FLOW COUPLING | slide 28 image28.png p1 | 3-1/2" FLOW COUPLING | visual transcription of the picture (x3 upscale), OCR cross-check |
| `components[s-flow-coupling-1].type` | flow-coupling | slide 28 image28.png p1 | 3-1/2" FLOW COUPLING | EXACT TYPE (Component_Types) |
| `components[s-flow-coupling-1].topMd` | 832.349 | slide 28 image28.png p1 | 253.70 | 253.7 m x 3.28084 = 832.349 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[s-flow-coupling-1].topMdM` | 253.7 | slide 28 image28.png p1 | 253.70 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[s-flow-coupling-1].innerDiameter` | 2.88 | slide 28 image28.png p1 | 2.880 | column 'MIN ID' = explicit ID |
| `components[s-flow-coupling-1].outerDiameter` | 3.92 | slide 28 image28.png p1 | 3.920 | column 'MAX OD' = explicit OD |
| `components[s-flow-coupling-1].string` | Short | slide 28 image28.png p1 |  | string column of the schematic (short / long / dual packer) as printed |
| `components[s-trsv-1].name` | 3-1/2" TRSSSV | slide 28 image28.png p1 | 3-1/2" TRSSSV | visual transcription of the picture (x3 upscale), OCR cross-check |
| `components[s-trsv-1].type` | trsv | slide 28 image28.png p1 | 3-1/2" TRSSSV | EXACT TYPE (Component_Types) |
| `components[s-trsv-1].topMd` | 835.105 | slide 28 image28.png p1 | 254.54 | 254.54 m x 3.28084 = 835.105 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[s-trsv-1].topMdM` | 254.54 | slide 28 image28.png p1 | 254.54 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[s-trsv-1].innerDiameter` | 2.813 | slide 28 image28.png p1 | 2.813 | column 'MIN ID' = explicit ID |
| `components[s-trsv-1].outerDiameter` | 5.03 | slide 28 image28.png p1 | 5.030 | column 'MAX OD' = explicit OD |
| `components[s-trsv-1].string` | Short | slide 28 image28.png p1 |  | string column of the schematic (short / long / dual packer) as printed |
| `components[s-flow-coupling-2].name` | 3-1/2" FLOW COUPLING | slide 28 image28.png p1 | 3-1/2" FLOW COUPLING | visual transcription of the picture (x3 upscale), OCR cross-check |
| `components[s-flow-coupling-2].type` | flow-coupling | slide 28 image28.png p1 | 3-1/2" FLOW COUPLING | EXACT TYPE (Component_Types) |
| `components[s-flow-coupling-2].topMd` | 839.993 | slide 28 image28.png p1 | 256.03 | 256.03 m x 3.28084 = 839.993 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[s-flow-coupling-2].topMdM` | 256.03 | slide 28 image28.png p1 | 256.03 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[s-flow-coupling-2].innerDiameter` | 2.88 | slide 28 image28.png p1 | 2.880 | column 'MIN ID' = explicit ID |
| `components[s-flow-coupling-2].outerDiameter` | 3.92 | slide 28 image28.png p1 | 3.920 | column 'MAX OD' = explicit OD |
| `components[s-flow-coupling-2].string` | Short | slide 28 image28.png p1 |  | string column of the schematic (short / long / dual packer) as printed |
| `components[s-spm-1].name` | 3-1/2" SPM | slide 28 image28.png p1 | 3-1/2" SPM | visual transcription of the picture (x3 upscale), OCR cross-check |
| `components[s-spm-1].type` | spm | slide 28 image28.png p1 | 3-1/2" SPM | EXACT TYPE (Component_Types) |
| `components[s-spm-1].topMd` | 1331.201 | slide 28 image28.png p1 | 405.75 | 405.75 m x 3.28084 = 1331.201 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[s-spm-1].topMdM` | 405.75 | slide 28 image28.png p1 | 405.75 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[s-spm-1].innerDiameter` | 2.9 | slide 28 image28.png p1 | 2.900 | column 'MIN ID' = explicit ID |
| `components[s-spm-1].outerDiameter` | 5.23 | slide 28 image28.png p1 | 5.230 | column 'MAX OD' = explicit OD |
| `components[s-spm-1].string` | Short | slide 28 image28.png p1 |  | string column of the schematic (short / long / dual packer) as printed |
| `components[s-spm-2].name` | 3-1/2" SPM | slide 28 image28.png p1 | 3-1/2" SPM | visual transcription of the picture (x3 upscale), OCR cross-check |
| `components[s-spm-2].type` | spm | slide 28 image28.png p1 | 3-1/2" SPM | EXACT TYPE (Component_Types) |
| `components[s-spm-2].topMd` | 2297.966 | slide 28 image28.png p1 | 700.42 | 700.42 m x 3.28084 = 2297.966 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[s-spm-2].topMdM` | 700.42 | slide 28 image28.png p1 | 700.42 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[s-spm-2].innerDiameter` | 2.9 | slide 28 image28.png p1 | 2.900 | column 'MIN ID' = explicit ID |
| `components[s-spm-2].outerDiameter` | 5.23 | slide 28 image28.png p1 | 5.230 | column 'MAX OD' = explicit OD |
| `components[s-spm-2].string` | Short | slide 28 image28.png p1 |  | string column of the schematic (short / long / dual packer) as printed |
| `components[s-spm-3].name` | 3-1/2" SPM | slide 28 image28.png p1 | 3-1/2" SPM | visual transcription of the picture (x3 upscale), OCR cross-check |
| `components[s-spm-3].type` | spm | slide 28 image28.png p1 | 3-1/2" SPM | EXACT TYPE (Component_Types) |
| `components[s-spm-3].topMd` | 3139.009 | slide 28 image28.png p1 | 956.77 | 956.77 m x 3.28084 = 3139.009 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[s-spm-3].topMdM` | 956.77 | slide 28 image28.png p1 | 956.77 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[s-spm-3].innerDiameter` | 2.9 | slide 28 image28.png p1 | 2.900 | column 'MIN ID' = explicit ID |
| `components[s-spm-3].outerDiameter` | 5.23 | slide 28 image28.png p1 | 5.230 | column 'MAX OD' = explicit OD |
| `components[s-spm-3].string` | Short | slide 28 image28.png p1 |  | string column of the schematic (short / long / dual packer) as printed |
| `components[s-spm-4].name` | 3-1-2" Orifice | slide 28 image28.png p1 | 3-1-2" Orifice | visual transcription of the picture (x3 upscale), OCR cross-check |
| `components[s-spm-4].type` | spm | slide 28 image28.png p1 | 3-1-2" Orifice | NEAREST TYPE (flagged) |
| `components[s-spm-4].topMd` | 3605.479 | slide 28 image28.png p1 | 1098.95 | 1098.95 m x 3.28084 = 3605.479 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[s-spm-4].topMdM` | 1098.95 | slide 28 image28.png p1 | 1098.95 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[s-spm-4].innerDiameter` | 2.9 | slide 28 image28.png p1 | 2.900 | column 'MIN ID' = explicit ID |
| `components[s-spm-4].outerDiameter` | 5.23 | slide 28 image28.png p1 | 5.230 | column 'MAX OD' = explicit OD |
| `components[s-spm-4].string` | Short | slide 28 image28.png p1 |  | string column of the schematic (short / long / dual packer) as printed |
| `components[s-nipple-x-1].name` | 3-1/2" X-Nipple | slide 28 image28.png p1 | 3-1/2" X-Nipple | visual transcription of the picture (x3 upscale), OCR cross-check |
| `components[s-nipple-x-1].type` | nipple (x) | slide 28 image28.png p1 | 3-1/2" X-Nipple | EXACT TYPE (Component_Types) |
| `components[s-nipple-x-1].topMd` | 3945.538 | slide 28 image28.png p1 | 1202.60 | 1202.6 m x 3.28084 = 3945.538 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[s-nipple-x-1].topMdM` | 1202.6 | slide 28 image28.png p1 | 1202.60 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[s-nipple-x-1].innerDiameter` | 2.813 | slide 28 image28.png p1 | 2.813 | column 'MIN ID' = explicit ID |
| `components[s-nipple-x-1].outerDiameter` | 3.96 | slide 28 image28.png p1 | 3.960 | column 'MAX OD' = explicit OD |
| `components[s-nipple-x-1].string` | Short | slide 28 image28.png p1 |  | string column of the schematic (short / long / dual packer) as printed |
| `components[packer-1].name` | 9-5/8" x 3-1/2" DUAL AHC PACKER | slide 28 image28.png p1 | 9-5/8" x 3-1/2" DUAL AHC PACKER | visual transcription of the picture (x3 upscale), OCR cross-check |
| `components[packer-1].type` | packer | slide 28 image28.png p1 | 9-5/8" x 3-1/2" DUAL AHC PACKER | EXACT TYPE (Component_Types) |
| `components[packer-1].topMd` | 4011.417 | slide 28 image28.png p1 | 1222.68 | 1222.68 m x 3.28084 = 4011.417 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[packer-1].topMdM` | 1222.68 | slide 28 image28.png p1 | 1222.68 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[packer-1].innerDiameter` | 2.885 | slide 28 image28.png p1 | 2.885 | column 'MIN ID' = explicit ID |
| `components[packer-1].outerDiameter` | 8.44 | slide 28 image28.png p1 | 8.440 | column 'MAX OD' = explicit OD |
| `components[packer-1].string` | Short & Long | slide 28 image28.png p1 |  | string column of the schematic (short / long / dual packer) as printed |
| `components[s-pup-joint-1].name` | ADJUSTABLE UNION | slide 28 image28.png p2 | ADJUSTABLE UNION | visual transcription of the picture (x3 upscale), OCR cross-check |
| `components[s-pup-joint-1].type` | pup-joint | slide 28 image28.png p2 | ADJUSTABLE UNION | UNRESOLVED / NEAREST TYPE (flagged) |
| `components[s-pup-joint-1].topMd` | 4084.646 | slide 28 image28.png p2 | 1245.00 | 1245.0 m x 3.28084 = 4084.646 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[s-pup-joint-1].topMdM` | 1245.0 | slide 28 image28.png p2 | 1245.00 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[s-pup-joint-1].string` | Short | slide 28 image28.png p2 |  | string column of the schematic (short / long / dual packer) as printed |
| `components[s-sliding-sleeve-1].name` | 3-1/2" SSD | slide 28 image28.png p2 | 3-1/2" SSD | visual transcription of the picture (x3 upscale), OCR cross-check |
| `components[s-sliding-sleeve-1].type` | sliding-sleeve | slide 28 image28.png p2 | 3-1/2" SSD | EXACT TYPE (Component_Types) |
| `components[s-sliding-sleeve-1].topMd` | 4168.307 | slide 28 image28.png p2 | 1270.50 | 1270.5 m x 3.28084 = 4168.307 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[s-sliding-sleeve-1].topMdM` | 1270.5 | slide 28 image28.png p2 | 1270.50 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[s-sliding-sleeve-1].innerDiameter` | 2.813 | slide 28 image28.png p2 | 2.813 | column 'MIN ID' = explicit ID |
| `components[s-sliding-sleeve-1].outerDiameter` | 4.55 | slide 28 image28.png p2 | 4.550 | column 'MAX OD' = explicit OD |
| `components[s-sliding-sleeve-1].string` | Short | slide 28 image28.png p2 |  | string column of the schematic (short / long / dual packer) as printed |
| `components[s-nipple-no-go-1].name` | 3-1/2" XN-NIPPLE | slide 28 image28.png p2 | 3-1/2" XN-NIPPLE | visual transcription of the picture (x3 upscale), OCR cross-check |
| `components[s-nipple-no-go-1].type` | nipple (no-go) | slide 28 image28.png p2 | 3-1/2" XN-NIPPLE | EXACT TYPE (Component_Types) |
| `components[s-nipple-no-go-1].topMd` | 4228.215 | slide 28 image28.png p2 | 1288.76 | 1288.76 m x 3.28084 = 4228.215 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[s-nipple-no-go-1].topMdM` | 1288.76 | slide 28 image28.png p2 | 1288.76 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[s-nipple-no-go-1].innerDiameter` | 2.69 | slide 28 image28.png p2 | 2.690 | column 'MIN ID' = explicit ID |
| `components[s-nipple-no-go-1].outerDiameter` | 3.96 | slide 28 image28.png p2 | 3.960 | column 'MAX OD' = explicit OD |
| `components[s-nipple-no-go-1].string` | Short | slide 28 image28.png p2 |  | string column of the schematic (short / long / dual packer) as printed |
| `components[l-flow-coupling-1].name` | 3-1/2" FLOW COUPLING | slide 28 image28.png p1 | 3-1/2" FLOW COUPLING | visual transcription of the picture (x3 upscale), OCR cross-check |
| `components[l-flow-coupling-1].type` | flow-coupling | slide 28 image28.png p1 | 3-1/2" FLOW COUPLING | EXACT TYPE (Component_Types) |
| `components[l-flow-coupling-1].topMd` | 799.639 | slide 28 image28.png p1 | 243.73 | 243.73 m x 3.28084 = 799.639 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[l-flow-coupling-1].topMdM` | 243.73 | slide 28 image28.png p1 | 243.73 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[l-flow-coupling-1].innerDiameter` | 2.88 | slide 28 image28.png p1 | 2.880 | column 'MIN ID' = explicit ID |
| `components[l-flow-coupling-1].outerDiameter` | 3.92 | slide 28 image28.png p1 | 3.920 | column 'MAX OD' = explicit OD |
| `components[l-flow-coupling-1].string` | Long | slide 28 image28.png p1 |  | string column of the schematic (short / long / dual packer) as printed |
| `components[l-trsv-1].name` | 3-1/2" TRSSSV | slide 28 image28.png p1 | 3-1/2" TRSSSV | visual transcription of the picture (x3 upscale), OCR cross-check |
| `components[l-trsv-1].type` | trsv | slide 28 image28.png p1 | 3-1/2" TRSSSV | EXACT TYPE (Component_Types) |
| `components[l-trsv-1].topMd` | 802.362 | slide 28 image28.png p1 | 244.56 | 244.56 m x 3.28084 = 802.362 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[l-trsv-1].topMdM` | 244.56 | slide 28 image28.png p1 | 244.56 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[l-trsv-1].innerDiameter` | 2.813 | slide 28 image28.png p1 | 2.813 | column 'MIN ID' = explicit ID |
| `components[l-trsv-1].outerDiameter` | 5.03 | slide 28 image28.png p1 | 5.030 | column 'MAX OD' = explicit OD |
| `components[l-trsv-1].string` | Long | slide 28 image28.png p1 |  | string column of the schematic (short / long / dual packer) as printed |
| `components[l-flow-coupling-2].name` | 3-1/2" FLOW COUPLING | slide 28 image28.png p1 | 3-1/2" FLOW COUPLING | visual transcription of the picture (x3 upscale), OCR cross-check |
| `components[l-flow-coupling-2].type` | flow-coupling | slide 28 image28.png p1 | 3-1/2" FLOW COUPLING | EXACT TYPE (Component_Types) |
| `components[l-flow-coupling-2].topMd` | 807.251 | slide 28 image28.png p1 | 246.05 | 246.05 m x 3.28084 = 807.251 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[l-flow-coupling-2].topMdM` | 246.05 | slide 28 image28.png p1 | 246.05 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[l-flow-coupling-2].innerDiameter` | 2.88 | slide 28 image28.png p1 | 2.880 | column 'MIN ID' = explicit ID |
| `components[l-flow-coupling-2].outerDiameter` | 3.92 | slide 28 image28.png p1 | 3.920 | column 'MAX OD' = explicit OD |
| `components[l-flow-coupling-2].string` | Long | slide 28 image28.png p1 |  | string column of the schematic (short / long / dual packer) as printed |
| `components[l-spm-1].name` | 3-1/2" SPM | slide 28 image28.png p1 | 3-1/2" SPM | visual transcription of the picture (x3 upscale), OCR cross-check |
| `components[l-spm-1].type` | spm | slide 28 image28.png p1 | 3-1/2" SPM | EXACT TYPE (Component_Types) |
| `components[l-spm-1].topMd` | 1361.844 | slide 28 image28.png p1 | 415.09 | 415.09 m x 3.28084 = 1361.844 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[l-spm-1].topMdM` | 415.09 | slide 28 image28.png p1 | 415.09 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[l-spm-1].innerDiameter` | 2.9 | slide 28 image28.png p1 | 2.900 | column 'MIN ID' = explicit ID |
| `components[l-spm-1].outerDiameter` | 5.23 | slide 28 image28.png p1 | 5.230 | column 'MAX OD' = explicit OD |
| `components[l-spm-1].string` | Long | slide 28 image28.png p1 |  | string column of the schematic (short / long / dual packer) as printed |
| `components[l-spm-2].name` | 3-1/2" SPM | slide 28 image28.png p1 | 3-1/2" SPM | visual transcription of the picture (x3 upscale), OCR cross-check |
| `components[l-spm-2].type` | spm | slide 28 image28.png p1 | 3-1/2" SPM | EXACT TYPE (Component_Types) |
| `components[l-spm-2].topMd` | 2234.58 | slide 28 image28.png p1 | 681.10 | 681.1 m x 3.28084 = 2234.58 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[l-spm-2].topMdM` | 681.1 | slide 28 image28.png p1 | 681.10 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[l-spm-2].innerDiameter` | 2.9 | slide 28 image28.png p1 | 2.900 | column 'MIN ID' = explicit ID |
| `components[l-spm-2].outerDiameter` | 5.23 | slide 28 image28.png p1 | 5.230 | column 'MAX OD' = explicit OD |
| `components[l-spm-2].string` | Long | slide 28 image28.png p1 |  | string column of the schematic (short / long / dual packer) as printed |
| `components[l-spm-3].name` | 3-1/2" SPM | slide 28 image28.png p1 | 3-1/2" SPM | visual transcription of the picture (x3 upscale), OCR cross-check |
| `components[l-spm-3].type` | spm | slide 28 image28.png p1 | 3-1/2" SPM | EXACT TYPE (Component_Types) |
| `components[l-spm-3].topMd` | 3014.501 | slide 28 image28.png p1 | 918.82 | 918.82 m x 3.28084 = 3014.501 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[l-spm-3].topMdM` | 918.82 | slide 28 image28.png p1 | 918.82 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[l-spm-3].innerDiameter` | 2.9 | slide 28 image28.png p1 | 2.900 | column 'MIN ID' = explicit ID |
| `components[l-spm-3].outerDiameter` | 5.23 | slide 28 image28.png p1 | 5.230 | column 'MAX OD' = explicit OD |
| `components[l-spm-3].string` | Long | slide 28 image28.png p1 |  | string column of the schematic (short / long / dual packer) as printed |
| `components[l-spm-4].name` | 3-1-2" Orifice | slide 28 image28.png p1 | 3-1-2" Orifice | visual transcription of the picture (x3 upscale), OCR cross-check |
| `components[l-spm-4].type` | spm | slide 28 image28.png p1 | 3-1-2" Orifice | NEAREST TYPE (flagged) |
| `components[l-spm-4].topMd` | 3636.352 | slide 28 image28.png p1 | 1108.36 | 1108.36 m x 3.28084 = 3636.352 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[l-spm-4].topMdM` | 1108.36 | slide 28 image28.png p1 | 1108.36 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[l-spm-4].innerDiameter` | 2.9 | slide 28 image28.png p1 | 2.900 | column 'MIN ID' = explicit ID |
| `components[l-spm-4].outerDiameter` | 5.23 | slide 28 image28.png p1 | 5.230 | column 'MAX OD' = explicit OD |
| `components[l-spm-4].string` | Long | slide 28 image28.png p1 |  | string column of the schematic (short / long / dual packer) as printed |
| `components[l-sliding-sleeve-1].name` | 3-1/2" SSD | slide 28 image28.png p1 | 3-1/2" SSD | visual transcription of the picture (x3 upscale), OCR cross-check |
| `components[l-sliding-sleeve-1].type` | sliding-sleeve | slide 28 image28.png p1 | 3-1/2" SSD | EXACT TYPE (Component_Types) |
| `components[l-sliding-sleeve-1].topMd` | 3918.242 | slide 28 image28.png p1 | 1194.28 | 1194.28 m x 3.28084 = 3918.242 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[l-sliding-sleeve-1].topMdM` | 1194.28 | slide 28 image28.png p1 | 1194.28 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[l-sliding-sleeve-1].innerDiameter` | 2.813 | slide 28 image28.png p1 | 2.813 | column 'MIN ID' = explicit ID |
| `components[l-sliding-sleeve-1].outerDiameter` | 4.55 | slide 28 image28.png p1 | 4.550 | column 'MAX OD' = explicit OD |
| `components[l-sliding-sleeve-1].string` | Long | slide 28 image28.png p1 |  | string column of the schematic (short / long / dual packer) as printed |
| `components[l-pup-joint-1].name` | PDG | slide 28 image28.png p2 | PDG | visual transcription of the picture (x3 upscale), OCR cross-check |
| `components[l-pup-joint-1].type` | pup-joint | slide 28 image28.png p2 | PDG | UNRESOLVED / NEAREST TYPE (flagged) |
| `components[l-pup-joint-1].topMd` | 4075.656 | slide 28 image28.png p2 | 1242.26 | 1242.26 m x 3.28084 = 4075.656 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[l-pup-joint-1].topMdM` | 1242.26 | slide 28 image28.png p2 | 1242.26 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[l-pup-joint-1].string` | Long | slide 28 image28.png p2 |  | string column of the schematic (short / long / dual packer) as printed |
| `components[crossover-1].name` | 3-1/2"x 3-1/2" x 4" x 2-7/8" Y-BLOCK ASSEMBLY | slide 28 image28.png p2 | 3-1/2"x 3-1/2" x 4" x 2-7/8" Y-BLOCK ASSEMBLY | visual transcription of the picture (x3 upscale), OCR cross-check |
| `components[crossover-1].type` | crossover | slide 28 image28.png p2 | 3-1/2"x 3-1/2" x 4" x 2-7/8" Y-BLOCK ASSEMBLY | NEAREST TYPE (flagged) |
| `components[crossover-1].topMd` | 4243.963 | slide 28 image28.png p2 | 1293.56 | 1293.56 m x 3.28084 = 4243.963 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[crossover-1].topMdM` | 1293.56 | slide 28 image28.png p2 | 1293.56 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[crossover-1].innerDiameter` | 2.441 | slide 28 image28.png p2 | 2.441 | column 'MIN ID' = explicit ID |
| `components[crossover-1].outerDiameter` | 8.32 | slide 28 image28.png p2 | 8.320 | column 'MAX OD' = explicit OD |
| `components[crossover-1].string` | Short & Long | slide 28 image28.png p2 |  | string column of the schematic (short / long / dual packer) as printed |
| `components[lc-packer-1].name` | 9-5/8" x 6" QUANTUM Packer | slide 28 image28.png p2 | 9-5/8" x 6" QUANTUM Packer | visual transcription of the picture (x3 upscale), OCR cross-check |
| `components[lc-packer-1].type` | packer | slide 28 image28.png p2 | 9-5/8" x 6" QUANTUM Packer | EXACT TYPE (Component_Types) |
| `components[lc-packer-1].topMd` | 4253.642 | slide 28 image28.png p2 | 1296.51 | 1296.51 m x 3.28084 = 4253.642 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[lc-packer-1].topMdM` | 1296.51 | slide 28 image28.png p2 | 1296.51 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[lc-packer-1].innerDiameter` | 6.0 | slide 28 image28.png p2 | 6.000 | column 'MIN ID' = explicit ID |
| `components[lc-packer-1].outerDiameter` | 8.445 | slide 28 image28.png p2 | 8.445 | column 'MAX OD' = explicit OD |
| `components[lc-tubing-1].name` | 5-1/2" Blank Pipes | slide 28 image28.png p2 | 5-1/2" Blank Pipes | visual transcription of the picture (x3 upscale), OCR cross-check |
| `components[lc-tubing-1].type` | tubing | slide 28 image28.png p2 | 5-1/2" Blank Pipes | NEAREST TYPE (flagged) |
| `components[lc-tubing-1].topMd` | 4287.369 | slide 28 image28.png p2 | 1306.79 | 1306.79 m x 3.28084 = 4287.369 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[lc-tubing-1].topMdM` | 1306.79 | slide 28 image28.png p2 | 1306.79 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[lc-tubing-1].innerDiameter` | 4.778 | slide 28 image28.png p2 | 4.778 | column 'MIN ID' = explicit ID |
| `components[lc-tubing-1].outerDiameter` | 8.4 | slide 28 image28.png p2 | 8.400 | column 'MAX OD' = explicit OD |
| `components[lc-screen-1].name` | 5-1/2" Wire Wrapped Screen + Shunt Tubes | slide 28 image28.png p2 | 5-1/2" Wire Wrapped Screen + Shunt Tubes | visual transcription of the picture (x3 upscale), OCR cross-check |
| `components[lc-screen-1].type` | screen | slide 28 image28.png p2 | 5-1/2" Wire Wrapped Screen + Shunt Tubes | EXACT TYPE (Component_Types) |
| `components[lc-screen-1].topMd` | 4364.305 | slide 28 image28.png p2 | 1330.24 | 1330.24 m x 3.28084 = 4364.305 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[lc-screen-1].topMdM` | 1330.24 | slide 28 image28.png p2 | 1330.24 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[lc-screen-1].innerDiameter` | 4.892 | slide 28 image28.png p2 | 4.892 | column 'MIN ID' = explicit ID |
| `components[lc-screen-1].outerDiameter` | 8.4 | slide 28 image28.png p2 | 8.400 | column 'MAX OD' = explicit OD |
| `components[lc-pbr-1].name` | 9-5/8" x 4" Polished Bore MZ ALLFRAC PACKER | slide 28 image28.png p2 | 9-5/8" x 4" Polished Bore MZ ALLFRAC PACKER | visual transcription of the picture (x3 upscale), OCR cross-check |
| `components[lc-pbr-1].type` | pbr | slide 28 image28.png p2 | 9-5/8" x 4" Polished Bore MZ ALLFRAC PACKER | EXACT TYPE (Component_Types) |
| `components[lc-pbr-1].topMd` | 4441.404 | slide 28 image28.png p2 | 1353.74 | 1353.74 m x 3.28084 = 4441.404 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[lc-pbr-1].topMdM` | 1353.74 | slide 28 image28.png p2 | 1353.74 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[lc-pbr-1].innerDiameter` | 4.0 | slide 28 image28.png p2 | 4.000 | column 'MIN ID' = explicit ID |
| `components[lc-pbr-1].outerDiameter` | 8.305 | slide 28 image28.png p2 | 8.305 | column 'MAX OD' = explicit OD |
| `components[lc-pbr-2].name` | 4" Seal Bore | slide 28 image28.png p2 | 4" Seal Bore | visual transcription of the picture (x3 upscale), OCR cross-check |
| `components[lc-pbr-2].type` | pbr | slide 28 image28.png p2 | 4" Seal Bore | NEAREST TYPE (flagged) |
| `components[lc-pbr-2].topMd` | 4448.721 | slide 28 image28.png p2 | 1355.97 | 1355.97 m x 3.28084 = 4448.721 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[lc-pbr-2].topMdM` | 1355.97 | slide 28 image28.png p2 | 1355.97 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[lc-pbr-2].innerDiameter` | 4.0 | slide 28 image28.png p2 | 4.000 | column 'MIN ID' = explicit ID |
| `components[lc-pbr-2].outerDiameter` | 7.0 | slide 28 image28.png p2 | 7.000 | column 'MAX OD' = explicit OD |
| `components[lc-tubing-2].name` | 5-1/2" Blank Pipes | slide 28 image28.png p2 | 5-1/2" Blank Pipes | visual transcription of the picture (x3 upscale), OCR cross-check |
| `components[lc-tubing-2].type` | tubing | slide 28 image28.png p2 | 5-1/2" Blank Pipes | NEAREST TYPE (flagged) |
| `components[lc-tubing-2].topMd` | 4452.362 | slide 28 image28.png p2 | 1357.08 | 1357.08 m x 3.28084 = 4452.362 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[lc-tubing-2].topMdM` | 1357.08 | slide 28 image28.png p2 | 1357.08 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[lc-tubing-2].innerDiameter` | 4.778 | slide 28 image28.png p2 | 4.778 | column 'MIN ID' = explicit ID |
| `components[lc-tubing-2].outerDiameter` | 8.4 | slide 28 image28.png p2 | 8.400 | column 'MAX OD' = explicit OD |
| `components[lc-screen-2].name` | 5-1/2" Wire Wrapped Screen + Shunt Tubes | slide 28 image28.png p3 | 5-1/2" Wire Wrapped Screen + Shunt Tubes | visual transcription of the picture (x3 upscale), OCR cross-check |
| `components[lc-screen-2].type` | screen | slide 28 image28.png p3 | 5-1/2" Wire Wrapped Screen + Shunt Tubes | EXACT TYPE (Component_Types) |
| `components[lc-screen-2].topMd` | 4721.851 | slide 28 image28.png p3 | 1439.22 | 1439.22 m x 3.28084 = 4721.851 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[lc-screen-2].topMdM` | 1439.22 | slide 28 image28.png p3 | 1439.22 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[lc-screen-2].innerDiameter` | 4.892 | slide 28 image28.png p3 | 4.892 | column 'MIN ID' = explicit ID |
| `components[lc-screen-2].outerDiameter` | 8.4 | slide 28 image28.png p3 | 8.400 | column 'MAX OD' = explicit OD |
| `components[sliding-sleeve-1].name` | Slim line 2-7/8" SSD | slide 28 image28.png p3 | Slim line 2-7/8" SSD | visual transcription of the picture (x3 upscale), OCR cross-check |
| `components[sliding-sleeve-1].type` | sliding-sleeve | slide 28 image28.png p3 | Slim line 2-7/8" SSD | EXACT TYPE (Component_Types) |
| `components[sliding-sleeve-1].topMd` | 4769.16 | slide 28 image28.png p3 | 1453.64 | 1453.64 m x 3.28084 = 4769.16 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[sliding-sleeve-1].topMdM` | 1453.64 | slide 28 image28.png p3 | 1453.64 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[sliding-sleeve-1].innerDiameter` | 2.313 | slide 28 image28.png p3 | 2.313 | column 'MIN ID' = explicit ID |
| `components[sliding-sleeve-1].outerDiameter` | 3.22 | slide 28 image28.png p3 | 3.220 | column 'MAX OD' = explicit OD |
| `components[lc-pbr-3].name` | 9-5/8" x 4" Polished Bore MZ ALLFRAC PACKER | slide 28 image28.png p3 | 9-5/8" x 4" Polished Bore MZ ALLFRAC PACKER | visual transcription of the picture (x3 upscale), OCR cross-check |
| `components[lc-pbr-3].type` | pbr | slide 28 image28.png p3 | 9-5/8" x 4" Polished Bore MZ ALLFRAC PACKER | EXACT TYPE (Component_Types) |
| `components[lc-pbr-3].topMd` | 4837.5 | slide 28 image28.png p3 | 1474.47 | 1474.47 m x 3.28084 = 4837.5 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[lc-pbr-3].topMdM` | 1474.47 | slide 28 image28.png p3 | 1474.47 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[lc-pbr-3].innerDiameter` | 4.0 | slide 28 image28.png p3 | 4.000 | column 'MIN ID' = explicit ID |
| `components[lc-pbr-3].outerDiameter` | 8.305 | slide 28 image28.png p3 | 8.305 | column 'MAX OD' = explicit OD |
| `components[lc-pbr-4].name` | 3.25" Seal Bore | slide 28 image28.png p3 | 3.25" Seal Bore | visual transcription of the picture (x3 upscale), OCR cross-check |
| `components[lc-pbr-4].type` | pbr | slide 28 image28.png p3 | 3.25" Seal Bore | NEAREST TYPE (flagged) |
| `components[lc-pbr-4].topMd` | 4844.849 | slide 28 image28.png p3 | 1476.71 | 1476.71 m x 3.28084 = 4844.849 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[lc-pbr-4].topMdM` | 1476.71 | slide 28 image28.png p3 | 1476.71 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[lc-pbr-4].innerDiameter` | 3.25 | slide 28 image28.png p3 | 3.250 | column 'MIN ID' = explicit ID |
| `components[lc-pbr-4].outerDiameter` | 7.0 | slide 28 image28.png p3 | 7.000 | column 'MAX OD' = explicit OD |
| `components[lc-tubing-3].name` | 5-1/2" Blank Pipes | slide 28 image28.png p3 | 5-1/2" Blank Pipes | visual transcription of the picture (x3 upscale), OCR cross-check |
| `components[lc-tubing-3].type` | tubing | slide 28 image28.png p3 | 5-1/2" Blank Pipes | NEAREST TYPE (flagged) |
| `components[lc-tubing-3].topMd` | 4848.557 | slide 28 image28.png p3 | 1477.84 | 1477.84 m x 3.28084 = 4848.557 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[lc-tubing-3].topMdM` | 1477.84 | slide 28 image28.png p3 | 1477.84 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[lc-tubing-3].innerDiameter` | 4.778 | slide 28 image28.png p3 | 4.778 | column 'MIN ID' = explicit ID |
| `components[lc-tubing-3].outerDiameter` | 8.4 | slide 28 image28.png p3 | 8.400 | column 'MAX OD' = explicit OD |
| `components[lc-screen-3].name` | 5-1/2" Wire Wrapped Screen + Shunt Tubes | slide 28 image28.png p3 | 5-1/2" Wire Wrapped Screen + Shunt Tubes | visual transcription of the picture (x3 upscale), OCR cross-check |
| `components[lc-screen-3].type` | screen | slide 28 image28.png p3 | 5-1/2" Wire Wrapped Screen + Shunt Tubes | EXACT TYPE (Component_Types) |
| `components[lc-screen-3].topMd` | 4887.139 | slide 28 image28.png p3 | 1489.60 | 1489.6 m x 3.28084 = 4887.139 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[lc-screen-3].topMdM` | 1489.6 | slide 28 image28.png p3 | 1489.60 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[lc-screen-3].innerDiameter` | 4.892 | slide 28 image28.png p3 | 4.892 | column 'MIN ID' = explicit ID |
| `components[lc-screen-3].outerDiameter` | 8.4 | slide 28 image28.png p3 | 8.400 | column 'MAX OD' = explicit OD |
| `components[nipple-no-go-1].name` | 2-7/8" XN-Nipple | slide 28 image28.png p3 | 2-7/8" XN-Nipple | visual transcription of the picture (x3 upscale), OCR cross-check |
| `components[nipple-no-go-1].type` | nipple (no-go) | slide 28 image28.png p3 | 2-7/8" XN-Nipple | EXACT TYPE (Component_Types) |
| `components[nipple-no-go-1].topMd` | 4864.272 | slide 28 image28.png p3 | 1482.63 | 1482.63 m x 3.28084 = 4864.272 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[nipple-no-go-1].topMdM` | 1482.63 | slide 28 image28.png p3 | 1482.63 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[nipple-no-go-1].innerDiameter` | 2.205 | slide 28 image28.png p3 | 2.205 | column 'MIN ID' = explicit ID |
| `components[nipple-no-go-1].outerDiameter` | 2.313 | slide 28 image28.png p3 | 2.313 | column 'MAX OD' = explicit OD |
| `components[mule-shoe-1].name` | Self Aligning Muleshoe | slide 28 image28.png | Self Aligning Muleshoe | visual transcription of the picture (x3 upscale), OCR cross-check |
| `components[mule-shoe-1].type` | mule-shoe | slide 28 image28.png | Self Aligning Muleshoe | EXACT TYPE (Component_Types) |
| `components[mule-shoe-1].topMd` | 4872.605 | slide 28 image28.png | 1485.17 | 1485.17 m x 3.28084 = 4872.605 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[mule-shoe-1].topMdM` | 1485.17 | slide 28 image28.png | 1485.17 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[mule-shoe-1].innerDiameter` | 2.36 | slide 28 image28.png | 2.360 | column 'MIN ID' = explicit ID |
| `components[mule-shoe-1].outerDiameter` | 3.22 | slide 28 image28.png | 3.220 | column 'MAX OD' = explicit OD |
| `components[lc-packer-2].name` | 9-5/8" x 6" Sump Packer | slide 28 image28.png p3 | 9-5/8" x 6" Sump Packer | visual transcription of the picture (x3 upscale), OCR cross-check |
| `components[lc-packer-2].type` | packer | slide 28 image28.png p3 | 9-5/8" x 6" Sump Packer | EXACT TYPE (Component_Types) |
| `components[lc-packer-2].topMd` | 4969.357 | slide 28 image28.png p3 | 1514.66 | 1514.66 m x 3.28084 = 4969.357 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[lc-packer-2].topMdM` | 1514.66 | slide 28 image28.png p3 | 1514.66 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[lc-packer-2].innerDiameter` | 6.0 | slide 28 image28.png p3 | 6.000 | column 'MIN ID' = explicit ID |
| `components[lc-packer-2].outerDiameter` | 8.463 | slide 28 image28.png p3 | 8.463 | column 'MAX OD' = explicit OD |
| `components[lc-bridge-plug-1].name` | PBTD / BRIDGE PLUG | slide 28 image28.png p3 | PBTD / BRIDGE PLUG | visual transcription of the picture (x3 upscale), OCR cross-check |
| `components[lc-bridge-plug-1].type` | bridge-plug | slide 28 image28.png p3 | PBTD / BRIDGE PLUG | EXACT TYPE (Component_Types) |
| `components[lc-bridge-plug-1].topMd` | 5085.302 | slide 28 image28.png p3 | 1550.00 | 1550.0 m x 3.28084 = 5085.302 ft (3 dp); m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[lc-bridge-plug-1].topMdM` | 1550.0 | slide 28 image28.png p3 | 1550.00 | original metres; m MDDF (column headers 'DEPTH m MDDF') (equipment depth column as printed) |
| `components[production-casing-1].name` | PRODUCTION CASING 9-5/8" 47 ppf L-80 CS VAM TOP | slide 28 image28.png header | PRODUCTION CASING 9-5/8" 47 ppf L-80 CS VAM TOP | visual transcription of the picture (x3 upscale), OCR cross-check |
| `components[production-casing-1].type` | casing | slide 28 image28.png header | PRODUCTION CASING 9-5/8" 47 ppf L-80 CS VAM TOP | EXACT TYPE (Component_Types) |
| `components[production-casing-1].innerDiameter` | 9.625 | slide 28 image28.png header | 9.625 | USER-DEFINED RULE: header size without ID/OD designation -> innerDiameter |
| `components[production-casing-1].weight` | 47.0 | slide 28 image28.png header | PRODUCTION CASING 9-5/8" 47 ppf L-80 CS VAM TOP | parsed from the label (ppf = lb/ft) |
| `components[production-casing-1].grade` | L-80 | slide 28 image28.png header | PRODUCTION CASING 9-5/8" 47 ppf L-80 CS VAM TOP | parsed from the label |
| `components[production-casing-1].bottomMd` | 5472.441 | slide 28 image28.png diagram | 9-5/8" CASING SHOE 1668.00 | 1668.0 m x 3.28084; diagram value preferred over header (None m) |
| `components[production-casing-1].bottomMdM` | 1668.0 | slide 28 image28.png diagram | 9-5/8" CASING SHOE 1668.00 | original metres; MDDF |
| `components[s-tubing-1].name` | 3-1/2" L80 13Cr 9.2 PPF JFE-FOX (Upper Completion) | slide 28 image28.png header | 3-1/2" L80 13Cr 9.2 PPF JFE-FOX (Upper Completion) | visual transcription of the picture (x3 upscale), OCR cross-check |
| `components[s-tubing-1].type` | tubing | slide 28 image28.png header | 3-1/2" L80 13Cr 9.2 PPF JFE-FOX (Upper Completion) | EXACT TYPE (Component_Types) |
| `components[s-tubing-1].innerDiameter` | 3.5 | slide 28 image28.png header | 3-1/2" L80 13Cr 9.2 PPF JFE-FOX (Upper Completion) | USER-DEFINED RULE: header size without ID/OD designation -> innerDiameter |
| `components[s-tubing-1].weight` | 9.2 | slide 28 image28.png header | 3-1/2" L80 13Cr 9.2 PPF JFE-FOX (Upper Completion) | parsed from the label (ppf = lb/ft) |
| `components[s-tubing-1].grade` | L80 | slide 28 image28.png header | 3-1/2" L80 13Cr 9.2 PPF JFE-FOX (Upper Completion) | parsed from the label |
| `components[s-tubing-1].string` | Short | slide 28 image28.png header |  | string column of the schematic (short / long / dual packer) as printed |
| `components[l-tubing-1].name` | 3-1/2" L80 13Cr 9.2 PPF JFE-FOX (Upper Completion) | slide 28 image28.png header | 3-1/2" L80 13Cr 9.2 PPF JFE-FOX (Upper Completion) | visual transcription of the picture (x3 upscale), OCR cross-check |
| `components[l-tubing-1].type` | tubing | slide 28 image28.png header | 3-1/2" L80 13Cr 9.2 PPF JFE-FOX (Upper Completion) | EXACT TYPE (Component_Types) |
| `components[l-tubing-1].innerDiameter` | 3.5 | slide 28 image28.png header | 3-1/2" L80 13Cr 9.2 PPF JFE-FOX (Upper Completion) | USER-DEFINED RULE: header size without ID/OD designation -> innerDiameter |
| `components[l-tubing-1].weight` | 9.2 | slide 28 image28.png header | 3-1/2" L80 13Cr 9.2 PPF JFE-FOX (Upper Completion) | parsed from the label (ppf = lb/ft) |
| `components[l-tubing-1].grade` | L80 | slide 28 image28.png header | 3-1/2" L80 13Cr 9.2 PPF JFE-FOX (Upper Completion) | parsed from the label |
| `components[l-tubing-1].string` | Long | slide 28 image28.png header |  | string column of the schematic (short / long / dual packer) as printed |
| `components[inner-string-tubing-1].name` | 2-7/8" SC L80 13Cr 6.4 PPF JFE-FOX (Inner String) | slide 28 image28.png header | 2-7/8" SC L80 13Cr 6.4 PPF JFE-FOX (Inner String) | visual transcription of the picture (x3 upscale), OCR cross-check |
| `components[inner-string-tubing-1].type` | tubing | slide 28 image28.png header | 2-7/8" SC L80 13Cr 6.4 PPF JFE-FOX (Inner String) | EXACT TYPE (Component_Types) |
| `components[inner-string-tubing-1].innerDiameter` | 2.875 | slide 28 image28.png header | 2-7/8" SC L80 13Cr 6.4 PPF JFE-FOX (Inner String) | USER-DEFINED RULE: header size without ID/OD designation -> innerDiameter |
| `components[inner-string-tubing-1].weight` | 6.4 | slide 28 image28.png header | 2-7/8" SC L80 13Cr 6.4 PPF JFE-FOX (Inner String) | parsed from the label (ppf = lb/ft) |
| `components[inner-string-tubing-1].grade` | L80 | slide 28 image28.png header | 2-7/8" SC L80 13Cr 6.4 PPF JFE-FOX (Inner String) | parsed from the label |
| `components[zone-e-10-11-1].name` | E 10-11 1336-1345 m MDDF | slide 28 image28.png p2 | E 10-11 1336-1345 m MDDF | visual transcription of the picture (x3 upscale), OCR cross-check |
| `components[zone-e-10-11-1].type` | perforation | slide 28 image28.png p2 | E 10-11 1336-1345 m MDDF | NEAREST TYPE (flagged) |
| `components[zone-e-10-11-1].topMd` | 4383.202 | slide 28 image28.png p2 | E 10-11 1336 -1345 m MDDF 1219 - 1227 m TVDDF 1226 psi 196 deg F | 1336.0 m x 3.28084 = 4383.202 ft (3 dp); m MDDF as printed |
| `components[zone-e-10-11-1].topMdM` | 1336.0 | slide 28 image28.png p2 | E 10-11 1336 -1345 m MDDF 1219 - 1227 m TVDDF 1226 psi 196 deg F | original metres; m MDDF as printed |
| `components[zone-e-10-11-1].bottomMd` | 4412.73 | slide 28 image28.png p2 | E 10-11 1336 -1345 m MDDF 1219 - 1227 m TVDDF 1226 psi 196 deg F | 1345.0 m x 3.28084 = 4412.73 ft (3 dp); m MDDF as printed |
| `components[zone-e-10-11-1].bottomMdM` | 1345.0 | slide 28 image28.png p2 | E 10-11 1336 -1345 m MDDF 1219 - 1227 m TVDDF 1226 psi 196 deg F | original metres; m MDDF as printed |
| `components[zone-e-10-11-1].zone` | E 10-11 | slide 28 image28.png p2 | E 10-11 | zone label as printed |
| `components[zone-e-32-34-1].name` | E 32-34 1445-1449 m MDDF | slide 28 image28.png p3 | E 32-34 1445-1449 m MDDF | visual transcription of the picture (x3 upscale), OCR cross-check |
| `components[zone-e-32-34-1].type` | perforation | slide 28 image28.png p3 | E 32-34 1445-1449 m MDDF | NEAREST TYPE (flagged) |
| `components[zone-e-32-34-1].topMd` | 4740.814 | slide 28 image28.png p3 | E 32-34 1445 - 1449 m MDDF 1316 - 1319.5 m TVDDF | 1445.0 m x 3.28084 = 4740.814 ft (3 dp); m MDDF as printed |
| `components[zone-e-32-34-1].topMdM` | 1445.0 | slide 28 image28.png p3 | E 32-34 1445 - 1449 m MDDF 1316 - 1319.5 m TVDDF | original metres; m MDDF as printed |
| `components[zone-e-32-34-1].bottomMd` | 4753.937 | slide 28 image28.png p3 | E 32-34 1445 - 1449 m MDDF 1316 - 1319.5 m TVDDF | 1449.0 m x 3.28084 = 4753.937 ft (3 dp); m MDDF as printed |
| `components[zone-e-32-34-1].bottomMdM` | 1449.0 | slide 28 image28.png p3 | E 32-34 1445 - 1449 m MDDF 1316 - 1319.5 m TVDDF | original metres; m MDDF as printed |
| `components[zone-e-32-34-1].zone` | E 32-34 | slide 28 image28.png p3 | E 32-34 | zone label as printed |
| `components[zone-e-32-34-2].name` | E 32-34 1465-1472 m MDDF | slide 28 image28.png p3 | E 32-34 1465-1472 m MDDF | visual transcription of the picture (x3 upscale), OCR cross-check |
| `components[zone-e-32-34-2].type` | perforation | slide 28 image28.png p3 | E 32-34 1465-1472 m MDDF | NEAREST TYPE (flagged) |
| `components[zone-e-32-34-2].topMd` | 4806.431 | slide 28 image28.png p3 | 1465 - 1472 m MDDF 1334 - 1340 m TVDDF 1573 psi 210 deg F | 1465.0 m x 3.28084 = 4806.431 ft (3 dp); m MDDF as printed |
| `components[zone-e-32-34-2].topMdM` | 1465.0 | slide 28 image28.png p3 | 1465 - 1472 m MDDF 1334 - 1340 m TVDDF 1573 psi 210 deg F | original metres; m MDDF as printed |
| `components[zone-e-32-34-2].bottomMd` | 4829.396 | slide 28 image28.png p3 | 1465 - 1472 m MDDF 1334 - 1340 m TVDDF 1573 psi 210 deg F | 1472.0 m x 3.28084 = 4829.396 ft (3 dp); m MDDF as printed |
| `components[zone-e-32-34-2].bottomMdM` | 1472.0 | slide 28 image28.png p3 | 1465 - 1472 m MDDF 1334 - 1340 m TVDDF 1573 psi 210 deg F | original metres; m MDDF as printed |
| `components[zone-e-32-34-2].zone` | E 32-34 | slide 28 image28.png p3 | E 32-34 | zone label as printed |
| `components[zone-e-40-1].name` | E 40 1499-1503 m MDDF | slide 28 image28.png p3 | E 40 1499-1503 m MDDF | visual transcription of the picture (x3 upscale), OCR cross-check |
| `components[zone-e-40-1].type` | perforation | slide 28 image28.png p3 | E 40 1499-1503 m MDDF | NEAREST TYPE (flagged) |
| `components[zone-e-40-1].topMd` | 4917.979 | slide 28 image28.png p3 | E 40 1499 - 1503 m MDDF 1364.5 - 1368 m TVDDF 1544 psi 212 deg F | 1499.0 m x 3.28084 = 4917.979 ft (3 dp); m MDDF as printed |
| `components[zone-e-40-1].topMdM` | 1499.0 | slide 28 image28.png p3 | E 40 1499 - 1503 m MDDF 1364.5 - 1368 m TVDDF 1544 psi 212 deg F | original metres; m MDDF as printed |
| `components[zone-e-40-1].bottomMd` | 4931.103 | slide 28 image28.png p3 | E 40 1499 - 1503 m MDDF 1364.5 - 1368 m TVDDF 1544 psi 212 deg F | 1503.0 m x 3.28084 = 4931.103 ft (3 dp); m MDDF as printed |
| `components[zone-e-40-1].bottomMdM` | 1503.0 | slide 28 image28.png p3 | E 40 1499 - 1503 m MDDF 1364.5 - 1368 m TVDDF 1544 psi 212 deg F | original metres; m MDDF as printed |
| `components[zone-e-40-1].zone` | E 40 | slide 28 image28.png p3 | E 40 | zone label as printed |
| `components[pbtd-1].name` | PBTD / BRIDGE PLUG 1550.00 (header PBTD : 1570 m MDDF) | slide 28 image28.png diagram | PBTD / BRIDGE PLUG 1550.00 (header PBTD : 1570 m MDDF) | visual transcription of the picture (x3 upscale), OCR cross-check |
| `components[pbtd-1].type` | pbtd | slide 28 image28.png diagram | PBTD / BRIDGE PLUG 1550.00 (header PBTD : 1570 m MDDF) | EXACT TYPE (Component_Types) |
| `components[pbtd-1].topMd` | 5085.302 | slide 28 image28.png diagram | PBTD / BRIDGE PLUG 1550.00 (header PBTD : 1570 m MDDF) | 1550.0 m x 3.28084 = 5085.302 ft (3 dp); m MDDF |
| `components[pbtd-1].topMdM` | 1550.0 | slide 28 image28.png diagram | PBTD / BRIDGE PLUG 1550.00 (header PBTD : 1570 m MDDF) | original metres; m MDDF |
| `components[int-mechanical-plug-1].name` | 2.31” PXN Plug @ 1470m MDTHF (Jul 2017) | slide 28 obj #6 | 2.31” PXN Plug @ 1470m MDTHF (Jul 2017) | visual transcription of the picture (x3 upscale), OCR cross-check |
| `components[int-mechanical-plug-1].type` | mechanical-plug | slide 28 obj #6 | 2.31” PXN Plug @ 1470m MDTHF (Jul 2017) | NEAREST TYPE (flagged) |
| `components[int-mechanical-plug-2].name` | 2.2” Slickplug @ 1455m MDTHF (Dec 2017) | slide 28 obj #7 | 2.2” Slickplug @ 1455m MDTHF (Dec 2017) | visual transcription of the picture (x3 upscale), OCR cross-check |
| `components[int-mechanical-plug-2].type` | mechanical-plug | slide 28 obj #7 | 2.2” Slickplug @ 1455m MDTHF (Dec 2017) | NEAREST TYPE (flagged) |
| `wellInfo.completionDate` | TBA | slide 28 image28.png header | TBA | kept as printed (not parseable) |
| `wellInfo.completionFluid` | 8.5 ppg Drill water + 3% KCL based fluid | slide 28 image28.png header | 8.5 ppg Drill water + 3% KCL based fluid | as printed |
| `wellInfo.completionType` | PROPOSED COMPLETION SCHEMATIC — SHORT STRING (OIL PRODUCER) / LONG STRING (OIL PRODUCER), Y-block dual completion | slide 28 image28.png header | PROPOSED COMPLETION SCHEMATIC — SHORT STRING (OIL PRODUCER) / LONG STRING (OIL PRODUCER), Y-block dual completion | as printed |
| `wellInfo.groundElevation` | None |  |  | not supported by the source -> template empty value ('-' text / null number) |
| `wellInfo.kbElevation` | None |  |  | not supported by the source -> template empty value ('-' text / null number) |
| `wellInfo.maxDeviation` | 36.06 DEG @ 504.41 m MDDF / 462.19 m TVDDF | slide 28 image28.png header | 36.06 DEG @ 504.41 m MDDF / 462.19 m TVDDF | as printed |
| `wellInfo.name` | Dulang_B28ST2 | slide 28 | WELL NAME: DULANG B-28ST2 (LOCATION : DULANG-B) | identifier rule Dulang_B<no><suffix>; original label kept in the summary |
| `wellInfo.packerFluid` | 8.5 ppg Drill water + 3% KCL based fluid + 0.1% vol. Bacteride + 2% vol. CONQOR 303A Corrosion Inhibitor + 1% vol. O2 scavenger | slide 28 image28.png header | 8.5 ppg Drill water + 3% KCL based fluid + 0.1% vol. Bacteride + 2% vol. CONQOR 303A Corrosion Inhibitor + 1% vol. O2 scavenger | as printed |
| `wellInfo.rig` | KM-1 | slide 28 image28.png header | KM-1 | as printed |
| `wellInfo.surfaceCasing` | - |  |  | not supported by the source -> template empty value ('-' text / null number) |
| `wellInfo.totalDepth` | None |  |  | not supported by the source -> template empty value ('-' text / null number) |
| `wellInfo.trajectoryType` | - |  |  | not supported by the source -> template empty value ('-' text / null number) |
| `wellInfo.tubing` | Dual String 3-1/2" L80 13Cr 9.2 PPF JFE-FOX (Upper Completion); 4" L80 13Cr 9.5 PPF VAM FJL (Outer String); 2-7/8" SC L80 13Cr 6.4 PPF JFE-FOX (Inner String); 9-5/8" Cased Hole Gravel Pack wt. 5-1/2" Wire Wrap Screen (Lower Completion) | slide 28 image28.png header | Dual String 3-1/2" L80 13Cr 9.2 PPF JFE-FOX (Upper Completion); 4" L80 13Cr 9.5 PPF VAM FJL (Outer String); 2-7/8" SC L80 13Cr 6.4 PPF JFE-FOX (Inner String); 9 | as printed |
| `wellInfo.wellhead` | - |  |  | not supported by the source -> template empty value ('-' text / null number) |

## Dulang_B30ST1 — slide 30

Original label: WELL NAME: B30 ST-1 (LOCATION / SLOT: DULANG B PLATFORM /SLOT 26)
Equipment depth datum: m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment')  ·  source: image30.png (visual transcription of a low-resolution picture (632x842 px, x3 upscale), OCR cross-check — many digits uncertain)

### Conflicts / missing / unclear

| kind | where | detail |
|---|---|---|
| Unclear | slide 30 image30.png header/diagram | TD 2086.0 m printed in MDDF; equipment datum m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') — no datum shift applied |
| Unclear | picture | very low-resolution picture (632x842 px): almost all depths, several IDs/ODs and the header values are uncertain reads (marked '?', flagged estimated); the Excel original in ECMS (named on the slide) should be used to confirm — reported as partly unreadable content |
| Unclear | picture | PROPOSED schematic (Rev 2) — treated as the well's schematic; the slide says the latest editable diagram is in ECMS (not part of this input) |

### Unmapped content (kept here, not in the JSON)

| where | text | reason |
|---|---|---|
| picture | E20 SAND (BCO) 'Future Perf Interval: 1846 m-MDDF - 1850 m-MDDF (4 meters)' — planned, not mapped as a zone; sand depths / BHP / BHT per zone (illegible in part); 'STMZ Gross Length: 216.0 meter'; Pip Tag #1/#2 and WAB labels (≈1576 / 1688 / 1870 m-MDDF, illegible); left table also lists 13-3/8" CSG | note (status / history / survey / other revision) |
| slide 30 | 'FOR B30 LATEST EDITABLE DIAGRAM, PLEASE REFER TO EXCEL FILE IN ECMS - PRAVIN NAIR 2023-' | note (status / history / survey / other revision) |
| slide 30 obj #23 | FOR B30 LATEST EDITABLE DIAGRAM, PLEASE REFER TO EXCEL FILE IN ECMS - PRAVIN NAIR 2023- | slide text box over/beside the picture (status / history / note) |

### Traceability (every populated value)

| JSON path | value | slide / object / location | original text | conversion / interpretation |
|---|---|---|---|---|
| `components[christmas-tree-1].name` | Solar Alert 3-1/8" 5K U/HH Trim | slide 30 image30.png header | Solar Alert 3-1/8" 5K U/HH Trim | visual transcription of a low-resolution picture (632x842 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[christmas-tree-1].type` | christmas-tree | slide 30 image30.png header | Solar Alert 3-1/8" 5K U/HH Trim | EXACT TYPE (Component_Types) |
| `components[tubing-hanger-1].name` | TUBING HANGER, 13-5/8" x 5"-3SA-2G LH x 3-1/2" 9.2# JFEBEAR, w/ 3" HBPV PROFILE | slide 30 image30.png right table | TUBING HANGER, 13-5/8" x 5"-3SA-2G LH x 3-1/2" 9.2# JFEBEAR, w/ 3" HBPV PROFILE | visual transcription of a low-resolution picture (632x842 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[tubing-hanger-1].type` | tubing-hanger | slide 30 image30.png right table | TUBING HANGER, 13-5/8" x 5"-3SA-2G LH x 3-1/2" 9.2# JFEBEAR, w/ 3" HBPV PROFILE | EXACT TYPE (Component_Types) |
| `components[tubing-hanger-1].topMd` | 68.898 | slide 30 image30.png right table | 21 | 21.0 m x 3.28084 = 68.898 ft (3 dp); m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[tubing-hanger-1].topMdM` | 21.0 | slide 30 image30.png right table | 21 | original metres; m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[tubing-hanger-1].innerDiameter` | 2.95 | slide 30 image30.png right table | 2.950 | column 'MIN ID' = explicit ID |
| `components[tubing-hanger-1].outerDiameter` | 13.47 | slide 30 image30.png right table | 13.470 | column 'MAX OD' = explicit OD |
| `components[flow-coupling-1].name` | FLOW COUPLING, 3-1/2", S13CR, JFEBEAR (SLB) | slide 30 image30.png right table | FLOW COUPLING, 3-1/2", S13CR, JFEBEAR (SLB) | visual transcription of a low-resolution picture (632x842 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[flow-coupling-1].type` | flow-coupling | slide 30 image30.png right table | FLOW COUPLING, 3-1/2", S13CR, JFEBEAR (SLB) | EXACT TYPE (Component_Types) |
| `components[flow-coupling-1].topMd` | 820.21 | slide 30 image30.png right table | 250 | 250.0 m x 3.28084 = 820.21 ft (3 dp); m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[flow-coupling-1].topMdM` | 250.0 | slide 30 image30.png right table | 250 | original metres; m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[flow-coupling-1].innerDiameter` | 2.992 | slide 30 image30.png right table | 2.992 | column 'MIN ID' = explicit ID |
| `components[flow-coupling-1].outerDiameter` | 3.885 | slide 30 image30.png right table | 3.885 | column 'MAX OD' = explicit OD |
| `components[scssv-1].name` | SCSSV, SELF-EQ, 3-1/2", S13CR, JFEBEAR (SLB) | slide 30 image30.png right table | SCSSV, SELF-EQ, 3-1/2", S13CR, JFEBEAR (SLB) | visual transcription of a low-resolution picture (632x842 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[scssv-1].type` | scssv | slide 30 image30.png right table | SCSSV, SELF-EQ, 3-1/2", S13CR, JFEBEAR (SLB) | EXACT TYPE (Component_Types) |
| `components[scssv-1].innerDiameter` | 2.813 | slide 30 image30.png right table | 2.813 | column 'MIN ID' = explicit ID |
| `components[scssv-1].outerDiameter` | 5.0 | slide 30 image30.png right table | 5.000 | column 'MAX OD' = explicit OD |
| `components[flow-coupling-2].name` | FLOW COUPLING, 3-1/2", S13CR, JFEBEAR (SLB) | slide 30 image30.png right table | FLOW COUPLING, 3-1/2", S13CR, JFEBEAR (SLB) | visual transcription of a low-resolution picture (632x842 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[flow-coupling-2].type` | flow-coupling | slide 30 image30.png right table | FLOW COUPLING, 3-1/2", S13CR, JFEBEAR (SLB) | EXACT TYPE (Component_Types) |
| `components[flow-coupling-2].innerDiameter` | 2.992 | slide 30 image30.png right table | 2.992 | column 'MIN ID' = explicit ID |
| `components[flow-coupling-2].outerDiameter` | 3.885 | slide 30 image30.png right table | 3.885 | column 'MAX OD' = explicit OD |
| `components[spm-1].name` | SPM, 3-1/2" 13CR JFEBEAR (SLB) | slide 30 image30.png right table | SPM, 3-1/2" 13CR JFEBEAR (SLB) | visual transcription of a low-resolution picture (632x842 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[spm-1].type` | spm | slide 30 image30.png right table | SPM, 3-1/2" 13CR JFEBEAR (SLB) | EXACT TYPE (Component_Types) |
| `components[spm-1].topMd` | 2503.281 | slide 30 image30.png right table | 763 | 763.0 m x 3.28084 = 2503.281 ft (3 dp); m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[spm-1].topMdM` | 763.0 | slide 30 image30.png right table | 763 | original metres; m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[spm-1].innerDiameter` | 2.9 | slide 30 image30.png right table | 2.900 | column 'MIN ID' = explicit ID |
| `components[spm-1].outerDiameter` | 5.23 | slide 30 image30.png right table | 5.230 | column 'MAX OD' = explicit OD |
| `components[spm-2].name` | SPM, 3-1/2" 13CR JFEBEAR (SLB) | slide 30 image30.png right table | SPM, 3-1/2" 13CR JFEBEAR (SLB) | visual transcription of a low-resolution picture (632x842 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[spm-2].type` | spm | slide 30 image30.png right table | SPM, 3-1/2" 13CR JFEBEAR (SLB) | EXACT TYPE (Component_Types) |
| `components[spm-2].topMd` | 3595.801 | slide 30 image30.png right table | 1096 | 1096.0 m x 3.28084 = 3595.801 ft (3 dp); m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[spm-2].topMdM` | 1096.0 | slide 30 image30.png right table | 1096 | original metres; m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[spm-2].innerDiameter` | 2.9 | slide 30 image30.png right table | 2.900 | column 'MIN ID' = explicit ID |
| `components[spm-2].outerDiameter` | 5.23 | slide 30 image30.png right table | 5.230 | column 'MAX OD' = explicit OD |
| `components[spm-3].name` | SPM, 3-1/2" 13CR JFEBEAR (SLB) | slide 30 image30.png right table | SPM, 3-1/2" 13CR JFEBEAR (SLB) | visual transcription of a low-resolution picture (632x842 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[spm-3].type` | spm | slide 30 image30.png right table | SPM, 3-1/2" 13CR JFEBEAR (SLB) | EXACT TYPE (Component_Types) |
| `components[spm-3].topMd` | 4271.654 | slide 30 image30.png right table | 1302 | 1302.0 m x 3.28084 = 4271.654 ft (3 dp); m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[spm-3].topMdM` | 1302.0 | slide 30 image30.png right table | 1302 | original metres; m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[spm-3].innerDiameter` | 2.9 | slide 30 image30.png right table | 2.900 | column 'MIN ID' = explicit ID |
| `components[spm-3].outerDiameter` | 5.23 | slide 30 image30.png right table | 5.230 | column 'MAX OD' = explicit OD |
| `components[spm-4].name` | SPM, 3-1/2" 13CR JFEBEAR (SLB) | slide 30 image30.png right table | SPM, 3-1/2" 13CR JFEBEAR (SLB) | visual transcription of a low-resolution picture (632x842 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[spm-4].type` | spm | slide 30 image30.png right table | SPM, 3-1/2" 13CR JFEBEAR (SLB) | EXACT TYPE (Component_Types) |
| `components[spm-4].topMd` | 4960.63 | slide 30 image30.png right table | 1512 | 1512.0 m x 3.28084 = 4960.63 ft (3 dp); m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[spm-4].topMdM` | 1512.0 | slide 30 image30.png right table | 1512 | original metres; m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[spm-4].innerDiameter` | 2.9 | slide 30 image30.png right table | 2.900 | column 'MIN ID' = explicit ID |
| `components[spm-4].outerDiameter` | 5.23 | slide 30 image30.png right table | 5.230 | column 'MAX OD' = explicit OD |
| `components[sliding-sleeve-1].name` | SLIMLINE SSD, 3-1/2" 13CR JFEBEAR (SLB) | slide 30 image30.png right table | SLIMLINE SSD, 3-1/2" 13CR JFEBEAR (SLB) | visual transcription of a low-resolution picture (632x842 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[sliding-sleeve-1].type` | sliding-sleeve | slide 30 image30.png right table | SLIMLINE SSD, 3-1/2" 13CR JFEBEAR (SLB) | EXACT TYPE (Component_Types) |
| `components[sliding-sleeve-1].topMd` | 5013.124 | slide 30 image30.png right table | 1528 | 1528.0 m x 3.28084 = 5013.124 ft (3 dp); m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[sliding-sleeve-1].topMdM` | 1528.0 | slide 30 image30.png right table | 1528 | original metres; m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[sliding-sleeve-1].innerDiameter` | 2.813 | slide 30 image30.png right table | 2.813 | column 'MIN ID' = explicit ID |
| `components[sliding-sleeve-1].outerDiameter` | 4.2 | slide 30 image30.png right table | 4.200 | column 'MAX OD' = explicit OD |
| `components[packer-1].name` | 9-5/8" X 3-1/2" FEED THRU PACKER, 25CR JFE BEAR, V0 RATED HYDRAULIC SET, CUT-TO-RELEASE (SLB) | slide 30 image30.png right table | 9-5/8" X 3-1/2" FEED THRU PACKER, 25CR JFE BEAR, V0 RATED HYDRAULIC SET, CUT-TO-RELEASE (SLB) | visual transcription of a low-resolution picture (632x842 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[packer-1].type` | packer | slide 30 image30.png right table | 9-5/8" X 3-1/2" FEED THRU PACKER, 25CR JFE BEAR, V0 RATED HYDRAULIC SET, CUT-TO-RELEASE (SLB) | EXACT TYPE (Component_Types) |
| `components[packer-1].topMd` | 5052.494 | slide 30 image30.png right table | 1540 | 1540.0 m x 3.28084 = 5052.494 ft (3 dp); m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[packer-1].topMdM` | 1540.0 | slide 30 image30.png right table | 1540 | original metres; m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[packer-1].innerDiameter` | 2.93 | slide 30 image30.png right table | 2.930 | column 'MIN ID' = explicit ID |
| `components[packer-1].outerDiameter` | 8.66 | slide 30 image30.png right table | 8.660 | column 'MAX OD' = explicit OD |
| `components[landing-nipple-1].name` | 2.813" X- LANDING NIPPLE (SLB) | slide 30 image30.png right table | 2.813" X- LANDING NIPPLE (SLB) | visual transcription of a low-resolution picture (632x842 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[landing-nipple-1].type` | landing-nipple | slide 30 image30.png right table | 2.813" X- LANDING NIPPLE (SLB) | NEAREST TYPE (flagged) |
| `components[landing-nipple-1].topMd` | 5062.336 | slide 30 image30.png right table | 1543 | 1543.0 m x 3.28084 = 5062.336 ft (3 dp); m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[landing-nipple-1].topMdM` | 1543.0 | slide 30 image30.png right table | 1543 | original metres; m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[landing-nipple-1].innerDiameter` | 2.813 | slide 30 image30.png right table | 2.813 | column 'MIN ID' = explicit ID |
| `components[landing-nipple-1].outerDiameter` | 4.04 | slide 30 image30.png right table | 4.040 | column 'MAX OD' = explicit OD |
| `components[pup-joint-1].name` | PDG, DUAL SENSOR (SLB) | slide 30 image30.png right table | PDG, DUAL SENSOR (SLB) | visual transcription of a low-resolution picture (632x842 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[pup-joint-1].type` | pup-joint | slide 30 image30.png right table | PDG, DUAL SENSOR (SLB) | UNRESOLVED / NEAREST TYPE (flagged) |
| `components[pup-joint-1].topMd` | 5085.302 | slide 30 image30.png right table | 1550 | 1550.0 m x 3.28084 = 5085.302 ft (3 dp); m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[pup-joint-1].topMdM` | 1550.0 | slide 30 image30.png right table | 1550 | original metres; m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[pup-joint-1].innerDiameter` | 2.915 | slide 30 image30.png right table | 2.915 | column 'MIN ID' = explicit ID |
| `components[pup-joint-1].outerDiameter` | 5.578 | slide 30 image30.png right table | 5.578 | column 'MAX OD' = explicit OD |
| `components[circulating-valve-1].name` | SELF-FILLING SUB, 3-1/2", 13CR JFE BEAR (SLB) | slide 30 image30.png right table | SELF-FILLING SUB, 3-1/2", 13CR JFE BEAR (SLB) | visual transcription of a low-resolution picture (632x842 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[circulating-valve-1].type` | circulating-valve | slide 30 image30.png right table | SELF-FILLING SUB, 3-1/2", 13CR JFE BEAR (SLB) | NEAREST TYPE (flagged) |
| `components[circulating-valve-1].topMd` | 5118.11 | slide 30 image30.png right table | 1560 | 1560.0 m x 3.28084 = 5118.11 ft (3 dp); m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[circulating-valve-1].topMdM` | 1560.0 | slide 30 image30.png right table | 1560 | original metres; m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[circulating-valve-1].innerDiameter` | 2.88 | slide 30 image30.png right table | 2.880 | column 'MIN ID' = explicit ID |
| `components[circulating-valve-1].outerDiameter` | 4.93 | slide 30 image30.png right table | 4.930 | column 'MAX OD' = explicit OD |
| `components[mechanical-plug-1].name` | i-MAGNUM DISC, 3-1/2", 13CR JFE BEAR (SLB) | slide 30 image30.png right table | i-MAGNUM DISC, 3-1/2", 13CR JFE BEAR (SLB) | visual transcription of a low-resolution picture (632x842 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[mechanical-plug-1].type` | mechanical-plug | slide 30 image30.png right table | i-MAGNUM DISC, 3-1/2", 13CR JFE BEAR (SLB) | NEAREST TYPE (flagged) |
| `components[mechanical-plug-1].topMd` | 5137.795 | slide 30 image30.png right table | 1566 | 1566.0 m x 3.28084 = 5137.795 ft (3 dp); m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[mechanical-plug-1].topMdM` | 1566.0 | slide 30 image30.png right table | 1566 | original metres; m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[mechanical-plug-1].innerDiameter` | 2.92 | slide 30 image30.png right table | 2.920 | column 'MIN ID' = explicit ID |
| `components[mechanical-plug-1].outerDiameter` | 5.0 | slide 30 image30.png right table | 5.000 | column 'MAX OD' = explicit OD |
| `components[st-slot-locator-1].name` | NO GO LOCATOR (HLB) | slide 30 image30.png right table | NO GO LOCATOR (HLB) | visual transcription of a low-resolution picture (632x842 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[st-slot-locator-1].type` | st-slot-locator | slide 30 image30.png right table | NO GO LOCATOR (HLB) | NEAREST TYPE (flagged) |
| `components[st-slot-locator-1].topMd` | 5164.042 | slide 30 image30.png right table | 1574 | 1574.0 m x 3.28084 = 5164.042 ft (3 dp); m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[st-slot-locator-1].topMdM` | 1574.0 | slide 30 image30.png right table | 1574 | original metres; m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[st-slot-locator-1].innerDiameter` | 2.963 | slide 30 image30.png right table | 2.963 | column 'MIN ID' = explicit ID |
| `components[st-slot-locator-1].outerDiameter` | 6.69 | slide 30 image30.png right table | 6.690 | column 'MAX OD' = explicit OD |
| `components[sliding-sleeve-2].name` | SLIMLINE SSD, 3-1/2" 13CR JFEBEAR (SLB) | slide 30 image30.png right table | SLIMLINE SSD, 3-1/2" 13CR JFEBEAR (SLB) | visual transcription of a low-resolution picture (632x842 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[sliding-sleeve-2].type` | sliding-sleeve | slide 30 image30.png right table | SLIMLINE SSD, 3-1/2" 13CR JFEBEAR (SLB) | EXACT TYPE (Component_Types) |
| `components[sliding-sleeve-2].topMd` | 5364.173 | slide 30 image30.png right table | 1635 | 1635.0 m x 3.28084 = 5364.173 ft (3 dp); m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[sliding-sleeve-2].topMdM` | 1635.0 | slide 30 image30.png right table | 1635 | original metres; m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[sliding-sleeve-2].innerDiameter` | 2.813 | slide 30 image30.png right table | 2.813 | column 'MIN ID' = explicit ID |
| `components[sliding-sleeve-2].outerDiameter` | 4.2 | slide 30 image30.png right table | 4.200 | column 'MAX OD' = explicit OD |
| `components[seal-assembly-1].name` | 4.75" MSN SEAL ASSY (HLB) | slide 30 image30.png right table | 4.75" MSN SEAL ASSY (HLB) | visual transcription of a low-resolution picture (632x842 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[seal-assembly-1].type` | seal-assembly | slide 30 image30.png right table | 4.75" MSN SEAL ASSY (HLB) | EXACT TYPE (Component_Types) |
| `components[seal-assembly-1].topMd` | 5462.599 | slide 30 image30.png right table | 1665 | 1665.0 m x 3.28084 = 5462.599 ft (3 dp); m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[seal-assembly-1].topMdM` | 1665.0 | slide 30 image30.png right table | 1665 | original metres; m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[seal-assembly-1].innerDiameter` | 2.963 | slide 30 image30.png right table | 2.963 | column 'MIN ID' = explicit ID |
| `components[seal-assembly-1].outerDiameter` | 4.75 | slide 30 image30.png right table | 4.750 | column 'MAX OD' = explicit OD |
| `components[sliding-sleeve-3].name` | SLIMLINE SSD, 3-1/2" 13CR JFEBEAR (SLB) | slide 30 image30.png right table | SLIMLINE SSD, 3-1/2" 13CR JFEBEAR (SLB) | visual transcription of a low-resolution picture (632x842 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[sliding-sleeve-3].type` | sliding-sleeve | slide 30 image30.png right table | SLIMLINE SSD, 3-1/2" 13CR JFEBEAR (SLB) | EXACT TYPE (Component_Types) |
| `components[sliding-sleeve-3].topMd` | 5469.16 | slide 30 image30.png right table | 1667 | 1667.0 m x 3.28084 = 5469.16 ft (3 dp); m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[sliding-sleeve-3].topMdM` | 1667.0 | slide 30 image30.png right table | 1667 | original metres; m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[sliding-sleeve-3].innerDiameter` | 2.813 | slide 30 image30.png right table | 2.813 | column 'MIN ID' = explicit ID |
| `components[sliding-sleeve-3].outerDiameter` | 4.2 | slide 30 image30.png right table | 4.200 | column 'MAX OD' = explicit OD |
| `components[landing-nipple-2].name` | 2.813" X- LANDING NIPPLE (SLB) | slide 30 image30.png right table | 2.813" X- LANDING NIPPLE (SLB) | visual transcription of a low-resolution picture (632x842 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[landing-nipple-2].type` | landing-nipple | slide 30 image30.png right table | 2.813" X- LANDING NIPPLE (SLB) | NEAREST TYPE (flagged) |
| `components[landing-nipple-2].topMd` | 5508.53 | slide 30 image30.png right table | 1679 | 1679.0 m x 3.28084 = 5508.53 ft (3 dp); m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[landing-nipple-2].topMdM` | 1679.0 | slide 30 image30.png right table | 1679 | original metres; m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[landing-nipple-2].innerDiameter` | 2.813 | slide 30 image30.png right table | 2.813 | column 'MIN ID' = explicit ID |
| `components[landing-nipple-2].outerDiameter` | 4.04 | slide 30 image30.png right table | 4.040 | column 'MAX OD' = explicit OD |
| `components[seal-assembly-2].name` | 4.75" MSN SEAL ASSY (HLB) | slide 30 image30.png right table | 4.75" MSN SEAL ASSY (HLB) | visual transcription of a low-resolution picture (632x842 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[seal-assembly-2].type` | seal-assembly | slide 30 image30.png right table | 4.75" MSN SEAL ASSY (HLB) | EXACT TYPE (Component_Types) |
| `components[seal-assembly-2].topMd` | 5623.36 | slide 30 image30.png right table | 1714 | 1714.0 m x 3.28084 = 5623.36 ft (3 dp); m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[seal-assembly-2].topMdM` | 1714.0 | slide 30 image30.png right table | 1714 | original metres; m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[seal-assembly-2].innerDiameter` | 2.963 | slide 30 image30.png right table | 2.963 | column 'MIN ID' = explicit ID |
| `components[seal-assembly-2].outerDiameter` | 4.75 | slide 30 image30.png right table | 4.750 | column 'MAX OD' = explicit OD |
| `components[sliding-sleeve-4].name` | SLIMLINE SSD, 3-1/2" 13CR JFEBEAR (SLB) | slide 30 image30.png right table | SLIMLINE SSD, 3-1/2" 13CR JFEBEAR (SLB) | visual transcription of a low-resolution picture (632x842 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[sliding-sleeve-4].type` | sliding-sleeve | slide 30 image30.png right table | SLIMLINE SSD, 3-1/2" 13CR JFEBEAR (SLB) | EXACT TYPE (Component_Types) |
| `components[sliding-sleeve-4].topMd` | 5764.436 | slide 30 image30.png right table | 1757 | 1757.0 m x 3.28084 = 5764.436 ft (3 dp); m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[sliding-sleeve-4].topMdM` | 1757.0 | slide 30 image30.png right table | 1757 | original metres; m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[sliding-sleeve-4].innerDiameter` | 2.813 | slide 30 image30.png right table | 2.813 | column 'MIN ID' = explicit ID |
| `components[sliding-sleeve-4].outerDiameter` | 4.2 | slide 30 image30.png right table | 4.200 | column 'MAX OD' = explicit OD |
| `components[landing-nipple-3].name` | 2.75" NO-GO LANDING NIPPLE (SLB) | slide 30 image30.png right table | 2.75" NO-GO LANDING NIPPLE (SLB) | visual transcription of a low-resolution picture (632x842 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[landing-nipple-3].type` | landing-nipple | slide 30 image30.png right table | 2.75" NO-GO LANDING NIPPLE (SLB) | NEAREST TYPE (flagged) |
| `components[landing-nipple-3].topMd` | 6007.218 | slide 30 image30.png right table | 1831 | 1831.0 m x 3.28084 = 6007.218 ft (3 dp); m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[landing-nipple-3].topMdM` | 1831.0 | slide 30 image30.png right table | 1831 | original metres; m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[landing-nipple-3].innerDiameter` | 2.635 | slide 30 image30.png right table | 2.635 | column 'MIN ID' = explicit ID |
| `components[landing-nipple-3].outerDiameter` | 4.04 | slide 30 image30.png right table | 4.040 | column 'MAX OD' = explicit OD |
| `components[mule-shoe-1].name` | 3-1/2" SELF ALIGNING MULE SHOE (HLB) | slide 30 image30.png right table | 3-1/2" SELF ALIGNING MULE SHOE (HLB) | visual transcription of a low-resolution picture (632x842 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[mule-shoe-1].type` | mule-shoe | slide 30 image30.png right table | 3-1/2" SELF ALIGNING MULE SHOE (HLB) | EXACT TYPE (Component_Types) |
| `components[mule-shoe-1].topMd` | 6036.746 | slide 30 image30.png right table | 1840 | 1840.0 m x 3.28084 = 6036.746 ft (3 dp); m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[mule-shoe-1].topMdM` | 1840.0 | slide 30 image30.png right table | 1840 | original metres; m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[mule-shoe-1].innerDiameter` | 2.94 | slide 30 image30.png right table | 2.940 | column 'MIN ID' = explicit ID |
| `components[mule-shoe-1].outerDiameter` | 3.96 | slide 30 image30.png right table | 3.960 | column 'MAX OD' = explicit OD |
| `components[lc-packer-1].name` | 9-5/8" 6" GRAVEL PACK PACKER (HLB) | slide 30 image30.png left table | 9-5/8" 6" GRAVEL PACK PACKER (HLB) | visual transcription of a low-resolution picture (632x842 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[lc-packer-1].type` | packer | slide 30 image30.png left table | 9-5/8" 6" GRAVEL PACK PACKER (HLB) | EXACT TYPE (Component_Types) |
| `components[lc-packer-1].topMd` | 5164.042 | slide 30 image30.png left table | 1574 | 1574.0 m x 3.28084 = 5164.042 ft (3 dp); m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[lc-packer-1].topMdM` | 1574.0 | slide 30 image30.png left table | 1574 | original metres; m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[lc-packer-1].innerDiameter` | 6.0 | slide 30 image30.png left table | 6.000 | column 'MIN ID' = explicit ID |
| `components[lc-packer-1].outerDiameter` | 8.45 | slide 30 image30.png left table | 8.450 | column 'MAX OD' = explicit OD |
| `components[lc-pbr-1].name` | UPPER SEALBORE EXTENSION (HLB) | slide 30 image30.png left table | UPPER SEALBORE EXTENSION (HLB) | visual transcription of a low-resolution picture (632x842 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[lc-pbr-1].type` | pbr | slide 30 image30.png left table | UPPER SEALBORE EXTENSION (HLB) | NEAREST TYPE (flagged) |
| `components[lc-pbr-1].topMd` | 5177.166 | slide 30 image30.png left table | 1578 | 1578.0 m x 3.28084 = 5177.166 ft (3 dp); m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[lc-pbr-1].topMdM` | 1578.0 | slide 30 image30.png left table | 1578 | original metres; m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[lc-pbr-1].innerDiameter` | 4.75 | slide 30 image30.png left table | 4.750 | column 'MIN ID' = explicit ID |
| `components[lc-pbr-1].outerDiameter` | 5.55 | slide 30 image30.png left table | 5.550 | column 'MAX OD' = explicit OD |
| `components[lc-sliding-sleeve-1].name` | MCS CLOSING SLEEVE (HLB) | slide 30 image30.png left table | MCS CLOSING SLEEVE (HLB) | visual transcription of a low-resolution picture (632x842 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[lc-sliding-sleeve-1].type` | sliding-sleeve | slide 30 image30.png left table | MCS CLOSING SLEEVE (HLB) | NEAREST TYPE (flagged) |
| `components[lc-sliding-sleeve-1].topMd` | 5183.727 | slide 30 image30.png left table | 1580 | 1580.0 m x 3.28084 = 5183.727 ft (3 dp); m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[lc-sliding-sleeve-1].topMdM` | 1580.0 | slide 30 image30.png left table | 1580 | original metres; m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[lc-sliding-sleeve-1].innerDiameter` | 4.75 | slide 30 image30.png left table | 4.750 | column 'MIN ID' = explicit ID |
| `components[lc-sliding-sleeve-1].outerDiameter` | 7.28 | slide 30 image30.png left table | 7.280 | column 'MAX OD' = explicit OD |
| `components[lc-tubing-1].name` | 5-1/2" BLANK PIPE, 13CR-80MY JFE BEAR | slide 30 image30.png left table | 5-1/2" BLANK PIPE, 13CR-80MY JFE BEAR | visual transcription of a low-resolution picture (632x842 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[lc-tubing-1].type` | tubing | slide 30 image30.png left table | 5-1/2" BLANK PIPE, 13CR-80MY JFE BEAR | NEAREST TYPE (flagged) |
| `components[lc-tubing-1].topMd` | 5200.131 | slide 30 image30.png left table | 1585 | 1585.0 m x 3.28084 = 5200.131 ft (3 dp); m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[lc-tubing-1].topMdM` | 1585.0 | slide 30 image30.png left table | 1585 | original metres; m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[lc-tubing-1].innerDiameter` | 4.879 | slide 30 image30.png left table | 4.879 | column 'MIN ID' = explicit ID |
| `components[lc-tubing-1].outerDiameter` | 5.5 | slide 30 image30.png left table | 5.500 | column 'MAX OD' = explicit OD |
| `components[lc-screen-1].name` | 5-1/2" WIRE WRAP SCREEN W/CENTRALIZER, 12 GA (HLB) | slide 30 image30.png left table | 5-1/2" WIRE WRAP SCREEN W/CENTRALIZER, 12 GA (HLB) | visual transcription of a low-resolution picture (632x842 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[lc-screen-1].type` | screen | slide 30 image30.png left table | 5-1/2" WIRE WRAP SCREEN W/CENTRALIZER, 12 GA (HLB) | EXACT TYPE (Component_Types) |
| `components[lc-screen-1].topMd` | 5403.543 | slide 30 image30.png left table | 1647 | 1647.0 m x 3.28084 = 5403.543 ft (3 dp); m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[lc-screen-1].topMdM` | 1647.0 | slide 30 image30.png left table | 1647 | original metres; m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[lc-screen-1].innerDiameter` | 4.888 | slide 30 image30.png left table | 4.888 | column 'MIN ID' = explicit ID |
| `components[lc-screen-1].outerDiameter` | 8.25 | slide 30 image30.png left table | 8.250 | column 'MAX OD' = explicit OD |
| `components[lc-packer-2].name` | 9-5/8" ISOLATION PACKER (HLB) | slide 30 image30.png left table | 9-5/8" ISOLATION PACKER (HLB) | visual transcription of a low-resolution picture (632x842 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[lc-packer-2].type` | packer | slide 30 image30.png left table | 9-5/8" ISOLATION PACKER (HLB) | EXACT TYPE (Component_Types) |
| `components[lc-packer-2].topMd` | 5436.352 | slide 30 image30.png left table | 1657 | 1657.0 m x 3.28084 = 5436.352 ft (3 dp); m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[lc-packer-2].topMdM` | 1657.0 | slide 30 image30.png left table | 1657 | original metres; m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[lc-packer-2].innerDiameter` | 4.75 | slide 30 image30.png left table | 4.750 | column 'MIN ID' = explicit ID |
| `components[lc-packer-2].outerDiameter` | 8.45 | slide 30 image30.png left table | 8.450 | column 'MAX OD' = explicit OD |
| `components[lc-sliding-sleeve-2].name` | MCS CLOSING SLEEVE (HLB) | slide 30 image30.png left table | MCS CLOSING SLEEVE (HLB) | visual transcription of a low-resolution picture (632x842 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[lc-sliding-sleeve-2].type` | sliding-sleeve | slide 30 image30.png left table | MCS CLOSING SLEEVE (HLB) | NEAREST TYPE (flagged) |
| `components[lc-sliding-sleeve-2].topMd` | 5442.914 | slide 30 image30.png left table | 1659 | 1659.0 m x 3.28084 = 5442.914 ft (3 dp); m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[lc-sliding-sleeve-2].topMdM` | 1659.0 | slide 30 image30.png left table | 1659 | original metres; m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[lc-sliding-sleeve-2].innerDiameter` | 4.75 | slide 30 image30.png left table | 4.750 | column 'MIN ID' = explicit ID |
| `components[lc-sliding-sleeve-2].outerDiameter` | 7.28 | slide 30 image30.png left table | 7.280 | column 'MAX OD' = explicit OD |
| `components[lc-tubing-2].name` | 5-1/2" BLANK PIPE, 13CR-80MY JFE BEAR | slide 30 image30.png left table | 5-1/2" BLANK PIPE, 13CR-80MY JFE BEAR | visual transcription of a low-resolution picture (632x842 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[lc-tubing-2].type` | tubing | slide 30 image30.png left table | 5-1/2" BLANK PIPE, 13CR-80MY JFE BEAR | NEAREST TYPE (flagged) |
| `components[lc-tubing-2].topMd` | 5465.879 | slide 30 image30.png left table | 1666 | 1666.0 m x 3.28084 = 5465.879 ft (3 dp); m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[lc-tubing-2].topMdM` | 1666.0 | slide 30 image30.png left table | 1666 | original metres; m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[lc-tubing-2].innerDiameter` | 4.879 | slide 30 image30.png left table | 4.879 | column 'MIN ID' = explicit ID |
| `components[lc-tubing-2].outerDiameter` | 5.5 | slide 30 image30.png left table | 5.500 | column 'MAX OD' = explicit OD |
| `components[lc-screen-2].name` | 5-1/2" WIRE WRAP SCREEN W/CENTRALIZER, 12 GA (HLB) | slide 30 image30.png left table | 5-1/2" WIRE WRAP SCREEN W/CENTRALIZER, 12 GA (HLB) | visual transcription of a low-resolution picture (632x842 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[lc-screen-2].type` | screen | slide 30 image30.png left table | 5-1/2" WIRE WRAP SCREEN W/CENTRALIZER, 12 GA (HLB) | EXACT TYPE (Component_Types) |
| `components[lc-screen-2].topMd` | 5577.428 | slide 30 image30.png left table | 1700 | 1700.0 m x 3.28084 = 5577.428 ft (3 dp); m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[lc-screen-2].topMdM` | 1700.0 | slide 30 image30.png left table | 1700 | original metres; m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[lc-screen-2].innerDiameter` | 4.888 | slide 30 image30.png left table | 4.888 | column 'MIN ID' = explicit ID |
| `components[lc-screen-2].outerDiameter` | 8.25 | slide 30 image30.png left table | 8.250 | column 'MAX OD' = explicit OD |
| `components[lc-packer-3].name` | 9-5/8" ISOLATION PACKER (HLB) | slide 30 image30.png left table | 9-5/8" ISOLATION PACKER (HLB) | visual transcription of a low-resolution picture (632x842 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[lc-packer-3].type` | packer | slide 30 image30.png left table | 9-5/8" ISOLATION PACKER (HLB) | EXACT TYPE (Component_Types) |
| `components[lc-packer-3].topMd` | 5692.257 | slide 30 image30.png left table | 1735 | 1735.0 m x 3.28084 = 5692.257 ft (3 dp); m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[lc-packer-3].topMdM` | 1735.0 | slide 30 image30.png left table | 1735 | original metres; m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[lc-packer-3].innerDiameter` | 4.75 | slide 30 image30.png left table | 4.750 | column 'MIN ID' = explicit ID |
| `components[lc-packer-3].outerDiameter` | 8.45 | slide 30 image30.png left table | 8.450 | column 'MAX OD' = explicit OD |
| `components[lc-sliding-sleeve-3].name` | MCS CLOSING SLEEVE (HLB) | slide 30 image30.png left table | MCS CLOSING SLEEVE (HLB) | visual transcription of a low-resolution picture (632x842 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[lc-sliding-sleeve-3].type` | sliding-sleeve | slide 30 image30.png left table | MCS CLOSING SLEEVE (HLB) | NEAREST TYPE (flagged) |
| `components[lc-sliding-sleeve-3].topMd` | 5698.819 | slide 30 image30.png left table | 1737 | 1737.0 m x 3.28084 = 5698.819 ft (3 dp); m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[lc-sliding-sleeve-3].topMdM` | 1737.0 | slide 30 image30.png left table | 1737 | original metres; m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[lc-sliding-sleeve-3].innerDiameter` | 4.75 | slide 30 image30.png left table | 4.750 | column 'MIN ID' = explicit ID |
| `components[lc-sliding-sleeve-3].outerDiameter` | 7.28 | slide 30 image30.png left table | 7.280 | column 'MAX OD' = explicit OD |
| `components[lc-screen-3].name` | 5-1/2" WIRE WRAP SCREEN W/CENTRALIZER, 12 GA (HLB) | slide 30 image30.png left table | 5-1/2" WIRE WRAP SCREEN W/CENTRALIZER, 12 GA (HLB) | visual transcription of a low-resolution picture (632x842 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[lc-screen-3].type` | screen | slide 30 image30.png left table | 5-1/2" WIRE WRAP SCREEN W/CENTRALIZER, 12 GA (HLB) | EXACT TYPE (Component_Types) |
| `components[lc-screen-3].topMd` | 5839.895 | slide 30 image30.png left table | 1780 | 1780.0 m x 3.28084 = 5839.895 ft (3 dp); m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[lc-screen-3].topMdM` | 1780.0 | slide 30 image30.png left table | 1780 | original metres; m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[lc-screen-3].innerDiameter` | 4.888 | slide 30 image30.png left table | 4.888 | column 'MIN ID' = explicit ID |
| `components[lc-screen-3].outerDiameter` | 8.25 | slide 30 image30.png left table | 8.250 | column 'MAX OD' = explicit OD |
| `components[lc-st-slot-locator-1].name` | SNAP LATCH LOCATOR (HLB) | slide 30 image30.png left table | SNAP LATCH LOCATOR (HLB) | visual transcription of a low-resolution picture (632x842 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[lc-st-slot-locator-1].type` | st-slot-locator | slide 30 image30.png left table | SNAP LATCH LOCATOR (HLB) | NEAREST TYPE (flagged) |
| `components[lc-st-slot-locator-1].topMd` | 5872.704 | slide 30 image30.png left table | 1790 | 1790.0 m x 3.28084 = 5872.704 ft (3 dp); m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[lc-st-slot-locator-1].topMdM` | 1790.0 | slide 30 image30.png left table | 1790 | original metres; m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[lc-st-slot-locator-1].innerDiameter` | 4.888 | slide 30 image30.png left table | 4.888 | column 'MIN ID' = explicit ID |
| `components[lc-st-slot-locator-1].outerDiameter` | 6.6 | slide 30 image30.png left table | 6.600 | column 'MAX OD' = explicit OD |
| `components[lc-packer-4].name` | 9-5/8" SUMP PACKER (HLB) | slide 30 image30.png left table | 9-5/8" SUMP PACKER (HLB) | visual transcription of a low-resolution picture (632x842 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[lc-packer-4].type` | packer | slide 30 image30.png left table | 9-5/8" SUMP PACKER (HLB) | EXACT TYPE (Component_Types) |
| `components[lc-packer-4].topMd` | 5872.704 | slide 30 image30.png left table | 1790 | 1790.0 m x 3.28084 = 5872.704 ft (3 dp); m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[lc-packer-4].topMdM` | 1790.0 | slide 30 image30.png left table | 1790 | original metres; m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[lc-packer-4].innerDiameter` | 6.0 | slide 30 image30.png left table | 6.000 | column 'MIN ID' = explicit ID |
| `components[lc-packer-4].outerDiameter` | 8.22 | slide 30 image30.png left table | 8.220 | column 'MAX OD' = explicit OD |
| `components[surface-casing-1].name` | INTERMEDIATE CASING 13-3/8 ?68 ppf K-55 BTC | slide 30 image30.png tubular table | INTERMEDIATE CASING 13-3/8 ?68 ppf K-55 BTC | visual transcription of a low-resolution picture (632x842 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[surface-casing-1].type` | casing | slide 30 image30.png tubular table | INTERMEDIATE CASING 13-3/8 ?68 ppf K-55 BTC | EXACT TYPE (Component_Types) |
| `components[surface-casing-1].bottomMd` | 1968.504 | slide 30 image30.png tubular table | 600 | 600.0 m x 3.28084 = 1968.504 ft (3 dp); m MDDF (tubular table DEPTH column) |
| `components[surface-casing-1].bottomMdM` | 600.0 | slide 30 image30.png tubular table | 600 | original metres; m MDDF (tubular table DEPTH column) |
| `components[surface-casing-1].innerDiameter` | 12.415 | slide 30 image30.png tubular table | 12.415 | tubular table 'ID (inch)' = explicit ID |
| `components[surface-casing-1].outerDiameter` | 13.375 | slide 30 image30.png tubular table | 13.375 | tubular table 'OD (inch)' = explicit OD |
| `components[surface-casing-1].weight` | 68.0 | slide 30 image30.png tubular table | INTERMEDIATE CASING 13-3/8 ?68 ppf K-55 BTC | parsed from the label (ppf = lb/ft) |
| `components[surface-casing-1].grade` | K-55 | slide 30 image30.png tubular table | INTERMEDIATE CASING 13-3/8 ?68 ppf K-55 BTC | parsed from the label |
| `components[production-casing-1].name` | PRODUCTION CASING 9-5/8 47.0 ppf L80 JFE BEAR | slide 30 image30.png tubular table | PRODUCTION CASING 9-5/8 47.0 ppf L80 JFE BEAR | visual transcription of a low-resolution picture (632x842 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[production-casing-1].type` | casing | slide 30 image30.png tubular table | PRODUCTION CASING 9-5/8 47.0 ppf L80 JFE BEAR | EXACT TYPE (Component_Types) |
| `components[production-casing-1].bottomMd` | 5413.386 | slide 30 image30.png tubular table | 1650 | 1650.0 m x 3.28084 = 5413.386 ft (3 dp); m MDDF (tubular table DEPTH column) |
| `components[production-casing-1].bottomMdM` | 1650.0 | slide 30 image30.png tubular table | 1650 | original metres; m MDDF (tubular table DEPTH column) |
| `components[production-casing-1].innerDiameter` | 8.681 | slide 30 image30.png tubular table | 8.681 | tubular table 'ID (inch)' = explicit ID |
| `components[production-casing-1].outerDiameter` | 9.625 | slide 30 image30.png tubular table | 9.625 | tubular table 'OD (inch)' = explicit OD |
| `components[production-casing-1].weight` | 47.0 | slide 30 image30.png tubular table | PRODUCTION CASING 9-5/8 47.0 ppf L80 JFE BEAR | parsed from the label (ppf = lb/ft) |
| `components[production-casing-1].grade` | L80 | slide 30 image30.png tubular table | PRODUCTION CASING 9-5/8 47.0 ppf L80 JFE BEAR | parsed from the label |
| `components[production-casing-2].name` | PRODUCTION CASING 9-5/8 47.0 ppf 13Cr-80 JFE BEAR (9-5/8" 47# 13CR-80 JFE BEAR CSG SHOE) | slide 30 image30.png tubular table | PRODUCTION CASING 9-5/8 47.0 ppf 13Cr-80 JFE BEAR (9-5/8" 47# 13CR-80 JFE BEAR CSG SHOE) | visual transcription of a low-resolution picture (632x842 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[production-casing-2].type` | casing | slide 30 image30.png tubular table | PRODUCTION CASING 9-5/8 47.0 ppf 13Cr-80 JFE BEAR (9-5/8" 47# 13CR-80 JFE BEAR CSG SHOE) | EXACT TYPE (Component_Types) |
| `components[production-casing-2].bottomMd` | 6843.832 | slide 30 image30.png tubular table | 2086 | 2086.0 m x 3.28084 = 6843.832 ft (3 dp); m MDDF (tubular table DEPTH column) |
| `components[production-casing-2].bottomMdM` | 2086.0 | slide 30 image30.png tubular table | 2086 | original metres; m MDDF (tubular table DEPTH column) |
| `components[production-casing-2].innerDiameter` | 8.681 | slide 30 image30.png tubular table | 8.681 | tubular table 'ID (inch)' = explicit ID |
| `components[production-casing-2].outerDiameter` | 9.625 | slide 30 image30.png tubular table | 9.625 | tubular table 'OD (inch)' = explicit OD |
| `components[production-casing-2].weight` | 47.0 | slide 30 image30.png tubular table | PRODUCTION CASING 9-5/8 47.0 ppf 13Cr-80 JFE BEAR (9-5/8" 47# 13CR-80 JFE BEAR CSG SHOE) | parsed from the label (ppf = lb/ft) |
| `components[production-casing-2].grade` | 13Cr-80 | slide 30 image30.png tubular table | PRODUCTION CASING 9-5/8 47.0 ppf 13Cr-80 JFE BEAR (9-5/8" 47# 13CR-80 JFE BEAR CSG SHOE) | parsed from the label |
| `components[lower-completion-string-1].name` | LOWER COMP STRING 5-1/2 17.0 ppf 13Cr-L80 JFE BEAR | slide 30 image30.png tubular table | LOWER COMP STRING 5-1/2 17.0 ppf 13Cr-L80 JFE BEAR | visual transcription of a low-resolution picture (632x842 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[lower-completion-string-1].type` | tubing | slide 30 image30.png tubular table | LOWER COMP STRING 5-1/2 17.0 ppf 13Cr-L80 JFE BEAR | EXACT TYPE (Component_Types) |
| `components[lower-completion-string-1].bottomMd` | 5872.704 | slide 30 image30.png tubular table | LOWER COMP STRING 5-1/2 17.0 ppf 13Cr-L80 JFE BEAR 1790 MDDF | 1790.0 m x 3.28084 = 5872.704 ft (3 dp); m MDDF (tubular table depth = string end) |
| `components[lower-completion-string-1].bottomMdM` | 1790.0 | slide 30 image30.png tubular table | LOWER COMP STRING 5-1/2 17.0 ppf 13Cr-L80 JFE BEAR 1790 MDDF | original metres; m MDDF (tubular table depth = string end) |
| `components[lower-completion-string-1].innerDiameter` | 4.892 | slide 30 image30.png tubular table | 4.892 | tubular table 'ID (inch)' = explicit ID |
| `components[lower-completion-string-1].outerDiameter` | 5.5 | slide 30 image30.png tubular table | 5.500 | tubular table 'OD (inch)' = explicit OD |
| `components[lower-completion-string-1].weight` | 17.0 | slide 30 image30.png tubular table | LOWER COMP STRING 5-1/2 17.0 ppf 13Cr-L80 JFE BEAR | parsed from the label (ppf = lb/ft) |
| `components[lower-completion-string-1].grade` | 13Cr-L80 | slide 30 image30.png tubular table | LOWER COMP STRING 5-1/2 17.0 ppf 13Cr-L80 JFE BEAR | parsed from the label |
| `components[tubing-1].name` | PRODUCTION TUBING #1 3-1/2 9.2 ppf HP2-13Cr-95 JFE BEAR | slide 30 image30.png tubular table | PRODUCTION TUBING #1 3-1/2 9.2 ppf HP2-13Cr-95 JFE BEAR | visual transcription of a low-resolution picture (632x842 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[tubing-1].type` | tubing | slide 30 image30.png tubular table | PRODUCTION TUBING #1 3-1/2 9.2 ppf HP2-13Cr-95 JFE BEAR | EXACT TYPE (Component_Types) |
| `components[tubing-1].bottomMd` | 820.21 | slide 30 image30.png tubular table | PRODUCTION TUBING #1 3-1/2 9.2 ppf HP2-13Cr-95 JFE BEAR 250 MDDF | 250.0 m x 3.28084 = 820.21 ft (3 dp); m MDDF (tubular table depth = string end) |
| `components[tubing-1].bottomMdM` | 250.0 | slide 30 image30.png tubular table | PRODUCTION TUBING #1 3-1/2 9.2 ppf HP2-13Cr-95 JFE BEAR 250 MDDF | original metres; m MDDF (tubular table depth = string end) |
| `components[tubing-1].innerDiameter` | 2.992 | slide 30 image30.png tubular table | 2.992 | tubular table 'ID (inch)' = explicit ID |
| `components[tubing-1].outerDiameter` | 3.5 | slide 30 image30.png tubular table | 3.500 | tubular table 'OD (inch)' = explicit OD |
| `components[tubing-1].weight` | 9.2 | slide 30 image30.png tubular table | PRODUCTION TUBING #1 3-1/2 9.2 ppf HP2-13Cr-95 JFE BEAR | parsed from the label (ppf = lb/ft) |
| `components[tubing-1].grade` | HP2-13Cr-95 | slide 30 image30.png tubular table | PRODUCTION TUBING #1 3-1/2 9.2 ppf HP2-13Cr-95 JFE BEAR | parsed from the label |
| `components[tubing-2].name` | PRODUCTION TUBING #2 3-1/2 9.2 ppf 13Cr-L80 JFE BEAR | slide 30 image30.png tubular table | PRODUCTION TUBING #2 3-1/2 9.2 ppf 13Cr-L80 JFE BEAR | visual transcription of a low-resolution picture (632x842 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[tubing-2].type` | tubing | slide 30 image30.png tubular table | PRODUCTION TUBING #2 3-1/2 9.2 ppf 13Cr-L80 JFE BEAR | EXACT TYPE (Component_Types) |
| `components[tubing-2].bottomMd` | 6036.746 | slide 30 image30.png tubular table | PRODUCTION TUBING #2 3-1/2 9.2 ppf 13Cr-L80 JFE BEAR 1840 MDDF | 1840.0 m x 3.28084 = 6036.746 ft (3 dp); m MDDF (tubular table depth = string end) |
| `components[tubing-2].bottomMdM` | 1840.0 | slide 30 image30.png tubular table | PRODUCTION TUBING #2 3-1/2 9.2 ppf 13Cr-L80 JFE BEAR 1840 MDDF | original metres; m MDDF (tubular table depth = string end) |
| `components[tubing-2].innerDiameter` | 2.992 | slide 30 image30.png tubular table | 2.992 | tubular table 'ID (inch)' = explicit ID |
| `components[tubing-2].outerDiameter` | 3.5 | slide 30 image30.png tubular table | 3.500 | tubular table 'OD (inch)' = explicit OD |
| `components[tubing-2].weight` | 9.2 | slide 30 image30.png tubular table | PRODUCTION TUBING #2 3-1/2 9.2 ppf 13Cr-L80 JFE BEAR | parsed from the label (ppf = lb/ft) |
| `components[tubing-2].grade` | 13Cr-L80 | slide 30 image30.png tubular table | PRODUCTION TUBING #2 3-1/2 9.2 ppf 13Cr-L80 JFE BEAR | parsed from the label |
| `components[zone-e7-1].name` | E7 ?1650.7-?1652.7 m MDDF | slide 30 image30.png left table | E7 ?1650.7-?1652.7 m MDDF | visual transcription of a low-resolution picture (632x842 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[zone-e7-1].type` | perforation | slide 30 image30.png left table | E7 ?1650.7-?1652.7 m MDDF | EXACT TYPE (Component_Types) |
| `components[zone-e7-1].topMd` | 5415.683 | slide 30 image30.png left table | E7 SAND - EXTENSION PACK: Perf Interval: 1650.7 m-MDDF - 1652.7 m-MDDF (2 meters) (low legibility) | 1650.7 m x 3.28084 = 5415.683 ft (3 dp); m MDDF as printed |
| `components[zone-e7-1].topMdM` | 1650.7 | slide 30 image30.png left table | E7 SAND - EXTENSION PACK: Perf Interval: 1650.7 m-MDDF - 1652.7 m-MDDF (2 meters) (low legibility) | original metres; m MDDF as printed |
| `components[zone-e7-1].bottomMd` | 5422.244 | slide 30 image30.png left table | E7 SAND - EXTENSION PACK: Perf Interval: 1650.7 m-MDDF - 1652.7 m-MDDF (2 meters) (low legibility) | 1652.7 m x 3.28084 = 5422.244 ft (3 dp); m MDDF as printed |
| `components[zone-e7-1].bottomMdM` | 1652.7 | slide 30 image30.png left table | E7 SAND - EXTENSION PACK: Perf Interval: 1650.7 m-MDDF - 1652.7 m-MDDF (2 meters) (low legibility) | original metres; m MDDF as printed |
| `components[zone-e7-1].zone` | E7 | slide 30 image30.png left table | E7 | zone label as printed |
| `components[zone-e10a-e11b-1].name` | E10A/E11B ?1702.55-?1704.55 m MDDF | slide 30 image30.png left table | E10A/E11B ?1702.55-?1704.55 m MDDF | visual transcription of a low-resolution picture (632x842 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[zone-e10a-e11b-1].type` | perforation | slide 30 image30.png left table | E10A/E11B ?1702.55-?1704.55 m MDDF | EXACT TYPE (Component_Types) |
| `components[zone-e10a-e11b-1].topMd` | 5585.794 | slide 30 image30.png left table | E10A/E11B SAND - EXTENSION PACK (HYBRID): Perf Interval: 1702.55 m-MDDF - 1704.55 m-MDDF (2 meters) (low legibility) | 1702.55 m x 3.28084 = 5585.794 ft (3 dp); m MDDF as printed |
| `components[zone-e10a-e11b-1].topMdM` | 1702.55 | slide 30 image30.png left table | E10A/E11B SAND - EXTENSION PACK (HYBRID): Perf Interval: 1702.55 m-MDDF - 1704.55 m-MDDF (2 meters) (low legibility) | original metres; m MDDF as printed |
| `components[zone-e10a-e11b-1].bottomMd` | 5592.356 | slide 30 image30.png left table | E10A/E11B SAND - EXTENSION PACK (HYBRID): Perf Interval: 1702.55 m-MDDF - 1704.55 m-MDDF (2 meters) (low legibility) | 1704.55 m x 3.28084 = 5592.356 ft (3 dp); m MDDF as printed |
| `components[zone-e10a-e11b-1].bottomMdM` | 1704.55 | slide 30 image30.png left table | E10A/E11B SAND - EXTENSION PACK (HYBRID): Perf Interval: 1702.55 m-MDDF - 1704.55 m-MDDF (2 meters) (low legibility) | original metres; m MDDF as printed |
| `components[zone-e10a-e11b-1].zone` | E10A/E11B | slide 30 image30.png left table | E10A/E11B | zone label as printed |
| `components[zone-e12b-1].name` | E12B ?1725-?1733 m MDDF | slide 30 image30.png left table | E12B ?1725-?1733 m MDDF | visual transcription of a low-resolution picture (632x842 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[zone-e12b-1].type` | perforation | slide 30 image30.png left table | E12B ?1725-?1733 m MDDF | EXACT TYPE (Component_Types) |
| `components[zone-e12b-1].topMd` | 5659.449 | slide 30 image30.png left table | E12B SAND - CIRCULATING PACK (HYBRID): Perf Interval: 1725 m-MDDF - 1733 m-MDDF (8 meters) (low legibility) | 1725.0 m x 3.28084 = 5659.449 ft (3 dp); m MDDF as printed |
| `components[zone-e12b-1].topMdM` | 1725.0 | slide 30 image30.png left table | E12B SAND - CIRCULATING PACK (HYBRID): Perf Interval: 1725 m-MDDF - 1733 m-MDDF (8 meters) (low legibility) | original metres; m MDDF as printed |
| `components[zone-e12b-1].bottomMd` | 5685.696 | slide 30 image30.png left table | E12B SAND - CIRCULATING PACK (HYBRID): Perf Interval: 1725 m-MDDF - 1733 m-MDDF (8 meters) (low legibility) | 1733.0 m x 3.28084 = 5685.696 ft (3 dp); m MDDF as printed |
| `components[zone-e12b-1].bottomMdM` | 1733.0 | slide 30 image30.png left table | E12B SAND - CIRCULATING PACK (HYBRID): Perf Interval: 1725 m-MDDF - 1733 m-MDDF (8 meters) (low legibility) | original metres; m MDDF as printed |
| `components[zone-e12b-1].zone` | E12B | slide 30 image30.png left table | E12B | zone label as printed |
| `components[zone-e14b-1].name` | E14B ?1783.65-?1785.65 m MDDF | slide 30 image30.png left table | E14B ?1783.65-?1785.65 m MDDF | visual transcription of a low-resolution picture (632x842 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[zone-e14b-1].type` | perforation | slide 30 image30.png left table | E14B ?1783.65-?1785.65 m MDDF | EXACT TYPE (Component_Types) |
| `components[zone-e14b-1].topMd` | 5851.87 | slide 30 image30.png left table | E14B SAND - EXTENSION PACK: Perf Interval: 1783.65 m-MDDF - 1785.65 m-MDDF (2 meters) (low legibility) | 1783.65 m x 3.28084 = 5851.87 ft (3 dp); m MDDF as printed |
| `components[zone-e14b-1].topMdM` | 1783.65 | slide 30 image30.png left table | E14B SAND - EXTENSION PACK: Perf Interval: 1783.65 m-MDDF - 1785.65 m-MDDF (2 meters) (low legibility) | original metres; m MDDF as printed |
| `components[zone-e14b-1].bottomMd` | 5858.432 | slide 30 image30.png left table | E14B SAND - EXTENSION PACK: Perf Interval: 1783.65 m-MDDF - 1785.65 m-MDDF (2 meters) (low legibility) | 1785.65 m x 3.28084 = 5858.432 ft (3 dp); m MDDF as printed |
| `components[zone-e14b-1].bottomMdM` | 1785.65 | slide 30 image30.png left table | E14B SAND - EXTENSION PACK: Perf Interval: 1783.65 m-MDDF - 1785.65 m-MDDF (2 meters) (low legibility) | original metres; m MDDF as printed |
| `components[zone-e14b-1].zone` | E14B | slide 30 image30.png left table | E14B | zone label as printed |
| `components[pbtd-1].name` | PBTD ?2016 (left table; header PBTD : ?2016 m-MDDF) | slide 30 image30.png diagram | PBTD ?2016 (left table; header PBTD : ?2016 m-MDDF) | visual transcription of a low-resolution picture (632x842 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[pbtd-1].type` | pbtd | slide 30 image30.png diagram | PBTD ?2016 (left table; header PBTD : ?2016 m-MDDF) | EXACT TYPE (Component_Types) |
| `components[pbtd-1].topMd` | 6614.173 | slide 30 image30.png diagram | PBTD ?2016 (left table; header PBTD : ?2016 m-MDDF) | 2016.0 m x 3.28084 = 6614.173 ft (3 dp); m MDDF |
| `components[pbtd-1].topMdM` | 2016.0 | slide 30 image30.png diagram | PBTD ?2016 (left table; header PBTD : ?2016 m-MDDF) | original metres; m MDDF |
| `wellInfo.completionDate` | 2022-05-30 | slide 30 image30.png header | ?30-May-22 | ISO date (two-digit year expanded) |
| `wellInfo.completionFluid` | 8.5ppg Brine (8.33 ppg DW + 3% KCl filtered to 30 NTU) | slide 30 image30.png header | 8.5ppg Brine (8.33 ppg DW + 3% KCl filtered to 30 NTU) | as printed |
| `wellInfo.completionType` | PROPOSED WELL COMPLETION SCHEMATIC (OIL PRODUCER #2) — Rev 2, 20-May-22; 9-5/8" STMZ - 3 ZONES | slide 30 image30.png header | PROPOSED WELL COMPLETION SCHEMATIC (OIL PRODUCER #2) — Rev 2, 20-May-22; 9-5/8" STMZ - 3 ZONES | as printed |
| `wellInfo.groundElevation` | None |  |  | not supported by the source -> template empty value ('-' text / null number) |
| `wellInfo.kbElevation` | None |  |  | not supported by the source -> template empty value ('-' text / null number) |
| `wellInfo.maxDeviation` | ?67.12 deg | slide 30 image30.png header | ?67.12 deg | as printed |
| `wellInfo.name` | Dulang_B30ST1 | slide 30 | WELL NAME: B30 ST-1 (LOCATION / SLOT: DULANG B PLATFORM /SLOT 26) | identifier rule Dulang_B<no><suffix>; original label kept in the summary |
| `wellInfo.packerFluid` | 8.5ppg Brine + 0.2% Biocide + 2% Oxygen Scavenger + 2% Corrosion Inhibitor | slide 30 image30.png header | 8.5ppg Brine + 0.2% Biocide + 2% Oxygen Scavenger + 2% Corrosion Inhibitor | as printed |
| `wellInfo.rig` | Sapura T9 (TADR) | slide 30 image30.png header | Sapura T9 (TADR) | as printed |
| `wellInfo.surfaceCasing` | - |  |  | not supported by the source -> template empty value ('-' text / null number) |
| `wellInfo.totalDepth` | 6843.832 | slide 30 image30.png header/diagram | TOTAL DEPTH : ?2086 m-MDDF | 2086.0 m x 3.28084 = 6843.832 ft; datum MDDF as printed (not shifted) |
| `wellInfo.trajectoryType` | - |  |  | not supported by the source -> template empty value ('-' text / null number) |
| `wellInfo.tubing` | PRODUCTION TUBING #1 3-1/2 9.2 HP2-13Cr-95 JFE BEAR ?250 m-MDDF; PRODUCTION TUBING #2 3-1/2 9.2 13Cr-L80 JFE BEAR ?1840 m-MDDF | slide 30 image30.png header | PRODUCTION TUBING #1 3-1/2 9.2 HP2-13Cr-95 JFE BEAR ?250 m-MDDF; PRODUCTION TUBING #2 3-1/2 9.2 13Cr-L80 JFE BEAR ?1840 m-MDDF | as printed |
| `wellInfo.wellhead` | Solar Alert | slide 30 image30.png header | Solar Alert | as printed |

## Dulang_B31ST1 — slide 31

Original label: WELL NAME: B31 ST-1 (LOCATION / SLOT: DULANG B PLATFORM /SLOT 14)
Equipment depth datum: m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment')  ·  source: image31.png (visual transcription of a low-resolution picture (728x817 px, x3 upscale), OCR cross-check — many digits uncertain)

### Conflicts / missing / unclear

| kind | where | detail |
|---|---|---|
| Unresolved | slide 31 image31.png right table | '4-1/2 SSSV, SELF-EQ, 25CR JFE BEAR': no Type_Key |
| Unclear | slide 31 image31.png header/diagram | TD 1735.0 m printed in MDDF; equipment datum m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') — no datum shift applied |
| Unclear | picture | very low-resolution picture: most depths and diameters are uncertain reads (marked '?', flagged estimated); the Excel original in ECMS (named on the slide) should be used to confirm — reported as partly unreadable content |
| Unclear | picture | header shows 4-1/2" completion equipment while the tubing table lists 5-1/2" 26 ppf tubing — both recorded as printed |

### Unmapped content (kept here, not in the JSON)

| where | text | reason |
|---|---|---|
| slide 31 image31.png right table | 4-1/2 SSSV, SELF-EQ, 25CR JFE BEAR | no Type_Key could be assigned |
| left table | 13-3/8" CSG SIDETRACK WINDOW ?531 m-MDDF / ?507 m-TVDDF; 9-5/8" 47# L80 JFE BEAR RS CSG ?1351 m-MDDF; 9-5/8" CASING PIP TAG ?1489 m-MDDF; sand tops/bottoms and BHP/BHT per zone (E6/E7 1000 psi 198.5 F; E31/E32 1090 psi 203.9 F; E34 1580 psi 206.7 F); 'Rathole (E34 bottom perf to PBTD): 25 m MDDF'; t | note (status / history / survey / other revision) |
| slide 31 | 'FOR B31 LATEST EDITABLE DIAGRAM, PLEASE REFER TO EXCEL FILE IN ECMS - PRAVIN NAIR 2023-' | note (status / history / survey / other revision) |
| slide 31 obj #19 | FOR B31 LATEST EDITABLE DIAGRAM, PLEASE REFER TO EXCEL FILE IN ECMS - PRAVIN NAIR 2023- | slide text box over/beside the picture (status / history / note) |

### Traceability (every populated value)

| JSON path | value | slide / object / location | original text | conversion / interpretation |
|---|---|---|---|---|
| `components[christmas-tree-1].name` | ?Solar Alert 5-1/8" 5K U/FF Trim | slide 31 image31.png header | ?Solar Alert 5-1/8" 5K U/FF Trim | visual transcription of a low-resolution picture (728x817 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[christmas-tree-1].type` | christmas-tree | slide 31 image31.png header | ?Solar Alert 5-1/8" 5K U/FF Trim | EXACT TYPE (Component_Types) |
| `components[tubing-hanger-1].name` | TUBING HANGER, 13-5/8" x 6.375"-4SA-2G LH x 5-1/2" 26# VAM21, w/5" HBPV PROFILE | slide 31 image31.png right table | TUBING HANGER, 13-5/8" x 6.375"-4SA-2G LH x 5-1/2" 26# VAM21, w/5" HBPV PROFILE | visual transcription of a low-resolution picture (728x817 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[tubing-hanger-1].type` | tubing-hanger | slide 31 image31.png right table | TUBING HANGER, 13-5/8" x 6.375"-4SA-2G LH x 5-1/2" 26# VAM21, w/5" HBPV PROFILE | EXACT TYPE (Component_Types) |
| `components[flow-coupling-1].name` | FLOW COUPLING, 4-1/2" 25CR JFEBEAR | slide 31 image31.png right table | FLOW COUPLING, 4-1/2" 25CR JFEBEAR | visual transcription of a low-resolution picture (728x817 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[flow-coupling-1].type` | flow-coupling | slide 31 image31.png right table | FLOW COUPLING, 4-1/2" 25CR JFEBEAR | EXACT TYPE (Component_Types) |
| `components[flow-coupling-1].topMd` | 711.942 | slide 31 image31.png right table | 217 | 217.0 m x 3.28084 = 711.942 ft (3 dp); m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[flow-coupling-1].topMdM` | 217.0 | slide 31 image31.png right table | 217 | original metres; m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[flow-coupling-1].innerDiameter` | 3.845 | slide 31 image31.png right table | 3.845 | column 'MIN ID' = explicit ID |
| `components[flow-coupling-1].outerDiameter` | 5.09 | slide 31 image31.png right table | 5.090 | column 'MAX OD' = explicit OD |
| `components[flow-coupling-2].name` | FLOW COUPLING | slide 31 image31.png right table | FLOW COUPLING | visual transcription of a low-resolution picture (728x817 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[flow-coupling-2].type` | flow-coupling | slide 31 image31.png right table | FLOW COUPLING | EXACT TYPE (Component_Types) |
| `components[flow-coupling-2].innerDiameter` | 3.845 | slide 31 image31.png right table | 3.845 | column 'MIN ID' = explicit ID |
| `components[flow-coupling-2].outerDiameter` | 5.09 | slide 31 image31.png right table | 5.090 | column 'MAX OD' = explicit OD |
| `components[sliding-sleeve-1].name` | SSD, 3.813" X-PROFILE, 4-1/2" 25CR JFEBEAR | slide 31 image31.png right table | SSD, 3.813" X-PROFILE, 4-1/2" 25CR JFEBEAR | visual transcription of a low-resolution picture (728x817 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[sliding-sleeve-1].type` | sliding-sleeve | slide 31 image31.png right table | SSD, 3.813" X-PROFILE, 4-1/2" 25CR JFEBEAR | EXACT TYPE (Component_Types) |
| `components[sliding-sleeve-1].topMd` | 5065.617 | slide 31 image31.png right table | 1544 | 1544.0 m x 3.28084 = 5065.617 ft (3 dp); m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[sliding-sleeve-1].topMdM` | 1544.0 | slide 31 image31.png right table | 1544 | original metres; m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[sliding-sleeve-1].innerDiameter` | 3.813 | slide 31 image31.png right table | 3.813 | column 'MIN ID' = explicit ID |
| `components[sliding-sleeve-1].outerDiameter` | 5.613 | slide 31 image31.png right table | 5.613 | column 'MAX OD' = explicit OD |
| `components[packer-1].name` | 9-5/8" X 4-1/2" PRODUCTION PACKER, 25CR, JFE BEAR, V0 RATED HYDRAULIC SET, CUT-TO-RELEASE | slide 31 image31.png right table | 9-5/8" X 4-1/2" PRODUCTION PACKER, 25CR, JFE BEAR, V0 RATED HYDRAULIC SET, CUT-TO-RELEASE | visual transcription of a low-resolution picture (728x817 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[packer-1].type` | packer | slide 31 image31.png right table | 9-5/8" X 4-1/2" PRODUCTION PACKER, 25CR, JFE BEAR, V0 RATED HYDRAULIC SET, CUT-TO-RELEASE | EXACT TYPE (Component_Types) |
| `components[packer-1].topMd` | 5127.953 | slide 31 image31.png right table | 1563 | 1563.0 m x 3.28084 = 5127.953 ft (3 dp); m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[packer-1].topMdM` | 1563.0 | slide 31 image31.png right table | 1563 | original metres; m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[packer-1].innerDiameter` | 3.9 | slide 31 image31.png right table | 3.900 | column 'MIN ID' = explicit ID |
| `components[packer-1].outerDiameter` | 8.5 | slide 31 image31.png right table | 8.500 | column 'MAX OD' = explicit OD |
| `components[sliding-sleeve-2].name` | SSD, 4-1/2" 25CR JFEBEAR | slide 31 image31.png right table | SSD, 4-1/2" 25CR JFEBEAR | visual transcription of a low-resolution picture (728x817 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[sliding-sleeve-2].type` | sliding-sleeve | slide 31 image31.png right table | SSD, 4-1/2" 25CR JFEBEAR | EXACT TYPE (Component_Types) |
| `components[sliding-sleeve-2].topMd` | 5137.795 | slide 31 image31.png right table | 1566 | 1566.0 m x 3.28084 = 5137.795 ft (3 dp); m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[sliding-sleeve-2].topMdM` | 1566.0 | slide 31 image31.png right table | 1566 | original metres; m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[sliding-sleeve-2].innerDiameter` | 3.813 | slide 31 image31.png right table | 3.813 | column 'MIN ID' = explicit ID |
| `components[sliding-sleeve-2].outerDiameter` | 5.613 | slide 31 image31.png right table | 5.613 | column 'MAX OD' = explicit OD |
| `components[sliding-sleeve-3].name` | PREMIUM PORT, 4-1/2" 25CR JFEBEAR | slide 31 image31.png right table | PREMIUM PORT, 4-1/2" 25CR JFEBEAR | visual transcription of a low-resolution picture (728x817 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[sliding-sleeve-3].type` | sliding-sleeve | slide 31 image31.png right table | PREMIUM PORT, 4-1/2" 25CR JFEBEAR | NEAREST TYPE (flagged) |
| `components[sliding-sleeve-3].topMd` | 5190.289 | slide 31 image31.png right table | 1582 | 1582.0 m x 3.28084 = 5190.289 ft (3 dp); m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[sliding-sleeve-3].topMdM` | 1582.0 | slide 31 image31.png right table | 1582 | original metres; m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[sliding-sleeve-3].innerDiameter` | 3.9 | slide 31 image31.png right table | 3.900 | column 'MIN ID' = explicit ID |
| `components[sliding-sleeve-3].outerDiameter` | 6.125 | slide 31 image31.png right table | 6.125 | column 'MAX OD' = explicit OD |
| `components[packer-2].name` | 9-5/8" X 4-1/2" BLUEPACK RH PACKER, 25CR, JFE BEAR, V3 RATED HYDRAULIC SET, PULL-TO-RELEASE | slide 31 image31.png right table | 9-5/8" X 4-1/2" BLUEPACK RH PACKER, 25CR, JFE BEAR, V3 RATED HYDRAULIC SET, PULL-TO-RELEASE | visual transcription of a low-resolution picture (728x817 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[packer-2].type` | packer | slide 31 image31.png right table | 9-5/8" X 4-1/2" BLUEPACK RH PACKER, 25CR, JFE BEAR, V3 RATED HYDRAULIC SET, PULL-TO-RELEASE | EXACT TYPE (Component_Types) |
| `components[packer-2].topMd` | 5252.625 | slide 31 image31.png right table | 1601 | 1601.0 m x 3.28084 = 5252.625 ft (3 dp); m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[packer-2].topMdM` | 1601.0 | slide 31 image31.png right table | 1601 | original metres; m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[packer-2].innerDiameter` | 3.9 | slide 31 image31.png right table | 3.900 | column 'MIN ID' = explicit ID |
| `components[packer-2].outerDiameter` | 8.5 | slide 31 image31.png right table | 8.500 | column 'MAX OD' = explicit OD |
| `components[sliding-sleeve-4].name` | PREMIUM PORT, 4-1/2" 25CR JFEBEAR | slide 31 image31.png right table | PREMIUM PORT, 4-1/2" 25CR JFEBEAR | visual transcription of a low-resolution picture (728x817 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[sliding-sleeve-4].type` | sliding-sleeve | slide 31 image31.png right table | PREMIUM PORT, 4-1/2" 25CR JFEBEAR | NEAREST TYPE (flagged) |
| `components[sliding-sleeve-4].topMd` | 5318.242 | slide 31 image31.png right table | 1621 | 1621.0 m x 3.28084 = 5318.242 ft (3 dp); m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[sliding-sleeve-4].topMdM` | 1621.0 | slide 31 image31.png right table | 1621 | original metres; m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[sliding-sleeve-4].innerDiameter` | 3.9 | slide 31 image31.png right table | 3.900 | column 'MIN ID' = explicit ID |
| `components[sliding-sleeve-4].outerDiameter` | 6.125 | slide 31 image31.png right table | 6.125 | column 'MAX OD' = explicit OD |
| `components[sliding-sleeve-5].name` | SSD, 3.813" X-PROFILE, 4-1/2" 25CR JFEBEAR | slide 31 image31.png right table | SSD, 3.813" X-PROFILE, 4-1/2" 25CR JFEBEAR | visual transcription of a low-resolution picture (728x817 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[sliding-sleeve-5].type` | sliding-sleeve | slide 31 image31.png right table | SSD, 3.813" X-PROFILE, 4-1/2" 25CR JFEBEAR | EXACT TYPE (Component_Types) |
| `components[sliding-sleeve-5].topMd` | 5341.208 | slide 31 image31.png right table | 1628 | 1628.0 m x 3.28084 = 5341.208 ft (3 dp); m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[sliding-sleeve-5].topMdM` | 1628.0 | slide 31 image31.png right table | 1628 | original metres; m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[sliding-sleeve-5].innerDiameter` | 3.813 | slide 31 image31.png right table | 3.813 | column 'MIN ID' = explicit ID |
| `components[sliding-sleeve-5].outerDiameter` | 5.613 | slide 31 image31.png right table | 5.613 | column 'MAX OD' = explicit OD |
| `components[packer-3].name` | 9-5/8" X 4-1/2" PRODUCTION PACKER, 25CR, JFE BEAR, V0 RATED HYDRAULIC SET, CUT-TO-RELEASE | slide 31 image31.png right table | 9-5/8" X 4-1/2" PRODUCTION PACKER, 25CR, JFE BEAR, V0 RATED HYDRAULIC SET, CUT-TO-RELEASE | visual transcription of a low-resolution picture (728x817 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[packer-3].type` | packer | slide 31 image31.png right table | 9-5/8" X 4-1/2" PRODUCTION PACKER, 25CR, JFE BEAR, V0 RATED HYDRAULIC SET, CUT-TO-RELEASE | EXACT TYPE (Component_Types) |
| `components[packer-3].topMd` | 5400.263 | slide 31 image31.png right table | 1646 | 1646.0 m x 3.28084 = 5400.263 ft (3 dp); m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[packer-3].topMdM` | 1646.0 | slide 31 image31.png right table | 1646 | original metres; m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[packer-3].innerDiameter` | 3.9 | slide 31 image31.png right table | 3.900 | column 'MIN ID' = explicit ID |
| `components[packer-3].outerDiameter` | 8.5 | slide 31 image31.png right table | 8.500 | column 'MAX OD' = explicit OD |
| `components[sliding-sleeve-6].name` | SSD, 3.813" X-PROFILE, 4-1/2" 25CR JFEBEAR | slide 31 image31.png right table | SSD, 3.813" X-PROFILE, 4-1/2" 25CR JFEBEAR | visual transcription of a low-resolution picture (728x817 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[sliding-sleeve-6].type` | sliding-sleeve | slide 31 image31.png right table | SSD, 3.813" X-PROFILE, 4-1/2" 25CR JFEBEAR | EXACT TYPE (Component_Types) |
| `components[sliding-sleeve-6].topMd` | 5410.105 | slide 31 image31.png right table | 1649 | 1649.0 m x 3.28084 = 5410.105 ft (3 dp); m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[sliding-sleeve-6].topMdM` | 1649.0 | slide 31 image31.png right table | 1649 | original metres; m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[sliding-sleeve-6].innerDiameter` | 3.813 | slide 31 image31.png right table | 3.813 | column 'MIN ID' = explicit ID |
| `components[sliding-sleeve-6].outerDiameter` | 5.613 | slide 31 image31.png right table | 5.613 | column 'MAX OD' = explicit OD |
| `components[sliding-sleeve-7].name` | PREMIUM PORT, 4-1/2" 25CR JFEBEAR | slide 31 image31.png right table | PREMIUM PORT, 4-1/2" 25CR JFEBEAR | visual transcription of a low-resolution picture (728x817 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[sliding-sleeve-7].type` | sliding-sleeve | slide 31 image31.png right table | PREMIUM PORT, 4-1/2" 25CR JFEBEAR | NEAREST TYPE (flagged) |
| `components[sliding-sleeve-7].topMd` | 5419.948 | slide 31 image31.png right table | 1652 | 1652.0 m x 3.28084 = 5419.948 ft (3 dp); m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[sliding-sleeve-7].topMdM` | 1652.0 | slide 31 image31.png right table | 1652 | original metres; m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[sliding-sleeve-7].innerDiameter` | 3.9 | slide 31 image31.png right table | 3.900 | column 'MIN ID' = explicit ID |
| `components[sliding-sleeve-7].outerDiameter` | 6.125 | slide 31 image31.png right table | 6.125 | column 'MAX OD' = explicit OD |
| `components[landing-nipple-1].name` | LANDING NIPPLE, 3.725" NO GO, 4-1/2" 25CR JFEBEAR | slide 31 image31.png right table | LANDING NIPPLE, 3.725" NO GO, 4-1/2" 25CR JFEBEAR | visual transcription of a low-resolution picture (728x817 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[landing-nipple-1].type` | landing-nipple | slide 31 image31.png right table | LANDING NIPPLE, 3.725" NO GO, 4-1/2" 25CR JFEBEAR | NEAREST TYPE (flagged) |
| `components[landing-nipple-1].innerDiameter` | 3.725 | slide 31 image31.png right table | 3.725 | column 'MIN ID' = explicit ID |
| `components[landing-nipple-1].outerDiameter` | 5.06 | slide 31 image31.png right table | 5.060 | column 'MAX OD' = explicit OD |
| `components[circulating-valve-1].name` | SELF-FILLING SUB, 4-1/2", 25CR JFE BEAR | slide 31 image31.png right table | SELF-FILLING SUB, 4-1/2", 25CR JFE BEAR | visual transcription of a low-resolution picture (728x817 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[circulating-valve-1].type` | circulating-valve | slide 31 image31.png right table | SELF-FILLING SUB, 4-1/2", 25CR JFE BEAR | NEAREST TYPE (flagged) |
| `components[mechanical-plug-1].name` | BULL PLUG, 4-1/2", 25CR JFE BEAR | slide 31 image31.png right table | BULL PLUG, 4-1/2", 25CR JFE BEAR | visual transcription of a low-resolution picture (728x817 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[mechanical-plug-1].type` | mechanical-plug | slide 31 image31.png right table | BULL PLUG, 4-1/2", 25CR JFE BEAR | NEAREST TYPE (flagged) |
| `components[mechanical-plug-1].topMd` | 5485.564 | slide 31 image31.png right table | 1672 | 1672.0 m x 3.28084 = 5485.564 ft (3 dp); m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[mechanical-plug-1].topMdM` | 1672.0 | slide 31 image31.png right table | 1672 | original metres; m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[lc-bridge-plug-1].name` | TOP OF BRIDGE PLUG | slide 31 image31.png left table | TOP OF BRIDGE PLUG | visual transcription of a low-resolution picture (728x817 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[lc-bridge-plug-1].type` | bridge-plug | slide 31 image31.png left table | TOP OF BRIDGE PLUG | EXACT TYPE (Component_Types) |
| `components[lc-bridge-plug-1].topMd` | 5538.058 | slide 31 image31.png left table | 1688 | 1688.0 m x 3.28084 = 5538.058 ft (3 dp); m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[lc-bridge-plug-1].topMdM` | 1688.0 | slide 31 image31.png left table | 1688 | original metres; m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[lc-bridge-plug-2].name` | CEMENT RETAINER | slide 31 image31.png left table | CEMENT RETAINER | visual transcription of a low-resolution picture (728x817 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[lc-bridge-plug-2].type` | bridge-plug | slide 31 image31.png left table | CEMENT RETAINER | NEAREST TYPE (flagged) |
| `components[lc-bridge-plug-2].topMd` | 5544.62 | slide 31 image31.png left table | 1690 | 1690.0 m x 3.28084 = 5544.62 ft (3 dp); m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[lc-bridge-plug-2].topMdM` | 1690.0 | slide 31 image31.png left table | 1690 | original metres; m MDDF (columns 'DEPTH m-MDDF'; 'Note: All depths are in meter from RKB to top of equipment') (equipment depth column as printed) |
| `components[surface-casing-1].name` | INTERMEDIATE CASING 13-3/8 54.5 ppf K-55 BTC | slide 31 image31.png tubular table | INTERMEDIATE CASING 13-3/8 54.5 ppf K-55 BTC | visual transcription of a low-resolution picture (728x817 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[surface-casing-1].type` | casing | slide 31 image31.png tubular table | INTERMEDIATE CASING 13-3/8 54.5 ppf K-55 BTC | EXACT TYPE (Component_Types) |
| `components[surface-casing-1].bottomMd` | 1738.845 | slide 31 image31.png tubular table | 530.00 | 530.0 m x 3.28084 = 1738.845 ft (3 dp); m MDDF (tubular table DEPTH column) |
| `components[surface-casing-1].bottomMdM` | 530.0 | slide 31 image31.png tubular table | 530.00 | original metres; m MDDF (tubular table DEPTH column) |
| `components[surface-casing-1].innerDiameter` | 12.615 | slide 31 image31.png tubular table | 12.615 | tubular table 'ID (inch)' = explicit ID |
| `components[surface-casing-1].outerDiameter` | 13.375 | slide 31 image31.png tubular table | 13.375 | tubular table 'OD (inch)' = explicit OD |
| `components[surface-casing-1].weight` | 54.5 | slide 31 image31.png tubular table | INTERMEDIATE CASING 13-3/8 54.5 ppf K-55 BTC | parsed from the label (ppf = lb/ft) |
| `components[surface-casing-1].grade` | K-55 | slide 31 image31.png tubular table | INTERMEDIATE CASING 13-3/8 54.5 ppf K-55 BTC | parsed from the label |
| `components[production-casing-1].name` | PRODUCTION CASING 9-5/8 47.0 ppf L80 JFE BEAR | slide 31 image31.png tubular table | PRODUCTION CASING 9-5/8 47.0 ppf L80 JFE BEAR | visual transcription of a low-resolution picture (728x817 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[production-casing-1].type` | casing | slide 31 image31.png tubular table | PRODUCTION CASING 9-5/8 47.0 ppf L80 JFE BEAR | EXACT TYPE (Component_Types) |
| `components[production-casing-1].bottomMd` | 4429.134 | slide 31 image31.png tubular table | 1350.00 | 1350.0 m x 3.28084 = 4429.134 ft (3 dp); m MDDF (tubular table DEPTH column) |
| `components[production-casing-1].bottomMdM` | 1350.0 | slide 31 image31.png tubular table | 1350.00 | original metres; m MDDF (tubular table DEPTH column) |
| `components[production-casing-1].innerDiameter` | 8.681 | slide 31 image31.png tubular table | 8.681 | tubular table 'ID (inch)' = explicit ID |
| `components[production-casing-1].outerDiameter` | 9.625 | slide 31 image31.png tubular table | 9.625 | tubular table 'OD (inch)' = explicit OD |
| `components[production-casing-1].weight` | 47.0 | slide 31 image31.png tubular table | PRODUCTION CASING 9-5/8 47.0 ppf L80 JFE BEAR | parsed from the label (ppf = lb/ft) |
| `components[production-casing-1].grade` | L80 | slide 31 image31.png tubular table | PRODUCTION CASING 9-5/8 47.0 ppf L80 JFE BEAR | parsed from the label |
| `components[production-casing-2].name` | PRODUCTION CASING 9-5/8 47.0 ppf 13CR-80 JFE BEAR (9-5/8" 47# 13CR-80 JFE BEAR CSG SHOE) | slide 31 image31.png tubular table | PRODUCTION CASING 9-5/8 47.0 ppf 13CR-80 JFE BEAR (9-5/8" 47# 13CR-80 JFE BEAR CSG SHOE) | visual transcription of a low-resolution picture (728x817 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[production-casing-2].type` | casing | slide 31 image31.png tubular table | PRODUCTION CASING 9-5/8 47.0 ppf 13CR-80 JFE BEAR (9-5/8" 47# 13CR-80 JFE BEAR CSG SHOE) | EXACT TYPE (Component_Types) |
| `components[production-casing-2].bottomMd` | 5692.257 | slide 31 image31.png tubular table | 1735.00 | 1735.0 m x 3.28084 = 5692.257 ft (3 dp); m MDDF (tubular table DEPTH column) |
| `components[production-casing-2].bottomMdM` | 1735.0 | slide 31 image31.png tubular table | 1735.00 | original metres; m MDDF (tubular table DEPTH column) |
| `components[production-casing-2].innerDiameter` | 8.681 | slide 31 image31.png tubular table | 8.681 | tubular table 'ID (inch)' = explicit ID |
| `components[production-casing-2].outerDiameter` | 9.625 | slide 31 image31.png tubular table | 9.625 | tubular table 'OD (inch)' = explicit OD |
| `components[production-casing-2].weight` | 47.0 | slide 31 image31.png tubular table | PRODUCTION CASING 9-5/8 47.0 ppf 13CR-80 JFE BEAR (9-5/8" 47# 13CR-80 JFE BEAR CSG SHOE) | parsed from the label (ppf = lb/ft) |
| `components[production-casing-2].grade` | 13CR-80 | slide 31 image31.png tubular table | PRODUCTION CASING 9-5/8 47.0 ppf 13CR-80 JFE BEAR (9-5/8" 47# 13CR-80 JFE BEAR CSG SHOE) | parsed from the label |
| `components[tubing-1].name` | PRODUCTION TUBING 5-1/2 26 ppf SM22Cr-125 VAM21 | slide 31 image31.png tubular table | PRODUCTION TUBING 5-1/2 26 ppf SM22Cr-125 VAM21 | visual transcription of a low-resolution picture (728x817 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[tubing-1].type` | tubing | slide 31 image31.png tubular table | PRODUCTION TUBING 5-1/2 26 ppf SM22Cr-125 VAM21 | EXACT TYPE (Component_Types) |
| `components[tubing-1].bottomMd` | 5485.564 | slide 31 image31.png tubular table | PRODUCTION TUBING 5-1/2 26 ppf SM22Cr-125 VAM21 1672.00 MDDF | 1672.0 m x 3.28084 = 5485.564 ft (3 dp); m MDDF (tubular table depth = string end) |
| `components[tubing-1].bottomMdM` | 1672.0 | slide 31 image31.png tubular table | PRODUCTION TUBING 5-1/2 26 ppf SM22Cr-125 VAM21 1672.00 MDDF | original metres; m MDDF (tubular table depth = string end) |
| `components[tubing-1].innerDiameter` | 4.548 | slide 31 image31.png tubular table | 4.548 | tubular table 'ID (inch)' = explicit ID |
| `components[tubing-1].outerDiameter` | 5.5 | slide 31 image31.png tubular table | 5.500 | tubular table 'OD (inch)' = explicit OD |
| `components[tubing-1].weight` | 26.0 | slide 31 image31.png tubular table | PRODUCTION TUBING 5-1/2 26 ppf SM22Cr-125 VAM21 | parsed from the label (ppf = lb/ft) |
| `components[tubing-1].grade` | SM22Cr-125 | slide 31 image31.png tubular table | PRODUCTION TUBING 5-1/2 26 ppf SM22Cr-125 VAM21 | parsed from the label |
| `components[zone-e6-e7-1].name` | E6 / E7 ?1570.02-?1574.5 m MDDF | slide 31 image31.png left table | E6 / E7 ?1570.02-?1574.5 m MDDF | visual transcription of a low-resolution picture (728x817 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[zone-e6-e7-1].type` | perforation | slide 31 image31.png left table | E6 / E7 ?1570.02-?1574.5 m MDDF | EXACT TYPE (Component_Types) |
| `components[zone-e6-e7-1].topMd` | 5150.984 | slide 31 image31.png left table | E6 / E7 SAND Top Perf: 1570.02 m-MDDF / 1243.6 m-TVDDF; Bottom Perf: 1574.5 m-MDDF / 1247 m-TVDDF (Perf Length: 4.5 m) (low legibility) | 1570.02 m x 3.28084 = 5150.984 ft (3 dp); m MDDF as printed |
| `components[zone-e6-e7-1].topMdM` | 1570.02 | slide 31 image31.png left table | E6 / E7 SAND Top Perf: 1570.02 m-MDDF / 1243.6 m-TVDDF; Bottom Perf: 1574.5 m-MDDF / 1247 m-TVDDF (Perf Length: 4.5 m) (low legibility) | original metres; m MDDF as printed |
| `components[zone-e6-e7-1].bottomMd` | 5165.683 | slide 31 image31.png left table | E6 / E7 SAND Top Perf: 1570.02 m-MDDF / 1243.6 m-TVDDF; Bottom Perf: 1574.5 m-MDDF / 1247 m-TVDDF (Perf Length: 4.5 m) (low legibility) | 1574.5 m x 3.28084 = 5165.683 ft (3 dp); m MDDF as printed |
| `components[zone-e6-e7-1].bottomMdM` | 1574.5 | slide 31 image31.png left table | E6 / E7 SAND Top Perf: 1570.02 m-MDDF / 1243.6 m-TVDDF; Bottom Perf: 1574.5 m-MDDF / 1247 m-TVDDF (Perf Length: 4.5 m) (low legibility) | original metres; m MDDF as printed |
| `components[zone-e6-e7-1].zone` | E6 / E7 | slide 31 image31.png left table | E6 / E7 | zone label as printed |
| `components[zone-e6-e7-2].name` | E6 / E7 ?1581-?1585 m MDDF | slide 31 image31.png left table | E6 / E7 ?1581-?1585 m MDDF | visual transcription of a low-resolution picture (728x817 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[zone-e6-e7-2].type` | perforation | slide 31 image31.png left table | E6 / E7 ?1581-?1585 m MDDF | EXACT TYPE (Component_Types) |
| `components[zone-e6-e7-2].topMd` | 5187.008 | slide 31 image31.png left table | Top Perf: 1581 m-MDDF / 1253 m-TVDDF; Bottom Perf: 1585 m-MDDF / 1257 m-TVDDF (Perf Length: 4 m) | 1581.0 m x 3.28084 = 5187.008 ft (3 dp); m MDDF as printed |
| `components[zone-e6-e7-2].topMdM` | 1581.0 | slide 31 image31.png left table | Top Perf: 1581 m-MDDF / 1253 m-TVDDF; Bottom Perf: 1585 m-MDDF / 1257 m-TVDDF (Perf Length: 4 m) | original metres; m MDDF as printed |
| `components[zone-e6-e7-2].bottomMd` | 5200.131 | slide 31 image31.png left table | Top Perf: 1581 m-MDDF / 1253 m-TVDDF; Bottom Perf: 1585 m-MDDF / 1257 m-TVDDF (Perf Length: 4 m) | 1585.0 m x 3.28084 = 5200.131 ft (3 dp); m MDDF as printed |
| `components[zone-e6-e7-2].bottomMdM` | 1585.0 | slide 31 image31.png left table | Top Perf: 1581 m-MDDF / 1253 m-TVDDF; Bottom Perf: 1585 m-MDDF / 1257 m-TVDDF (Perf Length: 4 m) | original metres; m MDDF as printed |
| `components[zone-e6-e7-2].zone` | E6 / E7 | slide 31 image31.png left table | E6 / E7 | zone label as printed |
| `components[zone-e31-e32-1].name` | E31 / E32 ?1608.5-?1612.5 m MDDF | slide 31 image31.png left table | E31 / E32 ?1608.5-?1612.5 m MDDF | visual transcription of a low-resolution picture (728x817 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[zone-e31-e32-1].type` | perforation | slide 31 image31.png left table | E31 / E32 ?1608.5-?1612.5 m MDDF | EXACT TYPE (Component_Types) |
| `components[zone-e31-e32-1].topMd` | 5277.231 | slide 31 image31.png left table | E31 / E32 SAND Top Perf: 1608.5 m-MDDF / 1276.65 m-TVDDF; Bottom Perf: 1612.5 m-MDDF / 1280.1 m-TVDDF (Perf Length: 4 m) | 1608.5 m x 3.28084 = 5277.231 ft (3 dp); m MDDF as printed |
| `components[zone-e31-e32-1].topMdM` | 1608.5 | slide 31 image31.png left table | E31 / E32 SAND Top Perf: 1608.5 m-MDDF / 1276.65 m-TVDDF; Bottom Perf: 1612.5 m-MDDF / 1280.1 m-TVDDF (Perf Length: 4 m) | original metres; m MDDF as printed |
| `components[zone-e31-e32-1].bottomMd` | 5290.355 | slide 31 image31.png left table | E31 / E32 SAND Top Perf: 1608.5 m-MDDF / 1276.65 m-TVDDF; Bottom Perf: 1612.5 m-MDDF / 1280.1 m-TVDDF (Perf Length: 4 m) | 1612.5 m x 3.28084 = 5290.355 ft (3 dp); m MDDF as printed |
| `components[zone-e31-e32-1].bottomMdM` | 1612.5 | slide 31 image31.png left table | E31 / E32 SAND Top Perf: 1608.5 m-MDDF / 1276.65 m-TVDDF; Bottom Perf: 1612.5 m-MDDF / 1280.1 m-TVDDF (Perf Length: 4 m) | original metres; m MDDF as printed |
| `components[zone-e31-e32-1].zone` | E31 / E32 | slide 31 image31.png left table | E31 / E32 | zone label as printed |
| `components[zone-e31-e32-2].name` | E31 / E32 ?1623-?1633.5 m MDDF | slide 31 image31.png left table | E31 / E32 ?1623-?1633.5 m MDDF | visual transcription of a low-resolution picture (728x817 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[zone-e31-e32-2].type` | perforation | slide 31 image31.png left table | E31 / E32 ?1623-?1633.5 m MDDF | EXACT TYPE (Component_Types) |
| `components[zone-e31-e32-2].topMd` | 5324.803 | slide 31 image31.png left table | Top Perf: 1623 m-MDDF / 1289 m-TVDDF; Bottom Perf: 1633.5 m-MDDF / 1298 m-TVDDF (Perf Length: 10.5 m) | 1623.0 m x 3.28084 = 5324.803 ft (3 dp); m MDDF as printed |
| `components[zone-e31-e32-2].topMdM` | 1623.0 | slide 31 image31.png left table | Top Perf: 1623 m-MDDF / 1289 m-TVDDF; Bottom Perf: 1633.5 m-MDDF / 1298 m-TVDDF (Perf Length: 10.5 m) | original metres; m MDDF as printed |
| `components[zone-e31-e32-2].bottomMd` | 5359.252 | slide 31 image31.png left table | Top Perf: 1623 m-MDDF / 1289 m-TVDDF; Bottom Perf: 1633.5 m-MDDF / 1298 m-TVDDF (Perf Length: 10.5 m) | 1633.5 m x 3.28084 = 5359.252 ft (3 dp); m MDDF as printed |
| `components[zone-e31-e32-2].bottomMdM` | 1633.5 | slide 31 image31.png left table | Top Perf: 1623 m-MDDF / 1289 m-TVDDF; Bottom Perf: 1633.5 m-MDDF / 1298 m-TVDDF (Perf Length: 10.5 m) | original metres; m MDDF as printed |
| `components[zone-e31-e32-2].zone` | E31 / E32 | slide 31 image31.png left table | E31 / E32 | zone label as printed |
| `components[zone-e34-1].name` | E34 ?1658-?1663 m MDDF | slide 31 image31.png left table | E34 ?1658-?1663 m MDDF | visual transcription of a low-resolution picture (728x817 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[zone-e34-1].type` | perforation | slide 31 image31.png left table | E34 ?1658-?1663 m MDDF | EXACT TYPE (Component_Types) |
| `components[zone-e34-1].topMd` | 5439.633 | slide 31 image31.png left table | E34 SAND Top Perf: 1658 m-MDDF / 1319 m-TVDDF; Bottom Perf: 1663 m-MDDF / 1323.3 m-TVDDF (Perf Length: 5 m) | 1658.0 m x 3.28084 = 5439.633 ft (3 dp); m MDDF as printed |
| `components[zone-e34-1].topMdM` | 1658.0 | slide 31 image31.png left table | E34 SAND Top Perf: 1658 m-MDDF / 1319 m-TVDDF; Bottom Perf: 1663 m-MDDF / 1323.3 m-TVDDF (Perf Length: 5 m) | original metres; m MDDF as printed |
| `components[zone-e34-1].bottomMd` | 5456.037 | slide 31 image31.png left table | E34 SAND Top Perf: 1658 m-MDDF / 1319 m-TVDDF; Bottom Perf: 1663 m-MDDF / 1323.3 m-TVDDF (Perf Length: 5 m) | 1663.0 m x 3.28084 = 5456.037 ft (3 dp); m MDDF as printed |
| `components[zone-e34-1].bottomMdM` | 1663.0 | slide 31 image31.png left table | E34 SAND Top Perf: 1658 m-MDDF / 1319 m-TVDDF; Bottom Perf: 1663 m-MDDF / 1323.3 m-TVDDF (Perf Length: 5 m) | original metres; m MDDF as printed |
| `components[zone-e34-1].zone` | E34 | slide 31 image31.png left table | E34 | zone label as printed |
| `components[pbtd-1].name` | PBTD ?1696 (left table; header PBTD : ?1696 m-MDDF) | slide 31 image31.png diagram | PBTD ?1696 (left table; header PBTD : ?1696 m-MDDF) | visual transcription of a low-resolution picture (728x817 px, x3 upscale), OCR cross-check — many digits uncertain |
| `components[pbtd-1].type` | pbtd | slide 31 image31.png diagram | PBTD ?1696 (left table; header PBTD : ?1696 m-MDDF) | EXACT TYPE (Component_Types) |
| `components[pbtd-1].topMd` | 5564.305 | slide 31 image31.png diagram | PBTD ?1696 (left table; header PBTD : ?1696 m-MDDF) | 1696.0 m x 3.28084 = 5564.305 ft (3 dp); m MDDF |
| `components[pbtd-1].topMdM` | 1696.0 | slide 31 image31.png diagram | PBTD ?1696 (left table; header PBTD : ?1696 m-MDDF) | original metres; m MDDF |
| `wellInfo.completionDate` | 2022-06-24 | slide 31 image31.png header | 24-Jun-22 | ISO date (two-digit year expanded) |
| `wellInfo.completionFluid` | 8.33 ppg DW + 3% KCl filtered to 30 NTU | slide 31 image31.png header | 8.33 ppg DW + 3% KCl filtered to 30 NTU | as printed |
| `wellInfo.completionType` | FINAL WELL COMPLETION SCHEMATIC (WATER INJECTOR) — Rev: Final, Date ?1-Jul-22 | slide 31 image31.png header | FINAL WELL COMPLETION SCHEMATIC (WATER INJECTOR) — Rev: Final, Date ?1-Jul-22 | as printed |
| `wellInfo.groundElevation` | None |  |  | not supported by the source -> template empty value ('-' text / null number) |
| `wellInfo.kbElevation` | None |  |  | not supported by the source -> template empty value ('-' text / null number) |
| `wellInfo.maxDeviation` | ?61.0 deg | slide 31 image31.png header | ?61.0 deg | as printed |
| `wellInfo.name` | Dulang_B31ST1 | slide 31 | WELL NAME: B31 ST-1 (LOCATION / SLOT: DULANG B PLATFORM /SLOT 14) | identifier rule Dulang_B<no><suffix>; original label kept in the summary |
| `wellInfo.packerFluid` | 8.33 ppg DW + 3% KCl + 0.2% Biocide + 2% Oxygen Scavenger + 2% Corrosion Inhibitor | slide 31 image31.png header | 8.33 ppg DW + 3% KCl + 0.2% Biocide + 2% Oxygen Scavenger + 2% Corrosion Inhibitor | as printed |
| `wellInfo.rig` | Sapura T9 (TADR) | slide 31 image31.png header | Sapura T9 (TADR) | as printed |
| `wellInfo.surfaceCasing` | - |  |  | not supported by the source -> template empty value ('-' text / null number) |
| `wellInfo.totalDepth` | 5692.257 | slide 31 image31.png header/diagram | TOTAL DEPTH : ?1735 m-MDDF | 1735.0 m x 3.28084 = 5692.257 ft; datum MDDF as printed (not shifted) |
| `wellInfo.trajectoryType` | - |  |  | not supported by the source -> template empty value ('-' text / null number) |
| `wellInfo.tubing` | PRODUCTION TUBING 5-1/2 26 ppf SM22Cr-125 VAM21 ?1672.00 m-MDDF | slide 31 image31.png header | PRODUCTION TUBING 5-1/2 26 ppf SM22Cr-125 VAM21 ?1672.00 m-MDDF | as printed |
| `wellInfo.wellhead` | Solar Alert | slide 31 image31.png header | Solar Alert | as printed |

