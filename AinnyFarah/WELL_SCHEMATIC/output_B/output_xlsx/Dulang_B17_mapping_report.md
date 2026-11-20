# Dulang_B17 — mapping and quality-check report

- Source workbook: `/Volumes/homes/public/DulangOSS/20261016/DULANG_WELLS/Well_Schematic/1. Well Schematic/Dulang B/Dulang B17 ESP Workover Schematic.xlsx` (read-only; worked on a copy in `_work/src/`)
- Internal well label: `COMPLETION SCHEMATIC : DULANG B-17` (B2) — agrees with the file name; output name `Dulang_B17` (user-approved)
- Worksheet mapped: `New B-17 (12.12.23)`; worksheets inspected: 'New B-17 (12.12.23)' (visible), 'MLE Space out' (visible)
- Reading: cells + formulas + cached values + drawing layer (shapes / text boxes / pictures) from the .xlsx XML; rendered with macOS Quick Look (`_work/render/ql/`) and visually compared; embedded images extracted to `_work/media/B17/` and OCR'd where they hold text
- Template: `/Users/farahyusof/Documents/GitHub/NoramiraSazuki/AinnyFarah/WELL_SCHEMATIC/well-schematic-template-json.rtf`; Type_Keys from `well-schematic-sample-template.xlsx` > Component_Types
- Output JSON: `Dulang_B17.json` — 39 components; surveys / cement / plugBackCement / centralizers / floatEquipment / formations / formationTops / fluidContacts left empty (no supported source data)
- Generated: 2026-10-29 10:30

## Depth units, reference and conversion

- Source depths are **m-MDTHF (MD from tubing hanger flange; column headers J16/Q16)**. JSON depths are **ft = m × 3.28084**, 3 decimals; original metres kept in `topMdM` / `bottomMdM`.
- The MD reference (THF) was not shifted to drill floor/RKB and no MD→TVD conversion was made.

| field | ref | m | unit | ft |
|---|---|---|---|---|
| trsv.topMd | Q22 | 224.2 | m-MDTHF (column header Q16) | 735.564 |
| glm-1.topMd | Q26 | 541.38 | m-MDTHF (column header Q16) | 1776.181 |
| glm-2.topMd | Q28 | 803.08 | m-MDTHF (column header Q16) | 2634.777 |
| ssd-1.topMd | Q30 | 811.3 | m-MDTHF (column header Q16) | 2661.745 |
| esp-packer.topMd | Q33 | 830.31 | m-MDTHF (column header Q16) | 2724.114 |
| ssd-2.topMd | Q37 | 836.99 | m-MDTHF (column header Q16) | 2746.03 |
| sand-guard.topMd | Q44 | 847.87 | m-MDTHF (column header Q16) | 2781.726 |
| esp.topMd | Q50 | 893.56 | m-MDTHF (column header Q16) | 2931.627 |
| eot.topMd | Q51 | 906.75 | m-MDTHF (column header Q16) | 2974.902 |
| ss-stub.topMd | J54 | 1219.2 | m-MDTHF (column header J16) | 4000.0 |
| ls-stub.topMd | Q54 | 1218.0 | m-MDTHF (column header Q16) | 3996.063 |
| dual-hyd-packer.topMd | J55 | 1220.7 | m-MDTHF (column header J16) | 4004.921 |
| xn-no-go-nipple-left.topMd | J58 | 1221.0 | m-MDTHF (column header J16) | 4005.906 |
| liner-7.topMd | Q61 | 1315.0 | m-MDTHF (column header Q16) | 4314.305 |
| liner-7.bottomMd | Q93 | 2194.0 | m-MDTHF (column header Q16) | 7198.163 |
| casing-9-5-8.bottomMd | Q64 | 1379.0 | m-MDTHF (column header Q16) | 4524.278 |
| ssd-3.topMd | Q65 | 1708.0 | m-MDTHF (column header Q16) | 5603.675 |
| bwd-packer-1.topMd | J67 | 1724.0 | m-MDTHF (column header J16) | 5656.168 |
| bwd-packer-2.topMd | J76 | 1792.0 | m-MDTHF (column header J16) | 5879.265 |
| ssd-4.topMd | Q80 | 1784.7 | m-MDTHF (column header Q16) | 5855.315 |
| bwd-packer-3.topMd | J86 | 1883.0 | m-MDTHF (column header J16) | 6177.822 |
| xn-no-go-nipple-right.topMd | Q89 | 1947.0 | m-MDTHF (column header Q16) | 6387.795 |
| conical-shoe.topMd | Q90 | 1950.0 | m-MDTHF (column header Q16) | 6397.638 |
| pbtd.topMd | L95 | 2171.7 | m (in text; MDTHF stated) | 7125.0 |
| perf-e12-13.topMd | C71 | 1734.7 | m (in text; MDTHF stated) | 5691.273 |
| perf-e12-13.bottomMd | C71 | 1763.7 | m (in text; MDTHF stated) | 5786.418 |
| perf-e14-1.topMd | C80 | 1796.7 | m (in text; MDTHF stated) | 5894.685 |
| perf-e14-1.bottomMd | C80 | 1816.7 | m (in text; MDTHF stated) | 5960.302 |
| perf-e14-2.topMd | C80 | 1824.7 | m (in text; MDTHF stated) | 5986.549 |
| perf-e14-2.bottomMd | C80 | 1832.7 | m (in text; MDTHF stated) | 6012.795 |
| perf-e34-36-1.topMd | C90 | 2012.7 | m (in text; MDTHF stated) | 6603.347 |
| perf-e34-36-1.bottomMd | C90 | 2042.7 | m (in text; MDTHF stated) | 6701.772 |
| perf-e34-36-2.topMd | C90 | 2052.7 | m (in text; MDTHF stated) | 6734.58 |
| perf-e34-36-2.bottomMd | C90 | 2067.7 | m (in text; MDTHF stated) | 6783.793 |
| perf-e40-42.topMd | C90 | 2102.7 | m (in text; MDTHF stated) | 6898.622 |
| perf-e40-42.bottomMd | C90 | 2120.7 | m (in text; MDTHF stated) | 6957.677 |

## wellInfo

| json_path | mapped_value | cell_or_object | original_value | method | note |
|---|---|---|---|---|---|
| wellInfo.completionDate | - |  |  | EMPTY | not supported by the visible source -> template empty value ('-' text / null number) |
| wellInfo.completionFluid | - |  |  | EMPTY | not supported by the visible source -> template empty value ('-' text / null number) |
| wellInfo.completionType | SINGLE OIL PRODUCER | B3 | ' SINGLE OIL PRODUCER ' | DIRECT | visible title row 3 ('SINGLE OIL PRODUCER'); ESP completion |
| wellInfo.groundElevation |  |  |  | EMPTY | not supported by the visible source -> template empty value ('-' text / null number) |
| wellInfo.kbElevation |  |  |  | EMPTY | not supported by the visible source -> template empty value ('-' text / null number) |
| wellInfo.maxDeviation | - |  |  | EMPTY | not supported by the visible source -> template empty value ('-' text / null number) |
| wellInfo.name | Dulang_B17 | B2 | 'COMPLETION SCHEMATIC : DULANG B-17' | STANDARDISED NAME | internal label B2 'COMPLETION SCHEMATIC : DULANG B-17' -> Dulang_B17 (user-approved naming) |
| wellInfo.packerFluid | - |  |  | EMPTY | not supported by the visible source -> template empty value ('-' text / null number) |
| wellInfo.rig | - |  |  | EMPTY | not supported by the visible source -> template empty value ('-' text / null number) |
| wellInfo.surfaceCasing | - |  |  | EMPTY | not supported by the visible source -> template empty value ('-' text / null number) |
| wellInfo.totalDepth |  |  |  | EMPTY | not supported by the visible source -> template empty value ('-' text / null number) |
| wellInfo.trajectoryType | - |  |  | EMPTY | not supported by the visible source -> template empty value ('-' text / null number) |
| wellInfo.tubing | - |  |  | EMPTY | not supported by the visible source -> template empty value ('-' text / null number) |
| wellInfo.wellhead | - |  |  | EMPTY | not supported by the visible source -> template empty value ('-' text / null number) |

## Components

| id | type | name | string | topMd_ft | bottomMd_ft | topMd_m | bottomMd_m | OD_in | ID_in | interval | zone | estimated |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| flow-coupling-1 | flow-coupling | Flow Coupling |  |  |  |  |  | 5.09 | 3.884 | 0.914 m |  |  |
| trsv | trsv | 4-1/2"TRSV |  | 735.564 |  | 224.2 |  | 6.923 | 3.812 | 1.625 m |  |  |
| flow-coupling-2 | flow-coupling | Flow Coupling |  |  |  |  |  | 5.09 | 3.884 | 0.914 m |  |  |
| glm-1 | spm | 4-1/2" KBMG GLM (DUMMY) |  | 1776.181 |  | 541.38 |  | 6.393 | 3.865 | 2.234 m |  | type |
| glm-2 | spm | 4-1/2" KBMG GLM (GLOV 14/64") |  | 2634.777 |  | 803.08 |  | 6.393 | 3.865 | 2.234 m |  | type |
| ssd-1 | sliding-sleeve | 4-1/2" SSD (CLOSED) |  | 2661.745 |  | 811.3 |  | 5.613 | 3.813 | 1.448 m |  |  |
| esp-packer | packer | 9-5/8" x 4-1/2" ESP PACKER |  | 2724.114 |  | 830.31 |  | 8.835 | 3.96 | 1.46 m |  |  |
| ssd-2 | sliding-sleeve | 4-1/2" SSD (OPEN) |  | 2746.03 |  | 836.99 |  | 5.613 | 3.813 | 1.448 m |  |  |
| crossover-1 | crossover | CROSSOVER SUB 4-1/2" JFE-BEAR BOX X 3-1/2" VAMTOP PIN |  |  |  |  |  | 5.0 | 2.959 | 0.31 m |  |  |
| sand-guard | sand-screen | 3-1/2" MULTILIFT 550 SAND GUARD |  | 2781.726 |  | 847.87 |  | 5.5 |  | 6.33 m |  | type |
| crossover-2 | crossover | CROSSOVER SUB 3-1/2" JFE-BEAR PIN X 4-1/2" VAMTOP BOX |  |  |  |  |  | 5.0 | 2.959 | 0.28 m |  |  |
| esp | esp | 4-1/2" ESP |  | 2931.627 |  | 893.56 |  | 6.132 |  | 21.9 m |  |  |
| eot | eot | EOT |  | 2974.902 |  | 906.75 |  |  |  |  |  |  |
| ss-stub | tubing | SS STUB | Short | 4000.0 |  | 1219.2 |  |  |  |  |  | type |
| ls-stub | tubing | LS STUB | Long | 3996.063 |  | 1218.0 |  |  |  |  |  | type |
| dual-hyd-packer | packer | 9-5/8" DUAL HYD. PKR. | Short & Long | 4004.921 |  | 1220.7 |  |  |  |  |  |  |
| xn-no-go-nipple-left | nipple (no-go) | XN' NO-GO NIPPLE |  | 4005.906 |  | 1221.0 |  |  |  |  |  |  |
| liner-7 | liner | 7" TOP OF LINER 7" LINER SHOE |  | 4314.305 | 7198.163 | 1315.0 | 2194.0 |  | 7.0 |  |  | innerDiameter |
| casing-9-5-8 | casing | 9-5/8" Casing Shoe |  |  | 4524.278 |  | 1379.0 |  | 9.625 |  |  | innerDiameter |
| ssd-3 | sliding-sleeve | SSD |  | 5603.675 |  | 1708.0 |  |  | 2.313 |  |  |  |
| bwd-packer-1 | packer | 7" X 4.5" X 3.88" BWD PKR. BORE |  | 5656.168 |  | 1724.0 |  |  | 3.88 |  |  |  |
| blank-pipe-1 | tubing | 4½" L-80 LCT BLANK PIPE |  |  |  |  |  |  | 3.25 |  |  | type |
| screen-1 | screen | 4" P110 NU WWSCREEN 8 ga |  |  |  |  |  |  | 4.0 |  |  | innerDiameter |
| seal-bore-assembly-1 | seal-assembly | 3.25" SEAL BORE ASSY. |  |  |  |  |  |  | 3.25 |  |  | innerDiameter |
| bwd-packer-2 | packer | 7" X 4.5" X 3.88" BWD PKR. BORE |  | 5879.265 |  | 1792.0 |  |  | 3.25 |  |  |  |
| ssd-4 | sliding-sleeve | SSD |  | 5855.315 |  | 1784.7 |  |  | 2.313 |  |  |  |
| blank-pipe-2 | tubing | 4½" L-80 LCT BLANK PIPE |  |  |  |  |  |  |  |  |  | type |
| screen-2 | screen | 4" J-55 NU WW SCREEN 8 ga |  |  |  |  |  |  | 4.0 |  |  | innerDiameter |
| seal-bore-assembly-2 | seal-assembly | 3.25" SEAL BORE ASSY. |  |  |  |  |  |  | 3.25 |  |  | innerDiameter |
| bwd-packer-3 | packer | 7" X 4.5" X 3.88" BWD PKR. BORE |  | 6177.822 |  | 1883.0 |  |  | 3.25 |  |  |  |
| xn-no-go-nipple-right | nipple (no-go) | XN' NO-GO NIPPLE |  | 6387.795 |  | 1947.0 |  |  | 2.225 |  |  |  |
| conical-shoe | mule-shoe | CONICAL SHOE |  | 6397.638 |  | 1950.0 |  |  |  |  |  | type |
| pbtd | pbtd | PBTD @ 2171.7 m MDTHF |  | 7125.0 |  | 2171.7 |  |  |  |  |  |  |
| perf-e12-13 | perforation | E12/13 (FB S2; WIR) perforation 1734.7 m - 1763.7 m |  | 5691.273 | 5786.418 | 1734.7 | 1763.7 |  |  |  | E12/13 (FB S2; WIR) |  |
| perf-e14-1 | perforation | E14 (FB S2; WIR) perforation 1796.7 m - 1816.7 m |  | 5894.685 | 5960.302 | 1796.7 | 1816.7 |  |  |  | E14 (FB S2; WIR) |  |
| perf-e14-2 | perforation | E14 (FB S2; WIR) perforation 1824.7 m - 1832.7 m |  | 5986.549 | 6012.795 | 1824.7 | 1832.7 |  |  |  | E14 (FB S2; WIR) |  |
| perf-e34-36-1 | perforation | E34/36 (FB S2; WIR) perforation 2012.7 m - 2042.7 m |  | 6603.347 | 6701.772 | 2012.7 | 2042.7 |  |  |  | E34/36 (FB S2; WIR) |  |
| perf-e34-36-2 | perforation | E34/36 (FB S2; WIR) perforation 2052.7 m - 2067.7 m |  | 6734.58 | 6783.793 | 2052.7 | 2067.7 |  |  |  | E34/36 (FB S2; WIR) |  |
| perf-e40-42 | perforation | E40/42 perforation 2102.7 m - 2120.7 m |  | 6898.622 | 6957.677 | 2102.7 | 2120.7 |  |  |  | E40/42 |  |

Cell-by-cell traceability: `Dulang_B17_traceability.csv` (every populated JSON path, source cell/object, original value, formula + cached value, unit, method).

## Approximate type mappings and assumptions

| id | type | actualType | note |
|---|---|---|---|
| glm-1 | spm | 4-1/2" KBMG GLM (DUMMY) | type: 'KBMG GLM' = gas-lift mandrel (side-pocket) -> SPM (Gas Lift) is the nearest type |
| glm-2 | spm | 4-1/2" KBMG GLM (GLOV 14/64") | type: 'KBMG GLM' = gas-lift mandrel (side-pocket) -> SPM (Gas Lift) is the nearest type |
| sand-guard | sand-screen | 3-1/2" MULTILIFT 550 SAND GUARD | type: ESP intake sand guard (filter) -> Sand Screen is the nearest type |
| ss-stub | tubing | SS STUB | type: tubing stub = top of the tubing left in hole -> Tubing (top = stub depth; bottom not printed) |
| ls-stub | tubing | LS STUB | type: tubing stub = top of the tubing left in hole -> Tubing (top = stub depth; bottom not printed) |
| liner-7 | liner |  | innerDiameter: size 7" in label T61 has no ID/OD designation -> recorded as ID (user-defined mapping rule) |
| casing-9-5-8 | casing |  | innerDiameter: size 9-5/8" in label T64 has no ID/OD designation -> recorded as ID (user-defined mapping rule) |
| blank-pipe-1 | tubing | 4½" L-80 LCT BLANK PIPE | type: blank (unperforated) pipe in the screen assembly -> Tubing is the nearest type |
| screen-1 | screen |  | innerDiameter: size 4" in label T71 has no ID/OD designation -> recorded as ID (user-defined mapping rule) |
| seal-bore-assembly-1 | seal-assembly |  | innerDiameter: size 3.25" in label T74 has no ID/OD designation -> recorded as ID (user-defined mapping rule) |
| blank-pipe-2 | tubing | 4½" L-80 LCT BLANK PIPE | type: blank (unperforated) pipe in the screen assembly -> Tubing is the nearest type |
| screen-2 | screen |  | innerDiameter: size 4" in label T83 has no ID/OD designation -> recorded as ID (user-defined mapping rule) |
| seal-bore-assembly-2 | seal-assembly |  | innerDiameter: size 3.25" in label T85 has no ID/OD designation -> recorded as ID (user-defined mapping rule) |
| conical-shoe | mule-shoe | CONICAL SHOE | type: conical (guide) shoe at the tubing end -> Mule Shoe is the nearest type |

## Items requiring review

### Conflict (1)

| location | detail |
|---|---|
| rows 4-14 (hidden) | Hidden header rows 4-14 (B4-B14, casing & tubing table N4:Y11, fluids N13/N14) are identical in B05, B17 and B18 and name another well: 'Field / Platform : Erb West / EWDP-A Platform', 'Well ID / SLOT NO : EW-119 / Slot #9', spud/completion 'XX XX 2017', rig 'XXX'. Casing depths there (e.g. 7" liner 7150-9959 ft-MDTHF, max deviation @ 10092 ft) exceed the well's TD. Treated as a copied template block: NOT mapped (wellInfo fields left empty, no casing components from it). |

### Missing (14)

| location | detail |
|---|---|
| T20, T24, T41, T46 | Flow couplings and crossovers have no depth printed. |
| T68, T71, T74, T82, T83, T85 | Blank pipe / screens / seal bore assemblies have no depth printed. |
| casing-9-5-8 top | Casing top not printed -> empty (not assumed 0). |
| TD | No TD label in B17 (PBTD 2171.7 m only) -> totalDepth empty. |
| T20 | flow-coupling-1: no depth printed for 'Flow Coupling' -> topMd/bottomMd empty |
| T24 | flow-coupling-2: no depth printed for 'Flow Coupling' -> topMd/bottomMd empty |
| T41+T42 | crossover-1: no depth printed for 'CROSSOVER SUB 4-1/2" JFE-BEAR BOX X 3-1/2" VAMTOP PIN' -> topMd/bottomMd empty |
| T46+T47 | crossover-2: no depth printed for 'CROSSOVER SUB 3-1/2" JFE-BEAR PIN X 4-1/2" VAMTOP BOX' -> topMd/bottomMd empty |
| T68 | blank-pipe-1: no depth printed for '4½" L-80 LCT BLANK PIPE' -> topMd/bottomMd empty |
| T71 | screen-1: no depth printed for '4" P110 NU WWSCREEN 8 ga' -> topMd/bottomMd empty |
| T74 | seal-bore-assembly-1: no depth printed for '3.25" SEAL BORE ASSY.' -> topMd/bottomMd empty |
| T82 | blank-pipe-2: no depth printed for '4½" L-80 LCT BLANK PIPE' -> topMd/bottomMd empty |
| T83 | screen-2: no depth printed for '4" J-55 NU WW SCREEN 8 ga' -> topMd/bottomMd empty |
| T85 | seal-bore-assembly-2: no depth printed for '3.25" SEAL BORE ASSY.' -> topMd/bottomMd empty |

### Unclear (6)

| location | detail |
|---|---|
| J16 / Q16 | All depths are m-MDTHF (from the tubing hanger flange), not drill floor / RKB. Converted to ft only; the reference was not shifted (the DF-to-THF offset '46 ft' exists only in the hidden copied block). Consumers expecting MD-RKB must add the offset. |
| column Q / J 'DEPTH' | The sheet does not say whether equipment depths are to the top of the item; mapped to topMd. Item lengths (column S, m) are kept in 'interval' and not used to derive bottoms. |
| D54 / T54 | 'SS STUB' / 'LS STUB' read as short / long string tubing stubs. |
| D55 | '9-5/8" DUAL HYD. PKR.' drawn in the left column only; recorded once with string 'Short & Long'. |
| sheet 'MLE Space out' | ESP space-out item list (lengths, no depths, no well name; identical in B18) — not mapped (well association not stated). |
| perforation datum | Perforation texts state 'm MDTHF' (same reference as the depth columns). |

## Unmapped information

| location | content | reason |
|---|---|---|
| rows 4-14 (hidden) | Erb West / EW-119 header, casing & tubing table, completion/packer fluid, max deviation | copied block from another well — see Conflict |
| K84 | 19.431 (no label, no unit) | meaning unknown |
| T94-T95 | PRAVIN NAIR (JUNE 2024) | document control (updated by / date) |
| T39 / Q39 | 4-1/2" IPV1 at 841.66 m-MDTHF (2761.352 ft) | UNRESOLVED TYPE: IPV1 valve — no suitable Component_Type |
| D60-D62 | HYDRO SUB WITH PERFORATED PUP BLANKING PLUG | UNRESOLVED: combined item (hydro sub + perforated pup + blanking plug), no depth printed |
| sheet 'MLE Space out' | ESP space-out: 15 numbered items (packer, IPV1, crossovers, sand guard, pump 538-5800, gas handler, gas separators, seals, motor, sensor ACE Xtreme, motor guide), total length F62 = 70.86 m (formula =SUM(F14:F61)), pump setting depth to packer F64 = 55.46 m (=SUM(F16:F49)) | no depths / no well name; lengths only |
| B3 | SINGLE OIL PRODUCER | used for completionType |
| images | Picture 57 equipment graphic, Graphic 754 tubing graphic, MLE sheet ESP sketch + 'UAL — An Uzma Company' logo | images without data |

## Quality checks

| check | result |
|---|---|
| valid JSON (re-parsed) | PASS |
| top-level sections and order = template | PASS |
| wellInfo keys and order = template | PASS |
| wellInfo types = template type or empty value | PASS |
| component keys = template keys + supplementary (topMdM, bottomMdM, actualType, string) | PASS |
| numeric fields numeric or null | PASS |
| every type is a Component_Types Type_Key | PASS |
| component ids unique | PASS |
| no template sample records (surveys, components, formationTops) | PASS |
| every depth ft = m x 3.28084 (3 dp) | PASS |
| cell depths re-read from the working copy match | PASS |
| JSON depths = conversion table; metres kept | PASS |
| closest-type mappings flagged with estimated / estimatedFields 'type' / actualType | PASS |
| cell-text names equal the source cells | PASS |
| no depth placeholder / default 100 ft | PASS |
| well identity: B2 label matches the output name | PASS |
| sources, template and component reference unchanged (SHA-256) | PASS |

## Source / template integrity (SHA-256)

| file | sha256 before | sha256 after | unchanged |
|---|---|---|---|
| /Users/farahyusof/Documents/GitHub/NoramiraSazuki/AinnyFarah/WELL_SCHEMATIC/well-schematic-template-json.rtf | c7db15efb1917dfa76fa3102e270fe17115a8561f24eb0a155239870d079e712 | c7db15efb1917dfa76fa3102e270fe17115a8561f24eb0a155239870d079e712 | YES |
| /Users/farahyusof/Documents/GitHub/NoramiraSazuki/AinnyFarah/WELL_SCHEMATIC/well-schematic-sample-template.xlsx | d377a6310d42fda8c95e85db03f2416cb4c1b6f37f53920052331ce814f6d494 | d377a6310d42fda8c95e85db03f2416cb4c1b6f37f53920052331ce814f6d494 | YES |
| /Volumes/homes/public/DulangOSS/20261016/DULANG_WELLS/Well_Schematic/1. Well Schematic/Dulang B/Dulang B05 ESP Workover Schematic.xlsx | baf2892b31bb98223d0867e5fbc710e4d980a6768426644945af9c2ebf6c5fe9 | baf2892b31bb98223d0867e5fbc710e4d980a6768426644945af9c2ebf6c5fe9 | YES |
| /Volumes/homes/public/DulangOSS/20261016/DULANG_WELLS/Well_Schematic/1. Well Schematic/Dulang B/Dulang B17 ESP Workover Schematic.xlsx | 46549fe8a9359f2b5779f0694c6540430f4545a23c0c188799337d0f8bdcc986 | 46549fe8a9359f2b5779f0694c6540430f4545a23c0c188799337d0f8bdcc986 | YES |
| /Volumes/homes/public/DulangOSS/20261016/DULANG_WELLS/Well_Schematic/1. Well Schematic/Dulang B/Dulang B18 ESP Workover Schematic Updated .xlsx | 9528f1c7d9a3cd37c846d34d7b3ddce25cd7ad798ad2da88be08da646d9b27e9 | 9528f1c7d9a3cd37c846d34d7b3ddce25cd7ad798ad2da88be08da646d9b27e9 | YES |
