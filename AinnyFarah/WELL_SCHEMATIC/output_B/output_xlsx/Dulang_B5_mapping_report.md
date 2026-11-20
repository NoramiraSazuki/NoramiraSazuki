# Dulang_B5 — mapping and quality-check report

- Source workbook: `/Volumes/homes/public/DulangOSS/20261016/DULANG_WELLS/Well_Schematic/1. Well Schematic/Dulang B/Dulang B05 ESP Workover Schematic.xlsx` (read-only; worked on a copy in `_work/src/`)
- Internal well label: `COMPLETION SCHEMATIC : DULANG B-05` (B2) — agrees with the file name; output name `Dulang_B5` (user-approved)
- Worksheet mapped: `B-05 v1`; worksheets inspected: 'B-05 v1' (visible)
- Reading: cells + formulas + cached values + drawing layer (shapes / text boxes / pictures) from the .xlsx XML; rendered with macOS Quick Look (`_work/render/ql/`) and visually compared; embedded images extracted to `_work/media/B5/` and OCR'd where they hold text
- Template: `/Users/farahyusof/Documents/GitHub/NoramiraSazuki/AinnyFarah/WELL_SCHEMATIC/well-schematic-template-json.rtf`; Type_Keys from `well-schematic-sample-template.xlsx` > Component_Types
- Output JSON: `Dulang_B5.json` — 40 components; surveys / cement / plugBackCement / centralizers / floatEquipment / formations / formationTops / fluidContacts left empty (no supported source data)
- Generated: 2026-10-29 10:30

## Depth units, reference and conversion

- Source depths are **m-MDTHF (MD from tubing hanger flange; column headers J16/Q16)**. JSON depths are **ft = m × 3.28084**, 3 decimals; original metres kept in `topMdM` / `bottomMdM`.
- The MD reference (THF) was not shifted to drill floor/RKB and no MD→TVD conversion was made.

| field | ref | m | unit | ft |
|---|---|---|---|---|
| flow-coupling-1.topMd | Q20 | 222.61 | m-MDTHF (column header Q16) | 730.348 |
| trsv.topMd | Q22 | 224.03 | m-MDTHF (column header Q16) | 735.007 |
| flow-coupling-2.topMd | Q24 | 225.57 | m-MDTHF (column header Q16) | 740.059 |
| glm-1.topMd | Q26 | 614.85 | m-MDTHF (column header Q16) | 2017.224 |
| glm-2.topMd | Q29 | 1517.63 | m-MDTHF (column header Q16) | 4979.101 |
| glm-3.topMd | Q32 | 1768.5 | m-MDTHF (column header Q16) | 5802.166 |
| ssd-1.topMd | Q34 | 1776.81 | m-MDTHF (column header Q16) | 5829.429 |
| esp-packer.topMd | Q36 | 1808.02 | m-MDTHF (column header Q16) | 5931.824 |
| ssd-2.topMd | Q39 | 1814.1 | m-MDTHF (column header Q16) | 5951.772 |
| crossover-1.topMd | Q43 | 1824.68 | m-MDTHF (column header Q16) | 5986.483 |
| sand-guard.topMd | Q46 | 1828.99 | m-MDTHF (column header Q16) | 6000.624 |
| crossover-2.topMd | Q48 | 1831.32 | m-MDTHF (column header Q16) | 6008.268 |
| esp.topMd | Q52 | 1858.07 | m-MDTHF (column header Q16) | 6096.03 |
| tubing-stub-left.topMd | J56 | 1952.7 | m-MDTHF (column header J16) | 6406.496 |
| tubing-stub-right.topMd | Q56 | 1951.7 | m-MDTHF (column header Q16) | 6403.215 |
| guib-packer.topMd | Q58 | 1953.7 | m-MDTHF (column header Q16) | 6409.777 |
| bridge-plug-1.topMd | Rectangle 605 | 2037.0 | m-MDTHF (in drawing text) | 6683.071 |
| dual-hyd-packer.topMd | J68 | 2047.2 | m-MDTHF (column header J16) | 6716.536 |
| bridge-plug-2.topMd | Rectangle 606 | 2050.0 | m-MDTHF (in drawing text) | 6725.722 |
| no-go-nipple-left.topMd | J72 | 2051.1 | m-MDTHF (column header J16) | 6729.331 |
| wl-entry-guide-left.topMd | J73 | 2051.8 | m-MDTHF (column header J16) | 6731.628 |
| ssd-3.topMd | Q77 | 2061.4 | m-MDTHF (column header Q16) | 6763.124 |
| blast-joint.topMd | Q79 | 2065.1 | m-MDTHF (column header Q16) | 6775.263 |
| wp1-nipple.topMd | Q82 | 2159.0 | m-MDTHF (column header Q16) | 7083.334 |
| locator-seal-assembly.topMd | Q86 | 2165.5 | m-MDTHF (column header Q16) | 7104.659 |
| retainer-packer.topMd | J87 | 2166.2 | m-MDTHF (column header J16) | 7106.956 |
| no-go-nipple-right.topMd | Q90 | 2173.0 | m-MDTHF (column header Q16) | 7129.265 |
| wl-entry-guide-right.topMd | Q92 | 2173.9 | m-MDTHF (column header Q16) | 7132.218 |
| perm-packer.topMd | J91 | 2198.2 | m-MDTHF (column header J16) | 7211.942 |
| perf-e7.topMd | C61 | 1974.74 | m (in text; datum not stated in the label, sheet depths are m-MDTHF) | 6478.806 |
| perf-e7.bottomMd | C61 | 1994.24 | m (in text; datum not stated in the label, sheet depths are m-MDTHF) | 6542.782 |
| perf-e12-13-1.topMd | C77 | 2067.24 | m (in text; datum not stated in the label, sheet depths are m-MDTHF) | 6782.284 |
| perf-e12-13-1.bottomMd | C77 | 2075.74 | m (in text; datum not stated in the label, sheet depths are m-MDTHF) | 6810.171 |
| perf-e12-13-2.topMd | C77 | 2077.74 | m (in text; datum not stated in the label, sheet depths are m-MDTHF) | 6816.733 |
| perf-e12-13-2.bottomMd | C77 | 2083.74 | m (in text; datum not stated in the label, sheet depths are m-MDTHF) | 6836.418 |
| perf-e14-1.topMd | C77 | 2087.74 | m (in text; datum not stated in the label, sheet depths are m-MDTHF) | 6849.541 |
| perf-e14-1.bottomMd | C77 | 2103.74 | m (in text; datum not stated in the label, sheet depths are m-MDTHF) | 6902.034 |
| perf-e14-2.topMd | C77 | 2106.24 | m (in text; datum not stated in the label, sheet depths are m-MDTHF) | 6910.236 |
| perf-e14-2.bottomMd | C77 | 2135.74 | m (in text; datum not stated in the label, sheet depths are m-MDTHF) | 7007.021 |
| perf-c89.topMd | C89 | 2180.24 | m (in text; datum not stated in the label, sheet depths are m-MDTHF) | 7153.019 |
| perf-c89.bottomMd | C89 | 2186.24 | m (in text; datum not stated in the label, sheet depths are m-MDTHF) | 7172.704 |
| perf-e20-32-1.topMd | C94 | 2215.24 | m (in text; datum not stated in the label, sheet depths are m-MDTHF) | 7267.848 |
| perf-e20-32-1.bottomMd | C94 | 2217.24 | m (in text; datum not stated in the label, sheet depths are m-MDTHF) | 7274.41 |
| perf-e20-32-2.topMd | C94 | 2223.74 | m (in text; datum not stated in the label, sheet depths are m-MDTHF) | 7295.735 |
| perf-e20-32-2.bottomMd | C94 | 2228.12 | m (in text; datum not stated in the label, sheet depths are m-MDTHF) | 7310.105 |
| perf-e20-32-3.topMd | C94 | 2243.24 | m (in text; datum not stated in the label, sheet depths are m-MDTHF) | 7359.712 |
| perf-e20-32-3.bottomMd | C94 | 2249.24 | m (in text; datum not stated in the label, sheet depths are m-MDTHF) | 7379.397 |
| perf-e20-32-4.topMd | C94 | 2258.24 | m (in text; datum not stated in the label, sheet depths are m-MDTHF) | 7408.924 |
| perf-e20-32-4.bottomMd | C94 | 2275.24 | m (in text; datum not stated in the label, sheet depths are m-MDTHF) | 7464.698 |
| perf-e20-32-5.topMd | C94 | 2276.24 | m (in text; datum not stated in the label, sheet depths are m-MDTHF) | 7467.979 |
| perf-e20-32-5.bottomMd | C94 | 2279.74 | m (in text; datum not stated in the label, sheet depths are m-MDTHF) | 7479.462 |
| wellInfo.totalDepth | L99 | 2424.8 | m-MDTHF (in text) | 7955.381 |

## wellInfo

| json_path | mapped_value | cell_or_object | original_value | method | note |
|---|---|---|---|---|---|
| wellInfo.completionDate | - |  |  | EMPTY | not supported by the visible source -> template empty value ('-' text / null number) |
| wellInfo.completionFluid | - |  |  | EMPTY | not supported by the visible source -> template empty value ('-' text / null number) |
| wellInfo.completionType | SINGLE OIL PRODUCER | B3 | ' SINGLE OIL PRODUCER ' | DIRECT | visible title row 3 ('SINGLE OIL PRODUCER'); ESP completion |
| wellInfo.groundElevation |  |  |  | EMPTY | not supported by the visible source -> template empty value ('-' text / null number) |
| wellInfo.kbElevation |  |  |  | EMPTY | not supported by the visible source -> template empty value ('-' text / null number) |
| wellInfo.maxDeviation | - |  |  | EMPTY | not supported by the visible source -> template empty value ('-' text / null number) |
| wellInfo.name | Dulang_B5 | B2 | 'COMPLETION SCHEMATIC : DULANG B-05' | STANDARDISED NAME | internal label B2 'COMPLETION SCHEMATIC : DULANG B-05' -> Dulang_B5 (user-approved naming) |
| wellInfo.packerFluid | - |  |  | EMPTY | not supported by the visible source -> template empty value ('-' text / null number) |
| wellInfo.rig | - |  |  | EMPTY | not supported by the visible source -> template empty value ('-' text / null number) |
| wellInfo.surfaceCasing | - |  |  | EMPTY | not supported by the visible source -> template empty value ('-' text / null number) |
| wellInfo.totalDepth | 7955.381 | L99 | 'TD @ 2424.8 m MDTHF' | CONVERTED | diagram label 'TD @ 2424.8 m MDTHF': 2424.8 m x 3.28084; datum MDTHF kept |
| wellInfo.trajectoryType | - |  |  | EMPTY | not supported by the visible source -> template empty value ('-' text / null number) |
| wellInfo.tubing | - |  |  | EMPTY | not supported by the visible source -> template empty value ('-' text / null number) |
| wellInfo.wellhead | - |  |  | EMPTY | not supported by the visible source -> template empty value ('-' text / null number) |

## Components

| id | type | name | string | topMd_ft | bottomMd_ft | topMd_m | bottomMd_m | OD_in | ID_in | interval | zone | estimated |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| flow-coupling-1 | flow-coupling | Flow Coupling |  | 730.348 |  | 222.61 |  | 5.09 | 3.884 | 1.412 m |  |  |
| trsv | trsv | 4-1/2"TRSV |  | 735.007 |  | 224.03 |  | 6.923 | 3.812 | 1.527 m |  |  |
| flow-coupling-2 | flow-coupling | Flow Coupling |  | 740.059 |  | 225.57 |  | 5.09 | 3.884 | 1.412 m |  |  |
| glm-1 | spm | 4-1/2" KBMG GLM (IPO 8/64, PTRO 1024 psi) |  | 2017.224 |  | 614.85 |  | 6.393 | 3.865 | 2.816 m |  | type |
| glm-2 | spm | 4-1/2" KBMG GLM (Dummy) |  | 4979.101 |  | 1517.63 |  | 6.393 | 3.865 | 2.816 m |  | type |
| glm-3 | spm | 4-1/2" KBMG GLM (GLOV 12/64) |  | 5802.166 |  | 1768.5 |  | 6.393 | 3.865 | 2.816 m |  | type |
| ssd-1 | sliding-sleeve | 4-1/2" SSD (CLOSED) |  | 5829.429 |  | 1776.81 |  | 5.613 | 3.813 | 1.423 m |  |  |
| esp-packer | packer | 9-5/8" x 4-1/2" ESP PACKER |  | 5931.824 |  | 1808.02 |  | 8.835 | 3.96 | 1.29 m |  |  |
| ssd-2 | sliding-sleeve | 4-1/2" SSD (OPEN) |  | 5951.772 |  | 1814.1 |  | 5.613 | 3.813 | 1.423 m |  |  |
| crossover-1 | crossover | CROSSOVER SUB 4-1/2" JFE-BEAR BOX X 3-1/2" VAMTOP PIN |  | 5986.483 |  | 1824.68 |  | 5.0 | 2.959 | 0.31 m |  |  |
| sand-guard | sand-screen | 3-1/2" MULTILIFT 550 SAND GUARD |  | 6000.624 |  | 1828.99 |  | 5.5 |  | 6.33 m |  | type |
| crossover-2 | crossover | CROSSOVER SUB 3-1/2" JFE-BEAR PIN X 4-1/2" VAMTOP BOX |  | 6008.268 |  | 1831.32 |  | 5.0 | 2.959 | 0.28 m |  |  |
| esp | esp | 4-1/2" ESP |  | 6096.03 |  | 1858.07 |  | 6.132 |  | 20.35 m |  |  |
| tubing-stub-left | tubing | TUBING STUB |  | 6406.496 |  | 1952.7 |  |  |  |  |  | type |
| tubing-stub-right | tubing | TUBING STUB |  | 6403.215 |  | 1951.7 |  |  |  |  |  | type |
| guib-packer | packer | GUIB Packer |  | 6409.777 |  | 1953.7 |  |  |  |  |  |  |
| bridge-plug-1 | bridge-plug | Permament Bridge plug @ 2037m-MDTHF |  | 6683.071 |  | 2037.0 |  |  |  |  |  |  |
| dual-hyd-packer | packer | BAKER 'T2-DSR' DUAL HYD. PKR |  | 6716.536 |  | 2047.2 |  |  |  |  |  |  |
| bridge-plug-2 | bridge-plug | Permament Bridge plug @ 2050m-MDTHF |  | 6725.722 |  | 2050.0 |  |  |  |  |  |  |
| no-go-nipple-left | nipple (no-go) | D' NO-GO NIPPLE |  | 6729.331 |  | 2051.1 |  |  |  |  |  |  |
| wl-entry-guide-left | mule-shoe | W/L ENTRY GUIDE |  | 6731.628 |  | 2051.8 |  |  |  |  |  | type |
| ssd-3 | sliding-sleeve | WPB-2DE' SSD |  | 6763.124 |  | 2061.4 |  |  | 2.312 |  |  |  |
| blast-joint | blast-joint | BLAST JOINT |  | 6775.263 |  | 2065.1 |  |  | 2.441 |  |  |  |
| wp1-nipple | landing-nipple | WP-1 NIPPLE |  | 7083.334 |  | 2159.0 |  |  | 2.312 |  |  | type |
| locator-seal-assembly | seal-assembly | LOC. SEAL ASSY. |  | 7104.659 |  | 2165.5 |  |  | 2.441 |  |  |  |
| retainer-packer | packer | BAKER 'D' RETAINER. PKR |  | 7106.956 |  | 2166.2 |  |  |  |  |  |  |
| no-go-nipple-right | nipple (no-go) | D' NO-GO NIPPLE |  | 7129.265 |  | 2173.0 |  |  | 2.25 |  |  |  |
| wl-entry-guide-right | mule-shoe | W/L ENTRY GUIDE |  | 7132.218 |  | 2173.9 |  |  | 2.441 |  |  | type |
| perm-packer | packer | GUIB MAGNUM HT' PERM. PKR |  | 7211.942 |  | 2198.2 |  |  |  |  |  |  |
| perf-e7 | perforation | E-7 (FB S3, WIR) perforation 1974.74 m - 1994.24 m |  | 6478.806 | 6542.782 | 1974.74 | 1994.24 |  |  |  | E-7 (FB S3, WIR) |  |
| perf-e12-13-1 | perforation | E-12/13 (FB S3, WIR) perforation 2067.24 m - 2075.74 m |  | 6782.284 | 6810.171 | 2067.24 | 2075.74 |  |  |  | E-12/13 (FB S3, WIR) |  |
| perf-e12-13-2 | perforation | E-12/13 (FB S3, WIR) perforation 2077.74 m - 2083.74 m |  | 6816.733 | 6836.418 | 2077.74 | 2083.74 |  |  |  | E-12/13 (FB S3, WIR) |  |
| perf-e14-1 | perforation | E-14 (FB S3, WIR) perforation 2087.74 m - 2103.74 m |  | 6849.541 | 6902.034 | 2087.74 | 2103.74 |  |  |  | E-14 (FB S3, WIR) |  |
| perf-e14-2 | perforation | E-14 (FB S3, WIR) perforation 2106.24 m - 2135.74 m |  | 6910.236 | 7007.021 | 2106.24 | 2135.74 |  |  |  | E-14 (FB S3, WIR) |  |
| perf-c89 | perforation | Perforation 2180.24 m - 2186.24 m |  | 7153.019 | 7172.704 | 2180.24 | 2186.24 |  |  |  |  |  |
| perf-e20-32-1 | perforation | E-20/22 & E-23/32 (FB S3, NDR) perforation 2215.24 m - 2217.24 m |  | 7267.848 | 7274.41 | 2215.24 | 2217.24 |  |  |  | E-20/22 & E-23/32 (FB S3, NDR) |  |
| perf-e20-32-2 | perforation | E-20/22 & E-23/32 (FB S3, NDR) perforation 2223.74 m - 2228.12 m |  | 7295.735 | 7310.105 | 2223.74 | 2228.12 |  |  |  | E-20/22 & E-23/32 (FB S3, NDR) |  |
| perf-e20-32-3 | perforation | E-20/22 & E-23/32 (FB S3, NDR) perforation 2243.24 m - 2249.24 m |  | 7359.712 | 7379.397 | 2243.24 | 2249.24 |  |  |  | E-20/22 & E-23/32 (FB S3, NDR) |  |
| perf-e20-32-4 | perforation | E-20/22 & E-23/32 (FB S3, NDR) perforation 2258.24 m - 2275.24 m |  | 7408.924 | 7464.698 | 2258.24 | 2275.24 |  |  |  | E-20/22 & E-23/32 (FB S3, NDR) |  |
| perf-e20-32-5 | perforation | E-20/22 & E-23/32 (FB S3, NDR) perforation 2276.24 m - 2279.74 m |  | 7467.979 | 7479.462 | 2276.24 | 2279.74 |  |  |  | E-20/22 & E-23/32 (FB S3, NDR) |  |

Cell-by-cell traceability: `Dulang_B5_traceability.csv` (every populated JSON path, source cell/object, original value, formula + cached value, unit, method).

## Approximate type mappings and assumptions

| id | type | actualType | note |
|---|---|---|---|
| glm-1 | spm | 4-1/2" KBMG GLM (IPO 8/64, PTRO 1024 psi) | type: 'KBMG GLM' = gas-lift mandrel (side-pocket) -> SPM (Gas Lift) is the nearest type |
| glm-2 | spm | 4-1/2" KBMG GLM (Dummy) | type: 'KBMG GLM' = gas-lift mandrel (side-pocket) -> SPM (Gas Lift) is the nearest type |
| glm-3 | spm | 4-1/2" KBMG GLM (GLOV 12/64) | type: 'KBMG GLM' = gas-lift mandrel (side-pocket) -> SPM (Gas Lift) is the nearest type |
| sand-guard | sand-screen | 3-1/2" MULTILIFT 550 SAND GUARD | type: ESP intake sand guard (filter) -> Sand Screen is the nearest type |
| tubing-stub-left | tubing | TUBING STUB | type: tubing stub = top of the tubing left in hole -> Tubing (top = stub depth; bottom not printed) |
| tubing-stub-right | tubing | TUBING STUB | type: tubing stub = top of the tubing left in hole -> Tubing (top = stub depth; bottom not printed) |
| wl-entry-guide-left | mule-shoe | W/L ENTRY GUIDE | type: wireline entry guide (WEG) at the tubing end -> Mule Shoe is the nearest type |
| wp1-nipple | landing-nipple | WP-1 NIPPLE | type: 'WP-1 NIPPLE' (nipple profile, no no-go stated) -> Landing Nipple is the nearest type |
| wl-entry-guide-right | mule-shoe | W/L ENTRY GUIDE | type: wireline entry guide (WEG) at the tubing end -> Mule Shoe is the nearest type |

## Items requiring review

### Conflict (1)

| location | detail |
|---|---|
| rows 4-14 (hidden) | Hidden header rows 4-14 (B4-B14, casing & tubing table N4:Y11, fluids N13/N14) are identical in B05, B17 and B18 and name another well: 'Field / Platform : Erb West / EWDP-A Platform', 'Well ID / SLOT NO : EW-119 / Slot #9', spud/completion 'XX XX 2017', rig 'XXX'. Casing depths there (e.g. 7" liner 7150-9959 ft-MDTHF, max deviation @ 10092 ft) exceed the well's TD. Treated as a copied template block: NOT mapped (wellInfo fields left empty, no casing components from it). |

### Missing (2)

| location | detail |
|---|---|
| TD only | No casing, liner or tubing intervals are printed in the visible schematic; no tubing component was created (no tubing label/size/interval). PBTD not printed. |
| wellInfo | Completion date, rig, fluids, elevations, wellhead not in the visible sheet. |

### Unclear (8)

| location | detail |
|---|---|
| J16 / Q16 | All depths are m-MDTHF (from the tubing hanger flange), not drill floor / RKB. Converted to ft only; the reference was not shifted (the DF-to-THF offset '46 ft' exists only in the hidden copied block). Consumers expecting MD-RKB must add the offset. |
| column Q / J 'DEPTH' | The sheet does not say whether equipment depths are to the top of the item; mapped to topMd. Item lengths (column S, m) are kept in 'interval' and not used to derive bottoms. |
| D56/J56 vs T56/Q56 | Two 'TUBING STUB' labels (left 1952.7 m, right 1951.7 m) — read as two strings of the old dual completion; the sheet does not label which is short/long, so 'string' is not set. |
| Rectangle 605 / 606 | Bridge plugs placed from drawing text boxes; string association by position only (605 left column, 606 right column). |
| C61 | 'PERFT' (spelling as printed) read as perforation. |
| C89 | Perforation 2180.24-2186.24 m has no sand label -> zone not set. |
| C94 | Five intervals listed under 'E-20/22 & E-23/32' without per-sand split -> each carries the combined zone label. |
| perforation datum | Perforation texts give 'm' without a datum; sheet depths are m-MDTHF (assumed same reference). |

### Anomaly (1)

| location | detail |
|---|---|
| J72/J73 vs J68 | Left-column depths: dual packer 2047.2 m, no-go 2051.1 m, entry guide 2051.8 m below bridge plug 2037 m — kept as printed. |

## Unmapped information

| location | content | reason |
|---|---|---|
| rows 4-14 (hidden) | Erb West / EW-119 header, casing & tubing table, completion/packer fluid, max deviation | copied block from another well — see Conflict |
| K88 | 19.431 (no label, no unit) | meaning unknown |
| T98-T99 | PRAVIN NAIR (AUGUST 2024) | document control (updated by / date) |
| T41 / Q41 / X41 / Y41 | 4-1/2" IPV1 at 1818.77 m-MDTHF (5967.093 ft), length 2.036 m, ID 2.813 in, OD 5.75 in | UNRESOLVED TYPE: IPV1 valve — no suitable Component_Type |
| T84 / Q84 / X84 | CENTRALIZER at 2164 m-MDTHF (7099.738 ft), ID 2.441 in | tubing centralizer: no Component_Type; template 'centralizers' are per casing (Casing_ID, MD) -> not created |
| B3 | SINGLE OIL PRODUCER | used for completionType |
| Picture 57 (L35:P37) | small equipment graphic (xl/media/image1.png) | image without data |

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
