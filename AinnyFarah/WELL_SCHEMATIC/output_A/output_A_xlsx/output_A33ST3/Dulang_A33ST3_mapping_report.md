# Dulang_A33ST3 — mapping and quality-check report

- Source workbook: `/Volumes/homes/public/DulangOSS/20261016/DULANG_WELLS/Well_Schematic/1. Well Schematic/Dulang A/Dulang A33ST3 Well Schematic.xls` (read-only; worked on a copy in `_work/src/`)
- Worksheet mapped: `SCHEM-1`
- Template: `/Users/farahyusof/Documents/GitHub/NoramiraSazuki/AinnyFarah/WELL_SCHEMATIC/well-schematic-template-json.rtf` (read-only); component Type_Keys from `well-schematic-sample-template.xlsx` > Component_Types
- Output JSON: `Dulang_A33ST3.json` — 26 components; surveys / cement / plugBackCement / centralizers / floatEquipment / formations / formationTops / fluidContacts left empty (no supported source data)
- Generated: 2026-10-27 10:54

## Depth units, reference and conversion

Depths: casing scheme S4 'DEPTH (m-MDDF)' (TOP/BOTTOM); equipment depths in K:L under L15/L16 'DEPTH m-MDDF'; diagram depths F40/F74/F75 under F15/F16 'DEPTH m-MDDF'.

- Every depth in the JSON is in **ft = m / 0.3048**, rounded to 3 decimals. Metres, unit evidence and cell are listed below.
- MD reference kept as printed (MDDF). No MD↔TVD or datum conversion was made. TVD values have no component field and are listed as unmapped.

| field | cell | m | unit | ft | kind |
|---|---|---|---|---|---|
| wellInfo.maxDeviation (depth in text) | B11 | 1800.0 | m MDDF (in text) | 5905.512 | direct |
| conductor.bottomMd | T6 | 181.46 | m-MDDF (column header S4 'DEPTH (m-MDDF)') | 595.341 | direct |
| casing-16.bottomMd | T7 | 188.46 | m-MDDF (column header S4 'DEPTH (m-MDDF)') | 618.307 | direct |
| casing-13-3-8.bottomMd | T8 | 443.26 | m-MDDF (column header S4 'DEPTH (m-MDDF)') | 1454.265 | direct |
| casing-9-5-8.bottomMd | T9 | 1355.46 | m-MDDF (column header S4 'DEPTH (m-MDDF)') | 4447.047 | direct |
| liner-7.topMd | F40 | 902.73 | m-MDDF (column header F16 'm-MDDF') | 2961.713 | direct |
| trsv.topMd | K25 | 254.61 | m-MDDF (column header L16 'm-MDDF') | 835.335 | direct |
| bluepack-packer.topMd | K52 | 1800.21 | m-MDDF (column header L16 'm-MDDF') | 5906.201 | direct |
| landing-nipple.topMd | K55 | 1803.95 | m-MDDF (column header L16 'm-MDDF') | 5918.471 | direct |
| muleshoe.topMd | K61 | 1836.02 | m-MDDF (column header L16 'm-MDDF') | 6023.688 | direct |
| pbtd.topMd | F74 | 1903.0 | m-MDDF (column header F16 'm-MDDF') | 6243.438 | direct |
| wellInfo.kbElevation | B9 | 30.1 | m-AMSL (elevation) | 98.753 | direct |
| wellInfo.totalDepth | F75 | 1944.6 | m-MDDF (F16) | 6379.921 | direct |

## wellInfo mapping

| json_path | mapped_value | cell | original_value | note |
|---|---|---|---|---|
| wellInfo.completionDate | 2020-04-03 | B10 | 'Completion Date : 3rd April 2020' | DIRECT: '3rd April 2020' written ISO |
| wellInfo.completionFluid | 8.5 ppg NaCl filtered to 30 NTU | I13 | 'Completion Fluid : 8.5 ppg NaCl filtered to 30 NTU' | DIRECT (label removed) |
| wellInfo.completionType | Single string | B12 | 'Wellhead : Surface Unihead, Solar Alert, Surface, Standard, Single' | CLOSEST: wellhead text '…Single' and one tubing string — review |
| wellInfo.groundElevation |  |  |  | NOT IN SOURCE -> template empty value; flagged |
| wellInfo.kbElevation | 98.753 | B9 | 'Drill floor height: 30.1m-AMSL' | CLOSEST FIELD: 'Drill floor height: 30.1m-AMSL' -> ft |
| wellInfo.maxDeviation | 58 deg @ 5905.512 ft MDDF | B11 | 'Maximum Deviation: 58deg @ 1,800m MDDF' | DIRECT text; depth 1,800 m MDDF -> ft |
| wellInfo.name | Dulang_A33ST3 | B2 | 'DULANG A33-ST3' | DIRECT: B2 'DULANG A33-ST3' (B4 says A33ST2 — see identity) |
| wellInfo.packerFluid | 8.5 ppg NaCl with 2% Corrosion Inhibitor + 2% Oxygen Scavenger + 0.2% Biocide | I14 | 'Packer Fluid: 8.5 ppg NaCl with 2% Corrosion Inhibitor + 2% Oxygen Scavenger + 0.2% Biocide' | DIRECT (label removed) |
| wellInfo.rig | Sapura Berani | B6 | 'Rig Name: Sapura Berani ' | DIRECT |
| wellInfo.surfaceCasing | - |  |  | NOT IN SOURCE -> template empty value; flagged |
| wellInfo.totalDepth | 6379.921 | F75 | 1944.6 | DIRECT: TD (label B75) 1944.6 m-MDDF (F16) -> ft |
| wellInfo.trajectoryType | Deviated | B11 | 'Maximum Deviation: 58deg @ 1,800m MDDF' | supported by maximum deviation 58 deg |
| wellInfo.tubing | Tubing 3-1/2" GRE lined 9.2 L80 JFEFOX; Tubing 3-1/2" - Below 7" AHC 9.2 15CR-125 JFE BEAR | I11 | 'Tubing 3-1/2" GRE lined' | DIRECT: header rows 11-12 joined |
| wellInfo.tubing | Tubing 3-1/2" GRE lined 9.2 L80 JFEFOX; Tubing 3-1/2" - Below 7" AHC 9.2 15CR-125 JFE BEAR | P11 | 9.2 | DIRECT: header rows 11-12 joined |
| wellInfo.tubing | Tubing 3-1/2" GRE lined 9.2 L80 JFEFOX; Tubing 3-1/2" - Below 7" AHC 9.2 15CR-125 JFE BEAR | Q11 | 'L80' | DIRECT: header rows 11-12 joined |
| wellInfo.tubing | Tubing 3-1/2" GRE lined 9.2 L80 JFEFOX; Tubing 3-1/2" - Below 7" AHC 9.2 15CR-125 JFE BEAR | R11 | 'JFEFOX' | DIRECT: header rows 11-12 joined |
| wellInfo.tubing | Tubing 3-1/2" GRE lined 9.2 L80 JFEFOX; Tubing 3-1/2" - Below 7" AHC 9.2 15CR-125 JFE BEAR | I12 | 'Tubing 3-1/2" - Below 7" AHC' | DIRECT: header rows 11-12 joined |
| wellInfo.tubing | Tubing 3-1/2" GRE lined 9.2 L80 JFEFOX; Tubing 3-1/2" - Below 7" AHC 9.2 15CR-125 JFE BEAR | P12 | 9.2 | DIRECT: header rows 11-12 joined |
| wellInfo.tubing | Tubing 3-1/2" GRE lined 9.2 L80 JFEFOX; Tubing 3-1/2" - Below 7" AHC 9.2 15CR-125 JFE BEAR | Q12 | '15CR-125' | DIRECT: header rows 11-12 joined |
| wellInfo.tubing | Tubing 3-1/2" GRE lined 9.2 L80 JFEFOX; Tubing 3-1/2" - Below 7" AHC 9.2 15CR-125 JFE BEAR | R12 | 'JFE BEAR' | DIRECT: header rows 11-12 joined |
| wellInfo.wellhead | Surface Unihead, Solar Alert, Surface, Standard, Single | B12 | 'Wellhead : Surface Unihead, Solar Alert, Surface, Standard, Single' | DIRECT (label removed) |

## Components

| id | type | name | topMd_ft | bottomMd_ft | OD_in | ID_in | weight | grade | interval | estimated |
|---|---|---|---|---|---|---|---|---|---|---|
| conductor | casing | 26" Conductor (Existing) |  | 595.341 | 26 |  | 202.0 | A-36 |  |  |
| casing-16 | casing | 16" Casing (Existing) |  | 618.307 | 16 |  | 75.0 | J-55 |  |  |
| casing-13-3-8 | casing | 13-3/8" Casing (Existing) |  | 1454.265 | 13.375 |  | 54.5 | K-55 |  |  |
| casing-9-5-8 | casing | 9-5/8" Casing (Existing) |  | 4447.047 | 9.625 |  | 43.5 | L-80 |  |  |
| liner-7 | liner | 7" Liner | 2961.713 |  | 7 |  |  |  |  |  |
| tubing-gre | tubing | TUBING 3-1/2" 9.2#, FOX GRE, R2 |  |  | 3.5 |  | 9.2 | L80 |  |  |
| tubing-15cr | tubing | TUBING 3-1/2" 9.2#, UHP 15Cr-125, JFE BEAR, R3 |  |  | 3.5 |  | 9.2 | 15CR-125 |  |  |
| xmas-tree | christmas-tree | X-Mas Tree : Solar Alert  5,000 psi |  |  |  |  |  |  |  |  |
| tubing-hanger | tubing-hanger | TUBING HANGER Bear (BOX Down) |  |  |  | 2.992 |  |  | 1.52 m |  |
| crossover-1 | crossover | XOVER 3-1/2" BEAR Pin x FOX DTP (BOX) |  |  | 3.833 | 2.545 |  |  | 0.46 m |  |
| pup-joint-12ft | pup-joint | 12 FT 3-1/2" GRE Lined Fox (Pin x Pin) Pups |  |  | 3.5 | 2.545 |  |  | 3.65742151782993 m |  |
| pup-joint-6ft-1 | pup-joint | 6 FT 3-1/2" GRE Lined Fox (Box x Pin) Pups |  |  | 3.878 | 2.545 |  |  | 1.8288 m |  |
| flow-coupling-1 | flow-coupling | OS, FLOW COUPLING, 3-1/2 9.2PPF JFEFOX DTP BOX X 2-7/8 6.4PPF JFEBEAR PIN, 25CR0[80] |  |  | 3.833 | 2.357 |  |  | 0.91 m |  |
| trsv | trsv | TRSV, 2-7/8 6.4PPF JFEBEAR BXP | 835.335 |  | 4.609 | 2.312 |  |  | 1.35 m |  |
| flow-coupling-2 | flow-coupling | OS, FLOW COUPLING, 2-7/8 6.4PPF JFEBEAR BOX X 3-1/2 9.2PPF JFEFOX DTP BOX |  |  | 3.833 | 2.357 |  |  | 0.91 m |  |
| pup-joint-4ft | pup-joint | 4 FT 3-1/2" GRE Lined Fox (Pin x Pin) Pups |  |  | 3.5 | 2.545 |  |  | 1.2191405059433098 m |  |
| pup-joint-6ft-2 | pup-joint | 6 FT 3-1/2" GRE Lined Fox (B x P) Pups |  |  | 3.878 | 2.545 |  |  | 1.83 m |  |
| crossover-2 | crossover | Crossover - 3-1/2" FOX DTP BOX x 3-1/2 Bear Pin, 25Cr-80KSI |  |  | 3.893 | 2.545 |  |  | 0.5 m |  |
| bluepack-packer | packer | 7 X 3-1/2 BLUEPACK MAX RH (29-32), 25CR (80), HNBR (90-90-90), 3-1/2 (9.2) JFE Bear BOX x 3-1/2 (9.2) JFE BEAR PIN | 5906.201 |  | 5.909 | 2.885 |  |  | 2.44 m |  |
| pup-joint-5ft | pup-joint | 5 FT 3 1/2" 9.2# JFE BEAR B x P, 25Cr Pups |  |  | 3.5 | 2.992 |  |  | 1.524 m |  |
| crossover-3 | crossover | Crossover - 3-1/2"  BEAR (B) x 2-7/8" Bear (P), 25CR-80KSI |  |  | 3.833 | 2.441 |  |  | 0.5 m |  |
| landing-nipple | landing-nipple | OS, LANDING NIPPLE, SELECTIVE, 2.313 OSX, 2-7/8 6.4PPF JFE BEAR BOX X PIN, 25CR[80], 3.385 X 2.313 X 15.1 | 5918.471 |  | 3.385 | 2.205 |  |  | 0.42 m |  |
| crossover-4 | crossover | Crossover - 2-7/8" Bear (B) X 3-1/2" Bear (P), 25CR-80KSI |  |  | 3.833 | 2.441 |  |  | 0.5 m |  |
| pup-joint-6ft-3 | pup-joint | 6 FT 3 1/2" 9.2# JFE BEAR B x P, 15Cr-125ksi Pups |  |  | 3.883 | 2.963 |  |  | 1.829 m |  |
| muleshoe | mule-shoe | 3-1/2" Muleshoe Guide / Drop Sub | 6023.688 |  | 4.0 | 2.9 |  |  |  | type |
| pbtd | pbtd | PBTD | 6243.438 |  |  |  |  |  |  |  |

Full cell-by-cell traceability (every mapped field, original value, unit, interpretation): `Dulang_A33ST3_traceability.csv`; each component's `source` string in the JSON names the cells too.

## Terminology → template type

| source term | cell(s) | template type | basis |
|---|---|---|---|
| 26" Conductor (Existing) | I6 | casing | matches source term |
| 16" Casing (Existing) | I7 | casing | matches source term |
| 13-3/8" Casing (Existing) | I8 | casing | matches source term |
| 9-5/8" Casing (Existing) | I9 | casing | matches source term |
| 7" Liner | I10 | liner | matches source term |
| TUBING 3-1/2" 9.2#, FOX GRE, R2 | O48 | tubing | matches source term |
| TUBING 3-1/2" 9.2#, UHP 15Cr-125, JFE BEAR, R3 | O59 | tubing | matches source term |
| X-Mas Tree : Solar Alert  5,000 psi | B13 | christmas-tree | matches source term |
| TUBING HANGER Bear (BOX Down) | O18 | tubing-hanger | matches source term |
| XOVER 3-1/2" BEAR Pin x FOX DTP (BOX) | O19 | crossover | matches source term |
| 12 FT 3-1/2" GRE Lined Fox (Pin x Pin) Pups | O20 | pup-joint | matches source term |
| 6 FT 3-1/2" GRE Lined Fox (Box x Pin) Pups | O23 | pup-joint | matches source term |
| OS, FLOW COUPLING, 3-1/2 9.2PPF JFEFOX DTP BOX X 2-7/8 6.4PPF JFEBEAR  | O24 | flow-coupling | matches source term |
| TRSV, 2-7/8 6.4PPF JFEBEAR BXP | O25 | trsv | matches source term |
| OS, FLOW COUPLING, 2-7/8 6.4PPF JFEBEAR BOX X 3-1/2 9.2PPF JFEFOX DTP  | O26 | flow-coupling | matches source term |
| 4 FT 3-1/2" GRE Lined Fox (Pin x Pin) Pups | O27 | pup-joint | matches source term |
| 6 FT 3-1/2" GRE Lined Fox (B x P) Pups | O50 | pup-joint | matches source term |
| Crossover - 3-1/2" FOX DTP BOX x 3-1/2 Bear Pin, 25Cr-80KSI | O51 | crossover | matches source term |
| 7 X 3-1/2 BLUEPACK MAX RH (29-32), 25CR (80), HNBR (90-90-90), 3-1/2 ( | O52 | packer | matches source term |
| 5 FT 3 1/2" 9.2# JFE BEAR B x P, 25Cr Pups | O53 | pup-joint | matches source term |
| Crossover - 3-1/2"  BEAR (B) x 2-7/8" Bear (P), 25CR-80KSI | O54 | crossover | matches source term |
| OS, LANDING NIPPLE, SELECTIVE, 2.313 OSX, 2-7/8 6.4PPF JFE BEAR BOX X  | O55 | landing-nipple | matches source term |
| Crossover - 2-7/8" Bear (B) X 3-1/2" Bear (P), 25CR-80KSI | O56 | crossover | matches source term |
| 6 FT 3 1/2" 9.2# JFE BEAR B x P, 15Cr-125ksi Pups | O57 | pup-joint | matches source term |
| 3-1/2" Muleshoe Guide / Drop Sub | O61 | mule-shoe | CLOSEST — 'Muleshoe Guide / Drop Sub' (combined item) -> Mule Shoe |
| PBTD | B74 | pbtd | matches source term |

## Items requiring review

### Identity (2)

| location | detail |
|---|---|
| B2 vs B4 | B2 'DULANG A33-ST3' vs B4 'Well ID / SLOT NO : Dulang A33ST2 WI' — wellbore suffix differs (ST3 vs ST2). Output named Dulang_A33ST3 (title + file name); confirm. |
| sheet 'OCTG' C1 | Sheet OCTG cell C1 = 'A23 , OP2' — refers to another well; OCTG sheets not used. |

### Conflict (2)

| location | detail |
|---|---|
| I12 vs O52 | Header I12 'Tubing 3-1/2" - Below 7" AHC' names an AHC packer; the diagram packer is '7 X 3-1/2 BLUEPACK MAX RH (29-32), 25CR '... Recorded; tubing grade/weight still from row 12. |
| B14 vs I13 | B14 'Completion Brine: 8.5ppg NaCl' vs I13 'Completion Fluid : 8.5 ppg NaCl filtered to 30 NTU' — I13 mapped to completionFluid (more detailed), B14 listed as unmapped. |

### Missing (21)

| location | detail |
|---|---|
| S6-S9 | Casing TOP printed as 'Surface' (no number); not converted to 0 because MDDF is measured from the drill floor -> topMd null. |
| row 10 | 7" liner header weight / grade / thread / top / bottom all 'TBC'; only the diagram TOL depth F40 is mapped. |
| O48, O59 | Tubing intervals not printed. Candidates (not mapped): GRE tubing between the TRSV assembly (K25 254.61) and packer assembly (K52 1800.21); 15Cr tubing below the packer to the muleshoe (K61 1836.02). |
| O18 | Tubing hanger depth not printed; row above says ELEVATION (DF to THF) = 15.8 m (THF depth below DF?) — not mapped. |
| T18 | Tubing hanger MAX OD 'TBC'. |
| groundElevation, surfaceCasing | Not in source -> template empty values. |
| O48 | tubing-gre: no depth printed for 'TUBING 3-1/2" 9.2#, FOX GRE, R2' -> topMd/bottomMd null |
| O59 | tubing-15cr: no depth printed for 'TUBING 3-1/2" 9.2#, UHP 15Cr-125, JFE BEAR, R3' -> topMd/bottomMd null |
| O18 | tubing-hanger: no depth printed for 'TUBING HANGER Bear (BOX Down)' -> topMd/bottomMd null |
| O19 | crossover-1: no depth printed for 'XOVER 3-1/2" BEAR Pin x FOX DTP (BOX)' -> topMd/bottomMd null |
| O20 | pup-joint-12ft: no depth printed for '12 FT 3-1/2" GRE Lined Fox (Pin x Pin) Pups' -> topMd/bottomMd null |
| O23 | pup-joint-6ft-1: no depth printed for '6 FT 3-1/2" GRE Lined Fox (Box x Pin) Pups' -> topMd/bottomMd null |
| O24 | flow-coupling-1: no depth printed for 'OS, FLOW COUPLING, 3-1/2 9.2PPF JFEFOX DTP BOX X 2-7/8 6.4PP' -> topMd/bottomMd null |
| O26 | flow-coupling-2: no depth printed for 'OS, FLOW COUPLING, 2-7/8 6.4PPF JFEBEAR BOX X 3-1/2 9.2PPF J' -> topMd/bottomMd null |
| O27 | pup-joint-4ft: no depth printed for '4 FT 3-1/2" GRE Lined Fox (Pin x Pin) Pups' -> topMd/bottomMd null |
| O50 | pup-joint-6ft-2: no depth printed for '6 FT 3-1/2" GRE Lined Fox (B x P) Pups' -> topMd/bottomMd null |
| O51 | crossover-2: no depth printed for 'Crossover - 3-1/2" FOX DTP BOX x 3-1/2 Bear Pin, 25Cr-80KSI' -> topMd/bottomMd null |
| O53 | pup-joint-5ft: no depth printed for '5 FT 3 1/2" 9.2# JFE BEAR B x P, 25Cr Pups' -> topMd/bottomMd null |
| O54 | crossover-3: no depth printed for 'Crossover - 3-1/2"  BEAR (B) x 2-7/8" Bear (P), 25CR-80KSI' -> topMd/bottomMd null |
| O56 | crossover-4: no depth printed for 'Crossover - 2-7/8" Bear (B) X 3-1/2" Bear (P), 25CR-80KSI' -> topMd/bottomMd null |
| O57 | pup-joint-6ft-3: no depth printed for '6 FT 3 1/2" 9.2# JFE BEAR B x P, 15Cr-125ksi Pups' -> topMd/bottomMd null |

### Unclear (4)

| location | detail |
|---|---|
| kbElevation B9 | B9 'Drill floor height: 30.1m-AMSL' mapped to kbElevation (closest field), converted to ft. |
| maxDeviation B11 | B11 'Maximum Deviation: 58deg @ 1,800m MDDF' — kept as text with the depth converted to ft. |
| completionType | 'Single string' from wellhead text B12 '…Single' and one tubing string on the diagram — confirm. |
| B67 | B67 = 6.47 (no unit; .xls formula not readable) — unmapped. |

## Unmapped information

Relevant source content with no suitable template field (or not mapped because its meaning/association is not established by the source):

| location | content | reason |
|---|---|---|
| B4 / B5 | Well ID / SLOT NO : Dulang A33ST2 WI; Field / Platform : Dulang A | slot / field — see identity |
| B7, B8 | Injection Rate:2500 blpd; Surface Injection Pressure: 1200psi | injection design data — no field |
| B14 | Completion Brine: 8.5ppg NaCl | see conflict |
| R6-R12 | threads: R6 Welded, R7 BTC, R8 BTC, R9 BTC, R11 JFEFOX, R12 JFE BEAR | no thread field |
| O17 / K17 / M17 | ELEVATION (DF to THF) 15.8 m | elevation, not equipment |
| N52 | angle 58.0 deg at packer K52 1800.21 m-MDDF (5906.201 ft) | no azimuth -> no survey station |
| B11 | Maximum Deviation: 58deg @ 1,800m MDDF | also only one survey value, no azimuth |
| O62 | ITCP Assembly | no depth, no Component_Type |
| text box | '2" LIB tagged HUD at 1858m WLTHF. Impression fish neck of ITCP assembly (Aug 2023)' (1858 m WLTHF = 6095.801 ft WLTHF; wireline datum, not converted to MDDF) | intervention note; sheet anchor not resolvable from the .xls |
| B63-B66 | Zone E34A: 1,846-1,856m MDDF (6056.43-6089.239 ft MDDF); 1,363.62-1,368.75m TVD; Reservoir Pressure: 1,503 psi, 93deg C | reservoir zone interval: not stated as perforation/formation top |
| B67 | 6.47 | no unit / label |
| O75 | Prepared by : Pravin Nair (Aug 2023) | document control |
| other sheets | OCTG, OCTG Movement, Summary OCTG Movement (tubular inventory/procurement); '   ' (symbol legend); 'Sheet1', 'XX2R&' (very hidden) empty | not schematic data |

## Quality checks

| check | result |
|---|---|
| valid JSON (re-parsed) | PASS |
| top-level sections and order = template | PASS |
| wellInfo keys and order = template (no extra keys) | PASS |
| wellInfo types = template type, or template empty value ('-' / null) | PASS |
| component keys are template component keys only | PASS |
| component value types (numbers/null for depths & diameters, text for names, bool flag) | PASS |
| every type is a Component_Types Type_Key | PASS |
| component ids unique | PASS |
| no template sample records copied | PASS |
| every depth: ft = source m / 0.3048 (re-read from the copy) | PASS |
| JSON depths = converted values in the conversion table | PASS |
| every depth has unit evidence (m-MDDF header or text) | PASS |
| component names = source cell text | PASS |
| derived / closest-type values flagged estimated | PASS |
| every component source names workbook > sheet > cells | PASS |
| well identity: B2 names this well | PASS |
| source workbook and template unchanged (SHA-256) | PASS |

## Source / template integrity

| file | sha256 before | sha256 after | unchanged |
|---|---|---|---|
| /Volumes/homes/public/DulangOSS/20261016/DULANG_WELLS/Well_Schematic/1. Well Schematic/Dulang A/Dulang A33ST3 Well Schematic.xls | b51284bc77e84503ffc3a7c07951b8b26d00021a16d041ac04b063a286736f8c | b51284bc77e84503ffc3a7c07951b8b26d00021a16d041ac04b063a286736f8c | YES |
| /Users/farahyusof/Documents/GitHub/NoramiraSazuki/AinnyFarah/WELL_SCHEMATIC/well-schematic-template-json.rtf | c7db15efb1917dfa76fa3102e270fe17115a8561f24eb0a155239870d079e712 | c7db15efb1917dfa76fa3102e270fe17115a8561f24eb0a155239870d079e712 | YES |
| /Users/farahyusof/Documents/GitHub/NoramiraSazuki/AinnyFarah/WELL_SCHEMATIC/well-schematic-sample-template.xlsx | d377a6310d42fda8c95e85db03f2416cb4c1b6f37f53920052331ce814f6d494 | d377a6310d42fda8c95e85db03f2416cb4c1b6f37f53920052331ce814f6d494 | YES |
