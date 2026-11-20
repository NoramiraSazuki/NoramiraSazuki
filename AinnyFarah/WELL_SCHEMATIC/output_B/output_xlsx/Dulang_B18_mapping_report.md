# Dulang_B18 — mapping and quality-check report

- Source workbook: `/Volumes/homes/public/DulangOSS/20261016/DULANG_WELLS/Well_Schematic/1. Well Schematic/Dulang B/Dulang B18 ESP Workover Schematic Updated .xlsx` (read-only; worked on a copy in `_work/src/`)
- Internal well label: `COMPLETION SCHEMATIC : DULANG B-18` (B2) — agrees with the file name; output name `Dulang_B18` (user-approved)
- Worksheet mapped: `B18 WO (Jan25)`; worksheets inspected: 'B18 WO (Jan25)' (visible), 'MLE Space out' (hidden)
- Reading: cells + formulas + cached values + drawing layer (shapes / text boxes / pictures) from the .xlsx XML; rendered with macOS Quick Look (`_work/render/ql/`) and visually compared; embedded images extracted to `_work/media/B18/` and OCR'd where they hold text
- Template: `/Users/farahyusof/Documents/GitHub/NoramiraSazuki/AinnyFarah/WELL_SCHEMATIC/well-schematic-template-json.rtf`; Type_Keys from `well-schematic-sample-template.xlsx` > Component_Types
- Output JSON: `Dulang_B18.json` — 38 components; surveys / cement / plugBackCement / centralizers / floatEquipment / formations / formationTops / fluidContacts left empty (no supported source data)
- Generated: 2026-10-29 10:30

## Depth units, reference and conversion

- Source depths are **m-MDTHF (MD from tubing hanger flange; column headers J16/Q16)**. JSON depths are **ft = m × 3.28084**, 3 decimals; original metres kept in `topMdM` / `bottomMdM`.
- The MD reference (THF) was not shifted to drill floor/RKB and no MD→TVD conversion was made.

| field | ref | m | unit | ft |
|---|---|---|---|---|
| flow-coupling-1.topMd | Q20 | 222.47 | m-MDTHF (column header Q16) | 729.888 |
| trsv.topMd | Q22 | 223.88 | m-MDTHF (column header Q16) | 734.514 |
| flow-coupling-2.topMd | Q24 | 225.41 | m-MDTHF (column header Q16) | 739.534 |
| spm-1.topMd | Q27 | 846.26 | m-MDTHF (column header Q16) | 2776.444 |
| spm-2.topMd | Q29 | 1250.19 | m-MDTHF (column header Q16) | 4101.673 |
| spm-3.topMd | Q32 | 1450.94 | m-MDTHF (column header Q16) | 4760.302 |
| sliding-sleeve-1.topMd | Q34 | 1459.15 | m-MDTHF (column header Q16) | 4787.238 |
| esp-packer.topMd | Q38 | 1464.75 | m-MDTHF (column header Q16) | 4805.61 |
| sliding-sleeve-2.topMd | Q41 | 1521.81 | m-MDTHF (column header Q16) | 4992.815 |
| esp-pump.topMd | Q49 | 1569.41 | m-MDTHF (column header Q16) | 5148.983 |
| end-of-esp-assy.topMd | Q51 | 1591.95 | m-MDTHF (column header Q16) | 5222.933 |
| ls-tubing-stub.topMd | J51 | 1616.0 | m-MDTHF (column header J16) | 5301.837 |
| ss-tubing-stub.topMd | J52 | 1620.0 | m-MDTHF (column header J16) | 5314.961 |
| guib-packer-1.topMd | J54 | 1627.7 | m-MDTHF (column header J16) | 5340.223 |
| telescopic-joint.topMd | J56 | 1628.8 | m-MDTHF (column header J16) | 5343.832 |
| ssd-left.topMd | J57 | 1633.3 | m-MDTHF (column header J16) | 5358.596 |
| ssd-right.topMd | Q57 | 1632.28 | m-MDTHF (column header Q16) | 5355.25 |
| blast-joint-1.topMd | Q59 | 1636.3 | m-MDTHF (column header Q16) | 5368.438 |
| wp1-nipple-left.topMd | J61 | 1677.6 | m-MDTHF (column header J16) | 5503.937 |
| wp1-nipple-right.topMd | Q61 | 1677.3 | m-MDTHF (column header Q16) | 5502.953 |
| guib-packer-2.topMd | J63 | 1684.1 | m-MDTHF (column header J16) | 5525.263 |
| no-go-nipple-left.topMd | J66 | 1687.0 | m-MDTHF (column header J16) | 5534.777 |
| wl-entry-guide-left.topMd | J67 | 1687.7 | m-MDTHF (column header J16) | 5537.074 |
| ssd-right-2.topMd | Q74 | 1694.4 | m-MDTHF (column header Q16) | 5559.055 |
| wp1-nipple-right-2.topMd | Q78 | 1737.3 | m-MDTHF (column header Q16) | 5699.803 |
| locator-seal-assembly.topMd | Q84 | 1739.0 | m-MDTHF (column header Q16) | 5705.381 |
| guib-packer-3.topMd | J84 | 1793.1 | m-MDTHF (column header J16) | 5882.874 |
| seal-bore-extension.topMd | Q85 | 1740.7 | m-MDTHF (column header Q16) | 5710.958 |
| no-go-nipple-right.topMd | Q88 | 1743.0 | m-MDTHF (column header Q16) | 5718.504 |
| weg-right.topMd | Q89 | 1743.6 | m-MDTHF (column header Q16) | 5720.473 |
| perf-e2.topMd | C58 | 1639.74 | m (in text; datum not stated in the label, sheet depths are m-MDTHF) | 5379.725 |
| perf-e2.bottomMd | C58 | 1643.74 | m (in text; datum not stated in the label, sheet depths are m-MDTHF) | 5392.848 |
| perf-e10-11-1.topMd | C76 | 1701.24 | m (in text; datum not stated in the label, sheet depths are m-MDTHF) | 5581.496 |
| perf-e10-11-1.bottomMd | C76 | 1715.24 | m (in text; datum not stated in the label, sheet depths are m-MDTHF) | 5627.428 |
| perf-e10-11-2.topMd | C76 | 1717.24 | m (in text; datum not stated in the label, sheet depths are m-MDTHF) | 5633.99 |
| perf-e10-11-2.bottomMd | C76 | 1722.24 | m (in text; datum not stated in the label, sheet depths are m-MDTHF) | 5650.394 |
| perf-e10-11-3.topMd | C76 | 1725.24 | m (in text; datum not stated in the label, sheet depths are m-MDTHF) | 5660.236 |
| perf-e10-11-3.bottomMd | C76 | 1735.24 | m (in text; datum not stated in the label, sheet depths are m-MDTHF) | 5693.045 |
| perf-e12-13.topMd | C88 | 1744.4 | m (in text; datum not stated in the label, sheet depths are m-MDTHF) | 5723.097 |
| perf-e12-13.bottomMd | C88 | 1760.24 | m (in text; datum not stated in the label, sheet depths are m-MDTHF) | 5775.066 |
| wellInfo.totalDepth | L95 | 1852.0 | m-MDTHF (in text) | 6076.116 |

## wellInfo

| json_path | mapped_value | cell_or_object | original_value | method | note |
|---|---|---|---|---|---|
| wellInfo.completionDate | - |  |  | EMPTY | not supported by the visible source -> template empty value ('-' text / null number) |
| wellInfo.completionFluid | - |  |  | EMPTY | not supported by the visible source -> template empty value ('-' text / null number) |
| wellInfo.completionType | SINGLE OIL PRODUCER | B3 | ' SINGLE OIL PRODUCER ' | DIRECT | visible title row 3 ('SINGLE OIL PRODUCER'); ESP completion |
| wellInfo.groundElevation |  |  |  | EMPTY | not supported by the visible source -> template empty value ('-' text / null number) |
| wellInfo.kbElevation |  |  |  | EMPTY | not supported by the visible source -> template empty value ('-' text / null number) |
| wellInfo.maxDeviation | 50.8 deg @ 2301.542 ft MD (THF) | R25 | 50.8 | DERIVED | maximum of the ANGLE column R (R15/R16 '( ° )'), at Q25 701.51 m-MDTHF |
| wellInfo.name | Dulang_B18 | B2 | 'COMPLETION SCHEMATIC : DULANG B-18' | STANDARDISED NAME | internal label B2 'COMPLETION SCHEMATIC : DULANG B-18' -> Dulang_B18 (user-approved naming) |
| wellInfo.packerFluid | - |  |  | EMPTY | not supported by the visible source -> template empty value ('-' text / null number) |
| wellInfo.rig | - |  |  | EMPTY | not supported by the visible source -> template empty value ('-' text / null number) |
| wellInfo.surfaceCasing | - |  |  | EMPTY | not supported by the visible source -> template empty value ('-' text / null number) |
| wellInfo.totalDepth | 6076.116 | L95 | 'TD @ 1852 m MDTHF' | CONVERTED | diagram label 'TD @ 1852 m MDTHF': 1852.0 m x 3.28084; datum MDTHF kept |
| wellInfo.trajectoryType | Deviated | R25 | 50.8 | DERIVED | rule: maximum printed inclination 50.8 deg > 10 deg -> 'Deviated' |
| wellInfo.tubing | 4-1/2" TBG (SINGLE) W/ESP |  | 'Picture 1 (xl/media/image4.png) title block' | OCR + visual check | casing programme line of the wellhead drawing |
| wellInfo.wellhead | SOLAR ALERT SDN. BHD. — WELL B18 (OEM: SOLAR/ICC); 9-5/8" CSG. X 4-1/2" TBG (SINGLE) W/ESP; 5,000 PSI WP |  | 'Picture 1 (xl/media/image4.png) title block' | OCR + visual check | embedded wellhead drawing, title block |

## Components

| id | type | name | string | topMd_ft | bottomMd_ft | topMd_m | bottomMd_m | OD_in | ID_in | interval | zone | estimated |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| christmas-tree | christmas-tree | ASSY, XMAS TREE, SINGLE, 4-1/16" 5K, MODEL 'SAE' SLAB GV |  |  |  |  |  |  |  |  |  |  |
| flow-coupling-1 | flow-coupling | 4-1/2", Flow Coupling, CP Groove, JFE Bear |  | 729.888 |  | 222.47 |  | 5.09 | 3.884 |  |  |  |
| trsv | trsv | 4-1/2" TRMAXX-5E, 3.812" X-TYPE, JFE Bear |  | 734.514 |  | 223.88 |  | 6.923 | 3.812 |  |  | type |
| flow-coupling-2 | flow-coupling | 4-1/2", Flow Coupling, CP Groove, JFE Bear |  | 739.534 |  | 225.41 |  | 5.09 | 3.884 |  |  |  |
| spm-1 | spm | 4-1/2" SPM 13-CR L80, JFE Bear, IPO (10/64), PTRO 1016 psi |  | 2776.444 |  | 846.26 |  | 6.393 | 3.865 |  |  |  |
| spm-2 | spm | 4-1/2" SPM 13-CR L80, JFE Bear (Dummy) |  | 4101.673 |  | 1250.19 |  | 6.393 | 3.865 |  |  |  |
| spm-3 | spm | 4-1/2" SPM 13-CR L80, JFE Bear GLOV (16/64) |  | 4760.302 |  | 1450.94 |  | 6.393 | 3.865 |  |  |  |
| sliding-sleeve-1 | sliding-sleeve | 4-1/2" SLIDING SLEEVE, OS(X)-3D, 3.813" OSX (CLOSED) |  | 4787.238 |  | 1459.15 |  | 5.613 | 3.813 |  |  |  |
| esp-packer | packer | 9-5/8" x 4-1/2" Viking ESP Packer |  | 4805.61 |  | 1464.75 |  | 8.835 | 3.96 |  |  |  |
| sliding-sleeve-2 | sliding-sleeve | 4-1/2" SLIDING SLEEVE, OS(X)-3D, 3.813" OSX (OPEN) |  | 4992.815 |  | 1521.81 |  | 5.613 | 3.813 |  |  |  |
| esp-pump | esp | ESP Pump: 538-7400.71.12: SSCA SM RM A |  | 5148.983 |  | 1569.41 |  |  |  |  |  |  |
| end-of-esp-assy | eot | End of ESP Assy |  | 5222.933 |  | 1591.95 |  |  |  |  |  | type |
| ls-tubing-stub | tubing | 2-7/8" 6.5ppf NK3-SB Long String Tubing Stub | Long | 5301.837 |  | 1616.0 |  |  | 2.875 |  |  | type,innerDiameter |
| ss-tubing-stub | tubing | 2-7/8" 6.5ppf NK3-SB Short String Tubing Stub | Short | 5314.961 |  | 1620.0 |  |  | 2.875 |  |  | type,innerDiameter |
| guib-packer-1 | packer | GUIB MAGNUM 'HT' PERM. PKR |  | 5340.223 |  | 1627.7 |  |  |  |  |  |  |
| telescopic-joint | expansion-joint | TELESCOPIC JOINT |  | 5343.832 |  | 1628.8 |  |  |  |  |  | type |
| ssd-left | sliding-sleeve | WPB-2DE' SSD (OPEN) |  | 5358.596 |  | 1633.3 |  |  | 2.312 |  |  |  |
| ssd-right | sliding-sleeve | WPB-2DE' SSD (CLOSED) |  | 5355.25 |  | 1632.28 |  |  | 2.312 |  |  |  |
| blast-joint-1 | blast-joint | BLAST JOINT |  | 5368.438 |  | 1636.3 |  |  | 2.441 |  |  |  |
| wp1-nipple-left | landing-nipple | WP-1 NIPPLE |  | 5503.937 |  | 1677.6 |  |  | 2.312 |  |  | type |
| wp1-nipple-right | landing-nipple | WP-1 NIPPLE |  | 5502.953 |  | 1677.3 |  |  | 2.312 |  |  | type |
| guib-packer-2 | packer | GUIB MAGNUM 'HT' PERM. PKR |  | 5525.263 |  | 1684.1 |  |  | 2.441 |  |  |  |
| no-go-nipple-left | nipple (no-go) | D' NO-GO NIPPLE |  | 5534.777 |  | 1687.0 |  |  | 2.25 |  |  |  |
| wl-entry-guide-left | mule-shoe | W/L ENTRY GUIDE |  | 5537.074 |  | 1687.7 |  |  | 2.441 |  |  | type |
| ssd-right-2 | sliding-sleeve | WPB-2DE' SSD (CLOSED) |  | 5559.055 |  | 1694.4 |  |  | 2.312 |  |  |  |
| blast-joint-2 | blast-joint | BLAST JOINT |  |  |  |  |  |  | 2.441 |  |  |  |
| wp1-nipple-right-2 | landing-nipple | WP-1 NIPPLE |  | 5699.803 |  | 1737.3 |  |  | 2.312 |  |  | type |
| locator-seal-assembly | seal-assembly | LOC. SEAL ASSY |  | 5705.381 |  | 1739.0 |  |  | 2.441 |  |  |  |
| guib-packer-3 | packer | GUIB MAGNUM 'HT' PERM. PKR |  | 5882.874 |  | 1793.1 |  |  | 4.0 |  |  |  |
| seal-bore-extension | pbr | SEAL BORE EXT |  | 5710.958 |  | 1740.7 |  |  | 4.0 |  |  | type |
| no-go-nipple-right | nipple (no-go) | D' NO-GO NIPPLE |  | 5718.504 |  | 1743.0 |  |  | 2.25 |  |  |  |
| weg-right | mule-shoe | WEG |  | 5720.473 |  | 1743.6 |  |  | 2.441 |  |  | type |
| casing-9-5-8 | casing | 9 5/8" CASING SHOE 9-5/8" 40PPF, N80 |  |  |  |  |  |  | 9.625 |  |  | innerDiameter |
| perf-e2 | perforation | E-2 (FB S2; NDR) perforation 1639.74 m - 1643.74 m |  | 5379.725 | 5392.848 | 1639.74 | 1643.74 |  |  |  | E-2 (FB S2; NDR) |  |
| perf-e10-11-1 | perforation | E-10/11 (FB S2; WIR) perforation 1701.24 m - 1715.24 m |  | 5581.496 | 5627.428 | 1701.24 | 1715.24 |  |  |  | E-10/11 (FB S2; WIR) |  |
| perf-e10-11-2 | perforation | E-10/11 (FB S2; WIR) perforation 1717.24 m - 1722.24 m |  | 5633.99 | 5650.394 | 1717.24 | 1722.24 |  |  |  | E-10/11 (FB S2; WIR) |  |
| perf-e10-11-3 | perforation | E-10/11 (FB S2; WIR) perforation 1725.24 m - 1735.24 m |  | 5660.236 | 5693.045 | 1725.24 | 1735.24 |  |  |  | E-10/11 (FB S2; WIR) |  |
| perf-e12-13 | perforation | E-12/13 (FB S2; WIR) perforation 1744.4 m - 1760.24 m |  | 5723.097 | 5775.066 | 1744.4 | 1760.24 |  |  |  | E-12/13 (FB S2; WIR) |  |

Cell-by-cell traceability: `Dulang_B18_traceability.csv` (every populated JSON path, source cell/object, original value, formula + cached value, unit, method).

## Approximate type mappings and assumptions

| id | type | actualType | note |
|---|---|---|---|
| trsv | trsv | 4-1/2" TRMAXX-5E, 3.812" X-TYPE, JFE Bear | type: TRMAXX-5E is a tubing-retrievable safety valve model (label does not say 'TRSV') -> TRSV |
| end-of-esp-assy | eot | End of ESP Assy | type: 'End of ESP Assy' = deepest point of the completion string -> EOT is the nearest type |
| ls-tubing-stub | tubing | 2-7/8" 6.5ppf NK3-SB Long String Tubing Stub | type: tubing stub = top of the tubing left in hole -> Tubing (top = stub depth; bottom not printed); innerDiameter: size 2-7/8" in label D51 has no ID/OD designation -> recorded as ID (user-defined mapping rule) |
| ss-tubing-stub | tubing | 2-7/8" 6.5ppf NK3-SB Short String Tubing Stub | type: tubing stub = top of the tubing left in hole -> Tubing (top = stub depth; bottom not printed); innerDiameter: size 2-7/8" in label D52 has no ID/OD designation -> recorded as ID (user-defined mapping rule) |
| telescopic-joint | expansion-joint | TELESCOPIC JOINT | type: telescopic joint -> Expansion Joint is the nearest type |
| wp1-nipple-left | landing-nipple | WP-1 NIPPLE | type: 'WP-1 NIPPLE' (nipple profile, no no-go stated) -> Landing Nipple is the nearest type |
| wp1-nipple-right | landing-nipple | WP-1 NIPPLE | type: 'WP-1 NIPPLE' (nipple profile, no no-go stated) -> Landing Nipple is the nearest type |
| wl-entry-guide-left | mule-shoe | W/L ENTRY GUIDE | type: wireline entry guide (WEG) at the tubing end -> Mule Shoe is the nearest type |
| wp1-nipple-right-2 | landing-nipple | WP-1 NIPPLE | type: 'WP-1 NIPPLE' (nipple profile, no no-go stated) -> Landing Nipple is the nearest type |
| seal-bore-extension | pbr | SEAL BORE EXT | type: sealbore extension -> PBR is the nearest type |
| weg-right | mule-shoe | WEG | type: wireline entry guide (WEG) at the tubing end -> Mule Shoe is the nearest type |
| casing-9-5-8 | casing |  | innerDiameter: size 9 5/8" in label D92 has no ID/OD designation -> recorded as ID (user-defined mapping rule) |

## Items requiring review

### Conflict (1)

| location | detail |
|---|---|
| rows 4-14 (hidden) | Hidden header rows 4-14 (B4-B14, casing & tubing table N4:Y11, fluids N13/N14) are identical in B05, B17 and B18 and name another well: 'Field / Platform : Erb West / EWDP-A Platform', 'Well ID / SLOT NO : EW-119 / Slot #9', spud/completion 'XX XX 2017', rig 'XXX'. Casing depths there (e.g. 7" liner 7150-9959 ft-MDTHF, max deviation @ 10092 ft) exceed the well's TD. Treated as a copied template block: NOT mapped (wellInfo fields left empty, no casing components from it). |

### Missing (4)

| location | detail |
|---|---|
| casing-9-5-8 | Casing shoe label D92 has no depth (J92 empty) -> top/bottom empty. |
| T76 | Second blast joint has no depth. |
| T76 | blast-joint-2: no depth printed for 'BLAST JOINT' -> topMd/bottomMd empty |
| D92+D93 | casing-9-5-8: no depth printed for '9 5/8" CASING SHOE 9-5/8" 40PPF, N80' -> topMd/bottomMd empty |

### Unclear (9)

| location | detail |
|---|---|
| J16 / Q16 | All depths are m-MDTHF (from the tubing hanger flange), not drill floor / RKB. Converted to ft only; the reference was not shifted (the DF-to-THF offset '46 ft' exists only in the hidden copied block). Consumers expecting MD-RKB must add the offset. |
| column Q / J 'DEPTH' | The sheet does not say whether equipment depths are to the top of the item; mapped to topMd. Item lengths (column S, m) are kept in 'interval' and not used to derive bottoms. |
| D63/D64, D66, D84/D85, D92 | Several left-column cells repeat a second line 'PERM. PKR' (visible on the sheet). Treated as stray text: D63/D84 keep 'GUIB MAGNUM 'HT' PERM. PKR'; D66 and D92 use only their first line. Confirm. |
| C76 / C88 | Perforation depths written with comma decimals ('1725,24', '1744,4', '1760,24') -> read as decimal points. |
| T51 | 'End of ESP Assy' mapped as EOT (nearest type). |
| maxDeviation / trajectoryType | Derived from the ANGLE column R (max 50.8 deg at Q25 701.51 m); rule: any printed inclination > 10 deg -> 'Deviated'. |
| Picture 1 (AD17:BL48) | Embedded wellhead drawing is a cropped screenshot of a PDF viewer: equipment table items 1-2 legible, items 3-9 cut off. Read by OCR + visual check. |
| sheet 'MLE Space out' (hidden) | Identical to B17's sheet (pump 538-5800 differs from B18's pump label T49 '538-7400') -> not mapped. |
| perforation datum | Perforation texts give 'm' without a datum; sheet depths are m-MDTHF (assumed same reference). |

## Unmapped information

| location | content | reason |
|---|---|---|
| rows 4-14 (hidden) | Erb West / EW-119 header, casing & tubing table, completion/packer fluid, max deviation | copied block from another well — see Conflict |
| K84 | 19.431 (no label, no unit) | meaning unknown |
| T94-T95 | DANISH (JAN 2025) | document control (updated by / date) |
| T25 / Q25 / R25 | ESP cable & Gas Vent control Line End point at 701.51 m-MDTHF (2301.542 ft), angle 50.8 deg | control-line end point: no Component_Type |
| T45 / Q45 | 5.75" x 3" IPV1 at 1526.47 m-MDTHF (5008.104 ft), ID 2.813 in, OD 5.75 in | UNRESOLVED TYPE: IPV1 valve |
| T50 / Q50 | ACE Xtreme Temperature Dual at 1591.39 m-MDTHF (5221.096 ft) | UNRESOLVED TYPE: downhole temperature/pressure sensor |
| T82 / Q82 | CENTRALIZER at 1739 m-MDTHF (5705.381 ft), ID 2.441 in | tubing centralizer: no Component_Type / casing-centralizer section |
| column R | inclinations 222.47 m: 13.5 deg, 701.51 m: 50.8 deg, 846.26 m: 50 deg, 1250.19 m: 50.2 deg, 1450.94 m: 48.7 deg, 1464.75 m: 48.7 deg, 1521.81 m: 48.2 deg, 1526.47 m: 48.2 deg, 1569.41 m: 48.2 deg, 1591.95 m: 48.2 deg | no azimuth -> no survey stations |
| Picture 1 item 2 | ASSY, TUBING HEAD ADAPTER, 4-1/16" 5K RX-39 ROTATING FLANGE TOP X 13-5/8" BX-160 SEGMENTED CLAMP BTM W/1 ESP PENETRATOR FEED THRU HOLE | OCR; no depth; tubing head adapter has no Component_Type |
| K37 | '`' | stray character |
| B3 | SINGLE OIL PRODUCER | used for completionType |

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
