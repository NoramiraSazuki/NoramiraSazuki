# Dulang_A17ST2 — mapping and quality-check report

- Source workbook: `/Volumes/homes/public/DulangOSS/20261016/DULANG_WELLS/Well_Schematic/1. Well Schematic/Dulang A/Dulang A17ST2 Well Schematic updated .xlsx` (read-only; worked on a copy in `_work/src/`)
- Worksheet mapped: `OP5 single A17 DIAL eSTMZ FID`
- Template: `/Users/farahyusof/Documents/GitHub/NoramiraSazuki/AinnyFarah/WELL_SCHEMATIC/well-schematic-template-json.rtf` (read-only); component Type_Keys from `well-schematic-sample-template.xlsx` > Component_Types
- Output JSON: `Dulang_A17ST2.json` — 31 components; surveys / cement / plugBackCement / centralizers / floatEquipment / formations / formationTops / fluidContacts left empty (no supported source data)
- Generated: 2026-10-27 10:54

## Depth units, reference and conversion

Depth columns K19/K20 and S19/S20 'DEPTH m-MDDF' (merged K:L, S:T); casing table AB9-AB11 carry 'm MDDF' in the text. Note 'Note: All depths are in meter from RKB to top of equipment.' — see conflict.

- Every depth in the JSON is in **ft = m / 0.3048**, rounded to 3 decimals. Metres, unit evidence and cell are listed below.
- MD reference kept as printed (MDDF). No MD↔TVD or datum conversion was made. TVD values have no component field and are listed as unmapped.

| field | cell | m | unit | ft | kind |
|---|---|---|---|---|---|
| conductor.bottomMd | AB9 | 182.0 | m MDDF (in cell text AB9) | 597.113 | direct |
| int-casing.bottomMd | AB10 | 421.0 | m MDDF (in cell text AB10) | 1381.234 | direct |
| casing.bottomMd | AB11 | 1513.0 | m MDDF (in cell text AB11) | 4963.911 | direct |
| liner-7.topMd | K34 | 450.0 | m-MDDF (column header K20 'm-MDDF') | 1476.378 | direct |
| liner-7.bottomMd | AB13 | 2283.0 | m-MDDF (column header K20 'm-MDDF') via formula =K110 | 7490.157 | direct |
| tubing-upper.bottomMd | AB14 | 250.738 | m-MDDF (column header S20 'm-MDDF') via formula =S30 | 822.631 | derived |
| tubing-lower.topMd | S30 | 250.738 | m-MDDF (column header S20 'm-MDDF') | 822.631 | derived |
| trsv.topMd | S30 | 250.738 | m-MDDF (column header S20 'm-MDDF') | 822.631 | direct |
| spm-1.topMd | S42 | 829.588 | m-MDDF (column header S20 'm-MDDF') | 2721.745 | direct |
| spm-dial-1.topMd | S44 | 1169.029 | m-MDDF (column header S20 'm-MDDF') | 3835.397 | direct |
| spm-2.topMd | S46 | 1191.982 | m-MDDF (column header S20 'm-MDDF') | 3910.702 | direct |
| spm-dial-2.topMd | S48 | 1810.45 | m-MDDF (column header S20 'm-MDDF') | 5939.797 | direct |
| spm-3.topMd | S50 | 1832.781 | m-MDDF (column header S20 'm-MDDF') | 6013.061 | direct |
| ahr-packer.topMd | S59 | 1853.118 | m-MDDF (column header S20 'm-MDDF') | 6079.783 | direct |
| landing-nipple.topMd | S62 | 1873.214 | m-MDDF (column header S20 'm-MDDF') | 6145.715 | direct |
| mule-shoe.topMd | S68 | 1911.54 | m-MDDF (column header S20 'm-MDDF') | 6271.457 | direct |
| vch-packer.topMd | S71 | 1909.947 | m-MDDF (column header S20 'm-MDDF') | 6266.23 | direct |
| mcsz-sleeve-1.topMd | S73 | 1921.164 | m-MDDF (column header S20 'm-MDDF') | 6303.031 | direct |
| screen-1-top-joint.topMd | S75 | 1937.236 | m-MDDF (column header S20 'm-MDDF') | 6355.761 | direct |
| screen-1-production-joint.topMd | S77 | 1946.339 | m-MDDF (column header S20 'm-MDDF') | 6385.627 | direct |
| fcs-1.topMd | S79 | 1955.444 | m-MDDF (column header S20 'm-MDDF') | 6415.499 | direct |
| tnt-isolation-packer.topMd | S83 | 1956.854 | m-MDDF (column header S20 'm-MDDF') | 6420.125 | direct |
| mcsz-sleeve-2.topMd | S88 | 1965.973 | m-MDDF (column header S20 'm-MDDF') | 6450.043 | direct |
| screen-2-top-joint.topMd | S90 | 1971.765 | m-MDDF (column header S20 'm-MDDF') | 6469.045 | direct |
| screen-2-production-joint.topMd | S92 | 1980.868 | m-MDDF (column header S20 'm-MDDF') | 6498.911 | direct |
| fcs-2.topMd | S94 | 1989.971 | m-MDDF (column header S20 'm-MDDF') | 6528.776 | direct |
| sump-seal-assembly.topMd | S98 | 1992.621 | m-MDDF (column header S20 'm-MDDF') | 6537.47 | direct |
| awv-sump-packer.topMd | S99 | 1992.0 | m-MDDF (column header S20 'm-MDDF') | 6535.433 | direct |
| pbtd.topMd | K108 | 2241.0 | m-MDDF (column header K20 'm-MDDF') | 7352.362 | direct |
| wellInfo.kbElevation | G14 | 33.0 | m-AMSL (elevation) | 108.268 | direct |
| wellInfo.totalDepth | G10 | 2283.0 | m-MDDF (column header K20 'm-MDDF') via formula =K110 | 7490.157 | direct |

## wellInfo mapping

| json_path | mapped_value | cell | original_value | note |
|---|---|---|---|---|
| wellInfo.completionDate | 2020-05-30 | G8 | '2020-05-30T00:00:00' | DIRECT (date cell, label B8 'DATE OF COMPLETION'), written ISO yyyy-mm-dd |
| wellInfo.completionFluid | 8.5 PPG NaCl (+ 3% KCl) filtered to 30 NTU | V16 | '8.5 PPG NaCl (+ 3% KCl) filtered to 30 NTU' | DIRECT (label P16) |
| wellInfo.completionType | Single string | U21 | '3-1/2" SINGLE STRING TUBING HANGER' | CLOSEST: U21 'SINGLE STRING TUBING HANGER' and sheet title 'OP5 single …' |
| wellInfo.groundElevation |  |  |  | NOT IN SOURCE -> template empty value; flagged |
| wellInfo.kbElevation | 108.268 | G14 | '33 m-AMSL' | CLOSEST FIELD: DRILL FLOOR HEIGHT '33 m-AMSL' -> 33 / 0.3048 ft |
| wellInfo.maxDeviation | 80.3 deg | G12 | 80.3 (formula =MAX(J24:J142), cached value) | DIRECT (formula =MAX(J24:J142); unit J20 'Degree') |
| wellInfo.maxDeviation | 80.3 deg | J20 | 'Degree' | DIRECT (formula =MAX(J24:J142); unit J20 'Degree') |
| wellInfo.name | Dulang_A17ST2 | B2 | 'WELL NAME: DULANG A17-ST2 OP5' | DIRECT: 'WELL NAME: DULANG A17-ST2 OP5' written in the output naming form |
| wellInfo.packerFluid | 8.5 PPG NaCl (+ 0.2% BIOCIDE + 2% OXYGEN SCAVANGER + 2% CORROSION INHIBITOR) | V17 | '8.5 PPG NaCl (+ 0.2% BIOCIDE + 2% OXYGEN SCAVANGER ' | DIRECT (two lines joined) |
| wellInfo.packerFluid | 8.5 PPG NaCl (+ 0.2% BIOCIDE + 2% OXYGEN SCAVANGER + 2% CORROSION INHIBITOR) | V18 | '+ 2% CORROSION INHIBITOR)' | DIRECT (two lines joined) |
| wellInfo.rig | SAPURA BERANI | G9 | 'SAPURA BERANI' | DIRECT |
| wellInfo.surfaceCasing | - |  |  | NOT IN SOURCE -> template empty value; flagged |
| wellInfo.totalDepth | 7490.157 | G10 | 2283 (formula =K110, cached value) | DIRECT: TD 2283.0 m-MDDF (formula =K110; K20 'm-MDDF') -> ft |
| wellInfo.trajectoryType | Deviated | G12 | 80.3 (formula =MAX(J24:J142), cached value) | supported by maximum deviation 80.3 deg |
| wellInfo.tubing | 3-1/2" 9.2 PPF HP2-13CR-95 JFEBEAR; 3-1/2" 9.2 PPF 13CR-L80 JFEBEAR | V14 | '3-1/2"' | DIRECT: header rows 14-15 joined |
| wellInfo.tubing | 3-1/2" 9.2 PPF HP2-13CR-95 JFEBEAR; 3-1/2" 9.2 PPF 13CR-L80 JFEBEAR | X14 | '9.2 PPF' | DIRECT: header rows 14-15 joined |
| wellInfo.tubing | 3-1/2" 9.2 PPF HP2-13CR-95 JFEBEAR; 3-1/2" 9.2 PPF 13CR-L80 JFEBEAR | Y14 | 'HP2-13CR-95' | DIRECT: header rows 14-15 joined |
| wellInfo.tubing | 3-1/2" 9.2 PPF HP2-13CR-95 JFEBEAR; 3-1/2" 9.2 PPF 13CR-L80 JFEBEAR | Z14 | 'JFEBEAR' | DIRECT: header rows 14-15 joined |
| wellInfo.tubing | 3-1/2" 9.2 PPF HP2-13CR-95 JFEBEAR; 3-1/2" 9.2 PPF 13CR-L80 JFEBEAR | V15 | '3-1/2"' | DIRECT: header rows 14-15 joined |
| wellInfo.tubing | 3-1/2" 9.2 PPF HP2-13CR-95 JFEBEAR; 3-1/2" 9.2 PPF 13CR-L80 JFEBEAR | X15 | '9.2 PPF' | DIRECT: header rows 14-15 joined |
| wellInfo.tubing | 3-1/2" 9.2 PPF HP2-13CR-95 JFEBEAR; 3-1/2" 9.2 PPF 13CR-L80 JFEBEAR | Y15 | '13CR-L80' | DIRECT: header rows 14-15 joined |
| wellInfo.tubing | 3-1/2" 9.2 PPF HP2-13CR-95 JFEBEAR; 3-1/2" 9.2 PPF 13CR-L80 JFEBEAR | Z15 | 'JFEBEAR' | DIRECT: header rows 14-15 joined |
| wellInfo.wellhead | SOLAR ALERT | G15 | 'SOLAR ALERT' | DIRECT |

## Components

| id | type | name | topMd_ft | bottomMd_ft | OD_in | ID_in | weight | grade | interval | estimated |
|---|---|---|---|---|---|---|---|---|---|---|
| conductor | casing | CONDUCTOR (EXISTING) |  | 597.113 | 26 |  | 202.0 | A-36 |  |  |
| int-casing | casing | INT. CASING (EXISTING) |  | 1381.234 | 13.375 |  | 54.5 | K-55 |  |  |
| casing | casing | CASING (EXISTING) |  | 4963.911 | 9.625 |  | 43.5 | L-80 |  |  |
| liner-7 | liner | LINER | 1476.378 | 7490.157 | 7 |  | 29.0 | L-80 |  |  |
| tubing-upper | tubing | 3-1/2", 9.2#, HP2-13CR-95, JFE Bear |  | 822.631 | 3.5 |  | 9.2 | HP2-13CR-95 |  | bottomMd |
| tubing-lower | tubing | 3-1/2" 9.2 PPF 13CR-L80 JFEBEAR | 822.631 |  | 3.5 |  | 9.2 | 13CR-L80 |  | topMd |
| tubing-hanger | tubing-hanger | 3-1/2" SINGLE STRING TUBING HANGER |  |  |  |  |  |  |  |  |
| flow-coupling-1 | flow-coupling | FLOW COUPLING, 3-1/2 9.2# JFEBEAR B X P, 3FT, 13CR-L80 |  |  | 4.141 | 2.88 |  |  |  |  |
| trsv | trsv | TRMAXX-5E, 3-1/2 9.2# JFEBEAR B X P, 2.812 X-TYPE, 13CR-L80 | 822.631 |  | 5.0 | 2.812 |  |  |  |  |
| flow-coupling-2 | flow-coupling | FLOW COUPLING, 3-1/2 9.2# JFEBEAR B X P, 3FT 13CR-L80 |  |  | 4.141 | 2.88 |  |  |  |  |
| spm-1 | spm | 3-1/2" SPM (GLV  : Dummy - Change out made on 20.10.2021) | 2721.745 |  | 5.23 | 2.9 |  |  |  |  |
| spm-dial-1 | spm | 3-1/2" AUTO SPM DIAL (JAMMED CLOSE) | 3835.397 |  | 5.15 | 2.963 |  |  |  | type |
| spm-2 | spm | 3-1/2" SPM (GLV: 8/64" PTRO: 1045 psi SN: PMA-DUL- IPO-0005) | 3910.702 |  | 5.23 | 2.9 |  |  |  |  |
| spm-dial-2 | spm | 3-1/2" AUTO SPM DIAL (JAMMED CLOSE) | 5939.797 |  | 5.15 | 2.963 |  |  |  | type |
| spm-3 | spm | 3-1/2" SPM (GLV: Empty - recovered DV but NOT ABLE to set GLOV 10/64"  due to low RW - high angle) | 6013.061 |  | 5.23 | 2.9 |  |  |  |  |
| ahr-packer | packer | 7 X 3-1/2 AHR PACKER, 3-1/2 9.2# JFEBEAR B X P, 13CR-L80 | 6079.783 |  | 6.0 | 2.897 |  |  |  |  |
| landing-nipple | landing-nipple | LANDING NIPPLE, 2.813 OSX,  3-1/2 9.2# JFEBEAR B X P, 13CR-L80 | 6145.715 |  | 4.04 | 2.813 |  |  |  |  |
| mule-shoe | mule-shoe | SELF ALIGN MULE SHOE, 3-1/2 9.2# VAMTOP SC80 PIN, 13CR-L80 | 6271.457 |  | 3.86 | 3.05 |  |  |  |  |
| vch-packer | packer | VCH PACKER, 7" 26 - 29#, 13CR-L80, HNBR, 10KSI RATING | 6266.23 |  | 6.015 | 3.88 |  |  |  |  |
| mcsz-sleeve-1 | sliding-sleeve | MCSZ CLOSING SLEEVE, 13CR, 10KSI RATING | 6303.031 |  | 5.735 | 2.75 |  |  |  | type |
| screen-1-top-joint | screen | TOP JOINT PG WRAP MODULAR SCREEN | 6355.761 |  | 4.301 | 2.918 |  |  |  |  |
| screen-1-production-joint | screen | PRODUCTION JOINT PG WRAP MODULAR SCREEN | 6385.627 |  | 4.301 | 2.918 |  |  |  |  |
| fcs-1 | circulating-valve | "FCS" FRACTURE / CIRCULATION SLEEVE, 13CR | 6415.499 |  | 5.05 | 2.87 |  |  |  | type |
| tnt-isolation-packer | packer | TNT PERMA-SERIES® ISOLATION PACKER, 13CR FLOW-WET, HNBR, 7KSI RATING | 6420.125 |  | 5.875 | 3.875 |  |  |  |  |
| mcsz-sleeve-2 | sliding-sleeve | MCSZ CLOSING SLEEVE, 13CR, 10KSI RATING | 6450.043 |  | 5.735 | 2.75 |  |  |  | type |
| screen-2-top-joint | screen | TOP JOINT PG WRAP MODULAR SCREEN | 6469.045 |  | 4.301 | 2.918 |  |  |  |  |
| screen-2-production-joint | screen | PRODUCTION JOINT PG WRAP MODULAR SCREEN | 6498.911 |  | 4.301 | 2.918 |  |  |  |  |
| fcs-2 | circulating-valve | "FCS" FRACTURE / CIRCULATION SLEEVE, 13CR | 6528.776 |  | 5.05 | 2.87 |  |  |  | type |
| sump-seal-assembly | seal-assembly | SUMP SEAL ASSEMBLY, For 3.88" Bore Packer, 13CR, HNBR | 6537.47 |  | 5.613 | 2.75 |  |  |  |  |
| awv-sump-packer | packer | AWV SUMP PACKER , 7 " 26 - 32#, 13CR-80, HNBR, 10KSI RATING | 6535.433 |  | 5.875 | 3.88 |  |  |  |  |
| pbtd | pbtd | PBTD | 7352.362 |  |  |  |  |  |  |  |

Full cell-by-cell traceability (every mapped field, original value, unit, interpretation): `Dulang_A17ST2_traceability.csv`; each component's `source` string in the JSON names the cells too.

## Terminology → template type

| source term | cell(s) | template type | basis |
|---|---|---|---|
| CONDUCTOR (EXISTING) | P9 | casing | matches source term |
| INT. CASING (EXISTING) | P10 | casing | matches source term |
| CASING (EXISTING) | P11 | casing | matches source term |
| LINER | P13 | liner | matches source term |
| 3-1/2", 9.2#, HP2-13CR-95, JFE Bear | U26 | tubing | matches source term |
| 3-1/2" 9.2 PPF 13CR-L80 JFEBEAR | V15+X15+Y15+Z15 | tubing | matches source term |
| 3-1/2" SINGLE STRING TUBING HANGER | U21 | tubing-hanger | matches source term |
| FLOW COUPLING, 3-1/2 9.2# JFEBEAR B X P, 3FT, 13CR-L80 | U29 | flow-coupling | matches source term |
| TRMAXX-5E, 3-1/2 9.2# JFEBEAR B X P, 2.812 X-TYPE, 13CR-L80 | U30 | trsv | matches source term |
| FLOW COUPLING, 3-1/2 9.2# JFEBEAR B X P, 3FT 13CR-L80 | U31 | flow-coupling | matches source term |
| 3-1/2" SPM (GLV  : Dummy - Change out made on 20.10.2021) | U42 | spm | matches source term |
| 3-1/2" AUTO SPM DIAL (JAMMED CLOSE) | U44 | spm | CLOSEST — 'AUTO SPM DIAL' is a side-pocket mandrel variant -> SPM (Gas Lift) |
| 3-1/2" SPM (GLV: 8/64" PTRO: 1045 psi SN: PMA-DUL- IPO-0005) | U46 | spm | matches source term |
| 3-1/2" AUTO SPM DIAL (JAMMED CLOSE) | U48 | spm | CLOSEST — 'AUTO SPM DIAL' is a side-pocket mandrel variant -> SPM (Gas Lift) |
| 3-1/2" SPM (GLV: Empty - recovered DV but NOT ABLE to set GLOV 10/64"  | U50+U51 | spm | matches source term |
| 7 X 3-1/2 AHR PACKER, 3-1/2 9.2# JFEBEAR B X P, 13CR-L80 | U59 | packer | matches source term |
| LANDING NIPPLE, 2.813 OSX,  3-1/2 9.2# JFEBEAR B X P, 13CR-L80 | U62 | landing-nipple | matches source term |
| SELF ALIGN MULE SHOE, 3-1/2 9.2# VAMTOP SC80 PIN, 13CR-L80 | U68 | mule-shoe | matches source term |
| VCH PACKER, 7" 26 - 29#, 13CR-L80, HNBR, 10KSI RATING | U71 | packer | matches source term |
| MCSZ CLOSING SLEEVE, 13CR, 10KSI RATING | U73 | sliding-sleeve | CLOSEST — closing sleeve (shifts to open/close ports) -> Sliding Sleeve is the closest type |
| TOP JOINT PG WRAP MODULAR SCREEN | U75 | screen | matches source term |
| PRODUCTION JOINT PG WRAP MODULAR SCREEN | U77 | screen | matches source term |
| "FCS" FRACTURE / CIRCULATION SLEEVE, 13CR | U79 | circulating-valve | CLOSEST — 'FRACTURE / CIRCULATION SLEEVE' -> Circulating Valve (closest by 'circulation'; Sliding Sleeve also possible) |
| TNT PERMA-SERIES® ISOLATION PACKER, 13CR FLOW-WET, HNBR, 7KSI RATING | U83 | packer | matches source term |
| MCSZ CLOSING SLEEVE, 13CR, 10KSI RATING | U88 | sliding-sleeve | CLOSEST — closing sleeve (shifts to open/close ports) -> Sliding Sleeve is the closest type |
| TOP JOINT PG WRAP MODULAR SCREEN | U90 | screen | matches source term |
| PRODUCTION JOINT PG WRAP MODULAR SCREEN | U92 | screen | matches source term |
| "FCS" FRACTURE / CIRCULATION SLEEVE, 13CR | U94 | circulating-valve | CLOSEST — 'FRACTURE / CIRCULATION SLEEVE' -> Circulating Valve (closest by 'circulation'; Sliding Sleeve also possible) |
| SUMP SEAL ASSEMBLY, For 3.88" Bore Packer, 13CR, HNBR | U98 | seal-assembly | matches source term |
| AWV SUMP PACKER , 7 " 26 - 32#, 13CR-80, HNBR, 10KSI RATING | U99 | packer | matches source term |
| PBTD | I108 | pbtd | matches source term |

## Items requiring review

### Conflict (4)

| location | detail |
|---|---|
| AB15; S71; S68 | Lower tubing bottom not mapped: header AB15 1660 formula =ROUNDUP(S71-S30,0) is a LENGTH ending at S71 (VCH packer, 1909.947 m), while the string physically ends at the mule shoe S68 1911.54. Two candidates -> left empty. |
| AB15 vs AB8 | Header column AB8 'DEPTH' holds a length in AB15 (1660, formula =ROUNDUP(S71-S30,0)) -> not used as a depth. |
| B18 vs K20/S20 | Note B18 'Note: All depths are in meter from RKB to top of equipment.' says depths are from RKB; the depth columns say 'm-MDDF' (drill floor). Kept as m-MDDF; datum not changed. |
| B76 vs B77 | B76 formula =1115.9/(0.0521*1302.63*3.281) uses 1115.9 psi, label B77 '1115.30 psi, 177.8 deg F' says 1115.30 psi (both unmapped). |

### Missing (6)

| location | detail |
|---|---|
| casing tops | Top depths of conductor / intermediate casing / 9-5/8" casing not printed -> topMd null. |
| U21; AB14 | Tubing hanger depth and the upper tubing top are not printed. |
| groundElevation, surfaceCasing | No ground elevation; no row labelled 'surface casing' (rows are CONDUCTOR / INT. CASING / CASING) -> template empty values. |
| U21 | tubing-hanger: no depth printed for '3-1/2" SINGLE STRING TUBING HANGER' -> topMd/bottomMd null |
| U29 | flow-coupling-1: no depth printed for 'FLOW COUPLING, 3-1/2 9.2# JFEBEAR B X P, 3FT, 13CR-L80' -> topMd/bottomMd null |
| U31 | flow-coupling-2: no depth printed for 'FLOW COUPLING, 3-1/2 9.2# JFEBEAR B X P, 3FT 13CR-L80' -> topMd/bottomMd null |

### Unclear (4)

| location | detail |
|---|---|
| AB13; H109 | Liner bottom = header AB13 2283 (formula =K110 = TD). Diagram label H109 '7" LINER SHOE' has no depth. Mapped, review. |
| AB9-AB11, AB13, AB14 | casing-table 'DEPTH' value mapped to bottomMd as the setting (shoe) depth — interpretation, review. Note B18 says depths are 'to top of equipment'; for casings the DEPTH column is read as the shoe depth. |
| kbElevation G14 | G14 '33 m-AMSL' (DRILL FLOOR HEIGHT) mapped to kbElevation (closest field) and converted to ft — confirm DF = KB for this rig. |
| X9 | conductor: weight 202 printed without a unit (other rows say PPF) — value kept, unit to confirm |

### Anomaly (2)

| location | detail |
|---|---|
| S67/S68 vs S71 | Depth order: no-go locator S67 1911.049 and mule shoe S68 1911.54 are deeper than the VCH packer S71 1909.947 listed below them (typical when the locator lands inside the packer). Kept as printed. |
| S98 vs S99 | S99 1992 (sump packer) is shallower than the seal assembly above it S98 1992.621. Kept as printed. |

## Unmapped information

Relevant source content with no suitable template field (or not mapped because its meaning/association is not established by the source):

| location | content | reason |
|---|---|---|
| B3 | LOCATION / SLOT: DULANG A / 06 | slot number — no template field |
| G13 | WATER DEPTH 77 m | no template field (groundElevation not stated) |
| B18 | Note: All depths are in meter from RKB to top of equipment. | depth convention note (see conflict) |
| Z9-Z15 | threads: Z9 WELDED, Z10 BTC, Z11 BTC, Z13 VAM TOP/BEAR, Z14 JFEBEAR, Z15 JFEBEAR | no thread field in template |
| F28 | Tie Back 7" Liner | tie-back liner label without depth or size data |
| U67 / S67 / Z67 / AB67 | NO-GO LOCATOR, 3-1/2 9.2# JFEBEAR B X VAMTOP SC80 P, 13CR-L80 at 1911.049 m-MDDF (6269.846 ft), ID 3.05 in, OD 4.798 in | no-go locator: no Component_Type with this meaning (ST Slot Locator / No-Go Nipple are different equipment) |
| U45, U47, U49, U52 | U45 2026-05-02, U47 2026-05-14, U49 2026-05-02, U52 2026-05-14 | dates printed under SPM rows (meaning not stated) — no field |
| U54-U56 | Run MLT from HUD at 1830m-MDTHF / Small leak (from SPM#3) observed by MLT Condition#2 run / 2026-05-01 | well-intervention note; HUD 1830 m is m-MDTHF (different datum) — not converted to MDDF |
| B73-B77 | Zone E22A: 1944 - 1951 m MDDF (7 m) (6377.953-6400.919 ft MDDF); 1302.63 - 1303.81 m TVD; B76 5.011 (EMW formula); 1115.30 psi, 177.8 deg F | reservoir zone interval: not stated as perforation/formation top; association with the screens is by position only -> not mapped |
| B91-B95 | Zone E22B: 1981 - 1988 m MDDF (7 m) (6499.344-6522.31 ft MDDF); 1308.89 - 1310.17 m TVD; B94 4.871 (EMW formula); 1089.80 psi, 177.8 deg F | as above |
| J75, J93, J109 | deviation 80.3 / 80.1 / 58 deg (column J 'Deviation Degree') | no MD/azimuth pair -> cannot form survey stations |
| B84 | Distance 30 m | no stated start/end |
| K72, K78, K90, K96 | K72 1934.429, K78 1961.783, K90 1978.395, K96 2015.815 m-MDDF | depths with no label in the row -> meaning unknown |
| U110-Z112 | Designed by: Mahirah/Hidayah (Completion Engineer); (Production Technologist); Date: 3rd June 2020; Rev: Draft | document control — no template field |
| drawing | packer picture (xl/media image) anchored beside the AHR packer row | image only, no data |

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
| /Volumes/homes/public/DulangOSS/20261016/DULANG_WELLS/Well_Schematic/1. Well Schematic/Dulang A/Dulang A17ST2 Well Schematic updated .xlsx | 55b59ceef97479498c74c6468ede7bf68944b24e690c279060633df8355ad121 | 55b59ceef97479498c74c6468ede7bf68944b24e690c279060633df8355ad121 | YES |
| /Users/farahyusof/Documents/GitHub/NoramiraSazuki/AinnyFarah/WELL_SCHEMATIC/well-schematic-template-json.rtf | c7db15efb1917dfa76fa3102e270fe17115a8561f24eb0a155239870d079e712 | c7db15efb1917dfa76fa3102e270fe17115a8561f24eb0a155239870d079e712 | YES |
| /Users/farahyusof/Documents/GitHub/NoramiraSazuki/AinnyFarah/WELL_SCHEMATIC/well-schematic-sample-template.xlsx | d377a6310d42fda8c95e85db03f2416cb4c1b6f37f53920052331ce814f6d494 | d377a6310d42fda8c95e85db03f2416cb4c1b6f37f53920052331ce814f6d494 | YES |
