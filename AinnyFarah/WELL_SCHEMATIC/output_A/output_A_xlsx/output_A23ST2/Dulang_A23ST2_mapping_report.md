# Dulang_A23ST2 — mapping and quality-check report

- Source workbook: `/Volumes/homes/public/DulangOSS/20261016/DULANG_WELLS/Well_Schematic/1. Well Schematic/Dulang A/Dulang A23ST2 Well Schematic.xlsx` (read-only; worked on a copy in `_work/src/`)
- Worksheet mapped: `OP2 dual A23St1 STMZ E51  FID`
- Template: `/Users/farahyusof/Documents/GitHub/NoramiraSazuki/AinnyFarah/WELL_SCHEMATIC/well-schematic-template-json.rtf` (read-only); component Type_Keys from `well-schematic-sample-template.xlsx` > Component_Types
- Output JSON: `Dulang_A23ST2.json` — 48 components; surveys / cement / plugBackCement / centralizers / floatEquipment / formations / formationTops / fluidContacts left empty (no supported source data)
- Generated: 2026-10-27 10:54

## Depth units, reference and conversion

Depth columns K19/K20 and S19/S20 'DEPTH m-MDDF'. Note 'Note: All depths are in meter from RKB to top of equipment.' — see conflict.

- Every depth in the JSON is in **ft = m / 0.3048**, rounded to 3 decimals. Metres, unit evidence and cell are listed below.
- MD reference kept as printed (MDDF). No MD↔TVD or datum conversion was made. TVD values have no component field and are listed as unmapped.

| field | cell | m | unit | ft | kind |
|---|---|---|---|---|---|
| short-tubing-upper.bottomMd | K32 | 245.6 | m-MDDF (column header K20 'm-MDDF') | 805.774 | derived |
| short-tubing-lower.topMd | K32 | 245.6 | m-MDDF (column header K20 'm-MDDF') | 805.774 | derived |
| short-tubing-lower.bottomMd | K73 | 1369.09 | m-MDDF (column header K20 'm-MDDF') | 4491.765 | derived |
| long-tubing-upper.bottomMd | S30 | 233.11 | m-MDDF (column header S20 'm-MDDF') | 764.797 | derived |
| long-tubing-lower.topMd | S30 | 233.11 | m-MDDF (column header S20 'm-MDDF') | 764.797 | derived |
| long-tubing-lower.bottomMd | S119 | 1521.9 | m-MDDF (column header S20 'm-MDDF') | 4993.11 | derived |
| short-trsv.topMd | K32 | 245.6 | m-MDDF (column header K20 'm-MDDF') | 805.774 | direct |
| short-spm-1.topMd | K40 | 673.01 | m-MDDF (column header K20 'm-MDDF') | 2208.038 | direct |
| short-spm-2.topMd | K43 | 976.05 | m-MDDF (column header K20 'm-MDDF') | 3202.264 | direct |
| short-spm-3.topMd | K46 | 1210.74 | m-MDDF (column header K20 'm-MDDF') | 3972.244 | direct |
| short-spm-4.topMd | K49 | 1335.83 | m-MDDF (column header K20 'm-MDDF') | 4382.644 | direct |
| short-sliding-sleeve.topMd | K54 | 1347.68 | m-MDDF (column header K20 'm-MDDF') | 4421.522 | direct |
| short-landing-nipple.topMd | K67 | 1366.01 | m-MDDF (column header K20 'm-MDDF') | 4481.66 | direct |
| short-pump-out-plug.topMd | K69 | 1367.59 | m-MDDF (column header K20 'm-MDDF') | 4486.844 | direct |
| short-eot.topMd | K73 | 1369.09 | m-MDDF (column header K20 'm-MDDF') | 4491.765 | direct |
| dual-packer.topMd | S60 | 1362.02 | m-MDDF (column header S20 'm-MDDF') | 4468.57 | direct |
| long-trsv.topMd | S30 | 233.11 | m-MDDF (column header S20 'm-MDDF') | 764.797 | direct |
| long-spm-1.topMd | S39 | 659.85 | m-MDDF (column header S20 'm-MDDF') | 2164.862 | direct |
| long-spm-2.topMd | S42 | 964.38 | m-MDDF (column header S20 'm-MDDF') | 3163.976 | direct |
| long-spm-3.topMd | S45 | 1198.29 | m-MDDF (column header S20 'm-MDDF') | 3931.398 | direct |
| long-spm-4.topMd | S48 | 1323.89 | m-MDDF (column header S20 'm-MDDF') | 4343.471 | direct |
| long-sliding-sleeve.topMd | S80 | 1377.18 | m-MDDF (column header S20 'm-MDDF') | 4518.307 | direct |
| gp-packer.topMd | S85 | 1387.0 | m-MDDF (column header S20 'm-MDDF') | 4550.525 | direct |
| mcs-sleeve-1.topMd | S86 | 1393.31 | m-MDDF (column header S20 'm-MDDF') | 4571.227 | direct |
| sealbore-extension-1.topMd | S87 | 1393.82 | m-MDDF (column header S20 'm-MDDF') | 4572.9 | direct |
| isolation-packer-1.topMd | S102 | 1425.1 | m-MDDF (column header S20 'm-MDDF') | 4675.525 | direct |
| mcs-sleeve-2.topMd | S103 | 1427.83 | m-MDDF (column header S20 'm-MDDF') | 4684.482 | direct |
| sealbore-extension-2.topMd | S104 | 1427.34 | m-MDDF (column header S20 'm-MDDF') | 4682.874 | direct |
| isolation-packer-2.topMd | S111 | 1492.87 | m-MDDF (column header S20 'm-MDDF') | 4897.867 | direct |
| mcs-sleeve-3.topMd | S112 | 1495.59 | m-MDDF (column header S20 'm-MDDF') | 4906.791 | direct |
| sealbore-extension-3.topMd | S113 | 1496.1 | m-MDDF (column header S20 'm-MDDF') | 4908.465 | direct |
| long-seal-assembly.topMd | S115 | 1495.83 | m-MDDF (column header S20 'm-MDDF') | 4907.579 | direct |
| long-landing-nipple.topMd | S116 | 1514.25 | m-MDDF (column header S20 'm-MDDF') | 4968.012 | direct |
| long-pump-out-plug.topMd | S117 | 1517.66 | m-MDDF (column header S20 'm-MDDF') | 4979.199 | direct |
| long-perforated-pup.topMd | S118 | 1517.93 | m-MDDF (column header S20 'm-MDDF') | 4980.085 | direct |
| long-mule-shoe.topMd | S119 | 1521.9 | m-MDDF (column header S20 'm-MDDF') | 4993.11 | direct |
| collet-seal-assembly.topMd | S127 | 1530.01 | m-MDDF (column header S20 'm-MDDF') | 5019.718 | direct |
| awg-sump-packer.topMd | S129 | 1528.0 | m-MDDF (column header S20 'm-MDDF') | 5013.123 | direct |
| pbtd.topMd | K137 | 1644.0 | m-MDDF (column header K20 'm-MDDF') | 5393.701 | direct |
| wellInfo.kbElevation | G14 | 33.0 | m-AMSL (elevation) | 108.268 | direct |
| wellInfo.totalDepth | G10 | 1670.0 | m-MDDF (column header K20 'm-MDDF') via formula =K140 | 5479.003 | direct |

## wellInfo mapping

| json_path | mapped_value | cell | original_value | note |
|---|---|---|---|---|
| wellInfo.completionDate | 2020-03-17 | G8 | '2020-03-17T00:00:00' | DIRECT (label B8), ISO |
| wellInfo.completionFluid | 8.5 PPG NaCl filtered to 30 NTU | V16 | '8.5 PPG NaCl filtered to 30 NTU' | DIRECT |
| wellInfo.completionType | Dual string | B19 | 'SHORT STRING' | CLOSEST: SHORT STRING / LONG STRING columns and 'DUAL STRING' packer |
| wellInfo.completionType | Dual string | U19 | 'LONG STRING' | CLOSEST: SHORT STRING / LONG STRING columns and 'DUAL STRING' packer |
| wellInfo.completionType | Dual string | U60 | '9-5/8" X 3-1/2" DUAL STRING, OSD-1, FEED THROUGH PACKER' | CLOSEST: SHORT STRING / LONG STRING columns and 'DUAL STRING' packer |
| wellInfo.groundElevation |  |  |  | NOT IN SOURCE -> template empty value; flagged |
| wellInfo.kbElevation | 108.268 | G14 | '33 m-AMSL' | CLOSEST FIELD: DRILL FLOOR HEIGHT '33 m-AMSL' -> ft |
| wellInfo.maxDeviation | 44.69 deg | G12 | 44.69 (formula =44.69, cached value) | DIRECT (G12 typed value; unit from label 'MAXIMUM DEVIATION' / column J 'Degree') |
| wellInfo.name | Dulang_A23ST2 | B2 | 'WELL NAME: DULANG A23 OP2' | as requested (output_A23ST2); B2 says 'DULANG A23 OP2' — see identity item |
| wellInfo.packerFluid | 8.5 PPG NaCl (+ 0.2% BIOCIDE + 2% OXYGEN SCAVANGER + 2% CORROSION INHIBITOR) | V17 | '8.5 PPG NaCl (+ 0.2% BIOCIDE + 2% OXYGEN SCAVANGER ' | DIRECT (two lines joined) |
| wellInfo.packerFluid | 8.5 PPG NaCl (+ 0.2% BIOCIDE + 2% OXYGEN SCAVANGER + 2% CORROSION INHIBITOR) | V18 | '+ 2% CORROSION INHIBITOR)' | DIRECT (two lines joined) |
| wellInfo.rig | SAPURA BERANI | G9 | 'SAPURA BERANI' | DIRECT |
| wellInfo.surfaceCasing | - |  |  | NOT IN SOURCE -> template empty value; flagged |
| wellInfo.totalDepth | 5479.003 | G10 | 1670 (formula =K140, cached value) | DIRECT: TD 1670.0 m-MDDF (formula =K140) -> ft |
| wellInfo.trajectoryType | Deviated | G12 | 44.69 (formula =44.69, cached value) | supported by maximum deviation 44.69 deg |
| wellInfo.tubing | 3-1/2" 9.2 PPF HP2-13CR-95 JFEBEAR; 3-1/2" 9.2 PPF 13CR-L80 JFEBEAR | V14 | '3-1/2"' | DIRECT: header rows 14-15 joined (row 15 grade conflicts with diagram labels — see report) |
| wellInfo.tubing | 3-1/2" 9.2 PPF HP2-13CR-95 JFEBEAR; 3-1/2" 9.2 PPF 13CR-L80 JFEBEAR | X14 | '9.2 PPF' | DIRECT: header rows 14-15 joined (row 15 grade conflicts with diagram labels — see report) |
| wellInfo.tubing | 3-1/2" 9.2 PPF HP2-13CR-95 JFEBEAR; 3-1/2" 9.2 PPF 13CR-L80 JFEBEAR | Y14 | 'HP2-13CR-95' | DIRECT: header rows 14-15 joined (row 15 grade conflicts with diagram labels — see report) |
| wellInfo.tubing | 3-1/2" 9.2 PPF HP2-13CR-95 JFEBEAR; 3-1/2" 9.2 PPF 13CR-L80 JFEBEAR | Z14 | 'JFEBEAR' | DIRECT: header rows 14-15 joined (row 15 grade conflicts with diagram labels — see report) |
| wellInfo.tubing | 3-1/2" 9.2 PPF HP2-13CR-95 JFEBEAR; 3-1/2" 9.2 PPF 13CR-L80 JFEBEAR | V15 | '3-1/2"' | DIRECT: header rows 14-15 joined (row 15 grade conflicts with diagram labels — see report) |
| wellInfo.tubing | 3-1/2" 9.2 PPF HP2-13CR-95 JFEBEAR; 3-1/2" 9.2 PPF 13CR-L80 JFEBEAR | X15 | '9.2 PPF' | DIRECT: header rows 14-15 joined (row 15 grade conflicts with diagram labels — see report) |
| wellInfo.tubing | 3-1/2" 9.2 PPF HP2-13CR-95 JFEBEAR; 3-1/2" 9.2 PPF 13CR-L80 JFEBEAR | Y15 | '13CR-L80' | DIRECT: header rows 14-15 joined (row 15 grade conflicts with diagram labels — see report) |
| wellInfo.tubing | 3-1/2" 9.2 PPF HP2-13CR-95 JFEBEAR; 3-1/2" 9.2 PPF 13CR-L80 JFEBEAR | Z15 | 'JFEBEAR' | DIRECT: header rows 14-15 joined (row 15 grade conflicts with diagram labels — see report) |
| wellInfo.wellhead | SOLAR ALERT | G15 | 'SOLAR ALERT' | DIRECT |

## Components

| id | type | name | topMd_ft | bottomMd_ft | OD_in | ID_in | weight | grade | interval | estimated |
|---|---|---|---|---|---|---|---|---|---|---|
| conductor | casing | CONDUCTOR |  |  | 26 |  |  |  |  |  |
| int-casing | casing | INT. CASING |  |  | 13.375 |  | 68.0 | N-80 |  |  |
| casing | casing | CASING |  |  | 9.625 |  | 47.0 | L-80 |  |  |
| liner | liner | LINER |  |  | 5.5 |  | 17.0 | L-80 |  |  |
| short-tubing-upper | short-string | 3-1/2", 9.2#, HP2-13CR-95, JFE Bear |  | 805.774 | 3.5 |  | 9.2 | HP2-13CR-95 |  | bottomMd |
| short-tubing-lower | short-string | 3-1/2" 9.2 PPF HP2 13Cr95 JFE BEAR | 805.774 | 4491.765 | 3.5 |  | 9.2 |  |  | topMd,bottomMd |
| long-tubing-upper | long-string | 3-1/2", 9.2#, HP2-13CR-95, JFE Bear |  | 764.797 | 3.5 |  | 9.2 | HP2-13CR-95 |  | bottomMd |
| long-tubing-lower | long-string | 3-1/2" 9.2 PPF HP2 13Cr95 JFE BEAR | 764.797 | 4993.11 | 3.5 |  | 9.2 |  |  | topMd,bottomMd |
| short-flow-coupling-1 | flow-coupling | FLOW COUPLING, 3-1/2 9.2# JFEBEAR B X P, 3FT, 13CR-L80 |  |  | 3.92 | 2.88 |  |  |  |  |
| short-trsv | trsv | TRMAXX-5E, 3-1/2 9.2# JFEBEAR B X P, 2.812 X-TYPE, 13CR-L80 | 805.774 |  | 5.03 | 2.813 |  |  |  |  |
| short-flow-coupling-2 | flow-coupling | FLOW COUPLING, 3-1/2 9.2# JFEBEAR B X P, 3FT, 13CR-L80 |  |  | 3.92 | 2.88 |  |  |  |  |
| short-spm-1 | spm | 3-1/2" SPM (Dummy) | 2208.038 |  | 5.23 | 2.9 |  |  |  |  |
| short-spm-2 | spm | 3-1/2" SPM (Dummy) | 3202.264 |  | 5.23 | 2.9 |  |  |  |  |
| short-spm-3 | spm | 3-1/2" SPM (Dummy) | 3972.244 |  | 5.23 | 2.9 |  |  |  |  |
| short-spm-4 | spm | 3-1/2" SPM (Dummy) | 4382.644 |  | 5.23 | 2.9 |  |  |  |  |
| short-sliding-sleeve | sliding-sleeve | SLIDING SLEEVE DOOR, 3-1/2 9.2# JFEBEAR B X P, 13CR-L80 | 4421.522 |  | 4.545 | 2.813 |  |  |  |  |
| short-landing-nipple | landing-nipple | LANDING NIPPLE, 2.750 OSN,  3-1/2 9.2# JFEBEAR B X P, 13CR-L80 | 4481.66 |  | 4.0 | 2.64 |  |  |  |  |
| short-pump-out-plug | mechanical-plug | PUMP OUT PLUG, PE-635, 3-1/2 9.2# JFEBEAR B X P, 13CR-L80 | 4486.844 |  | 4.545 | 1.485 |  |  |  | type |
| short-perforated-pup | pup-joint | PERFORATED PUP |  |  | 3.903 | 2.953 |  |  |  | type |
| short-eot | eot | EOT | 4491.765 |  |  |  |  |  |  |  |
| dual-packer | packer | 9-5/8" X 3-1/2" DUAL STRING, OSD-1, FEED THROUGH PACKER | 4468.57 |  | 8.5 | 2.93 |  |  |  |  |
| long-flow-coupling-1 | flow-coupling | FLOW COUPLING, 3-1/2 9.2# JFEBEAR B X P, 3FT, 13CR-L80 |  |  | 3.92 | 2.88 |  |  |  |  |
| long-trsv | trsv | TRMAXX-5E, 3-1/2 9.2# JFEBEAR B X P, 2.812 X-TYPE, 13CR-L80 | 764.797 |  | 5.03 | 2.813 |  |  |  |  |
| long-flow-coupling-2 | flow-coupling | 3-1/2" FLOW COUPLING |  |  | 3.92 | 2.88 |  |  |  |  |
| long-spm-1 | spm | 3-1/2" SPM (GLV  : PPO 10/64" ) - 1st | 2164.862 |  | 5.23 | 2.9 |  |  |  |  |
| long-spm-2 | spm | 3-1/2" SPM (GLOV : PPO10/64") - 2nd | 3163.976 |  | 5.23 | 2.9 |  |  |  |  |
| long-spm-3 | spm | 3-1/2" SPM (GLOV : PPO 10/64") - 3rd | 3931.398 |  | 5.23 | 2.9 |  |  |  |  |
| long-spm-4 | spm | 3-1/2" SPM (ORIFICE : 12/64) -4th | 4343.471 |  | 5.23 | 2.9 |  |  |  |  |
| long-sliding-sleeve | sliding-sleeve | SLIDING SLEEVE DOOR, 3-1/2 9.2# JFEBEAR B X P, 13CR-L80 | 4518.307 |  | 4.545 | 2.813 |  |  |  |  |
| gp-packer | packer | 9 5/8" VGH VERSA-TRIEVE GP PACKER | 4550.525 |  | 8.45 | 6.0 |  |  |  |  |
| mcs-sleeve-1 | sliding-sleeve | MCS CLOSING SLEEVE, 13CR-L80 | 4571.227 |  | 7.28 | 4.75 |  |  |  | type |
| sealbore-extension-1 | pbr | SEALBORE EXTENSION, 13CR-L80 (0.65 m length) | 4572.9 |  | 6.02 | 4.75 |  |  |  | type |
| screen-1 | screen | 5-1/2" WIRE-WRAP SCREENS |  |  | 8.25 | 4.818 |  |  |  |  |
| isolation-packer-1 | packer | 9-5/8 ISOLATION PACKER 13CR-L80 | 4675.525 |  | 8.45 | 4.75 |  |  |  |  |
| mcs-sleeve-2 | sliding-sleeve | MCS CLOSING SLEEVE, 13CR-L80 | 4684.482 |  | 7.28 | 4.75 |  |  |  | type |
| sealbore-extension-2 | pbr | SEALBORE EXTENSION, 13CR-L80 (0.65 m length) | 4682.874 |  | 6.02 | 4.75 |  |  |  | type |
| isolation-packer-2 | packer | 9-5/8 ISOLATION PACKER 13CR-L80 | 4897.867 |  | 8.45 | 4.75 |  |  |  |  |
| mcs-sleeve-3 | sliding-sleeve | MCS CLOSING SLEEVE, 13CR-L80 | 4906.791 |  | 7.28 | 4.75 |  |  |  | type |
| sealbore-extension-3 | pbr | SEALBORE EXTENSION, 13CR-L80 (0.65 m length) | 4908.465 |  | 6.02 | 4.75 |  |  |  | type |
| long-seal-assembly | seal-assembly | SEAL ASSEMBLY, 4.750", 3-1/2 9.2# JFEBEAR B X P, 13CR-L80 (5ea) | 4907.579 |  | 4.71 | 2.953 |  |  |  |  |
| long-landing-nipple | landing-nipple | LANDING NIPPLE, 2.750 OSN,  3-1/2 9.2# JFEBEAR B X P, 13CR-L80 | 4968.012 |  | 4.0 | 2.64 |  |  |  |  |
| long-pump-out-plug | mechanical-plug | PUMP OUT PLUG, PE-635, 3-1/2 9.2# JFEBEAR B X P, 13CR-L80 | 4979.199 |  | 4.545 | 1.485 |  |  |  | type |
| long-perforated-pup | pup-joint | PERFORATED PUP | 4980.085 |  | 3.903 | 2.953 |  |  |  | type |
| long-mule-shoe | mule-shoe | SELF ALIGNING MULE SHOE, 3-1/2 9.2# JFEBEAR B X P, 13CR-L80 | 4993.11 |  | 3.97 | 2.94 |  |  |  |  |
| screen-2 | screen | 5-1/2" WIRE-WRAP SCREENS |  |  | 8.25 | 4.818 |  |  |  |  |
| collet-seal-assembly | seal-assembly | COLLET SEAL ASSEMBLY | 5019.718 |  | 6.6 | 4.818 |  |  |  |  |
| awg-sump-packer | packer | AWG SUMP PACKER | 5013.123 |  | 8.22 | 6.0 |  |  |  |  |
| pbtd | pbtd | PBTD | 5393.701 |  |  |  |  |  |  |  |

Full cell-by-cell traceability (every mapped field, original value, unit, interpretation): `Dulang_A23ST2_traceability.csv`; each component's `source` string in the JSON names the cells too.

## Terminology → template type

| source term | cell(s) | template type | basis |
|---|---|---|---|
| CONDUCTOR | P9 | casing | matches source term |
| INT. CASING | P11 | casing | matches source term |
| CASING | P12 | casing | matches source term |
| LINER | P13 | liner | matches source term |
| 3-1/2", 9.2#, HP2-13CR-95, JFE Bear | F26 | short-string | matches source term |
| 3-1/2" 9.2 PPF HP2 13Cr95 JFE BEAR | F34 | short-string | matches source term |
| 3-1/2", 9.2#, HP2-13CR-95, JFE Bear | U26 | long-string | matches source term |
| 3-1/2" 9.2 PPF HP2 13Cr95 JFE BEAR | U32 | long-string | matches source term |
| FLOW COUPLING, 3-1/2 9.2# JFEBEAR B X P, 3FT, 13CR-L80 | F31 | flow-coupling | matches source term |
| TRMAXX-5E, 3-1/2 9.2# JFEBEAR B X P, 2.812 X-TYPE, 13CR-L80 | F32 | trsv | matches source term |
| FLOW COUPLING, 3-1/2 9.2# JFEBEAR B X P, 3FT, 13CR-L80 | F33 | flow-coupling | matches source term |
| 3-1/2" SPM (Dummy) | F40 | spm | matches source term |
| 3-1/2" SPM (Dummy) | F43 | spm | matches source term |
| 3-1/2" SPM (Dummy) | F46 | spm | matches source term |
| 3-1/2" SPM (Dummy) | F49 | spm | matches source term |
| SLIDING SLEEVE DOOR, 3-1/2 9.2# JFEBEAR B X P, 13CR-L80 | F54 | sliding-sleeve | matches source term |
| LANDING NIPPLE, 2.750 OSN,  3-1/2 9.2# JFEBEAR B X P, 13CR-L80 | F67 | landing-nipple | matches source term |
| PUMP OUT PLUG, PE-635, 3-1/2 9.2# JFEBEAR B X P, 13CR-L80 | F69 | mechanical-plug | CLOSEST — pump-out plug (plug removed by pressure) -> Mechanical Plug is the closest type |
| PERFORATED PUP | F70 | pup-joint | CLOSEST — perforated pup -> Pup Joint (a pup joint with holes) |
| EOT | F73 | eot | matches source term |
| 9-5/8" X 3-1/2" DUAL STRING, OSD-1, FEED THROUGH PACKER | U60 | packer | matches source term |
| FLOW COUPLING, 3-1/2 9.2# JFEBEAR B X P, 3FT, 13CR-L80 | U29 | flow-coupling | matches source term |
| TRMAXX-5E, 3-1/2 9.2# JFEBEAR B X P, 2.812 X-TYPE, 13CR-L80 | U30 | trsv | matches source term |
| 3-1/2" FLOW COUPLING | U31 | flow-coupling | matches source term |
| 3-1/2" SPM (GLV  : PPO 10/64" ) - 1st | U39 | spm | matches source term |
| 3-1/2" SPM (GLOV : PPO10/64") - 2nd | U42 | spm | matches source term |
| 3-1/2" SPM (GLOV : PPO 10/64") - 3rd | U45 | spm | matches source term |
| 3-1/2" SPM (ORIFICE : 12/64) -4th | U48 | spm | matches source term |
| SLIDING SLEEVE DOOR, 3-1/2 9.2# JFEBEAR B X P, 13CR-L80 | U80 | sliding-sleeve | matches source term |
| 9 5/8" VGH VERSA-TRIEVE GP PACKER | U85 | packer | matches source term |
| MCS CLOSING SLEEVE, 13CR-L80 | U86 | sliding-sleeve | CLOSEST — closing sleeve (shifts to open/close ports) -> Sliding Sleeve is the closest type |
| SEALBORE EXTENSION, 13CR-L80 (0.65 m length) | U87 | pbr | CLOSEST — sealbore extension (polished seal bore) -> PBR is the closest type |
| 5-1/2" WIRE-WRAP SCREENS | U95 | screen | matches source term |
| 9-5/8 ISOLATION PACKER 13CR-L80 | U102 | packer | matches source term |
| MCS CLOSING SLEEVE, 13CR-L80 | U103 | sliding-sleeve | CLOSEST — closing sleeve (shifts to open/close ports) -> Sliding Sleeve is the closest type |
| SEALBORE EXTENSION, 13CR-L80 (0.65 m length) | U104 | pbr | CLOSEST — sealbore extension (polished seal bore) -> PBR is the closest type |
| 9-5/8 ISOLATION PACKER 13CR-L80 | U111 | packer | matches source term |
| MCS CLOSING SLEEVE, 13CR-L80 | U112 | sliding-sleeve | CLOSEST — closing sleeve (shifts to open/close ports) -> Sliding Sleeve is the closest type |
| SEALBORE EXTENSION, 13CR-L80 (0.65 m length) | U113 | pbr | CLOSEST — sealbore extension (polished seal bore) -> PBR is the closest type |
| SEAL ASSEMBLY, 4.750", 3-1/2 9.2# JFEBEAR B X P, 13CR-L80 (5ea) | U115 | seal-assembly | matches source term |
| LANDING NIPPLE, 2.750 OSN,  3-1/2 9.2# JFEBEAR B X P, 13CR-L80 | U116 | landing-nipple | matches source term |
| PUMP OUT PLUG, PE-635, 3-1/2 9.2# JFEBEAR B X P, 13CR-L80 | U117 | mechanical-plug | CLOSEST — pump-out plug (plug removed by pressure) -> Mechanical Plug is the closest type |
| PERFORATED PUP | U118 | pup-joint | CLOSEST — perforated pup -> Pup Joint (a pup joint with holes) |
| SELF ALIGNING MULE SHOE, 3-1/2 9.2# JFEBEAR B X P, 13CR-L80 | U119 | mule-shoe | matches source term |
| 5-1/2" WIRE-WRAP SCREENS | U121 | screen | matches source term |
| COLLET SEAL ASSEMBLY | U127 | seal-assembly | matches source term |
| AWG SUMP PACKER | U129 | packer | matches source term |
| PBTD | I137 | pbtd | matches source term |

## Items requiring review

### Identity (1)

| location | detail |
|---|---|
| file name / B2 / sheet title | File says A23ST2, B2 'WELL NAME: DULANG A23 OP2', sheet title 'OP2 dual A23St1 STMZ E51  FID' says 'A23St1'. Output named Dulang_A23ST2 per the requested folder — confirm which wellbore this schematic is. |

### Conflict (4)

| location | detail |
|---|---|
| Y15 vs F34 / U32 | Lower tubing grade: header Y15 '13CR-L80' vs diagram labels F34 '3-1/2" 9.2 PPF HP2 13Cr95 JFE BEAR' and U32 '3-1/2" 9.2 PPF HP2 13Cr95 JFE BEAR' -> grade left empty on both lower tubing components. |
| AB13 / AB14 / AB15 | Header 'DEPTH' column holds lengths/sums: AB13 144 (=ROUNDUP(S127-S85,0)), AB14 478.71 (=K32+S30), AB15 2413 (=ROUNDUP((K73-K32)+(S119-S30),0)) -> none used as depths. |
| G17 vs U144 | G17 'PRAVIN NAIR VENUGOPALAN (FEB 2023)' vs U144 'Date: Feb 2024' — revision date differs. |
| B18 vs K20/S20 | Note B18 'Note: All depths are in meter from RKB to top of equipment.' (RKB) vs depth columns 'm-MDDF'. Kept as m-MDDF. |

### Missing (15)

| location | detail |
|---|---|
| AB9, AB11, AB12 | Casing depths printed as placeholder 'xx m MDDF' -> null. |
| P10 | SURF. CASING row has no size/weight/grade/depth -> no component; wellInfo.surfaceCasing '-'. |
| liner | No liner top (TOL) or shoe depth printed -> null. |
| groundElevation | Not in source -> null. |
| P9 | conductor: no depth printed for 'CONDUCTOR' -> topMd/bottomMd null |
| P11 | int-casing: no depth printed for 'INT. CASING' -> topMd/bottomMd null |
| P12 | casing: no depth printed for 'CASING' -> topMd/bottomMd null |
| P13 | liner: no depth printed for 'LINER' -> topMd/bottomMd null |
| F31 | short-flow-coupling-1: no depth printed for 'FLOW COUPLING, 3-1/2 9.2# JFEBEAR B X P, 3FT, 13CR-L80' -> topMd/bottomMd null |
| F33 | short-flow-coupling-2: no depth printed for 'FLOW COUPLING, 3-1/2 9.2# JFEBEAR B X P, 3FT, 13CR-L80' -> topMd/bottomMd null |
| F70 | short-perforated-pup: no depth printed for 'PERFORATED PUP' -> topMd/bottomMd null |
| U29 | long-flow-coupling-1: no depth printed for 'FLOW COUPLING, 3-1/2 9.2# JFEBEAR B X P, 3FT, 13CR-L80' -> topMd/bottomMd null |
| U31 | long-flow-coupling-2: no depth printed for '3-1/2" FLOW COUPLING' -> topMd/bottomMd null |
| U95 | screen-1: no depth printed for '5-1/2" WIRE-WRAP SCREENS' -> topMd/bottomMd null |
| U121 | screen-2: no depth printed for '5-1/2" WIRE-WRAP SCREENS' -> topMd/bottomMd null |

### Unclear (5)

| location | detail |
|---|---|
| G12 | maxDeviation 44.69 is a typed constant (=44.69); the deviation column shows at most 33.4 (J95/J121/J140). Mapped as printed in G12. |
| V9 | Conductor size V9 '26" DP': meaning of 'DP' not stated; OD 26 in mapped. |
| U60 / U61 | U61 '3-1/2 9.2# JFEBEAR B X P, 3FT, 13CR-L80' below the dual packer may be part of the packer description or a separate 3 ft item -> not mapped. |
| kbElevation G14 | G14 '33 m-AMSL' (DRILL FLOOR HEIGHT) mapped to kbElevation (closest field), converted to ft. |
| short/long tubing depths | Tubing intervals derived from the AB15 length formula and the 'From Tubing Hanger to TRSV' labels (flagged estimated). |

### Anomaly (3)

| location | detail |
|---|---|
| S103 vs S104 | S104 1427.34 shallower than S103 1427.83 listed above it. Kept as printed. |
| S127 vs S129 | Sump packer S129 1528 shallower than collet seal S127 1530.01 above it. Kept as printed. |
| D71 | Ball catcher MIN ID D71 0 (0.000). |

## Unmapped information

Relevant source content with no suitable template field (or not mapped because its meaning/association is not established by the source):

| location | content | reason |
|---|---|---|
| B3 | LOCATION / SLOT: DULANG A / 22 | slot number |
| G13 | WATER DEPTH 77 m | no template field |
| B20 / U20 | SHORT STRING (OIL PRODUCER); LONG STRING (OIL PRODUCER) | well function — no template field |
| Z11-Z15 | threads: Z11 BTC, Z12 VAM TOP, Z13 JFE BEAR, Z14 JFEBEAR, Z15 JFEBEAR | no thread field |
| F41, F44, F47, F50 | DV (in DROID, set on 21  Feb 2022) (same text under each short-string SPM) | valve status — no field |
| U40, U43, U46 | TRO : 650, TRO : 715, TRO : 770 | gas-lift valve TRO settings — no field |
| I56 / K56 | Casing PIP Tag at 1356.6 m-MDDF (4450.787 ft) | depth marker — no Component_Type |
| F71 / K71 / B71 / D71 | BALL CATCHER at 1367.86 m-MDDF (4487.73 ft), OD 3.903 in, ID 0 in | no Component_Type for a ball catcher |
| U82 / S82 | NO-GO LOCTR,6.000 BORE, 3-1/2 9.2# JFEBEAR B X P, 13CR-L80 at 1385.17 m-MDDF (4544.521 ft), ID 2.953 in, OD 6.49 in | no-go locator: no matching Component_Type |
| U61 | 3-1/2 9.2# JFEBEAR B X P, 3FT, 13CR-L80 | see Unclear |
| drawing K79:O81 | text box 'CLOSED (Nov 2024)' beside the long-string sliding sleeve U80 (row 80) | status associated by position only |
| B93-B97 | Zone E14B: 1416m - 1423 m MDDF (7 m) (4645.669-4668.635 ft MDDF); 1259.11 m - 1264.98 m TVD; B96 6.120 (EMW formula); 1317.16 psi, 194 deg F; J95 33 deg | reservoir zone interval: not stated as perforation/formation top; association with screens by position only |
| B119-B123 | Zone E34A: 1517 m - 1524 m MDDF (7m) (4977.034-5000.0 ft MDDF); 1343.93 m - 1349.81 m TVD; B122 6.307 (EMW formula); 1448.81 psi, 206 deg F; J121 32.8 deg | as above |
| K92, K98, K118, K124 | K92 1411.934, K98 1425.1, K118 1508.72, K124 1530.01 m-MDDF | depths with no label in the row |
| I139 | 9 5/8" Casing Shoe | casing shoe label without a depth |
| J140 | deviation 33.4 deg at TD row | no azimuth -> no survey station |
| G17, U142, U144 | PRAVIN NAIR VENUGOPALAN (FEB 2023); Updated By: Pravin Nair Venugopalan; Date: Feb 2024 | document control |

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
| /Volumes/homes/public/DulangOSS/20261016/DULANG_WELLS/Well_Schematic/1. Well Schematic/Dulang A/Dulang A23ST2 Well Schematic.xlsx | 62c51c2ca32a18c562531c69e642e372db23b7014c28350a9ada06c28cf17c5c | 62c51c2ca32a18c562531c69e642e372db23b7014c28350a9ada06c28cf17c5c | YES |
| /Users/farahyusof/Documents/GitHub/NoramiraSazuki/AinnyFarah/WELL_SCHEMATIC/well-schematic-template-json.rtf | c7db15efb1917dfa76fa3102e270fe17115a8561f24eb0a155239870d079e712 | c7db15efb1917dfa76fa3102e270fe17115a8561f24eb0a155239870d079e712 | YES |
| /Users/farahyusof/Documents/GitHub/NoramiraSazuki/AinnyFarah/WELL_SCHEMATIC/well-schematic-sample-template.xlsx | d377a6310d42fda8c95e85db03f2416cb4c1b6f37f53920052331ce814f6d494 | d377a6310d42fda8c95e85db03f2416cb4c1b6f37f53920052331ce814f6d494 | YES |
