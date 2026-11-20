# DULANG-A26S1.wsch.json — extraction notes

Input: `Dulang A26ST1 Well Schematic.xlsx` (read only, SHA-256 7af7a66f291cf2f33b2f3ba4cf2a6ac417ce5967e6571f5c37549b0c960ef2c6)  
Template: `THF M DULANG-B30ST1-2-7.wsch 1.json` (read only, SHA-256 214c1e85160c4b581cae1406146fb538d6c9f3ec5413bafdba1d3a19c05c6257)  
Well names: `EDAFY-Data_Dulang_Well.xlsx` (read only, SHA-256 eb874ebc996350ae1d25f3cb1020d6668d7887418d1066788473d04030fa3dac)

## Information with no slot in the template format

- Well name: input label 'WELL NAME: A26ST1' -> EDAFY name 'DULANG-A26S1' (UWI 1000000730; EDAFY base depth 1768 m = input TD 1768 m).
- Units: all depths converted m -> ft (x 3.28084, 3 dp): topMd, bottomMd, totalDepth, dfToThf, floatEquipment md. topMdM / bottomMdM keep the printed metres; diameters stay in inches; source texts quote the original metre values.
- Water depth: 76.5 m = 250.984 ft (F8) — no wellInfo field in the template.
- PBTD: label only (E87), no depth printed — no pbtd component created.
- Tubular OD / ID / drift from the header table are in the string components' source text; drift has no field.
- m-MDTHF depth column (R) is MDDF - DFE-THF; kept in each component's source text.
- Updated by: Pravin Nair Venugopalan (U91); date cell U92 is =TODAY() (no fixed date).
- Completion type / schematic title: not stated in the input -> '-'.
- Pictures: equipment/perforation symbol images and a radioactive-hazard symbol (no data).
- Items with no depth printed are kept with topMd/bottomMd null (flow couplings, tubing sections, NO-GO LOCATOR, 4-1/2" x 7" VTA packer, 3.88" seal bore extension, 3-way sub, SDG mule shoe, tubing strings).
- Row 31: MIN ID 2.88 / MAX OD 5.313 printed on a row with no equipment label (between row 29 'SPM, 3-1/2", 13CR L80 JFEFOX BXP (PTRO 767.99psi, 10/64")' and row 32 'SPM, 3-1/2", 13CR L80 JFEFOX BXP (PTRO 782.04 psi, 10/64")'); not assigned to a component.
- Row 35: MIN ID 2.88 / MAX OD 5.313 printed on a row with no equipment label (between row 32 'SPM, 3-1/2", 13CR L80 JFEFOX BXP (PTRO 782.04 psi, 10/64")' and row 36 'SPM, 3-1/2", 13CR L80 JFEFOX BXP (792.80 psi, 10/64")'); not assigned to a component.
- Row 58: MIN ID 2.813 / MAX OD 4.04 printed on a row with no equipment label (between row 57 '3 Way Sub: 4-1/2" x 4-1/2" x 2-7/8"' and row 59 'OIL SWELL PACKER'); not assigned to a component.
- Drawing text box at row 67 not mapped to a component: '2.29” Gauge Ring & 2.25" LIB HUD @1578m WLTHF(~1593m MDDF). Bailer recovered sand (Oct 2023)' (depth is wireline WLTHF, not MD).
