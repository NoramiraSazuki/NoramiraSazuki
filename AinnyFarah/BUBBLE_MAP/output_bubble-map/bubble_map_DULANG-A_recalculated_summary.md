# bubble_map_DULANG-A_recalculated

- Source (unchanged): `bubble_map_DULANG-A.xlsx`; output: `bubble_map_DULANG-A_recalculated.xlsx`
- Entities 140, data rows 28607, separator rows 139 (R/S/T left empty)
- Formulas written: R 28607, S 28607, T 28607
- Duplicate entity/month records: none; rows out of month order: none
- Opening cumulative volumes: none documented - every entity starts from 0 at its first record, so totals cover only the history in the sheet.

## Np (R)
- entities complete: 120
- gap: producing days missing: 14
- entities blank from a gap onward: 20
- no producing days before history starts: 2016
- gap: rate missing with PD > 0: 3
- missing rate on first record: 3

## Wp (S)
- entities complete: 96
- gap: rate missing with PD > 0: 25
- entities blank from a gap onward: 44
- gap: producing days missing: 14
- no producing days before history starts: 2016
- missing rate on first record: 5

## Productivity (T = R/L)
- error results: {'#DIV/0!': 606, '#VALUE!': 5616} (#VALUE! where Np is blank, #DIV/0! where cumulative PD = 0)

## Column L check
- rows with a discrepancy: 834
  - row 6813 DULG-A010S02-LS-S3-E12/13: formula `=IF(AND(B6813=B6812,D6814=D6812,A6813=A6812),L6812,0)+N(K6813)`, value 13, expected 13.0
  - row 6814 DULG-A010S02-LS-S3-E12/13: formula `=IF(AND(B6814=B6813,D6815=D6814,A6814=A6813),L6813,0)+N(K6814)`, value 42, expected 42.0
  - row 6815 DULG-A010S02-LS-S3-E12/13: formula `=IF(AND(B6815=B6814,D6816=D6815,A6815=A6814),L6814,0)+N(K6815)`, value 73, expected 73.0
  - row 6816 DULG-A010S02-LS-S3-E12/13: formula `=IF(AND(B6816=B6815,D6817=D6816,A6816=A6815),L6815,0)+N(K6816)`, value 102, expected 102.0
  - row 6817 DULG-A010S02-LS-S3-E12/13: formula `=IF(AND(B6817=B6816,D6818=D6817,A6817=A6816),L6816,0)+N(K6817)`, value 133, expected 133.0
  - row 6818 DULG-A010S02-LS-S3-E12/13: formula `=IF(AND(B6818=B6817,D6819=D6818,A6818=A6817),L6817,0)+N(K6818)`, value 161, expected 161.0
  - row 6819 DULG-A010S02-LS-S3-E12/13: formula `=IF(AND(B6819=B6818,D6820=D6819,A6819=A6818),L6818,0)+N(K6819)`, value 190, expected 190.0
  - row 6820 DULG-A010S02-LS-S3-E12/13: formula `=IF(AND(B6820=B6819,D6821=D6820,A6820=A6819),L6819,0)+N(K6820)`, value 190, expected 190.0
  - row 6821 DULG-A010S02-LS-S3-E12/13: formula `=IF(AND(B6821=B6820,D6822=D6821,A6821=A6820),L6820,0)+N(K6821)`, value 190, expected 190.0
  - row 6822 DULG-A010S02-LS-S3-E12/13: formula `=IF(AND(B6822=B6821,D6823=D6822,A6822=A6821),L6821,0)+N(K6822)`, value 190, expected 190.0
  - row 6823 DULG-A010S02-LS-S3-E12/13: formula `=IF(AND(B6823=B6822,D6824=D6823,A6823=A6822),L6822,0)+N(K6823)`, value 190, expected 190.0
  - row 6824 DULG-A010S02-LS-S3-E12/13: formula `=IF(AND(B6824=B6823,D6825=D6824,A6824=A6823),L6823,0)+N(K6824)`, value 219, expected 219.0
  - row 6825 DULG-A010S02-LS-S3-E12/13: formula `=IF(AND(B6825=B6824,D6826=D6825,A6825=A6824),L6824,0)+N(K6825)`, value 246, expected 246.0
  - row 6826 DULG-A010S02-LS-S3-E12/13: formula `=IF(AND(B6826=B6825,D6827=D6826,A6826=A6825),L6825,0)+N(K6826)`, value 273, expected 273.0
  - row 6827 DULG-A010S02-LS-S3-E12/13: formula `=IF(AND(B6827=B6826,D6828=D6827,A6827=A6826),L6826,0)+N(K6827)`, value 304, expected 304.0
  - row 6828 DULG-A010S02-LS-S3-E12/13: formula `=IF(AND(B6828=B6827,D6829=D6828,A6828=A6827),L6827,0)+N(K6828)`, value 314, expected 314.0
  - row 6829 DULG-A010S02-LS-S3-E12/13: formula `=IF(AND(B6829=B6828,D6830=D6829,A6829=A6828),L6828,0)+N(K6829)`, value 318, expected 318.0
  - row 6830 DULG-A010S02-LS-S3-E12/13: formula `=IF(AND(B6830=B6829,D6831=D6830,A6830=A6829),L6829,0)+N(K6830)`, value 346, expected 346.0
  - row 6831 DULG-A010S02-LS-S3-E12/13: formula `=IF(AND(B6831=B6830,D6832=D6831,A6831=A6830),L6830,0)+N(K6831)`, value 374, expected 374.0
  - row 6832 DULG-A010S02-LS-S3-E12/13: formula `=IF(AND(B6832=B6831,D6833=D6832,A6832=A6831),L6831,0)+N(K6832)`, value 405, expected 405.0
  - row 6833 DULG-A010S02-LS-S3-E12/13: formula `=IF(AND(B6833=B6832,D6834=D6833,A6833=A6832),L6832,0)+N(K6833)`, value 435, expected 435.0
  - row 6834 DULG-A010S02-LS-S3-E12/13: formula `=IF(AND(B6834=B6833,D6835=D6834,A6834=A6833),L6833,0)+N(K6834)`, value 466, expected 466.0
  - row 6835 DULG-A010S02-LS-S3-E12/13: formula `=IF(AND(B6835=B6834,D6836=D6835,A6835=A6834),L6834,0)+N(K6835)`, value 496, expected 496.0
  - row 6836 DULG-A010S02-LS-S3-E12/13: formula `=IF(AND(B6836=B6835,D6837=D6836,A6836=A6835),L6835,0)+N(K6836)`, value 527, expected 527.0
  - row 6837 DULG-A010S02-LS-S3-E12/13: formula `=IF(AND(B6837=B6836,D6838=D6837,A6837=A6836),L6836,0)+N(K6837)`, value 542, expected 542.0
  - row 6838 DULG-A010S02-LS-S3-E12/13: formula `=IF(AND(B6838=B6837,D6839=D6838,A6838=A6837),L6837,0)+N(K6838)`, value 566, expected 566.0
  - row 6839 DULG-A010S02-LS-S3-E12/13: formula `=IF(AND(B6839=B6838,D6840=D6839,A6839=A6838),L6838,0)+N(K6839)`, value 597, expected 597.0
  - row 6840 DULG-A010S02-LS-S3-E12/13: formula `=IF(AND(B6840=B6839,D6841=D6840,A6840=A6839),L6839,0)+N(K6840)`, value 627, expected 627.0
  - row 6841 DULG-A010S02-LS-S3-E12/13: formula `=IF(AND(B6841=B6840,D6842=D6841,A6841=A6840),L6840,0)+N(K6841)`, value 658, expected 658.0
  - row 6842 DULG-A010S02-LS-S3-E12/13: formula `=IF(AND(B6842=B6841,D6843=D6842,A6842=A6841),L6841,0)+N(K6842)`, value 689, expected 689.0
  - row 6843 DULG-A010S02-LS-S3-E12/13: formula `=IF(AND(B6843=B6842,D6844=D6843,A6843=A6842),L6842,0)+N(K6843)`, value 717, expected 717.0
  - row 6844 DULG-A010S02-LS-S3-E12/13: formula `=IF(AND(B6844=B6843,D6845=D6844,A6844=A6843),L6843,0)+N(K6844)`, value 748, expected 748.0
  - row 6845 DULG-A010S02-LS-S3-E12/13: formula `=IF(AND(B6845=B6844,D6846=D6845,A6845=A6844),L6844,0)+N(K6845)`, value 778, expected 778.0
  - row 6846 DULG-A010S02-LS-S3-E12/13: formula `=IF(AND(B6846=B6845,D6847=D6846,A6846=A6845),L6845,0)+N(K6846)`, value 805, expected 805.0
  - row 6847 DULG-A010S02-LS-S3-E12/13: formula `=IF(AND(B6847=B6846,D6848=D6847,A6847=A6846),L6846,0)+N(K6847)`, value 824, expected 824.0
  - row 6848 DULG-A010S02-LS-S3-E12/13: formula `=IF(AND(B6848=B6847,D6849=D6848,A6848=A6847),L6847,0)+N(K6848)`, value 855, expected 855.0
  - row 6849 DULG-A010S02-LS-S3-E12/13: formula `=IF(AND(B6849=B6848,D6850=D6849,A6849=A6848),L6848,0)+N(K6849)`, value 886, expected 886.0
  - row 6850 DULG-A010S02-LS-S3-E12/13: formula `=IF(AND(B6850=B6849,D6851=D6850,A6850=A6849),L6849,0)+N(K6850)`, value 916, expected 916.0
  - row 6851 DULG-A010S02-LS-S3-E12/13: formula `=IF(AND(B6851=B6850,D6852=D6851,A6851=A6850),L6850,0)+N(K6851)`, value 947, expected 947.0
  - row 6852 DULG-A010S02-LS-S3-E12/13: formula `=IF(AND(B6852=B6851,D6853=D6852,A6852=A6851),L6851,0)+N(K6852)`, value 977, expected 977.0
  - row 6853 DULG-A010S02-LS-S3-E12/13: formula `=IF(AND(B6853=B6852,D6854=D6853,A6853=A6852),L6852,0)+N(K6853)`, value 1008, expected 1008.0
  - row 6854 DULG-A010S02-LS-S3-E12/13: formula `=IF(AND(B6854=B6853,D6855=D6854,A6854=A6853),L6853,0)+N(K6854)`, value 1039, expected 1039.0
  - row 6855 DULG-A010S02-LS-S3-E12/13: formula `=IF(AND(B6855=B6854,D6856=D6855,A6855=A6854),L6854,0)+N(K6855)`, value 1061, expected 1061.0
  - row 6856 DULG-A010S02-LS-S3-E12/13: formula `=IF(AND(B6856=B6855,D6857=D6856,A6856=A6855),L6855,0)+N(K6856)`, value 1081, expected 1081.0
  - row 6857 DULG-A010S02-LS-S3-E12/13: formula `=IF(AND(B6857=B6856,D6858=D6857,A6857=A6856),L6856,0)+N(K6857)`, value 1111, expected 1111.0
  - row 6858 DULG-A010S02-LS-S3-E12/13: formula `=IF(AND(B6858=B6857,D6859=D6858,A6858=A6857),L6857,0)+N(K6858)`, value 1138, expected 1138.0
  - row 6859 DULG-A010S02-LS-S3-E12/13: formula `=IF(AND(B6859=B6858,D6860=D6859,A6859=A6858),L6858,0)+N(K6859)`, value 1168, expected 1168.0
  - row 6860 DULG-A010S02-LS-S3-E12/13: formula `=IF(AND(B6860=B6859,D6861=D6860,A6860=A6859),L6859,0)+N(K6860)`, value 1199, expected 1199.0
  - row 6861 DULG-A010S02-LS-S3-E12/13: formula `=IF(AND(B6861=B6860,D6862=D6861,A6861=A6860),L6860,0)+N(K6861)`, value 1212, expected 1212.0
  - row 6862 DULG-A010S02-LS-S3-E12/13: formula `=IF(AND(B6862=B6861,D6863=D6862,A6862=A6861),L6861,0)+N(K6862)`, value 1242, expected 1242.0
