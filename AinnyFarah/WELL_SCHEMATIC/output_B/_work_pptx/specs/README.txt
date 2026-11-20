Visual transcription files (one per well). Read from the upscaled crops in media_png/ (x3 LANCZOS of the part of each
picture shown on the slide); OCR (tesseract) used only as a cross-check. Format: one record per line, fields split by ' | '.

WELL  | key | slides | media | original label
SRC   | slide | media | method                    (source picture(s) for location strings)
DATUM | depth reference of the equipment columns as printed
INFO  | wellInfo key | value as printed            (completionDate, rig, tubing, xmas, packerFluid, completionFluid, rte, maxDeviation, completionType, totalDepth)
CASING| name as printed | size_in | weight | grade | shoe_m | datum [| top_m | top source text]   (top only for a documented TOL)
TUBING| name as printed | size_in | weight | grade | string(Short/Long/-/inner) | bottom_m (blank = derive from EOT/end item)
ROW   | side | name | top_m | bottom_m | ID | OD | status | note | part
        side: S short string, L long string, D shared/dual (one item), X single-string well, Lw lower completion (no string)
        part: picture area where the row is read (p1 = upper part, p2 = lower part)
        a value prefixed '?' = low-legibility read (flagged), '' = not printed
ZONE  | zone | top_m | bottom_m | datum | original text | part | kind (perf / interval)
PBTD  | value_m | datum | original text
TD    | value_m | datum | original text
SHOE  | name | depth_m | datum | tvd_m | original text
INT   | text | depth_m | datum | slide obj id        (intervention note: plug / insert valve / storm choke, from slide text boxes)
NOTE  | where | text                               (unmapped content: status boxes, HUD, history, pressures)
ISSUE | kind | where | text                        (conflict / missing / unclear found during the visual review)
CEMENT| name | top_m | bottom_m | datum | original text      (plug-back cement / TOC)
FLOAT | name | md_m | datum | original text              (float equipment, e.g. float shoe -> floatEquipment)
CASINGX| name | od_in | id_in | weight | grade | shoe_m | datum [| top_m | top source]   (tubular table with explicit OD / ID columns: no user-defined size rule)
TUBINGX| name | od_in | id_in | weight | grade | string | bottom_m datum
