# Oracle run 2026-09-27T18:25:19.685Z (audit-fixes-writer)

Host MacBook-Air.local, Node v24.21.0. Cell = reader verdict against the fixture's expected dump (PASS = every cell equal after policies; DIFF n% = share of matching cells; OPENED = no ground truth; ERROR = reader threw).

| fixture                              | validator | libreoffice | openpyxl | calamine | sheetjs  | simple-excel |
| ------------------------------------ | --------- | ----------- | -------- | -------- | -------- | ------------ |
| golden-canonical-simple-excel        | PASS      | PASS        | DIFF 94% | DIFF 95% | DIFF 93% | DIFF 99%     |
| golden-canonical-simple-excel-inline | PASS      | PASS        | DIFF 94% | DIFF 95% | DIFF 93% | DIFF 99%     |
| golden-canonical-simple-excel-zip64  | PASS      | PASS        | DIFF 94% | DIFF 95% | ERROR    | DIFF 99%     |
