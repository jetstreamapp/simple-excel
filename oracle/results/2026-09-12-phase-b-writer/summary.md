# Oracle run 2026-09-12T17:27:25.181Z (phase-b-writer)

Host MacBook-Air.local, Node v24.18.0. Cell = reader verdict against the fixture's expected dump (PASS = every cell equal after policies; DIFF n% = share of matching cells; OPENED = no ground truth; ERROR = reader threw).

| fixture | validator | sheetjs | openpyxl | calamine | libreoffice | excel |
|---|---|---|---|---|---|---|
| golden-canonical-simple-excel | PASS | DIFF 93% | DIFF 94% | DIFF 95% | PASS | PASS |
| golden-canonical-simple-excel-inline | PASS | DIFF 93% | DIFF 94% | DIFF 95% | PASS | PASS |
| golden-canonical-simple-excel-zip64 | PASS | ERROR | DIFF 94% | DIFF 95% | PASS | PASS |
