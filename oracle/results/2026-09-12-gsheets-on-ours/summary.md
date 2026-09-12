# Oracle run 2026-09-12T22:07:06.097Z (gsheets-on-ours)

Host MacBook-Air.local, Node v24.18.0. Cell = reader verdict against the fixture's expected dump (PASS = every cell equal after policies; DIFF n% = share of matching cells; OPENED = no ground truth; ERROR = reader threw).

| fixture | simple-excel | sheetjs | validator | calamine |
|---|---|---|---|---|
| golden-canonical-gsheets-resave-simple-excel | DIFF 96% | DIFF 91% | PASS | DIFF 93% |
| golden-canonical-gsheets-resave-simple-excel-inline | DIFF 96% | DIFF 91% | PASS | DIFF 93% |
