# Oracle run 2026-09-12T22:34:47.823Z (numbers-on-ours)

Host MacBook-Air.local, Node v24.18.0. Cell = reader verdict against the fixture's expected dump (PASS = every cell equal after policies; DIFF n% = share of matching cells; OPENED = no ground truth; ERROR = reader threw).

| fixture                                             | simple-excel | sheetjs  | validator | calamine | excel    |
| --------------------------------------------------- | ------------ | -------- | --------- | -------- | -------- |
| golden-canonical-numbers-resave-simple-excel        | DIFF 83%     | DIFF 82% | PASS      | DIFF 83% | REPAIRED |
| golden-canonical-numbers-resave-simple-excel-inline | DIFF 88%     | DIFF 87% | PASS      | DIFF 88% | REPAIRED |
