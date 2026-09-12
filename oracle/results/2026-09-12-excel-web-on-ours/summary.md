# Oracle run 2026-09-12T23:13:59.157Z (excel-web-on-ours)

Host MacBook-Air.local, Node v24.18.0. Cell = reader verdict against the fixture's expected dump (PASS = every cell equal after policies; DIFF n% = share of matching cells; OPENED = no ground truth; ERROR = reader threw).

| fixture                                               | simple-excel | sheetjs  | validator | calamine | excel |
| ----------------------------------------------------- | ------------ | -------- | --------- | -------- | ----- |
| golden-canonical-excel-web-resave-simple-excel        | DIFF 99%     | DIFF 94% | PASS      | DIFF 95% | PASS  |
| golden-canonical-excel-web-resave-simple-excel-inline | DIFF 99%     | DIFF 94% | PASS      | DIFF 95% | PASS  |
