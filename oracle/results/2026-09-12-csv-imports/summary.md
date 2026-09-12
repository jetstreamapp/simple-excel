# Oracle run 2026-09-12T15:38:42.898Z (csv-imports)

Host MacBook-Air.local, Node v24.18.0. Cell = reader verdict against the fixture's expected dump (PASS = every cell equal after policies; DIFF n% = share of matching cells; OPENED = no ground truth; ERROR = reader threw).

| fixture                               | sheetjs  | office-kit | calamine |
| ------------------------------------- | -------- | ---------- | -------- |
| golden-canonical-libreoffice-from-csv | DIFF 75% | DIFF 75%   | DIFF 76% |
| golden-canonical-excel-365-from-csv   | DIFF 63% | DIFF 64%   | DIFF 67% |
| golden-canonical-gsheets-from-csv     | DIFF 75% | DIFF 76%   | DIFF 81% |
