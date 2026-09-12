# Oracle run 2026-09-12T15:41:27.883Z (excel-365)

Host MacBook-Air.local, Node v24.18.0. Cell = reader verdict against the fixture's expected dump (PASS = every cell equal after policies; DIFF n% = share of matching cells; OPENED = no ground truth; ERROR = reader threw).

| fixture                   | sheetjs  | office-kit | openpyxl | calamine |
| ------------------------- | -------- | ---------- | -------- | -------- |
| hostile-biff8-xls-renamed | ACCEPTED | REJECTED   | REJECTED | REJECTED |
