# Oracle run 2026-09-12T15:42:31.907Z (hostile)

Host MacBook-Air.local, Node v24.18.0. Cell = reader verdict against the fixture's expected dump (PASS = every cell equal after policies; DIFF n% = share of matching cells; OPENED = no ground truth; ERROR = reader threw).

| fixture                   | sheetjs  | office-kit | office-kit-stream | openpyxl | calamine | validator              |
| ------------------------- | -------- | ---------- | ----------------- | -------- | -------- | ---------------------- |
| hostile-biff8-xls-renamed | ACCEPTED | REJECTED   | REJECTED          | REJECTED | REJECTED | FAIL (1 schema errors) |
| hostile-xlsb-renamed      | ACCEPTED | REJECTED   | REJECTED          | REJECTED | REJECTED | FAIL (1 schema errors) |
