# Oracle run 2026-09-12T04:07:38.145Z (goldens)

Host MacBook-Air.local, Node v24.18.0. Cell = reader verdict against the fixture's expected dump (PASS = every cell equal after policies; DIFF n% = share of matching cells; OPENED = no ground truth; ERROR = reader threw).

| fixture                               | excel    |
| ------------------------------------- | -------- |
| golden-canonical-sheetjs              | PASS     |
| golden-canonical-exceljs              | PASS     |
| golden-canonical-write-excel-file     | PASS     |
| golden-canonical-openpyxl             | REPAIRED |
| golden-canonical-xlsxwriter           | PASS     |
| golden-canonical-libreoffice-resave   | PASS     |
| golden-canonical-libreoffice-from-csv | REPAIRED |
| golden-canonical-office-kit           | PASS     |
