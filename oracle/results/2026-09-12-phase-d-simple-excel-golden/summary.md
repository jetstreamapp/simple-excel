# Oracle run 2026-09-12T18:23:58.527Z (phase-d-simple-excel-golden)

Host MacBook-Air.local, Node v24.18.0. Cell = reader verdict against the fixture's expected dump (PASS = every cell equal after policies; DIFF n% = share of matching cells; OPENED = no ground truth; ERROR = reader threw).

| fixture                                      | simple-excel |
| -------------------------------------------- | ------------ |
| golden-canonical-sheetjs                     | DIFF 92%     |
| golden-canonical-exceljs                     | DIFF 80%     |
| golden-canonical-write-excel-file            | DIFF 77%     |
| golden-canonical-openpyxl                    | DIFF 89%     |
| golden-canonical-xlsxwriter                  | DIFF 95%     |
| golden-canonical-libreoffice-resave          | DIFF 77%     |
| golden-canonical-libreoffice-from-csv        | DIFF 76%     |
| golden-canonical-office-kit                  | DIFF 85%     |
| golden-canonical-excel-365-from-csv          | DIFF 70%     |
| golden-canonical-excel-365-resave            | DIFF 80%     |
| golden-canonical-excel-365-strict            | DIFF 80%     |
| golden-canonical-excel-365-resave-office-kit | DIFF 86%     |
| golden-canonical-gsheets-resave              | DIFF 79%     |
| golden-canonical-numbers                     | DIFF 76%     |
| golden-sfdc-report-details                   | OPENED       |
| golden-sfdc-report-formatted                 | OPENED       |
| golden-canonical-gsheets-from-csv            | DIFF 80%     |
| golden-canonical-excel-365-resave-sheetjs    | DIFF 92%     |
| golden-canonical-excel-365-1904              | DIFF 80%     |
| golden-canonical-simple-excel                | DIFF 99%     |
| golden-canonical-simple-excel-inline         | DIFF 99%     |
| golden-canonical-simple-excel-zip64          | DIFF 99%     |
