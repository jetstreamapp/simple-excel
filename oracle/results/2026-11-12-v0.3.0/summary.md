# Oracle run 2026-11-12T15:01:58.555Z (v0.3.0)

Host darwin-arm64, Node v24.22.1. Cell = reader verdict against the fixture's expected dump (PASS = every cell equal after policies; DIFF n% = share of matching cells; OPENED = no ground truth; ERROR = reader threw).

| fixture                                               | simple-excel | sheetjs  | office-kit | office-kit-stream | openpyxl  | calamine | validator                | libreoffice | excel    |
| ----------------------------------------------------- | ------------ | -------- | ---------- | ----------------- | --------- | -------- | ------------------------ | ----------- | -------- |
| golden-canonical-sheetjs                              | DIFF 92%     | DIFF 96% | DIFF 91%   | DIFF 91%          | DIFF 88%  | DIFF 92% | PASS                     | PASS        | PASS     |
| golden-canonical-exceljs                              | DIFF 80%     | DIFF 79% | DIFF 79%   | DIFF 80%          | DIFF 81%  | DIFF 77% | FAIL (1 schema errors)   | PASS        | PASS     |
| golden-canonical-write-excel-file                     | DIFF 77%     | DIFF 76% | DIFF 77%   | DIFF 77%          | DIFF 78%  | DIFF 77% | FAIL (1 schema errors)   | PASS        | PASS     |
| golden-canonical-openpyxl                             | DIFF 89%     | DIFF 85% | ERROR      | ERROR             | DIFF 90%  | DIFF 88% | FAIL (1 schema errors)   | PASS        | REPAIRED |
| golden-canonical-xlsxwriter                           | DIFF 95%     | DIFF 88% | ERROR      | ERROR             | DIFF 91%  | DIFF 94% | PASS                     | PASS        | PASS     |
| golden-canonical-libreoffice-resave                   | DIFF 77%     | DIFF 77% | DIFF 76%   | DIFF 77%          | DIFF 72%  | DIFF 76% | PASS                     | PASS        | PASS     |
| golden-canonical-libreoffice-from-csv                 | DIFF 76%     | DIFF 75% | DIFF 76%   | DIFF 76%          | DIFF 73%  | DIFF 76% | FAIL (1 schema errors)   | PASS        | REPAIRED |
| golden-canonical-office-kit                           | DIFF 85%     | DIFF 84% | DIFF 85%   | DIFF 85%          | DIFF 80%  | DIFF 82% | FAIL (2 schema errors)   | PASS        | PASS     |
| golden-canonical-excel-365-from-csv                   | DIFF 70%     | DIFF 65% | DIFF 65%   | DIFF 65%          | DIFF 66%  | DIFF 68% | PASS                     | PASS        | PASS     |
| golden-canonical-excel-365-resave                     | DIFF 80%     | DIFF 79% | DIFF 79%   | DIFF 80%          | DIFF 79%  | DIFF 77% | PASS                     | PASS        | PASS     |
| golden-canonical-excel-365-strict                     | DIFF 80%     | DIFF 79% | ERROR      | ERROR             | DIFF 100% | ERROR    | FAIL (133 schema errors) | PASS        | PASS     |
| golden-canonical-excel-365-resave-office-kit          | DIFF 86%     | DIFF 85% | DIFF 85%   | DIFF 86%          | DIFF 81%  | DIFF 82% | PASS                     | PASS        | PASS     |
| golden-canonical-gsheets-resave                       | DIFF 79%     | DIFF 79% | DIFF 79%   | DIFF 79%          | DIFF 79%  | DIFF 77% | PASS                     | PASS        | PASS     |
| golden-canonical-numbers                              | DIFF 76%     | DIFF 75% | ERROR      | ERROR             | ERROR     | DIFF 76% | PASS                     | PASS        | REPAIRED |
| golden-sfdc-report-details                            | OPENED       | OPENED   | OPENED     | OPENED            | OPENED    | OPENED   | FAIL (20 schema errors)  | PASS        | PASS     |
| golden-sfdc-report-formatted                          | OPENED       | OPENED   | OPENED     | OPENED            | OPENED    | OPENED   | FAIL (39 schema errors)  | PASS        | PASS     |
| golden-canonical-gsheets-from-csv                     | DIFF 80%     | DIFF 75% | DIFF 76%   | DIFF 76%          | DIFF 78%  | DIFF 81% | PASS                     | PASS        | REPAIRED |
| golden-canonical-excel-365-resave-sheetjs             | DIFF 92%     | DIFF 96% | DIFF 92%   | DIFF 93%          | DIFF 88%  | DIFF 93% | PASS                     | PASS        | PASS     |
| golden-canonical-excel-365-1904                       | DIFF 80%     | DIFF 79% | DIFF 79%   | DIFF 80%          | DIFF 79%  | DIFF 77% | PASS                     | PASS        | PASS     |
| golden-canonical-simple-excel                         | DIFF 99%     | DIFF 93% | DIFF 94%   | DIFF 94%          | DIFF 94%  | DIFF 95% | PASS                     | PASS        | PASS     |
| golden-canonical-simple-excel-inline                  | DIFF 99%     | DIFF 93% | DIFF 94%   | DIFF 94%          | DIFF 94%  | DIFF 95% | PASS                     | PASS        | PASS     |
| golden-canonical-simple-excel-zip64                   | DIFF 99%     | ERROR    | DIFF 94%   | DIFF 94%          | DIFF 94%  | DIFF 95% | PASS                     | PASS        | PASS     |
| golden-canonical-gsheets-resave-simple-excel          | DIFF 96%     | DIFF 91% | DIFF 92%   | DIFF 92%          | DIFF 94%  | DIFF 93% | PASS                     | PASS        | PASS     |
| golden-canonical-gsheets-resave-simple-excel-inline   | DIFF 96%     | DIFF 91% | DIFF 92%   | DIFF 92%          | DIFF 94%  | DIFF 93% | PASS                     | PASS        | PASS     |
| golden-canonical-numbers-resave-simple-excel          | DIFF 83%     | DIFF 82% | ERROR      | ERROR             | ERROR     | DIFF 83% | PASS                     | PASS        | REPAIRED |
| golden-canonical-numbers-resave-simple-excel-inline   | DIFF 88%     | DIFF 87% | ERROR      | ERROR             | ERROR     | DIFF 88% | PASS                     | PASS        | REPAIRED |
| golden-canonical-excel-web-resave-simple-excel        | DIFF 99%     | DIFF 94% | DIFF 94%   | DIFF 94%          | DIFF 94%  | DIFF 95% | PASS                     | PASS        | PASS     |
| golden-canonical-excel-web-resave-simple-excel-inline | DIFF 99%     | DIFF 94% | DIFF 94%   | DIFF 94%          | DIFF 94%  | DIFF 95% | PASS                     | PASS        | PASS     |
| golden-header-offset-sheetjs                          | OPENED       | OPENED   | OPENED     | OPENED            | OPENED    | OPENED   | PASS                     | PASS        | PASS     |
| golden-header-offset-openpyxl                         | OPENED       | OPENED   | OPENED     | OPENED            | OPENED    | OPENED   | FAIL (1 schema errors)   | PASS        | PASS     |
| golden-header-offset-libreoffice                      | OPENED       | OPENED   | OPENED     | OPENED            | OPENED    | OPENED   | PASS                     | PASS        | PASS     |
