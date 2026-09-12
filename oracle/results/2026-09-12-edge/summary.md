# Oracle run 2026-09-12T04:13:13.844Z (edge)

Host MacBook-Air.local, Node v24.18.0. Cell = reader verdict against the fixture's expected dump (PASS = every cell equal after policies; DIFF n% = share of matching cells; OPENED = no ground truth; ERROR = reader threw).

| fixture                           | sheetjs | office-kit | office-kit-stream | openpyxl | calamine |
| --------------------------------- | ------- | ---------- | ----------------- | -------- | -------- |
| edge-empty-v-element              | PASS    | PASS       | PASS              | DIFF 86% | DIFF 78% |
| edge-inline-strings-cdata         | PASS    | PASS       | DIFF 83%          | PASS     | PASS     |
| edge-no-shared-strings-part       | PASS    | PASS       | PASS              | PASS     | PASS     |
| edge-hidden-and-veryhidden-sheets | PASS    | PASS       | PASS              | DIFF 91% | PASS     |
