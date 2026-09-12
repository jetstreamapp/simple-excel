# Oracle run 2026-09-12T18:24:20.149Z (phase-d-simple-excel-edge)

Host MacBook-Air.local, Node v24.18.0. Cell = reader verdict against the fixture's expected dump (PASS = every cell equal after policies; DIFF n% = share of matching cells; OPENED = no ground truth; ERROR = reader threw).

| fixture                               | simple-excel |
| ------------------------------------- | ------------ |
| edge-baseline-minimal                 | PASS         |
| edge-no-dimension                     | PASS         |
| edge-missing-r-attributes             | PASS         |
| edge-empty-v-element                  | PASS         |
| edge-inline-strings-cdata             | PASS         |
| edge-prefixed-elements                | PASS         |
| edge-strict-namespaces                | PASS         |
| edge-absolute-rel-targets             | PASS         |
| edge-backslash-rel-targets            | PASS         |
| edge-sst-after-sheet-data-descriptors | PASS         |
| edge-stored-entries                   | PASS         |
| edge-no-shared-strings-part           | PASS         |
| edge-sst-index-out-of-range           | DIFF 67%     |
| edge-date1904                         | DIFF 89%     |
| edge-hidden-and-veryhidden-sheets     | PASS         |
