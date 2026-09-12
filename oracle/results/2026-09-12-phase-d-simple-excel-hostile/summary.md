# Oracle run 2026-09-12T18:26:24.594Z (phase-d-simple-excel-hostile)

Host MacBook-Air.local, Node v24.18.0. Cell = reader verdict against the fixture's expected dump (PASS = every cell equal after policies; DIFF n% = share of matching cells; OPENED = no ground truth; ERROR = reader threw).

| fixture                              | simple-excel               |
| ------------------------------------ | -------------------------- |
| hostile-xxe-doctype-in-sharedstrings | PASS (XML_DOCTYPE)         |
| hostile-truncated-central-directory  | PASS (ZIP_TRUNCATED)       |
| hostile-crc-mismatch                 | PASS (ZIP_CRC_MISMATCH)    |
| hostile-duplicate-sheet-entries      | PASS (ZIP_DUPLICATE_ENTRY) |
| hostile-deeply-nested-rich-text      | PASS (LIMIT_EXCEEDED)      |
| hostile-csv-bytes-renamed            | PASS (NOT_XLSX)            |
| hostile-ods-renamed                  | PASS (ODS)                 |
| hostile-biff8-xls-renamed            | PASS (LEGACY_XLS)          |
| hostile-encrypted-password-test      | PASS (ENCRYPTED)           |
| hostile-not-a-zip                    | PASS (NOT_XLSX)            |
| hostile-zip-bomb-30mb-sheet          | PASS (ZIP_BOMB)            |
| hostile-xlsb-renamed                 | PASS (XLSB)                |
