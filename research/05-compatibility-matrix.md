# 05 - Compatibility matrix

Which readers open which writers' files, and how faithfully. Rendered from the newest
`oracle/results/<run>/results.json` by `oracle/render-matrix.mjs`; the oracle procedure and
PASS definition are in `README.md`.

_Accurate as of 2026-09-12. Regenerate with `research/regenerate.sh`._

## Readers

| Reader              | What it represents                                                                                                  |
| ------------------- | ------------------------------------------------------------------------------------------------------------------- |
| `simple-excel`      | This package, through its public API only (`dates:'utc'`, `errors:'object'`, `rows({ blankRows: true })`)           |
| `sheetjs`           | SheetJS CE 0.20.3 with Jetstream's exact read options (`cellDates:true, cellText:false`) - what the app sees today  |
| `office-kit`        | `@office-kit/xlsx` document API (`loadWorkbook`), UTC-field Dates resolved through the stylesheet                   |
| `office-kit-stream` | `@office-kit/xlsx` streaming API (`loadWorkbookStream` + `iterRows`), same date resolution                          |
| `openpyxl`          | Python openpyxl 3.1 (`rich_text=True`)                                                                              |
| `calamine`          | Rust calamine via python-calamine (`skip_empty_area=False`)                                                         |
| `validator`         | Open XML SDK schema validation (`@xarsh/ooxml-validator`, Microsoft365 profile)                                     |
| `libreoffice`       | LibreOffice 26.2 headless re-save succeeds                                                                          |
| `excel`             | Microsoft Excel 16.112 (read-only license): opens without a repair prompt or recovery log, sentinel cells read back |

Percentages compare the reader's typed dump with the canonical ground truth (`fixtures/canonical/canonical.json`),
so a low number can be the _writer's_ fault (the file never contained the value) or the _reader's_; the two
largest mismatch categories point at which (see 04 for the mapping). Goldens tagged `policy:truncate-32767`
are compared with Jetstream's truncation applied.

<!-- generated:start (node oracle/render-matrix.mjs) -->

### Run `csv-imports` (2026-09-12-csv-imports)

2026-09-12T15:38:42.898Z, MacBook-Air.local, Node v24.18.0.

| fixture (generator)                                                     | sheetjs                                             | office-kit                                          | calamine                                    |
| ----------------------------------------------------------------------- | --------------------------------------------------- | --------------------------------------------------- | ------------------------------------------- |
| `golden-canonical-libreoffice-from-csv`<br>LibreOffice 26.2.5.2         | 75% (text-as-number 42, date-as-text 30)            | 75% (text-as-number 42, date-as-text 30)            | 76% (text-as-number 42, date-as-text 30)    |
| `golden-canonical-excel-365-from-csv`<br>Microsoft Excel 16.112.4 macOS | 63% (string-mismatch 58, number-as-text 36)         | 64% (string-mismatch 54, number-as-text 36)         | 67% (string-mismatch 63, number-as-text 36) |
| `golden-canonical-gsheets-from-csv`<br>Google Sheets                    | 75% (text-as-number 38, time-only-has-date-part 24) | 76% (text-as-number 38, time-only-has-date-part 24) | 81% (text-as-number 38, number-as-text 16)  |

### Run `edge` (2026-09-12-edge)

2026-09-12T04:13:13.844Z, MacBook-Air.local, Node v24.18.0.

| fixture (generator)                               | sheetjs | office-kit | office-kit-stream       | openpyxl             | calamine                      |
| ------------------------------------------------- | ------- | ---------- | ----------------------- | -------------------- | ----------------------------- |
| `edge-empty-v-element`<br>hand-built              | PASS    | PASS       | PASS                    | 86% (cell-missing 1) | 78% (blank-vs-empty-string 2) |
| `edge-inline-strings-cdata`<br>hand-built         | PASS    | PASS       | 83% (string-mismatch 1) | PASS                 | PASS                          |
| `edge-no-shared-strings-part`<br>hand-built       | PASS    | PASS       | PASS                    | PASS                 | PASS                          |
| `edge-hidden-and-veryhidden-sheets`<br>hand-built | PASS    | PASS       | PASS                    | 91% (cell-missing 1) | PASS                          |

### Run `excel-365` (2026-09-12-excel-365)

2026-09-12T15:41:27.883Z, MacBook-Air.local, Node v24.18.0.

| fixture (generator)                                                                | sheetjs                                                                                              | office-kit                                                                                           | openpyxl                                     | calamine                                                     |
| ---------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------- | -------------------------------------------- | ------------------------------------------------------------ |
| `hostile-biff8-xls-renamed`<br>Microsoft Excel 16.112.4 macOS<br>expect `NOT_XLSX` | ACCEPTED (Data: 31 rows, Features: 12 rows, It's a very long sheet name 001: 2 rows, Hidden: 2 rows) | REJECTED: OpenXmlNotImplementedError: Encrypted xlsx is not supported. Decrypt with msoffcrypto-tool | REJECTED: BadZipFile: File is not a zip file | REJECTED: ZipError: invalid Zip archive: Could not find EOCD |

### Run `goldens` (2026-09-12-goldens)

2026-09-12T04:07:38.145Z, MacBook-Air.local, Node v24.18.0.

| fixture (generator)                                             | excel              |
| --------------------------------------------------------------- | ------------------ |
| `golden-canonical-sheetjs`<br>SheetJS 0.20.3                    | PASS               |
| `golden-canonical-exceljs`<br>ExcelJS 4.4.0                     | PASS               |
| `golden-canonical-write-excel-file`<br>write-excel-file 2.3.10  | PASS               |
| `golden-canonical-openpyxl`<br>openpyxl 3.1.5                   | REPAIRED (see log) |
| `golden-canonical-xlsxwriter`<br>XlsxWriter 3.2.9               | PASS               |
| `golden-canonical-libreoffice-resave`<br>LibreOffice 26.2.5.2   | PASS               |
| `golden-canonical-libreoffice-from-csv`<br>LibreOffice 26.2.5.2 | REPAIRED (see log) |
| `golden-canonical-office-kit`<br>@office-kit/xlsx 0.11.0        | PASS               |

#### What Excel holds in the sentinel cells (raw `value`; dates shown as Excel reports them)

Column `E` control chars, `F` `_x` escape literals, `G` formula-like text, `I` 17-digit integer, `J11` max double, `M` dates (1899-12-31 / 1900-02-28 / 1900-03-01), `N` time-only 12:34:56.789, `O` datetimes (2024-03-10 02:30 DST gap / 1900-01-01 12:00). Caveat: AppleScript date values pass through macOS local-time normalization, so 02:30 in a DST gap reads as 03:30.

| fixture                | E2         | F2        | F3        | F7        | G2         | I2           | J11           | M2                               | M3                               | M4                               | N2                | O2                             | O6                               |
| ---------------------- | ---------- | --------- | --------- | --------- | ---------- | ------------ | ------------- | -------------------------------- | -------------------------------- | -------------------------------- | ----------------- | ------------------------------ | -------------------------------- |
| `sheetjs`              | `a\u0001b` | `A`       | `_x0041_` | `_X0041_` | `=SUM(A1)` | `9.0072E+15` | `1.798E+308`  | `SunDecember 31, 1899 12:00:00 ` | `ThursMarch 1, 1900 12:00:00 AM` | `ThursMarch 1, 1900 12:00:00 AM` | `-0.475731608796` | `SunMarch 10, 2024 3:30:00 AM` | `MonJanuary 1, 1900 12:00:00 PM` |
| `exceljs`              | `ab`       | `A`       | `_x0041_` | `_X0041_` | `=SUM(A1)` | `9.01E+15`   | `1.8E+308`    | `MonJanuary 1, 1900 7:00:00 AM`  | `ThursMarch 1, 1900 7:00:00 AM`  | `ThursMarch 1, 1900 7:00:00 AM`  | `0.815935057872`  | `SunMarch 10, 2024 9:30:00 AM` | `TuesJanuary 2, 1900 7:00:00 PM` |
| `write-excel-file`     | `ab`       | `A`       | `_x0041_` | `_X0041_` | `=SUM(A1)` | `9.0072E+15` | `1.798E+308`  | `MonJanuary 1, 1900 7:00:00 AM`  | `ThursMarch 1, 1900 7:00:00 AM`  | `ThursMarch 1, 1900 7:00:00 AM`  | `0.815935057872`  | `SunMarch 10, 2024 9:30:00 AM` | `TuesJanuary 2, 1900 7:00:00 PM` |
| `openpyxl`             | ``         | `A`       | `_x0041_` | `_X0041_` | `0`        | `9.01E+15`   | `#NULL!`      | `SunDecember 31, 1899 12:00:00 ` | `WednesFebruary 28, 1900 12:00:` | `ThursMarch 1, 1900 12:00:00 AM` | `0.524268391204`  | `SunMarch 10, 2024 3:30:00 AM` | `MonJanuary 1, 1900 12:00:00 PM` |
| `xlsxwriter`           | `a\u0001b` | `_x0041_` | `_x005FA` | `_X0041_` | `=SUM(A1)` | `9.01E+15`   | `#NULL!`      | `SunDecember 31, 1899 12:00:00 ` | `WednesFebruary 28, 1900 12:00:` | `ThursMarch 1, 1900 12:00:00 AM` | `0.524268391204`  | `SunMarch 10, 2024 3:30:00 AM` | `SunDecember 31, 1899 12:00:00 ` |
| `libreoffice-resave`   | `ab`       | `A`       | `_x0041_` | `_X0041_` | `=SUM(A1)` | `9E+15`      | `2E+308`      | `MonJanuary 1, 1900 7:00:00 AM`  | `ThursMarch 1, 1900 7:00:00 AM`  | `ThursMarch 1, 1900 7:00:00 AM`  | `0.815935057872`  | `SunMarch 10, 2024 9:30:00 AM` | `TuesJanuary 2, 1900 7:00:00 PM` |
| `libreoffice-from-csv` | `a\u0001b` | `_x0041_` | `_x0041_` | `_X0041_` | `0`        | `9.0072E+15` | `1.7977E+308` | `MonJanuary 1, 1900 12:00:00 AM` | `ThursMarch 1, 1900 12:00:00 AM` | `ThursMarch 1, 1900 12:00:00 AM` | `12:34:56.789`    | `SunMarch 10, 2024 3:30:00 AM` | `TuesJanuary 2, 1900 12:00:00 P` |
| `office-kit`           | `a\u0001b` | `_x0041_` | `_x005FA` | `_X0041_` | `=SUM(A1)` | `9.01E+15`   | `1.8E+308`    | `SunDecember 31, 1899 7:00:00 A` | `WednesFebruary 28, 1900 7:00:0` | `ThursMarch 1, 1900 7:00:00 AM`  | `-0.18406494213`  | `SunMarch 10, 2024 9:30:00 AM` | `MonJanuary 1, 1900 7:00:00 PM`  |

### Run `gsheets-on-ours` (2026-09-12-gsheets-on-ours)

2026-09-12T22:07:06.097Z, MacBook-Air.local, Node v24.18.0.

| fixture (generator)                                                    | simple-excel                                         | sheetjs                                                      | validator | calamine                                                   |
| ---------------------------------------------------------------------- | ---------------------------------------------------- | ------------------------------------------------------------ | --------- | ---------------------------------------------------------- |
| `golden-canonical-gsheets-resave-simple-excel`<br>Google Sheets        | 96% (escape-sequence-mangled 14, dst-gap-shift-1h 5) | 91% (time-only-has-date-part 24, escape-sequence-mangled 14) | PASS      | 93% (blank-vs-empty-string 19, escape-sequence-mangled 14) |
| `golden-canonical-gsheets-resave-simple-excel-inline`<br>Google Sheets | 96% (escape-sequence-mangled 14, dst-gap-shift-1h 5) | 91% (time-only-has-date-part 24, escape-sequence-mangled 14) | PASS      | 93% (blank-vs-empty-string 19, escape-sequence-mangled 14) |

### Run `hostile` (2026-09-12-hostile)

2026-09-12T15:42:31.907Z, MacBook-Air.local, Node v24.18.0.

| fixture (generator)                                                                | sheetjs                                                                                              | office-kit                                                                                                                                                                 | office-kit-stream                                                                                                                                                          | openpyxl                                                | calamine                                                                      | validator            |
| ---------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------- | ----------------------------------------------------------------------------- | -------------------- |
| `hostile-biff8-xls-renamed`<br>Microsoft Excel 16.112.4 macOS<br>expect `NOT_XLSX` | ACCEPTED (Data: 31 rows, Features: 12 rows, It's a very long sheet name 001: 2 rows, Hidden: 2 rows) | REJECTED: OpenXmlNotImplementedError: Encrypted xlsx is not supported. Decrypt with msoffcrypto-tool                                                                       | REJECTED: OpenXmlNotImplementedError: Encrypted xlsx is not supported. Decrypt with msoffcrypto-tool                                                                       | REJECTED: OSError: File contains no valid workbook part | REJECTED: CalamineError: File not found 'theme/theme/_rels/workbook.xml.rels' | FAIL: 1 schema error |
| `hostile-xlsb-renamed`<br>Microsoft Excel 16.112.4 macOS<br>expect `NOT_XLSX`      | ACCEPTED (Data: 31 rows, Features: 12 rows, It's a very long sheet name 001: 2 rows, Hidden: 2 rows) | REJECTED: Error: readTagExp returned undefined at position 857. Context: "\u0000\u0000m\u0000i\u0000c\u0000r\u0000o\u0000s\u0000o\u0000f\u0000t\u0000.\u0000c\u0000o\u0000 | REJECTED: Error: readTagExp returned undefined at position 857. Context: "\u0000\u0000m\u0000i\u0000c\u0000r\u0000o\u0000s\u0000o\u0000f\u0000t\u0000.\u0000c\u0000o\u0000 | REJECTED: OSError: File contains no valid workbook part | REJECTED: CalamineError: File not found 'xl/_rels/workbook.xml.rels'          | FAIL: 1 schema error |

### Run `phase-b-writer` (2026-09-12-phase-b-writer)

2026-09-12T17:27:25.181Z, MacBook-Air.local, Node v24.18.0.

| fixture (generator)                                          | validator | sheetjs                                               | openpyxl                                             | calamine                                           | libreoffice | excel |
| ------------------------------------------------------------ | --------- | ----------------------------------------------------- | ---------------------------------------------------- | -------------------------------------------------- | ----------- | ----- |
| `golden-canonical-simple-excel`<br>simple-excel 0.0.0        | PASS      | 93% (time-only-has-date-part 24, temporal-mismatch 6) | 94% (escape-sequence-mangled 30, dst-gap-shift-1h 5) | 95% (blank-vs-empty-string 19, dst-gap-shift-1h 5) | PASS        | PASS  |
| `golden-canonical-simple-excel-inline`<br>simple-excel 0.0.0 | PASS      | 93% (time-only-has-date-part 24, temporal-mismatch 6) | 94% (escape-sequence-mangled 30, dst-gap-shift-1h 5) | 95% (blank-vs-empty-string 19, dst-gap-shift-1h 5) | PASS        | PASS  |
| `golden-canonical-simple-excel-zip64`<br>simple-excel 0.0.0  | PASS      | ERROR                                                 | 94% (escape-sequence-mangled 30, dst-gap-shift-1h 5) | 95% (blank-vs-empty-string 19, dst-gap-shift-1h 5) | PASS        | PASS  |

### Run `phase-d-office-kit-on-ours` (2026-09-12-phase-d-office-kit-on-ours)

2026-09-12T18:26:43.444Z, MacBook-Air.local, Node v24.18.0.

| fixture (generator)                                          | office-kit                                            | office-kit-stream                                            |
| ------------------------------------------------------------ | ----------------------------------------------------- | ------------------------------------------------------------ |
| `golden-canonical-simple-excel`<br>simple-excel 0.0.0        | 94% (time-only-has-date-part 24, temporal-mismatch 6) | 94% (time-only-has-date-part 24, temporal-mismatch 6)        |
| `golden-canonical-simple-excel-inline`<br>simple-excel 0.0.0 | 94% (time-only-has-date-part 24, temporal-mismatch 6) | 88% (escape-sequence-mangled 39, time-only-has-date-part 24) |
| `golden-canonical-simple-excel-zip64`<br>simple-excel 0.0.0  | 94% (time-only-has-date-part 24, temporal-mismatch 6) | 94% (time-only-has-date-part 24, temporal-mismatch 6)        |

### Run `phase-d-simple-excel-edge` (2026-09-12-phase-d-simple-excel-edge)

2026-09-12T18:24:20.149Z, MacBook-Air.local, Node v24.18.0.

| fixture (generator)                                   | simple-excel              |
| ----------------------------------------------------- | ------------------------- |
| `edge-baseline-minimal`<br>hand-built                 | PASS                      |
| `edge-no-dimension`<br>hand-built                     | PASS                      |
| `edge-missing-r-attributes`<br>hand-built             | PASS                      |
| `edge-empty-v-element`<br>hand-built                  | PASS                      |
| `edge-inline-strings-cdata`<br>hand-built             | PASS                      |
| `edge-prefixed-elements`<br>hand-built                | PASS                      |
| `edge-strict-namespaces`<br>hand-built                | PASS                      |
| `edge-absolute-rel-targets`<br>hand-built             | PASS                      |
| `edge-backslash-rel-targets`<br>hand-built            | PASS                      |
| `edge-sst-after-sheet-data-descriptors`<br>hand-built | PASS                      |
| `edge-stored-entries`<br>hand-built                   | PASS                      |
| `edge-no-shared-strings-part`<br>hand-built           | PASS                      |
| `edge-sst-index-out-of-range`<br>hand-built           | 67% (string-mismatch 3)   |
| `edge-date1904`<br>hand-built                         | 89% (temporal-mismatch 1) |
| `edge-hidden-and-veryhidden-sheets`<br>hand-built     | PASS                      |

### Run `phase-d-simple-excel-golden` (2026-09-12-phase-d-simple-excel-golden)

2026-09-12T18:23:58.527Z, MacBook-Air.local, Node v24.18.0.

| fixture (generator)                                                              | simple-excel                                          |
| -------------------------------------------------------------------------------- | ----------------------------------------------------- |
| `golden-canonical-sheetjs`<br>SheetJS 0.20.3                                     | 92% (date-as-serial 30, escape-sequence-mangled 14)   |
| `golden-canonical-exceljs`<br>ExcelJS 4.4.0                                      | 80% (temporal-mismatch 68, control-chars-stripped 22) |
| `golden-canonical-write-excel-file`<br>write-excel-file 2.3.10                   | 77% (temporal-mismatch 68, control-chars-stripped 22) |
| `golden-canonical-openpyxl`<br>openpyxl 3.1.5                                    | 89% (cell-missing 32, empty-string-vs-blank 15)       |
| `golden-canonical-xlsxwriter`<br>XlsxWriter 3.2.9                                | 95% (empty-string-vs-blank 15, temporal-mismatch 8)   |
| `golden-canonical-libreoffice-resave`<br>LibreOffice 26.2.5.2                    | 77% (temporal-mismatch 68, control-chars-stripped 22) |
| `golden-canonical-libreoffice-from-csv`<br>LibreOffice 26.2.5.2                  | 76% (text-as-number 42, date-as-text 30)              |
| `golden-canonical-office-kit`<br>@office-kit/xlsx 0.11.0                         | 85% (temporal-mismatch 39, date-tz-offset-shift 27)   |
| `golden-canonical-excel-365-from-csv`<br>Microsoft Excel 16.112.4 macOS          | 70% (string-mismatch 54, text-as-number 43)           |
| `golden-canonical-excel-365-resave`<br>Microsoft Excel 16.112.4 macOS            | 80% (temporal-mismatch 68, control-chars-stripped 22) |
| `golden-canonical-excel-365-strict`<br>Microsoft Excel 365 macOS 16.03           | 80% (temporal-mismatch 60, date-tz-offset-shift 30)   |
| `golden-canonical-excel-365-resave-office-kit`<br>Microsoft Excel 16.112.4 macOS | 86% (temporal-mismatch 39, date-tz-offset-shift 27)   |
| `golden-canonical-gsheets-resave`<br>Google Sheets                               | 79% (temporal-mismatch 68, control-chars-stripped 22) |
| `golden-canonical-numbers`<br>Numbers 14.4                                       | 76% (temporal-mismatch 68, control-chars-stripped 22) |
| `golden-sfdc-report-details`<br>Salesforce report export                         | OPENED                                                |
| `golden-sfdc-report-formatted`<br>Salesforce report export                       | OPENED                                                |
| `golden-canonical-gsheets-from-csv`<br>Google Sheets                             | 80% (text-as-number 38, number-as-text 16)            |
| `golden-canonical-excel-365-resave-sheetjs`<br>Microsoft Excel 16.112.4 macOS    | 92% (date-as-serial 30, escape-sequence-mangled 14)   |
| `golden-canonical-excel-365-1904`<br>Microsoft Excel 16.112.4 macOS              | 80% (temporal-mismatch 90, control-chars-stripped 22) |
| `golden-canonical-simple-excel`<br>simple-excel 0.0.0                            | 99% (dst-gap-shift-1h 5, date-as-text 3)              |
| `golden-canonical-simple-excel-inline`<br>simple-excel 0.0.0                     | 99% (dst-gap-shift-1h 5, date-as-text 3)              |
| `golden-canonical-simple-excel-zip64`<br>simple-excel 0.0.0                      | 99% (dst-gap-shift-1h 5, date-as-text 3)              |

### Run `phase-d-simple-excel-hostile` (2026-09-12-phase-d-simple-excel-hostile)

2026-09-12T18:26:24.594Z, MacBook-Air.local, Node v24.18.0.

| fixture (generator)                                                                | simple-excel              |
| ---------------------------------------------------------------------------------- | ------------------------- |
| `hostile-xxe-doctype-in-sharedstrings`<br>hand-built<br>expect `XML_DOCTYPE`       | PASS: XML_DOCTYPE         |
| `hostile-truncated-central-directory`<br>hand-built<br>expect `TRUNCATED`          | PASS: ZIP_TRUNCATED       |
| `hostile-crc-mismatch`<br>hand-built<br>expect `CRC_MISMATCH`                      | PASS: ZIP_CRC_MISMATCH    |
| `hostile-duplicate-sheet-entries`<br>hand-built<br>expect `DUPLICATE_ENTRY`        | PASS: ZIP_DUPLICATE_ENTRY |
| `hostile-deeply-nested-rich-text`<br>hand-built<br>expect `LIMIT_EXCEEDED`         | PASS: LIMIT_EXCEEDED      |
| `hostile-csv-bytes-renamed`<br>hand-built<br>expect `NOT_XLSX`                     | PASS: NOT_XLSX            |
| `hostile-ods-renamed`<br>hand-built<br>expect `NOT_XLSX`                           | PASS: ODS                 |
| `hostile-biff8-xls-renamed`<br>Microsoft Excel 16.112.4 macOS<br>expect `NOT_XLSX` | PASS: LEGACY_XLS          |
| `hostile-encrypted-password-test`<br>hand-built<br>expect `ENCRYPTED`              | PASS: ENCRYPTED           |
| `hostile-not-a-zip`<br>hand-built<br>expect `NOT_XLSX`                             | PASS: NOT_XLSX            |
| `hostile-zip-bomb-30mb-sheet`<br>hand-built<br>expect `ZIP_BOMB`                   | PASS: ZIP_BOMB            |
| `hostile-xlsb-renamed`<br>Microsoft Excel 16.112.4 macOS<br>expect `NOT_XLSX`      | PASS: XLSB                |

### Run `phase-d-simple-excel-jetstream` (2026-09-12-phase-d-simple-excel-jetstream)

2026-09-12T18:26:30.992Z, MacBook-Air.local, Node v24.18.0.

| fixture (generator)                                                                                 | simple-excel |
| --------------------------------------------------------------------------------------------------- | ------------ |
| `jetstream-multi-object-template-gsheets`<br>Google Sheets xlsx export, workbook created 2021-07-10 | OPENED       |
| `jetstream-records-product2-sheetjs`<br>SheetJS unknown                                             | OPENED       |
| `jetstream-records-product2-csv`<br>hand-authored CSV<br>expect `NOT_XLSX`                          | ERROR        |

### Run `user-goldens` (2026-09-12-user-goldens)

2026-09-12T15:30:31.725Z, MacBook-Air.local, Node v24.18.0.

| fixture (generator)                        | excel              |
| ------------------------------------------ | ------------------ |
| `golden-canonical-numbers`<br>Numbers 14.4 | REPAIRED (see log) |

Legend: **PASS** every compared cell equal after fixture policies; **n%** share of matching cells with the two largest mismatch categories (04 maps categories to catalog entries); **FAIL: n schema errors** Open XML SDK validator; **REPAIRED** Excel wrote a recovery log while opening (silent repair under automation); **OPENED** no ground truth; **REJECTED/ACCEPTED** hostile fixture outcome (a classified rejection is the goal); **ERROR** the reader threw on a fixture it should read.

<!-- generated:end -->

## Reading the current matrix

Run `user-goldens` holds the files only a person could produce (Excel 365 desktop, Google Sheets, Numbers,
Salesforce report exports). Its headline rows: office-kit cannot open either Salesforce export (POI's
`<u val="none"/>`), Excel's own Strict save is unreadable for office-kit, calamine and openpyxl, the Numbers
export carries `<v>inf</v>` (openpyxl crashes, Excel repairs it), and Excel's CSV import turned the UTF-8
sample into mojibake and evaluated `=SUM(A1)`. Run `csv-imports` compares the two CSV-derived goldens with a
Data-only ground truth (`expected/canonical-csv-import.json`) to show each application's typing heuristics.

- **Excel opens every golden**, including the ones the validator rejects; the validator is a strictness gate,
  not a compatibility predictor. The two Excel repairs seen (openpyxl and the LibreOffice CSV import) both come
  from formula-looking text evaluated as formulas.
- **SheetJS reading SheetJS output** is the reference for "what Jetstream sees today"; its residual mismatches
  are all SheetJS conventions (serial 60 returned as a number, `_X0041_` decoded, CRLF collapsed, DST gap).
- **office-kit reading office-kit output** loses only the DST-gap and escape-overlap cells, but reading any
  _other_ generator's output shows its two reader gaps: `_xHHHH_` escapes are not decoded and dates need the
  consumer to resolve number formats.
- **Goldens missing features** (no freeze panes in SheetJS CE, no cached formula results in openpyxl, no
  error cells in XlsxWriter) show up as structural gaps in the per-fixture sidecars
  (`fixtures/golden/<generator>/canonical.features.json`), not in the percentages.

## Our own reader (runs `phase-d-simple-excel-*`)

`simple-excel` reads through the public API and nothing else (`oracle/simple-excel/read-dump.mjs`, the plain-Node
twin of `test/helpers/dump.ts`), so an oracle verdict and a corpus-suite verdict cannot drift apart. The four runs
cover every fixture in the manifest: 22 goldens, 15 edge cases, 12 hostile files, 3 Jetstream assets.

**Hostile: 12 of 12 rejected with the code the manifest asks for.** Every one threw a classified `XlsxError`; none
crashed, hung or was accepted. The renamed `.ods`, `.xls` and `.xlsb` come back as `ODS`, `LEGACY_XLS` and `XLSB`
rather than a flat `NOT_XLSX` (the mapping in `run.mjs` and `test/hostile.test.ts` accepts any of those for
`NOT_XLSX`), which is what lets a host application tell the user what the file actually is - SheetJS parses those
same two files as if nothing were wrong. The zip bomb is read under an 8 MiB per-part cap, the same limits the
hostile suite uses: our default cap is 1 GiB per part, and under it the 30 MB sheet is simply a large file that
streams fine. The guard is the byte cap, not a compression-ratio heuristic - a real Salesforce export routinely
compresses 100:1, so a ratio guard would reject honest files.

**Edge: 13 of 15 PASS**, including the ones that catch other readers - the empty `<v/>` (openpyxl loses the cell),
inline strings with CDATA and rich-text runs (office-kit's stream reader mangles them), a missing `<dimension>`,
missing `r` attributes, prefixed elements, Strict namespaces, absolute and backslash rel targets, stored entries,
data descriptors with the SST after the sheet, and hidden/very-hidden sheets (openpyxl loses a cell). Both DIFFs
are the fixture's own point and are recorded in `test/corpus-policies.json`:

| Fixture                       | DIFF                    | Why                                                                                                                                                                                                                                         |
| ----------------------------- | ----------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `edge-sst-index-out-of-range` | 67% (string-mismatch 3) | three cells point past a two-entry shared-string table; Excel opens the file with them blank and we match it (EC-SST-INDEX-OUT-OF-RANGE), while the shared expected dump keeps the strings. SheetJS and office-kit refuse the file outright |
| `edge-date1904`               | 89% (temporal-mismatch) | the fixture is the baseline workbook with only `workbookPr/@date1904` flipped, so its one serial now means 2028-03-01 while the shared `expected-minimal.json` still says 2024-02-29 (EC-DATE-1904). SheetJS reads it the same way we do    |

**Goldens: no DIFF category that the fixture's policy entry does not already justify.** Every category on all 22
goldens was checked against that fixture's `allowedCategories` in `test/corpus-policies.json`; nothing new
appeared, so every difference is a deviation the generator baked into the file, not a reader defect. Grouped by
cause:

| Cause (whose)                                        | Where                                                                                                 | Categories                                                                                                      |
| ---------------------------------------------------- | ----------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------- |
| Dates written from UTC fields                        | ExcelJS and everything derived from it (LibreOffice re-save, Excel re-save + Strict, Sheets, Numbers) | `temporal-mismatch` 68, `date-tz-offset-shift` 22-30                                                            |
| Control characters dropped at write time             | the same family                                                                                       | `control-chars-stripped` 22, `string-mismatch` 4 (`delx`)                                                      |
| Escape-shaped text written unescaped or escaped once | most generators                                                                                       | `escape-sequence-mangled` 14 (or 5)                                                                             |
| Time-only and pre-1900 serials                       | SheetJS, office-kit, XlsxWriter                                                                       | `date-as-serial` 24-30, `temporal-mismatch`                                                                     |
| No support for the value at all                      | write-excel-file, XlsxWriter, openpyxl, Numbers                                                       | `empty-string-vs-blank` 15, `error-as-text`, `formula-dropped`, `sheet-missing`                                 |
| CSV import typing heuristics                         | the three `source:csv` goldens (ground truth is `expected/canonical-csv-import.json`)                 | `text-as-number`, `date-as-text`, `boolean-as-text-or-number`, `whitespace-trimmed`, mojibake `string-mismatch` |

Where another reader has read the same fixture (run `csv-imports`), the category sets line up: we report the same
deviations minus that reader's own quirks - no `time-only-has-date-part` (we hang a sub-1 serial on 1899-12-30,
which is what the ground truth uses), no `blank-vs-empty-string` (calamine's), no `date-as-serial` on files whose
serials are honest (SheetJS's).

**Our own three goldens read back at 99% (650/658)** with exactly the two categories every other reader reports on
them - `date-as-text` 3 (pre-1900 dates written as ISO text, ADR-003) and `dst-gap-shift-1h` 5. Those are writer
conventions, recorded in run `phase-b-writer`, not read defects.

**Jetstream assets**: both workbooks open and stream (`OPENED`). `jetstream-records-product2-csv` is `ERROR
NOT_XLSX` by design - it is CSV bytes that Jetstream feeds to papaparse, and `openWorkbook` refuses non-xlsx input
(ADR-007). The manifest gives it no `expectedError` because it is not a hostile fixture, so the runner has no way
to score it as expected; `test/corpus-policies.json` skips it for the same reason.

**office-kit on our output** (run `phase-d-office-kit-on-ours`) completes the reader set over our three goldens,
which run `phase-b-writer` had covered with validator, Excel, LibreOffice, SheetJS, openpyxl and calamine. The
document API reads all three at 94%: our two writer conventions plus 30 cells of office-kit's own time handling
(24 `time-only-has-date-part` and 6 `temporal-mismatch`, all the time-only column, which it hangs on 1899-12-31,
midnight included). The streaming API matches it on the shared-string and zip64 variants and drops to 88% on the
inline-string variant, where it does not decode `_xHHHH_` in inline strings (`a_x0001_b`, `_x005F_x0041_`) - the
office-kit reader gap 04 already records, visible here only because that variant puts every string inline.
