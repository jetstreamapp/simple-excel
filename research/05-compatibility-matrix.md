# 05 - Compatibility matrix

Which readers open which writers' files, and how faithfully. Rendered from the newest
`oracle/results/<run>/results.json` by `oracle/render-matrix.mjs`; the oracle procedure and
PASS definition are in `README.md`.

_Accurate as of 2026-09-12. Regenerate with `research/regenerate.sh`._

## Readers

| Reader              | What it represents                                                                                                  |
| ------------------- | ------------------------------------------------------------------------------------------------------------------- |
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

| fixture (generator) | sheetjs | office-kit | calamine |
|---|---|---|---|
| `golden-canonical-libreoffice-from-csv`<br>LibreOffice 26.2.5.2 | 75% (text-as-number 42, date-as-text 30) | 75% (text-as-number 42, date-as-text 30) | 76% (text-as-number 42, date-as-text 30) |
| `golden-canonical-excel-365-from-csv`<br>Microsoft Excel 16.112.4 macOS | 63% (string-mismatch 58, number-as-text 36) | 64% (string-mismatch 54, number-as-text 36) | 67% (string-mismatch 63, number-as-text 36) |
| `golden-canonical-gsheets-from-csv`<br>Google Sheets | 75% (text-as-number 38, time-only-has-date-part 24) | 76% (text-as-number 38, time-only-has-date-part 24) | 81% (text-as-number 38, number-as-text 16) |

### Run `edge` (2026-09-12-edge)

2026-09-12T04:13:13.844Z, MacBook-Air.local, Node v24.18.0.

| fixture (generator) | sheetjs | office-kit | office-kit-stream | openpyxl | calamine |
|---|---|---|---|---|---|
| `edge-empty-v-element`<br>hand-built | PASS | PASS | PASS | 86% (cell-missing 1) | 78% (blank-vs-empty-string 2) |
| `edge-inline-strings-cdata`<br>hand-built | PASS | PASS | 83% (string-mismatch 1) | PASS | PASS |
| `edge-no-shared-strings-part`<br>hand-built | PASS | PASS | PASS | PASS | PASS |
| `edge-hidden-and-veryhidden-sheets`<br>hand-built | PASS | PASS | PASS | 91% (cell-missing 1) | PASS |

### Run `excel-365` (2026-09-12-excel-365)

2026-09-12T15:41:27.883Z, MacBook-Air.local, Node v24.18.0.

| fixture (generator) | sheetjs | office-kit | openpyxl | calamine |
|---|---|---|---|---|
| `hostile-biff8-xls-renamed`<br>Microsoft Excel 16.112.4 macOS<br>expect `NOT_XLSX` | ACCEPTED (Data: 31 rows, Features: 12 rows, It's a very long sheet name 001: 2 rows, Hidden: 2 rows) | REJECTED: OpenXmlNotImplementedError: Encrypted xlsx is not supported. Decrypt with msoffcrypto-tool | REJECTED: BadZipFile: File is not a zip file | REJECTED: ZipError: invalid Zip archive: Could not find EOCD |

### Run `goldens` (2026-09-12-goldens)

2026-09-12T04:07:38.145Z, MacBook-Air.local, Node v24.18.0.

| fixture (generator) | excel |
|---|---|
| `golden-canonical-sheetjs`<br>SheetJS 0.20.3 | PASS |
| `golden-canonical-exceljs`<br>ExcelJS 4.4.0 | PASS |
| `golden-canonical-write-excel-file`<br>write-excel-file 2.3.10 | PASS |
| `golden-canonical-openpyxl`<br>openpyxl 3.1.5 | REPAIRED (see log) |
| `golden-canonical-xlsxwriter`<br>XlsxWriter 3.2.9 | PASS |
| `golden-canonical-libreoffice-resave`<br>LibreOffice 26.2.5.2 | PASS |
| `golden-canonical-libreoffice-from-csv`<br>LibreOffice 26.2.5.2 | REPAIRED (see log) |
| `golden-canonical-office-kit`<br>@office-kit/xlsx 0.11.0 | PASS |

#### What Excel holds in the sentinel cells (raw `value`; dates shown as Excel reports them)

Column `E` control chars, `F` `_x` escape literals, `G` formula-like text, `I` 17-digit integer, `J11` max double, `M` dates (1899-12-31 / 1900-02-28 / 1900-03-01), `N` time-only 12:34:56.789, `O` datetimes (2024-03-10 02:30 DST gap / 1900-01-01 12:00). Caveat: AppleScript date values pass through macOS local-time normalization, so 02:30 in a DST gap reads as 03:30.

| fixture | E2 | F2 | F3 | F7 | G2 | I2 | J11 | M2 | M3 | M4 | N2 | O2 | O6 |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| `sheetjs` | `a\u0001b` | `A` | `_x0041_` | `_X0041_` | `=SUM(A1)` | `9.0072E+15` | `1.798E+308` | `SunDecember 31, 1899 12:00:00 ` | `ThursMarch 1, 1900 12:00:00 AM` | `ThursMarch 1, 1900 12:00:00 AM` | `-0.475731608796` | `SunMarch 10, 2024 3:30:00 AM` | `MonJanuary 1, 1900 12:00:00 PM` |
| `exceljs` | `ab` | `A` | `_x0041_` | `_X0041_` | `=SUM(A1)` | `9.01E+15` | `1.8E+308` | `MonJanuary 1, 1900 7:00:00 AM` | `ThursMarch 1, 1900 7:00:00 AM` | `ThursMarch 1, 1900 7:00:00 AM` | `0.815935057872` | `SunMarch 10, 2024 9:30:00 AM` | `TuesJanuary 2, 1900 7:00:00 PM` |
| `write-excel-file` | `ab` | `A` | `_x0041_` | `_X0041_` | `=SUM(A1)` | `9.0072E+15` | `1.798E+308` | `MonJanuary 1, 1900 7:00:00 AM` | `ThursMarch 1, 1900 7:00:00 AM` | `ThursMarch 1, 1900 7:00:00 AM` | `0.815935057872` | `SunMarch 10, 2024 9:30:00 AM` | `TuesJanuary 2, 1900 7:00:00 PM` |
| `openpyxl` | `` | `A` | `_x0041_` | `_X0041_` | `0` | `9.01E+15` | `#NULL!` | `SunDecember 31, 1899 12:00:00 ` | `WednesFebruary 28, 1900 12:00:` | `ThursMarch 1, 1900 12:00:00 AM` | `0.524268391204` | `SunMarch 10, 2024 3:30:00 AM` | `MonJanuary 1, 1900 12:00:00 PM` |
| `xlsxwriter` | `a\u0001b` | `_x0041_` | `_x005FA` | `_X0041_` | `=SUM(A1)` | `9.01E+15` | `#NULL!` | `SunDecember 31, 1899 12:00:00 ` | `WednesFebruary 28, 1900 12:00:` | `ThursMarch 1, 1900 12:00:00 AM` | `0.524268391204` | `SunMarch 10, 2024 3:30:00 AM` | `SunDecember 31, 1899 12:00:00 ` |
| `libreoffice-resave` | `ab` | `A` | `_x0041_` | `_X0041_` | `=SUM(A1)` | `9E+15` | `2E+308` | `MonJanuary 1, 1900 7:00:00 AM` | `ThursMarch 1, 1900 7:00:00 AM` | `ThursMarch 1, 1900 7:00:00 AM` | `0.815935057872` | `SunMarch 10, 2024 9:30:00 AM` | `TuesJanuary 2, 1900 7:00:00 PM` |
| `libreoffice-from-csv` | `a\u0001b` | `_x0041_` | `_x0041_` | `_X0041_` | `0` | `9.0072E+15` | `1.7977E+308` | `MonJanuary 1, 1900 12:00:00 AM` | `ThursMarch 1, 1900 12:00:00 AM` | `ThursMarch 1, 1900 12:00:00 AM` | `12:34:56.789` | `SunMarch 10, 2024 3:30:00 AM` | `TuesJanuary 2, 1900 12:00:00 P` |
| `office-kit` | `a\u0001b` | `_x0041_` | `_x005FA` | `_X0041_` | `=SUM(A1)` | `9.01E+15` | `1.8E+308` | `SunDecember 31, 1899 7:00:00 A` | `WednesFebruary 28, 1900 7:00:0` | `ThursMarch 1, 1900 7:00:00 AM` | `-0.18406494213` | `SunMarch 10, 2024 9:30:00 AM` | `MonJanuary 1, 1900 7:00:00 PM` |

### Run `hostile` (2026-09-12-hostile)

2026-09-12T15:42:31.907Z, MacBook-Air.local, Node v24.18.0.

| fixture (generator) | sheetjs | office-kit | office-kit-stream | openpyxl | calamine | validator |
|---|---|---|---|---|---|---|
| `hostile-biff8-xls-renamed`<br>Microsoft Excel 16.112.4 macOS<br>expect `NOT_XLSX` | ACCEPTED (Data: 31 rows, Features: 12 rows, It's a very long sheet name 001: 2 rows, Hidden: 2 rows) | REJECTED: OpenXmlNotImplementedError: Encrypted xlsx is not supported. Decrypt with msoffcrypto-tool | REJECTED: OpenXmlNotImplementedError: Encrypted xlsx is not supported. Decrypt with msoffcrypto-tool | REJECTED: OSError: File contains no valid workbook part | REJECTED: CalamineError: File not found 'theme/theme/_rels/workbook.xml.rels' | FAIL: 1 schema error |
| `hostile-xlsb-renamed`<br>Microsoft Excel 16.112.4 macOS<br>expect `NOT_XLSX` | ACCEPTED (Data: 31 rows, Features: 12 rows, It's a very long sheet name 001: 2 rows, Hidden: 2 rows) | REJECTED: Error: readTagExp returned undefined at position 857. Context: "\u0000\u0000m\u0000i\u0000c\u0000r\u0000o\u0000s\u0000o\u0000f\u0000t\u0000.\u0000c\u0000o\u0000 | REJECTED: Error: readTagExp returned undefined at position 857. Context: "\u0000\u0000m\u0000i\u0000c\u0000r\u0000o\u0000s\u0000o\u0000f\u0000t\u0000.\u0000c\u0000o\u0000 | REJECTED: OSError: File contains no valid workbook part | REJECTED: CalamineError: File not found 'xl/_rels/workbook.xml.rels' | FAIL: 1 schema error |

### Run `phase-b-writer` (2026-09-12-phase-b-writer)

2026-09-12T17:27:25.181Z, MacBook-Air.local, Node v24.18.0.

| fixture (generator) | validator | sheetjs | openpyxl | calamine | libreoffice | excel |
|---|---|---|---|---|---|---|
| `golden-canonical-simple-excel`<br>simple-excel 0.0.0 | PASS | 93% (time-only-has-date-part 24, temporal-mismatch 6) | 94% (escape-sequence-mangled 30, dst-gap-shift-1h 5) | 95% (blank-vs-empty-string 19, dst-gap-shift-1h 5) | PASS | PASS |
| `golden-canonical-simple-excel-inline`<br>simple-excel 0.0.0 | PASS | 93% (time-only-has-date-part 24, temporal-mismatch 6) | 94% (escape-sequence-mangled 30, dst-gap-shift-1h 5) | 95% (blank-vs-empty-string 19, dst-gap-shift-1h 5) | PASS | PASS |
| `golden-canonical-simple-excel-zip64`<br>simple-excel 0.0.0 | PASS | ERROR | 94% (escape-sequence-mangled 30, dst-gap-shift-1h 5) | 95% (blank-vs-empty-string 19, dst-gap-shift-1h 5) | PASS | PASS |

### Run `user-goldens` (2026-09-12-user-goldens)

2026-09-12T15:30:31.725Z, MacBook-Air.local, Node v24.18.0.

| fixture (generator) | excel |
|---|---|
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
