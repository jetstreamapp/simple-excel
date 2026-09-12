# Google Sheets golden (manual)

1. In Google Drive, upload `fixtures/golden/exceljs/canonical.xlsx` and open it with Google Sheets
   (this converts it to a native Sheet). Note any "features lost" notice.
2. Check the sentinel cells on the `Data` sheet: E2 (control char), F2/F3 (`_x0041_` literals), I2 (17-digit),
   M2/N2/O2 (date/time/datetime), R31 (32,767 chars), S30 (truncated 40,000-char cell), and the `Features`
   sheet (merged header, freeze, autofilter, hyperlink B7, rich text B8, note B9, validation B10, cond. format B11,
   hidden row 12, hidden column D, hidden sheet).
3. File > Download > Microsoft Excel (.xlsx) → `fixtures/golden/gsheets/canonical.from-exceljs.xlsx`.
4. Also upload `fixtures/canonical/canonical.csv`, open as a Sheet, download as .xlsx →
   `fixtures/golden/gsheets/canonical.from-csv.xlsx` (Google's CSV typing heuristics; expect no `<dimension>`).
5. Register:
   `node fixtures/register.mjs golden/gsheets/canonical.from-exceljs.xlsx --id golden-canonical-gsheets-resave --generator "Google Sheets" --provenance "manual:golden/gsheets/STEPS.md" --license "Google Sheets export" --tags kind:golden,generator:gsheets --expected canonical/canonical.json --notes "<notices seen>"`
6. Run `npm run oracle -- --tag generator:gsheets --label gsheets`.

Notices observed (2026-09-12): no 'features lost' banner.

> (Resolved 2026-09-12: the redo below was done; the imported sheet is named `canonical.csv`.)
> The first `canonical.from-csv.xlsx` was NOT kept: the file downloaded for step 4 contained the same four sheets as the
> ExcelJS import (Data, Features, …, Hidden, plus the `_xlnm._FilterDatabase` name), so it was a second export of
> the converted workbook rather than a CSV import. To redo step 4: create a NEW blank Sheet, File > Import >
> Upload canonical.csv (Replace spreadsheet), then File > Download > .xlsx.

`fixtures/golden/exceljs/canonical.xlsx`:
M2/N2/O2 (date/time/datetime) -> these cells did not show a specified "format" (Format > Number > nothing was checked in the list)

S30 -> ended with ...(truncated) (32767 chars)

validation
Features: condftm "Input must be an item on the specified list" (was dropdown with Red,Green,Blue)
