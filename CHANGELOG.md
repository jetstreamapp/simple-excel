# Changelog

All notable changes to this project will be documented in this file.

## [Unreleased]

## [0.2.0] - 2026-09-28

### Added

- `OpenOptions.onWarning`, with the `ReadWarning` and `ReadWarningCode` types: the reader reports damage it repairs
  instead of refusing, once per code per sheet. The first codes are a shared-string index past the end of the table
  and a shared-string cell in a workbook with no shared-string part; both cells read blank, as in Excel.
- `openWorkbook` accepts any typed array or `DataView` and a `SharedArrayBuffer`, and recognises buffers from another
  realm (an iframe, a worker, an Electron bridge).

### Changed

- Object mode (`rows({ mode: 'object' })` and `toObjects()`) finds the header row: without `headerRow` it is the first
  row at or after `startRow` that holds a value, so a sheet whose first rows are blank reads like `sheet_to_json`
  instead of naming every column `__EMPTY`. Columns start at the sheet's `<dimension>`, so an all-blank column A no
  longer becomes `__EMPTY`, and `toObjects()` gives every record every key. An explicit `headerRow` behaves as before.
- A header cell holding an error value names its column with the error text (`#N/A`), as in SheetJS.
- `onCellTruncated` fires once, after `close()`, with the workbook's total. It used to fire at every sheet close
  with a running total.
- `rowCount` is only a zip64 sizing hint: the writer no longer writes `<dimension>` from it. A hint lower than the
  rows written made SheetJS and pandas read only the declared rows.
- The writer refuses input that used to produce a file Excel repairs, with `WRITER_STATE`: a workbook with no sheet or
  with every sheet hidden, style ids `registerStyle` did not return, error values that are not a `#` code, overlapping
  or duplicate merges (checked when `merge()` is called), style fields outside Excel's ranges, non-numeric or negative
  widths, `freeze` and `rowCount` values, a row that is not an array, and a `properties.created` outside the years
  1-9999. Column widths above 255 are written as 255. Unsupported cell values name the sheet, cell and type.
- Any failure while writing aborts the sink once, and every later call rejects with that original error.
- A `Date` written with a style that has no date format keeps that style's font, fill, border and alignment.
- An error value outside `CellErrorCode` that is still a `#` code (a newer Excel error such as `#SPILL!`, which the
  reader can return) is written as its text instead of as an error cell Excel would repair.
- A `Date` from another realm is accepted as a cell value and as `properties.created`.
- Control characters in sheet names become `_` and are dropped from the title and creator; a font name or number
  format code holding one is refused.
- A serial past 9999-12-31 in a date-formatted cell reads as a number, and a `Date` after 9999 is written as ISO text.
- `sniff` classifies Windows-1252, Shift-JIS and UTF-16 (with a byte-order mark) text as `'text'`.
- `engines.node` is `>=20.12`, the first Node whose `CompressionStream` and `DecompressionStream` accept `deflate-raw`.

### Fixed

- A column headed `__proto__` lost its values, and a `Date` under it replaced the record's prototype.
- A workbook whose part names differ in case from its relationships (`xl/SharedStrings.xml`) read every string as
  blank; part names now match case-insensitively.
- `writeRow({ Id, Name })` silently wrote an empty row.
- Control characters in a sheet name or document property produced a part that LibreOffice and openpyxl refuse.
- Truncation to 32,767 characters could split an emoji.
- A `properties.created` that was an invalid `Date` silently left out `docProps/core.xml`.
- A row array reused and changed before the previous `writeRow` settled could be written with the later values.
- The active tab now always points at a visible sheet when the first sheet is hidden.
- `openWorkbook` threw a bare `TypeError` for input that is not bytes; it now throws `NOT_XLSX`. Reading on a Node
  without `DecompressionStream('deflate-raw')` throws `UNSUPPORTED_ENVIRONMENT`, and reading a sheet after
  `workbook.close()` throws `ABORTED` for every kind of source.
- A numeric `<v>` of only whitespace read as 0, radix-prefixed text such as `0x1A` read as a number, and a cell
  reference past column XFD wrapped into the next column (it now fails with `LIMIT_EXCEEDED`).
- Reading a part whose header declared a huge uncompressed size allocated that much memory before the size limit
  was checked; the limit now comes first and the up-front allocation is capped at 16 MiB.
- A test that failed in time zones observing summer time on the day it ran.

## [0.1.1] - 2026-09-26

### Changed

- Updated the development toolchain (vitest 5, oxfmt 0.70, oxlint 1.85, and the office-kit, read-excel-file and
  write-excel-file versions the corpus and benchmarks compare against). There are no runtime dependency changes, and
  the published bundles are byte-identical to 0.1.0.

## [0.1.0] - 2026-09-12

### Added

- Streaming workbook writer (`createWorkbookWriter`): one open sheet at a time, rows pushed to a sink as they are
  written, inline strings by default with an optional bounded shared-string table, styles (fonts, fills, borders, alignment, number
  formats), header rows, column widths, freeze panes, autofilter, merges, hidden sheets, deterministic output,
  progress and abort, 32,767-character cell policy, zip64 on demand, and a stored-parts fallback where
  `CompressionStream` is unavailable.
- Streaming workbook reader (`openWorkbook`): format sniffing with classified errors (password-protected, legacy
  `.xls`, `.xlsb`, `.ods`, csv/html/text, empty), sheet listing without parsing, array and object row modes with
  SheetJS-compatible header naming, `toObjects()` and `head()`, typed values (numbers, booleans, Dates from
  number formats, errors), Strict Open XML, missing `r` attributes, rich and inline strings, `_xHHHH_` decoding, and
  configurable limits against hostile input (zip bombs, DOCTYPE, deep nesting, truncated or duplicated entries).
- Sinks for browsers and Node: `collectToBlob`, `collectToBytes`, `fromWritableStream`, `toWritableStream`, and the
  `/node` entry with `fromFile`, `toFile`, `toWritable` and a zlib-backed `nodeDeflater`.
- `npm run size`: a bundle-size gate on the browser entry (`dist/esm/index.mjs`), which fails above 40 KB brotli.
  It currently measures 24.9 KB.
- The research corpus (goldens from Excel 365, Google Sheets, Numbers, LibreOffice, Salesforce reports, SheetJS,
  ExcelJS, openpyxl, XlsxWriter, office-kit, write-excel-file, plus edge and hostile fixtures), the compatibility
  oracle (Excel, LibreOffice, Open XML SDK validator, SheetJS, openpyxl, calamine, office-kit) and the benchmark
  harness the library is verified against.

### Changed

- Reading a workbook that has no shared-strings part (SheetJS writes one for every export Jetstream makes) no
  longer scans every chunk of every sheet for `t="s"` cells that cannot exist. Typed reads of a 100,000-row
  Salesforce-shaped export went from 2.48 s to 2.10 s.
- Faster hot paths, with byte-identical output: CRC-32 is slicing-by-16 (560 MB/s to 2.8 GB/s), cell text finds
  the "nothing to escape" case with one precompiled pattern instead of a per-character walk (and no longer treats
  an emoji as work), each flush encodes through a reused buffer with `encodeInto`, and the XML tokenizer parses a
  tag's attributes once instead of once per lookup. Writing 100,000 Salesforce-shaped rows went from 4.56 s to
  2.77 s (1.71 s with `nodeDeflater(1)`), 1,000,000 rows from 45.8 s to 31.9 s, and typed reads of the same
  100,000 rows from 2.10 s to 1.62 s.

[Unreleased]: https://github.com/jetstreamapp/simple-excel/compare/0.2.0...HEAD
[0.2.0]: https://github.com/jetstreamapp/simple-excel/compare/0.1.1...0.2.0
[0.1.1]: https://github.com/jetstreamapp/simple-excel/compare/0.1.0...0.1.1
[0.1.0]: https://github.com/jetstreamapp/simple-excel/releases/tag/0.1.0
