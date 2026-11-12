# Changelog

All notable changes to this project will be documented in this file.

## [Unreleased]

## [0.3.0] - 2026-11-12

### Added

- `limits.maxPartBytes` (default 32 MiB): the most a part the reader holds whole may declare or inflate to. That is
  `[Content_Types].xml`, the package and workbook relationships and the workbook part; worksheets, the shared-string
  table and `styles.xml` are streamed and stay under `maxInflatedBytes`.
- `limits.maxRowChars` (default 64 Mi UTF-16 units): the most text one row of a sheet may hold, all its cells
  together.
- `IO_ERROR`, a new `XlsxErrorCode` for failures underneath the library: a sink's `write` or `close`, a Node stream
  or file (including an HTTP response whose client disconnected), a `RandomAccessSource` read, the compressor
  (`CompressionStream`, `nodeDeflater`, a custom deflater) and a throwing `onProgress`, `onCellTruncated` or
  `onWarning` callback. The original error is the standard `cause`.
- `XlsxError` carries a `Symbol.for('@jetstreamapp/simple-excel/XlsxError')` brand and a `Symbol.hasInstance`, so
  `instanceof XlsxError` and `isXlsxError()` hold for an error from any copy of the library: the core and `/node`
  entries, and the ESM and CommonJS builds loaded side by side.
- The `XlsxError` constructor takes an optional fourth `options` argument (`{ cause }`, the standard `ErrorOptions`);
  `new XlsxError(code, message, detail)` is unchanged.
- `limits.maxObjectCells` (default 100,000,000): the record cells (records × keys) `toObjects()` holds at once.
  Past it `toObjects()` throws `LIMIT_EXCEEDED` with a message pointing at the streaming `rows({ mode: 'object' })`.
  A record cell costs about 40 to 60 bytes in Node, so a legitimate wide load (200,000 rows × 300 columns is 60M)
  fits and a 16,384-column header over 65k rows stops a tenth of the way in.
- `ObjectRowsOptions.headers`: name the columns yourself (`headers[i]` keys the i-th column from A) instead of
  reading a header row; every record has exactly those keys.
- `DATA_BEYOND_HEADER`, a new `ReadWarningCode`: object mode left out values to the right of the last header.
  Reported once per sheet when the read ends, counting only the records the read returned (a `toObjects()` that
  fails reports nothing), with `detail.columns` (such as `'E:XFD'`), `columnCount`, `cellCount` and
  `lastHeaderColumn`.
- `CELL_FORMATS_IGNORED`, a new `ReadWarningCode`: the styles part lists more than 65,490 `cellXfs` or `numFmts`
  and a number cell uses one past that ceiling, so it reads as a plain number (a date there comes back as its
  serial).
- `ReadWarning.detail`: structured facts about a warning, for a UI to show (`DATA_BEYOND_HEADER` and
  `CELL_FORMATS_IGNORED` set it).
- `sheet.rows()` accepts a `mode` that is only known at run time (`'array' | 'object'`, the new `AnyModeRowsOptions`
  type) and yields `V[] | Record<string, V>`.

### Changed

- **Breaking:** `limits.maxTextLength` defaults to 1,048,576 characters instead of 64 Mi, and now caps what one
  cell's value, inline string or formula text, or one shared-string item, accumulates however many text runs build
  it, not only each run. Excel holds 32,767 characters in a cell and Salesforce's longest text field 131,072; pass a
  larger `maxTextLength` to read generated files with longer text.
- **Breaking:** `ABORTED` now means only that the caller cancelled (an `AbortSignal` or `abort()`), or read from a
  workbook or source they had closed. A sink, stream, file or source that fails on its own is `IO_ERROR`; it used to
  be reported as `ABORTED` (sinks) or escape as a bare platform `Error` (source reads, `CompressionStream` and zlib
  failures, `fromFile` on a missing file).
- **Breaking:** the original error of a classified failure is the standard `Error.cause`, no longer
  `detail.cause`; the writer's abort reason is the `cause` too, no longer `detail.reason`. `detail` only carries
  structured facts (limits, part names, paths).
- **Breaking:** `limits.maxSharedStringChars` counts 16 UTF-16 units per shared-string item on top of its text, so
  the default 256 Mi admits about 16.7M items. A table with more items than that, or with many items near it and
  long text, needs a larger `maxSharedStringChars`.
- **Breaking:** object mode (`rows({ mode: 'object' })` and `toObjects()`) keys each record only up to the last
  header cell that holds a value, or exactly the caller's `headers`. Values to the right of it are left out and
  reported as `DATA_BEYOND_HEADER`, where SheetJS `sheet_to_json` (and earlier versions) named every column out to
  the widest row and the sheet's `<dimension>`. Blank header cells inside the header keep their `__EMPTY`,
  `__EMPTY_1`, ... names; a dimension wider than the data no longer adds trailing `__EMPTY` keys; a merged header
  names only its first column; and an explicit `headerRow` on a blank row names no columns, so its records are empty.
- **Breaking:** `isXlsxError()` checks the brand instead of `name === 'XlsxError'`, so a plain object that only has
  that name no longer passes.
- **Breaking:** `toFile(path)` writes to a hidden temporary file in the same directory
  (`.<name>.<random>.partial`) and renames it over `path` when `close()` resolves, after the descriptor is closed.
  The target never holds a partial file, an existing file there is replaced only by a complete one, and on abort or
  failure the temporary file is deleted. The directory must allow creating files, and the new file gets the default
  mode rather than the replaced file's.
- **Breaking:** `openWorkbook` takes ownership of the source it is given: when opening fails for any reason (sniffed
  format, truncated zip, cancelled signal) the source is closed before the promise rejects.
- `openWorkbook(input, options)` with an `errors` that is not a literal (`const options: OpenOptions = ...`) is now
  typed `Workbook<ReadValue>` (`CellValue | CellError`), matching what it yields at run time; it was typed
  `Workbook<CellValue>`.
- The `/node` bundles import the core bundle (`@jetstreamapp/simple-excel`) instead of carrying their own copy of it.

### Security

- `openWorkbook()` refuses a package part that lists more `<sheet>`, `<Relationship>`, `<Override>` or `<Default>`
  elements than `limits.maxEntries` with `LIMIT_EXCEEDED`. A 58 KB file listing 5M `<sheet/>` elements reached
  1.9 GB and ran Node out of memory.
- `styles.xml` keeps at most 65,490 `cellXfs` and `numFmts` entries (Excel's own ceiling) and ignores the rest, so a
  flood of `<xf/>` elements no longer grows the reader's tables. The workbook still reads: some non-Excel writers
  emit one xf per cell, and a number cell that uses an ignored xf reads as a plain number with a
  `CELL_FORMATS_IGNORED` warning.
- The parts read whole are capped by the new `maxPartBytes` before they are inflated and again as they inflate.
  `styles.xml` is streamed through the tokenizer instead, keeping only its bounded format tables, so a heavily
  styled workbook's tens of MB of styles read in flat memory. A 698 KB file with 700 MiB of whitespace in
  `styles.xml` reached 2.2 GB and then failed with a bare `ERR_STRING_TOO_LONG` instead of an `XlsxError`; it now
  stops at `maxTextLength` after 1 Mi characters.
- `maxTextLength` holds for a whole cell and shared-string item, and `maxRowChars` for a row, so text split into
  many runs or spread over many cells no longer builds unbounded strings (about 1.4 GB from 1 MB of upload). A
  string past the engine's maximum length now surfaces as `LIMIT_EXCEEDED` rather than a `RangeError`.
- A shared-string table of empty `<si/>` items is held to `maxSharedStringChars`: every item counts 16 units for
  its slot. 100M of them (a 713 KB upload) read "successfully" at 2 to 3.6 GB RSS; the default now stops after
  about 16.7M items, in about a second.
- The XML tokenizer no longer rescans a tag from its start on every chunk while it waits for the `>`: a 1 MB file of
  1 MiB tags cost about 69 s of CPU per read, against about 3 s for a benign file of the same inflated size.
- Rows are yielded in batches bounded by the cells they hold (about 1M slots) rather than one batch per decoded
  chunk, so rows with a cell at column XFD (a 16,384-slot array each) no longer pile up before the first yield: a
  146 KB file reached 1.26 GB. `maxColumns` still defaults to 16,384, so a stray far-right cell reads as before.
- Object mode no longer widens every record to a stray far-right value: one value at XFD gave every record 16,384
  keys, so `toObjects()` on a 703 KB upload of 65k rows ran Node out of a 12 GB heap and streaming its records took
  104 s. Records now stop at the header (the same file reads in well under a second), and `toObjects()` stops at
  `maxObjectCells`.
- With `blankRows: true`, a jump in row numbers is filled under the same batch budget instead of queuing every blank
  row first (about 1M rows from a 1.8 KB file).

### Fixed

- Object-mode header keys built from a non-integer number are spelled the way Excel's General format shows them
  (`4.27289175`), as SheetJS names them, instead of with `String(value)`, which leaked float noise
  (`4.2728917499999985`). Integers are unchanged. Header keys are part of the `sheet_to_json` contract this restores,
  so this is a fix rather than a breaking change; only keys that carried float noise or more than 11 characters of
  fraction change.
- `toWritable` no longer hangs when the stream is destroyed without an error (an HTTP client disconnecting
  mid-download): a write waiting for `'drain'` or a close waiting for `'finish'` now settles on `'close'` and rejects
  with `IO_ERROR`, a write to a destroyed or ended stream is refused at once, and `close()` resolves for a stream that
  already finished.
- `workbook.abort()` cancels a write stuck on back-pressure: the abort reaches the sink at once instead of queueing
  behind the pending write, so `abort()` resolves and the pending `writeRow` rejects with `ABORTED`, for a stalled
  `WritableStream` and a paused Node stream alike. The same holds for a `WritableStream` whose `close()` never
  finishes: `abort()` resolves and the pending `workbook.close()` rejects with `ABORTED`.
- `openWorkbook` closes a `RandomAccessSource` (such as `fromFile`) when the sniff step or the zip directory read
  rejects; an encrypted, `.xls`, CSV, empty or truncated upload used to leak the file handle.
- `toFile` no longer leaves a partial file at the target after an abort or a failed write.
- A failing `CompressionStream` or zlib stream, and the `/node` helpers' internal stream errors, are reported as an
  `XlsxError` (`IO_ERROR`) instead of a bare `Error`.
- Object-mode header keys built from date cells no longer shift with the server's time zone under `dates: 'utc'`.
- An archive with exactly 65,535 entries (or a central-directory offset or size of exactly 0xFFFFFFFF) gets its zip64
  end records: the writer compared against the 0xFFFF / 0xFFFFFFFF markers with `>` and wrote the marker with no
  zip64 record behind it. The reader also accepts such archives from other writers, reading the 32-bit values when no
  zip64 locator is present and they describe a directory inside the file, as Python and Info-ZIP do.
- `collectToBytes()` stays failed after `LIMIT_EXCEEDED`, so a caller that catches the limit and keeps writing cannot
  `close()` a file with a chunk missing.

## [0.2.1] - 2026-10-31

### Security

- Object mode (`rows({ mode: 'object' })` and `toObjects()`) no longer takes quadratic time to look for a worksheet's
  `<dimension>`. A crafted sheet whose prologue was repeated unterminated `<dimension` starts made the pattern rescan
  the whole chunk from every start (about 4 s for a 256 KiB stored sheet, far worse for larger chunks), blocking the
  event loop or browser main thread. The match is now linear in the prologue.

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

[Unreleased]: https://github.com/jetstreamapp/simple-excel/compare/0.3.0...HEAD
[0.3.0]: https://github.com/jetstreamapp/simple-excel/compare/0.2.1...0.3.0
[0.2.1]: https://github.com/jetstreamapp/simple-excel/compare/0.2.0...0.2.1
[0.2.0]: https://github.com/jetstreamapp/simple-excel/compare/0.1.1...0.2.0
[0.1.1]: https://github.com/jetstreamapp/simple-excel/compare/0.1.0...0.1.1
[0.1.0]: https://github.com/jetstreamapp/simple-excel/releases/tag/0.1.0
