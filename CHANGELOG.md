# Changelog

All notable changes to this project will be documented in this file.

## [Unreleased]

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

[Unreleased]: https://github.com/jetstreamapp/simple-excel/compare/0.1.0...HEAD
[0.1.0]: https://github.com/jetstreamapp/simple-excel/releases/tag/0.1.0
