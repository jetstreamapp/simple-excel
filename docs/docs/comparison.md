---
id: comparison
title: Comparison
description: How simple-excel differs from SheetJS, ExcelJS, @office-kit/xlsx and write-excel-file, and when to use each.
---

# Comparison

There are good reasons to use another library. This page is the honest version of the README table: what each
one does well, where it breaks, and which of those breakages matter for your case.

Everything here comes from two documents in the repository — a cross-language survey of how xlsx libraries
achieve bounded memory ([`research/03-library-landscape.md`](https://github.com/jetstreamapp/simple-excel/blob/main/research/03-library-landscape.md))
and a detailed evaluation of `@office-kit/xlsx` against this corpus
([`research/10-office-kit-evaluation.md`](https://github.com/jetstreamapp/simple-excel/blob/main/research/10-office-kit-evaluation.md)) —
plus the benchmark runs in
[`research/06-performance-baseline.md`](https://github.com/jetstreamapp/simple-excel/blob/main/research/06-performance-baseline.md).

## At a glance

|                                             | simple-excel           | SheetJS CE                        | ExcelJS             | @office-kit/xlsx                   | write-excel-file / read-excel-file |
| ------------------------------------------- | ---------------------- | --------------------------------- | ------------------- | ---------------------------------- | ---------------------------------- |
| Streaming write (flat memory at 1M rows)    | yes                    | no                                | yes                 | no (buffers until `finalize()`)    | no                                 |
| Streaming read                              | yes                    | no                                | yes                 | yes                                | no                                 |
| Runtime dependencies                        | 0                      | 0                                 | 9                   | 3                                  | 3 each                             |
| Bundle size (min+brotli)                    | 24.3 KB                | not measured                      | not measured        | ≤ 120 KB (its README)              | not measured                       |
| Maintained on npm                           | yes                    | no (fixes on the vendor CDN only) | inactive since 2023 | yes (pre-1.0)                      | yes                                |
| Verified against Excel with a public corpus | yes                    | no                                | no                  | no (validator + fixtures in CI)    | no                                 |
| Styles                                      | write                  | Pro only                          | yes                 | yes                                | basic, write only                  |
| Formulas, charts, pivots, editing           | no                     | partial                           | partial             | yes                                | no                                 |
| Legacy formats (.xls, .xlsb, .ods, CSV)     | no, detected and named | yes, silently                     | CSV only            | no, every CFB reported "encrypted" | no                                 |

## Measured performance

100k rows × 20 columns of Salesforce-shaped data (ids, unicode names, decimals, booleans, dates, long text with
newlines, 10% nulls — about 1.7 KB of text per row), Node v24.18.0 on an Apple M4, median of three runs, each
engine in a fresh process. "Footprint" is the peak RSS above the pre-run baseline; "first byte" is how long until
the sink receives anything. Every number below is from one run,
[`2026-09-12-macbook-air-phase-d-baseline`](https://github.com/jetstreamapp/simple-excel/tree/main/bench/results/2026-09-12-macbook-air-phase-d-baseline),
plus its `phase-d-scale` companion for the 1M column.

| Engine                  | Write time | Write footprint | Read time  | Read footprint | 1M × 20 write | First byte |
| ----------------------- | ---------- | --------------- | ---------- | -------------- | ------------- | ---------- |
| SheetJS 0.20.3          | 9.43 s     | 2,896 MB        | 6.44 s     | 2,037 MB       | `RangeError`  | n/a        |
| ExcelJS 4.4.0           | 2.52 s     | 317 MB          | 11.07 s    | 178 MB         | not run       | 1.62 s     |
| @office-kit/xlsx 0.11.0 | 4.70 s     | 981 MB          | 6.70 s     | 452 MB         | `RangeError`  | 2.9 ms     |
| write-excel-file 2.3.10 | 3.62 s     | 1,117 MB        | 9.04 s     | 4,484 MB       | not run       | n/a        |
| **simple-excel**        | **4.56 s** | **108 MB**      | **2.10 s** | **881 MB**     | **45.8 s**    | **0.4 ms** |

Two footnotes on that row, in both directions. simple-excel's 881 MB read footprint is the 100,000 materialized
records the typed read returns, not the parser: streaming the same file with `sheet.rows()` costs 1.97 s and
75 MB. And its write time is the platform `CompressionStream`; in Node, `nodeDeflater(1)` brings it to 3.32 s for
a file about 20% larger. The 1M × 20 write grows the JS heap by 75 MB, and the 18M-cell shape (900,000 × 20)
completes in 40.8 s for 78 MB.

The `RangeError` is `Invalid string length`: V8 refuses to build a string past about 512 MiB, and both SheetJS and
office-kit build one — the sheet XML in SheetJS's case, the joined shared-string table in office-kit's. It is not
a tuning problem; it is the architecture. In a Chrome module worker the same split shows up as a crash: writing
1,000,000 × 20 to a Blob costs simple-excel 51.7 s and +253 MB of renderer RSS, while SheetJS takes the renderer
down with it.

## SheetJS CE

The default choice for a decade, and still the most tolerant reader in existence: it will read `.xls`, `.ods`,
CSV bytes and SpreadsheetML 2003 without being asked.

**Where it breaks.** It materializes the whole workbook, and builds the sheet XML as one string. That fails at a
predictable size — around 958k rows or 18M cells — with `RangeError: Invalid string length`. There is no
streaming read or write in the community edition.

**Maintenance.** npm is frozen at 0.18.5, which carries CVE-2023-30533 (prototype pollution) and CVE-2024-22363
(ReDoS). The fixes exist, but only on the vendor's own CDN, so `npm install xlsx` gets you the vulnerable
version and a lockfile that cannot be audited normally.

**Use it** for small files where its format tolerance is the feature — a "paste anything in" import path. Be
aware that reading `.xls` and CSV _silently_ means users get different type coercion depending on what they
uploaded, with no signal to you or to them.

**Migrating.** Object mode reproduces `sheet_to_json`'s key naming exactly (`__EMPTY`, `__EMPTY_1`, duplicate
`_1`/`_2` suffixes, `defval`, `blankrows: false`), and `dates: 'local'` matches its `cellDates: true` output. The
[reading page](./reading.md#header-naming) has the details. Two documented differences remain: SheetJS returns
the raw number for the fake 1900-02-29, collapses CRLF to LF, and decodes `_X0041_` with a capital X (none of
which this library does).

:::caution
SheetJS 0.20.3 cannot open a zip64 archive at all. If it is a consumer of your files, leave `zip64: 'auto'` alone.
:::

## ExcelJS

The most complete of the open-source JavaScript libraries: full styles, formulas, images, data validation, both a
DOM API and genuinely streaming reader and writer classes.

**Where it breaks.** The streaming façade sits over DOM code, and every intermediate string is not bounded:
`Invalid string length` on write is reported even through the stream API
([#1868](https://github.com/exceljs/exceljs/issues/1868)), and stream reads have OOM'd on a 6M-row file
([#355](https://github.com/exceljs/exceljs/issues/355)). Its measured first byte at 100k rows is 1.62 s, so the
"streaming" writer is doing a great deal before anything reaches the sink. Its shared-string reader once pushed
each rich-text run as a separate entry, shifting every later index
([#1431](https://github.com/exceljs/exceljs/issues/1431)).

**Maintenance.** 4.4.0 shipped in October 2023; the maintainers have said they are inactive
([#2969](https://github.com/exceljs/exceljs/issues/2969)). Nine runtime dependencies, including `archiver`,
`unzipper` and `jszip`.

**Use it** when you need the feature surface — editing workbooks, charts, images, conditional formatting — and
your files are of ordinary size.

## @office-kit/xlsx

The most interesting alternative, and the one this library was seriously evaluated against. MIT, TypeScript,
about 2,100 tests, an openpyxl fixture corpus as a submodule, and a three-tier validator in CI. One responsive
maintainer shipping same-day fixes.

**Where it breaks.** The streaming writer is streaming in API shape only. `appendRow` accumulates the worksheet
XML in memory and the zip entry is deflated and handed to the sink at `finalize()` — with `toFile()`, the output
file is 0 bytes until the very end. Every string is interned in an unbounded shared-string table that
`finalize()` serializes with one `Array.join`, which is the `RangeError` above. On the Salesforce-shaped dataset
that means a 981 MB footprint for 100k rows.

Two reader gaps also mattered here: it does not decode `_xHHHH_` escapes, so control characters and CRs come back
as literal `_x000d_` text; and it cannot open Salesforce report exports at all, because Apache POI writes
`<u val="none"/>` — a legal `ST_UnderlineValues` member — which its parser rejects for the whole workbook.
Dates come back as raw serials from both APIs, so the consumer has to resolve number formats itself.

**Use it** when you need its feature surface (styles, conditional formatting, data validation, charts, pivot and
VBA passthrough) on files that fit comfortably in memory, and you are not reading POI-generated files. It is
pre-1.0 and says its APIs may shift.

## write-excel-file / read-excel-file

A pair of small, focused, actively maintained libraries with a pleasant API. `read-excel-file` in particular has
a schema-mapping feature that is genuinely nice.

**Where they break.** Neither streams. `write-excel-file` builds XML from a fully materialized array of rows;
`read-excel-file`'s own README scopes it to "small to medium" files, and it measured a 4,484 MB footprint reading
100k × 20 — 2.2× SheetJS. `write-excel-file` also writes 1900-02-28 as serial 60, the leap-bug off-by-one.

**Use them** for reports of a known, modest size where the API and the schema mapping earn their place.

## What simple-excel gives up

Being honest about the other direction. There are no formulas, charts, pivot tables, images, comments,
hyperlinks, conditional formatting or data validation, on either side. You cannot open a workbook, change three
cells and save it — reading and writing are separate operations, and unknown parts are not preserved. Number
formats are stored but never rendered, so there is no display text, only values. Rich text is flattened on read
and cannot be written. `.xls`, `.xlsb`, `.ods` and CSV are detected and named, never parsed. And it is pre-1.0,
with one maintaining organisation.

The full list is [`research/07-reference-architecture.md`](https://github.com/jetstreamapp/simple-excel/blob/main/research/07-reference-architecture.md) §9.

## Choosing

| If…                                                                 | Then                          |
| ------------------------------------------------------------------- | ----------------------------- |
| The file might be large, or you cannot bound it                     | simple-excel                  |
| You are in a browser tab, worker or extension and memory is finite  | simple-excel                  |
| You need a dependency-free path for untrusted uploads               | simple-excel                  |
| You need to edit an existing workbook in place                      | ExcelJS or `@office-kit/xlsx` |
| You need charts, pivots, conditional formatting or data validation  | ExcelJS or `@office-kit/xlsx` |
| You need to accept `.xls`, `.ods` or CSV through the same code path | SheetJS, knowingly            |
| The file is small and a schema mapper would save you work           | read-excel-file               |
