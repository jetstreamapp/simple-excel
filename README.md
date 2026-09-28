# simple-excel

> Streaming, dependency-free xlsx reader and writer for browsers and Node.

**Status: 0.1, published on npm.** The API may still change before 1.0; `CHANGELOG.md` records every release. The
library is built against a committed corpus of real files and a compatibility oracle that includes Excel itself.

`@jetstreamapp/simple-excel` writes a workbook by pushing rows into a sink as they arrive, and reads one by pulling
rows through an async iterator, so memory stays flat whether the sheet has a thousand rows or a million. There are
no runtime dependencies: the zip container, the deflate stream, the XML tokenizer and the SpreadsheetML layer are
all in the package, and compression is the platform's own `CompressionStream`.

## Why another xlsx library

|                                             | simple-excel           | SheetJS CE                        | ExcelJS             | @office-kit/xlsx                | write-excel-file / read-excel-file |
| ------------------------------------------- | ---------------------- | --------------------------------- | ------------------- | ------------------------------- | ---------------------------------- |
| Streaming write (flat memory at 1M rows)    | yes                    | no                                | yes                 | no (buffers until `finalize()`) | no                                 |
| Streaming read                              | yes                    | no                                | yes                 | yes                             | no                                 |
| Runtime dependencies                        | 0                      | 0                                 | 9                   | 3                               | 1 / 4                              |
| Bundle size (min+brotli)                    | 29.0 KB                | not measured                      | not measured        | ≤ 120 KB (its README)           | not measured                       |
| Maintained on npm                           | yes                    | no (fixes on the vendor CDN only) | inactive since 2023 | yes (pre-1.0)                   | yes                                |
| Verified against Excel with a public corpus | yes                    | no                                | no                  | no (validator + fixtures in CI) | no                                 |
| Styles (fonts, fills, borders, number fmts) | write                  | Pro only                          | yes                 | yes                             | basic, write only                  |
| Formulas, charts, pivots, editing workbooks | no                     | partial                           | partial             | yes                             | no                                 |
| Legacy formats (.xls, .xlsb, .ods, CSV)     | no, detected and named | yes, silently                     | CSV only            | no; names `.xls`, not `.xlsb`   | no                                 |

Sources: dependency counts are each package's `dependencies` field at the versions pinned in `package.json`;
maintenance and architecture come from [`research/03-library-landscape.md`](research/03-library-landscape.md); the
office-kit column is [`research/10-office-kit-evaluation.md`](research/10-office-kit-evaluation.md), which found
that its streaming writer buffers the worksheet and an unbounded shared-string table until `finalize()`
(`RangeError: Invalid string length` at 1M × 20 — the same wall SheetJS hits). That evaluation was of 0.11.0; the
legacy-format cell was re-checked at 0.23.0. Only office-kit publishes a bundle figure; the other libraries have not
been measured here. Ours is `dist/esm/index.mjs` compressed at brotli quality 11, checked on every build by
`npm run size`. "Verified against Excel with a public corpus" means a committed fixture corpus plus an oracle that
opens every written file in Excel and fails on a repair prompt — see
[`research/05-compatibility-matrix.md`](research/05-compatibility-matrix.md).

Pick simple-excel when you export or import tabular data and it has to be large, fast and correct. Pick a
full-featured library when you edit existing workbooks or need charts, pivots and formulas.

## Measured

100,000 rows × 20 columns of Salesforce-shaped data (ids, unicode names, decimals, booleans, dates, long text,
JSON blobs, 10% nulls — about 1.7 KB of text per row), Apple M4, Node v24.18.0, median of three runs in a fresh
process. Full tables and method: [`research/06-performance-baseline.md`](research/06-performance-baseline.md),
runs `bench/results/2026-09-12-macbook-air-inline-default` (100k rows, the default configuration) and
`…-phase-e-optimized` / `…-phase-e-scale` / `…-phase-e-chrome` (1M rows, 18M cells and Chrome, measured with the
optional bounded shared-string table, `strings: 'auto'`, before inline became the default).

|                             | simple-excel                 | SheetJS 0.20.3                      |
| --------------------------- | ---------------------------- | ----------------------------------- |
| Write                       | 3.08 s                       | 9.75 s                              |
| Write peak memory over idle | 58 MB                        | 3,313 MB                            |
| Read (typed records)        | 1.64 s                       | 6.17 s                              |
| Read peak memory over idle  | 810 MB                       | 1,837 MB                            |
| First byte at the sink      | 0.5 ms                       | n/a (nothing until the end)         |
| 1,000,000 × 20 write        | 31.9 s, 77 MB of heap growth | `RangeError: Invalid string length` |
| 18M cells (900,000 × 20)    | 27.0 s, 76 MB of heap growth | `RangeError: Invalid string length` |

The 810 MB read figure is the 100,000 materialized records the typed read returns, not the parser: streaming the
same file with `sheet.rows()` costs 1.48 s and 46 MB. With `strings: 'auto'` (the bounded shared-string table)
the write is 2.77 s: fewer bytes reach the compressor, which is the bottleneck.

In a Chrome module worker (Chromium 153, writing to `collectToBlob()`), simple-excel writes 100k × 20 in 3.88 s
for +236 MB of renderer RSS and 1,000,000 × 20 in 39.6 s for +254 MB — flat across a 10× row increase. SheetJS
takes 7.64 s and +3,039 MB at 100k and crashes the renderer at 1M
(`bench/results/2026-09-12-macbook-air-phase-e-chrome`).

In Node, passing the `/node` entry's `nodeDeflater(1)` instead of the platform `CompressionStream` cuts write time
by about 44% (1.73 s at 100k, 19.1 s at 1M) for a file about 20% larger.

## Install

```sh
npm install @jetstreamapp/simple-excel
```

## Write

```ts
import { collectToBlob, createWorkbookWriter } from '@jetstreamapp/simple-excel';

const sink = collectToBlob();
const workbook = createWorkbookWriter(sink);
const sheet = workbook.addSheet('Accounts', {
  header: ['Id', 'Name', 'Created'],
  freeze: { rows: 1 },
  autoFilter: true,
});

for await (const record of records) {
  await sheet.writeRow([record.Id, record.Name, record.CreatedDate]);
}

await sheet.close();
const result = await workbook.close();
const blob = await sink.result();
console.log(`${result.bytes} bytes, ${result.sheets[0]?.rows} rows`);
```

One sheet is open at a time: close it before adding the next. Cell values are `string | number | boolean | Date |
null`, plus `undefined`, `bigint` and `{ error: '#N/A' }`.

## Read

```ts
import { openWorkbook } from '@jetstreamapp/simple-excel';

const workbook = await openWorkbook(file); // File | Blob | ArrayBuffer | Uint8Array | RandomAccessSource

for (const { name, index, kind, hidden } of workbook.sheets) {
  console.log(index, name, kind, hidden);
}

const sheet = workbook.sheet('Accounts');
for await (const row of sheet.rows({ mode: 'object' })) {
  // row is Record<string, string | number | boolean | Date | null>
}

await workbook.close();
```

`rows()` streams; nothing but the shared-string table stays resident. `sheet.head(5)` returns the first rows as an
A1-keyed map without inflating the rest of the sheet, which is how you inspect a template before reading it.

## Node

```ts
import { createWorkbookWriter, fromFile, nodeDeflater, toFile } from '@jetstreamapp/simple-excel/node';

const workbook = createWorkbookWriter(toFile('accounts.xlsx'), { deflater: nodeDeflater(1) });
const source = await fromFile('upload.xlsx');
```

The `/node` entry re-exports everything from the core entry and adds file-backed sources and sinks plus a
zlib-backed deflater with a compression level (the platform `CompressionStream` has no level).

## Browser support

The core entry needs Web Streams and `CompressionStream`/`DecompressionStream('deflate-raw')`: **Chrome 103+,
Firefox 113+, Safari 16.4+ and Node 20.12+** (21.2+ on the 21 line). Where `CompressionStream` is missing or rejects
`deflate-raw`, the writer stores parts uncompressed; the file is still a valid xlsx, only larger. On an older Node,
pass the Node entry's `nodeDeflater()` to keep files compressed. Reading has no fallback: it needs
`DecompressionStream('deflate-raw')`.

No DOM API is used beyond `TextEncoder`/`TextDecoder` and a duck-typed `Blob`, so the library runs unchanged in a
Web Worker, an MV3 service worker and an Electron renderer.

## What it does not do

Not in v1 (from [`research/07-reference-architecture.md`](research/07-reference-architecture.md) §9): writing or
evaluating formulas; rendering number formats; hyperlinks, comments, images, charts, tables, conditional
formatting, data validation; editing an existing workbook or preserving unknown parts through a round trip;
opening encrypted files (they are detected and named, not decrypted); `.xls`, `.xlsb`, `.ods`, SpreadsheetML 2003
and HTML; CSV parsing; theme or indexed colour resolution on read; reading from a non-seekable stream; a
JavaScript inflate/deflate; rich-text writing; print setup; sheet protection.

## Documentation

The user documentation is at [simple-excel.getjetstream.app](https://simple-excel.getjetstream.app) (source in
[`docs/`](docs/)). The design behind the library is in [`research/`](research/):
the format primer, the edge-case catalog, the compatibility matrix across eight readers, the performance baseline
and the architecture decision records.

## License

MIT
