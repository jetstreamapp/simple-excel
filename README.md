# simple-excel

> Streaming, dependency-free xlsx reader and writer for browsers and Node.

**Status: pre-release, under active development.** The API below is the target; `CHANGELOG.md` records what has
landed. The library is built against a committed corpus of real files and a compatibility oracle that includes
Excel itself.

`@jetstreamapp/simple-excel` writes a workbook by pushing rows into a sink as they arrive, and reads one by pulling
rows through an async iterator, so memory stays flat whether the sheet has a thousand rows or a million. There are
no runtime dependencies: the zip container, the deflate stream, the XML tokenizer and the SpreadsheetML layer are
all in the package, and compression is the platform's own `CompressionStream`.

## Why another xlsx library

|                                             | simple-excel            | SheetJS CE                        | ExcelJS             | @office-kit/xlsx                   | write-excel-file / read-excel-file |
| ------------------------------------------- | ----------------------- | --------------------------------- | ------------------- | ---------------------------------- | ---------------------------------- |
| Streaming write (flat memory at 1M rows)    | yes                     | no                                | yes                 | no (buffers until `finalize()`)    | no                                 |
| Streaming read                              | yes                     | no                                | yes                 | yes                                | no                                 |
| Runtime dependencies                        | 0                       | 0                                 | 9                   | 3                                  | 3 each                             |
| Bundle size (min+brotli)                    | _(measured in Phase D)_ | not measured                      | not measured        | ≤ 120 KB (its README)              | not measured                       |
| Maintained on npm                           | yes                     | no (fixes on the vendor CDN only) | inactive since 2023 | yes (pre-1.0)                      | yes                                |
| Verified against Excel with a public corpus | yes                     | no                                | no                  | no (validator + fixtures in CI)    | no                                 |
| Styles (fonts, fills, borders, number fmts) | write                   | Pro only                          | yes                 | yes                                | basic, write only                  |
| Formulas, charts, pivots, editing workbooks | no                      | partial                           | partial             | yes                                | no                                 |
| Legacy formats (.xls, .xlsb, .ods, CSV)     | no, detected and named  | yes, silently                     | CSV only            | no, every CFB reported "encrypted" | no                                 |

Sources: dependency counts are each package's `dependencies` field at the versions pinned in `package.json`;
maintenance and architecture come from [`research/03-library-landscape.md`](research/03-library-landscape.md); the
office-kit column is [`research/10-office-kit-evaluation.md`](research/10-office-kit-evaluation.md), which found
that its streaming writer buffers the worksheet and an unbounded shared-string table until `finalize()`
(`RangeError: Invalid string length` at 1M × 20 — the same wall SheetJS hits). Only office-kit publishes a bundle
figure; the other libraries have not been measured here, and ours is measured in Phase D. "Verified against Excel
with a public corpus" means a committed fixture corpus plus an oracle that opens every written file in Excel and
fails on a repair prompt — see [`research/05-compatibility-matrix.md`](research/05-compatibility-matrix.md).

Pick simple-excel when you export or import tabular data and it has to be large, fast and correct. Pick a
full-featured library when you edit existing workbooks or need charts, pivots and formulas.

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

The core entry needs Web Streams and `CompressionStream('deflate-raw')`: **Chrome 103+, Firefox 113+, Safari 16.4+
and Node 20+**. Where `CompressionStream` is absent entirely the writer falls back to storing parts uncompressed —
the file is still a valid xlsx, only larger. Node only accepts the `deflate-raw` format from 21.2, so on Node 20 the
writer falls back to stored parts too; pass the Node entry's `nodeDeflater()` there to keep files compressed.

No DOM API is used beyond `TextEncoder`/`TextDecoder` and a duck-typed `Blob`, so the library runs unchanged in a
Web Worker, an MV3 service worker and an Electron renderer.

## What it does not do

Not in v1 (from [`research/07-reference-architecture.md`](research/07-reference-architecture.md) §9): writing or
evaluating formulas; rendering number formats; hyperlinks, comments, images, charts, tables, conditional
formatting, data validation; editing an existing workbook or preserving unknown parts through a round trip;
opening encrypted files (they are detected and named, not decrypted); `.xls`, `.xlsb`, `.ods`, SpreadsheetML 2003
and HTML; CSV parsing; theme or indexed colour resolution on read; reading from a non-seekable stream; a
JavaScript inflate/deflate; rich-text writing; column-level styles; print setup; sheet protection.

## Documentation

The user documentation site is in [`docs/`](docs/). The design behind the library is in [`research/`](research/):
the format primer, the edge-case catalog, the compatibility matrix across eight readers, the performance baseline
and the architecture decision records.

## License

MIT
