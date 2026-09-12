---
id: writing
title: Writing
description: createWorkbookWriter, sinks, sheets, rows, styles and every writer option.
---

# Writing

A write has three moving parts: a **sink** that receives bytes, a **workbook writer** that owns the zip container,
and one **sheet writer** at a time that turns rows into XML.

```ts
import { collectToBlob, createWorkbookWriter } from '@jetstreamapp/simple-excel';

const sink = collectToBlob();
const workbook = createWorkbookWriter(sink);

const sheet = workbook.addSheet('Accounts', {
  header: ['Id', 'Name', 'Amount', 'Created'],
  freeze: { rows: 1 },
  autoFilter: true,
  columns: [{ width: 22 }, { width: 40 }, { width: 12 }, { width: 20 }],
});

await sheet.writeRow(['001xx000003DGb2AAG', 'Acme', 1234.5, new Date(2024, 2, 10, 9, 30)]);
await sheet.close();

const result = await workbook.close();
const blob = await sink.result();
```

Bytes reach the sink while rows are still being written: the first chunk of `xl/worksheets/sheet1.xml` is handed
over long before `close()`. Nothing is buffered except the zip central directory (a few KB) and whatever the
shared-string budget allows.

## `createWorkbookWriter(sink, options?)`

```ts
function createWorkbookWriter(sink: ByteSink | WritableStream<Uint8Array>, options?: WorkbookWriterOptions): WorkbookWriter;
```

A `WritableStream<Uint8Array>` is accepted anywhere a `ByteSink` is; it is wrapped with `fromWritableStream` for
you.

`WorkbookWriter` is:

| Member                     | Meaning                                                                                     |
| -------------------------- | ------------------------------------------------------------------------------------------- |
| `addSheet(name, options?)` | Opens a sheet. One at a time — close it before adding the next                              |
| `registerStyle(style)`     | Interns a `CellStyle` and returns a `StyleId` you pass to `writeRow`                        |
| `close()`                  | Finalizes shared strings, styles, workbook parts and the central directory; closes the sink |
| `abort(reason?)`           | Tears the whole thing down and aborts the sink                                              |

## Sinks

A sink is three methods:

```ts
interface ByteSink {
  write(chunk: Uint8Array): Promise<void>;
  close(): Promise<void>;
  abort(reason?: unknown): Promise<void>;
}
```

`write` may resolve late, and the writer waits, so a slow sink slows the whole pipeline down instead of queueing
in memory. Four built-in sinks cover the common cases.

| Sink                           | Result                         | Use it for                                                       |
| ------------------------------ | ------------------------------ | ---------------------------------------------------------------- |
| `collectToBlob(type?)`         | `result(): Promise<Blob>`      | Browser downloads. Folds chunks into sub-Blobs every 32 MiB      |
| `collectToBytes({ maxBytes })` | `result(): Uint8Array` (sync)  | Tests, small files, contexts with no `Blob`. Default cap 1 GiB   |
| `fromWritableStream(stream)`   | whatever the stream does       | File System Access, OPFS, `Writable.toWeb()`, a network response |
| `toWritableStream(sink)`       | a `WritableStream<Uint8Array>` | The other direction, for `pipeTo` consumers                      |

```ts
import { collectToBlob, collectToBytes, fromWritableStream } from '@jetstreamapp/simple-excel';

// browser download
const blobSink = collectToBlob();
// ... write ...
const blob = await blobSink.result();

// in memory, with an explicit ceiling
const bytesSink = collectToBytes({ maxBytes: 64 * 1024 * 1024 });
// ... write ...
const bytes = bytesSink.result();

// straight to disk through the File System Access API
const handle = await window.showSaveFilePicker({ suggestedName: 'accounts.xlsx' });
const fileSink = fromWritableStream(await handle.createWritable());
```

`collectToBlob` and `collectToBytes` expose `bytesWritten` while the write is running. `collectToBytes` throws
`LIMIT_EXCEEDED` past `maxBytes`; `collectToBlob` throws `UNSUPPORTED_ENVIRONMENT` at construction if the platform
has no `Blob`.

[Streaming and memory](./streaming-and-memory.md) covers which sink to choose for a given size and platform, and
the Node file sinks are in [Node](./node.md).

## Sheets

```ts
const sheet = workbook.addSheet('Accounts', options);
```

| `SheetOptions` | Type                       | Default | Notes                                                                            |
| -------------- | -------------------------- | ------- | -------------------------------------------------------------------------------- |
| `header`       | `readonly CellInput[]`     | —       | Written as row 1, bold unless you say otherwise                                  |
| `headerStyle`  | `StyleId \| false`         | bold    | `false` writes the header unstyled; a `StyleId` replaces the default             |
| `columns`      | `readonly ColumnOptions[]` | —       | `{ width?, hidden?, style? }` per column, left to right                          |
| `freeze`       | `{ rows?, cols? }`         | —       | `{ rows: 1 }` freezes the header row                                             |
| `autoFilter`   | `boolean`                  | `false` | Filter over the header row and every written column                              |
| `hidden`       | `boolean`                  | `false` | `state="hidden"` in the workbook part                                            |
| `rowCount`     | `number`                   | —       | Data rows, excluding the header. Enables `<dimension>` and up-front zip64 sizing |

`ColumnOptions.width` is Excel's character-width unit, and Excel's schema caps it at 255 — a wider value produces
a file Excel opens but the Open XML validator rejects (`EC-COLS-WIDTH-OVER-255`).

`SheetWriter` is:

| Member                      | Meaning                                                                       |
| --------------------------- | ----------------------------------------------------------------------------- |
| `name`                      | The name actually used, after sanitizing and de-duplication                   |
| `nextRow`                   | 1-based index of the row the next `writeRow` will produce                     |
| `writeRow(values, styles?)` | One row                                                                       |
| `writeRows(rows)`           | An `Iterable` or `AsyncIterable` of rows, written in order with back-pressure |
| `merge(range)`              | Registers a merged range in A1 notation; emitted at sheet close               |
| `close()`                   | Resolves to a `SheetWriteSummary` (`{ name, rows, columns }`)                 |

### Sheet names

Excel's rules are enforced for you, so `addSheet` never fails on a name a user typed. `: \ / ? * [ ]` become `_`,
leading and trailing apostrophes are stripped, the name is trimmed to 31 characters, the reserved name `History`
(any casing) becomes `History_`, an empty name becomes `Sheet<n>`, and a case-insensitive collision gets ` (2)`,
` (3)` … fitted inside the 31-character budget. Read `sheet.name` if you need to know what was used — the mapping
matters when you later look the sheet up by name.

## Rows and values

```ts
await sheet.writeRow(['001xx000003DGb2AAG', 'Acme', 1234.5, true, new Date(), null]);

await sheet.writeRows(records.map(record => [record.Id, record.Name]));
await sheet.writeRows(asyncGeneratorOfRows());
```

Rows are strictly sequential: row _n_ is serialized and gone before row _n+1_ starts. There is no way to revisit a
written row, and there are no cell coordinates in the API — position in the array is the column.

`CellInput` is `string | number | boolean | Date | null | undefined | bigint | CellError`:

| Input                | What is written                                                                                   |
| -------------------- | ------------------------------------------------------------------------------------------------- |
| `string`             | Shared string or inline string (see below); `_xHHHH_`-escaped, `xml:space="preserve"` when needed |
| `number`             | `<v>` with the shortest round-trip form. `NaN` and `±Infinity` become the error cell `#NUM!`      |
| `boolean`            | `t="b"` with `1` / `0`                                                                            |
| `Date`               | A serial under a date number format. An **invalid `Date` writes nothing** — no `<c>` at all       |
| `null` / `undefined` | No `<c>` element, unless a non-default style was given for that cell (then an empty styled cell)  |
| `bigint`             | A number while `\|value\| <= 2^53`; a string beyond that, so no digits are silently lost          |
| `{ error: '#N/A' }`  | A real error cell (`t="e"`), not the text `#N/A`                                                  |

Text that starts with `=` stays text. The writer never infers a formula from a value — that is both a
formula-injection surface and a reliable way to make Excel show its repair dialog
(`EC-FORMULA-LIKE-TEXT-WRITTEN-AS-FORMULA`).

[Dates and values](./dates-and-values.md) has the full value model, including what Excel does with control
characters, CRLF and pre-1900 dates.

### Per-cell and per-row styles

`writeRow`'s second argument is either one `StyleId` for the whole row or one per cell:

```ts
const currency = workbook.registerStyle({ numFmt: '#,##0.00' });
const warn = workbook.registerStyle({ fill: { color: '#FFF3CD' } });

await sheet.writeRow([id, name, amount], [undefined, undefined, currency]);
await sheet.writeRow([id, name, amount], warn); // whole row
```

`undefined` in the per-cell array means the default style.

## Styles

```ts
const styleId = workbook.registerStyle({
  font: { bold: true, italic: false, size: 11, color: '#1B4F72', name: 'Calibri' },
  fill: { color: '#EAF2F8' },
  border: { top: 'thin', bottom: 'thin', color: '#B3B6B7' },
  alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
  numFmt: '#,##0.00',
});
```

| `CellStyle` field | Shape                                                                                                 |
| ----------------- | ----------------------------------------------------------------------------------------------------- |
| `font`            | `{ bold?, italic?, underline?, strike?, size?, color?, name? }`; `color` is `#RRGGBB`                 |
| `fill`            | `{ color: '#RRGGBB' }` — solid fill only                                                              |
| `border`          | A line style for all four sides, or `{ top?, bottom?, left?, right?, color? }`                        |
| `alignment`       | `{ horizontal?: 'left' \| 'center' \| 'right', vertical?: 'top' \| 'center' \| 'bottom', wrapText? }` |
| `numFmt`          | A format code (`'0.00'`, `'yyyy-mm-dd'`) or a built-in numFmt id                                      |

Line styles are `'thin' | 'medium' | 'thick' | 'dashed' | 'dotted' | 'double' | 'hair'`.

`registerStyle` interns: two identical `CellStyle` objects return the same `StyleId`, so calling it inside a row
loop is safe but pointless. `StyleId` 0 is the default style. Register every style you need up front and keep the
ids.

### Number formats

A `numFmt` string that matches a built-in code resolves to the built-in id and emits no `<numFmt>` element — so
`'0.00'` becomes id 2, `'@'` becomes id 49. Anything else is allocated a custom id from 164 upward. You can also
pass a built-in id directly (`numFmt: 14`).

Number formats are **not rendered** on read: a cell's value is the number, never the displayed text. See
[Dates and values](./dates-and-values.md#numbers).

### The default date format

A `Date` written without a date style is given one automatically, because a bare serial under the General format
shows as a five-digit number in Excel. The default code is **`yyyy-mm-dd hh:mm:ss`**, registered once per
workbook. If the style you pass for that cell already carries a date or time number format, yours is used instead:

```ts
const dateOnly = workbook.registerStyle({ numFmt: 'yyyy-mm-dd' });
await sheet.writeRow([new Date()], dateOnly); // your format wins
await sheet.writeRow([new Date()]); // yyyy-mm-dd hh:mm:ss
```

## Merges

```ts
sheet.merge('A1:C1');
```

A1 notation, registered while the sheet is open and emitted at close in the order added. A single-cell range
(`'A1:A1'`) throws `WRITER_STATE`: Excel opens the repair dialog for one-cell merges.

## Shared strings

Excel can store strings in a workbook-wide table (`xl/sharedStrings.xml`) and have cells point at indexes, or
inline them in the sheet. The table shrinks files with repeated values, and grows in memory with every unique one
— which is exactly how a streaming writer runs out of heap on a column of record ids.

The default is a **bounded hybrid** (ADR-001): strings are interned while the table is under budget, then the map
is frozen — existing entries still resolve, new strings go inline — and memory stops growing.

| Option                | Values                           | Default  | Meaning                                                                               |
| --------------------- | -------------------------------- | -------- | ------------------------------------------------------------------------------------- |
| `strings`             | `'auto' \| 'inline' \| 'shared'` | `'auto'` | `'inline'` never builds a table; `'shared'` interns everything (unbounded, for tests) |
| `sstBudget.maxUnique` | number                           | 65,536   | Stop interning once this many unique strings exist                                    |
| `sstBudget.maxChars`  | number                           | 16 Mi    | Stop interning once the interned text totals this many UTF-16 units                   |
| `sstBudget.maxLength` | number                           | 256      | Strings longer than this are always inline                                            |

`'shared'` switches the budget off entirely, so it will run out of memory on a large unique-heavy sheet. Use it
only when you are testing something about the table itself.

The result of `workbook.close()` reports what happened:

```ts
const result = await workbook.close();
result.sharedStrings; // { count, uniqueCount, frozen }
```

`frozen: true` means the budget was reached and the rest of the workbook was written inline. Mixed `s` and
`inlineStr` cells in one sheet are legal, and every reader in the compatibility matrix accepts them.

## Zip64

An entry larger than 4 GiB needs zip64 headers, and Excel only accepts them when they were declared **in the local
file header** — that is, before the first byte of the entry, when the writer cannot yet know how big it will be.
So the decision has to be made up front (ADR-002).

| `zip64`            | Behaviour                                                                                                                                       |
| ------------------ | ----------------------------------------------------------------------------------------------------------------------------------------------- |
| `'auto'` (default) | A worksheet gets zip64 only when its declared `rowCount` makes a > 4 GiB part plausible (about 40M cells). Everything else stays a standard zip |
| `true`             | Every streamed part gets zip64 headers                                                                                                          |
| `false`            | Never emitted                                                                                                                                   |

:::caution
`'auto'` is deliberately conservative because **SheetJS 0.20.3 cannot open a zip64 archive at all** ("Unsupported
ZIP file" even for a small one) and **Google Drive fails to convert one to a Google Sheet** (`EC-ZIP64-SMALL`).
Excel, LibreOffice, openpyxl and calamine all read them. Only turn `zip64: true` on when you know your consumers
can handle it.
:::

A sheet written without a `rowCount` stays 32-bit, and if a part does pass 4 GiB the writer fails with
`ENTRY_TOO_LARGE` naming `zip64: true` as the fix. Announcing `rowCount` is the way to get zip64 exactly when it
is needed:

```ts
const sheet = workbook.addSheet('Huge', { header, rowCount: records.length });
```

## Compression and the deflater

| Option        | Values                 | Default     | Meaning                                                    |
| ------------- | ---------------------- | ----------- | ---------------------------------------------------------- |
| `compression` | `'deflate' \| 'store'` | `'deflate'` | `'store'` writes uncompressed entries: faster, much larger |
| `deflater`    | `DeflaterFactory`      | platform    | Swap in another compressor                                 |

The default compressor is `CompressionStream('deflate-raw')`, which has no compression level. Where
`CompressionStream` is unavailable the writer silently switches to `'store'` rather than claiming method 8 over
uncompressed bytes. In Node, `nodeDeflater(level)` gives you zlib and a level — level 1 is roughly 2–3× faster
than the default on xlsx-shaped XML:

```ts
import { nodeDeflater, toFile } from '@jetstreamapp/simple-excel/node';

const workbook = createWorkbookWriter(toFile('out.xlsx'), { deflater: nodeDeflater(1) });
```

## Deterministic output

```ts
const workbook = createWorkbookWriter(sink, { deterministic: true });
```

Fixes the zip entry timestamps and the `docProps` dates, so the same rows produce byte-identical output. The
repository's golden-bytes suite depends on this; it is also what you want for content-addressed caching or for
diffing two exports.

## Progress, abort and cancellation

```ts
const controller = new AbortController();

const workbook = createWorkbookWriter(sink, {
  signal: controller.signal,
  onProgress: ({ sheet, rows, bytesOut }) => updateUi(sheet, rows, bytesOut),
  onCellTruncated: count => console.warn(`${count} cells truncated`),
});
```

- `onProgress` fires every 5,000 rows and once at each sheet close.
- `signal` is checked on every `writeRow`, `addSheet`, `registerStyle` and `close`. Once it has fired, the next
  call throws `ABORTED` and the sink is torn down.
- `workbook.abort(reason?)` does the same thing imperatively. Call it in a `catch` so a half-written stream is not
  left open.

```ts
try {
  await writeEverything();
} catch (error) {
  await workbook.abort(error);
  throw error;
}
```

## Long cells: the 32,767-character limit

Excel refuses a cell longer than 32,767 characters. The writer applies a policy rather than producing a file that
opens with a repair prompt (`EC-CELL-32767-LIMIT`):

| Option             | Values                    | Default            |
| ------------------ | ------------------------- | ------------------ |
| `cellOverflow`     | `'truncate' \| 'throw'`   | `'truncate'`       |
| `truncationSuffix` | string                    | `'...(truncated)'` |
| `onCellTruncated`  | `(count: number) => void` | —                  |

With `'truncate'`, the text is cut so that the suffix still fits inside the limit, and the count is reported
through `onCellTruncated` and in the final result. With `'throw'` a long cell raises `CELL_TOO_LONG`.

## The result

`workbook.close()` resolves to a `WorkbookWriteResult`:

```ts
interface WorkbookWriteResult {
  readonly bytes: number;
  readonly sheets: readonly { name: string; rows: number; columns: number }[];
  readonly truncatedCells: number;
  readonly sharedStrings: { count: number; uniqueCount: number; frozen: boolean };
}
```

`bytes` is the size of the finished archive, `sheets` carries the sanitized name and the actual extent of each
sheet, and `truncatedCells` is the workbook-wide truncation count. `sheet.close()` returns the same per-sheet
summary earlier, if you need it before the workbook finishes.

## All writer options

| Option             | Type                                    | Default                |
| ------------------ | --------------------------------------- | ---------------------- |
| `strings`          | `'auto' \| 'inline' \| 'shared'`        | `'auto'`               |
| `sstBudget`        | `{ maxUnique?, maxChars?, maxLength? }` | 65,536 / 16 Mi / 256   |
| `zip64`            | `'auto' \| boolean`                     | `'auto'`               |
| `compression`      | `'deflate' \| 'store'`                  | `'deflate'`            |
| `deflater`         | `DeflaterFactory`                       | platform               |
| `deterministic`    | `boolean`                               | `false`                |
| `dates`            | `'local' \| 'utc'`                      | `'local'`              |
| `date1904`         | `boolean`                               | `false`                |
| `cellOverflow`     | `'truncate' \| 'throw'`                 | `'truncate'`           |
| `truncationSuffix` | `string`                                | `'...(truncated)'`     |
| `onCellTruncated`  | `(count: number) => void`               | —                      |
| `onProgress`       | `(progress: WriteProgress) => void`     | —                      |
| `signal`           | `AbortSignal`                           | —                      |
| `properties`       | `{ creator?, title?, created? }`        | creator `simple-excel` |
