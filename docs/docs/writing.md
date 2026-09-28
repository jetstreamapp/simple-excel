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
over long before `close()`. Nothing is buffered except the zip central directory (a few KB) and, under
`strings: 'auto'`, whatever the shared-string budget allows.

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
has no `Blob`. `collectToBytes().result()` is synchronous and throws `WRITER_STATE` until `workbook.close()` has
resolved; `collectToBlob().result()` can be called early and resolves when the workbook closes. After an abort, or a
failed write (which aborts the sink), both fail with the abort reason.

[Streaming and memory](./streaming-and-memory.md) covers which sink to choose for a given size and platform, and
the Node file sinks are in [Node](./node.md).

## Sheets

```ts
const sheet = workbook.addSheet('Accounts', options);
```

| `SheetOptions` | Type                       | Default | Notes                                                                                                                      |
| -------------- | -------------------------- | ------- | -------------------------------------------------------------------------------------------------------------------------- |
| `header`       | `readonly CellInput[]`     | —       | Written as row 1, bold unless you say otherwise                                                                            |
| `headerStyle`  | `StyleId \| false`         | bold    | `false` writes the header unstyled; a `StyleId` replaces the default                                                       |
| `columns`      | `readonly ColumnOptions[]` | —       | `{ width?, hidden?, style? }` per column, left to right                                                                    |
| `freeze`       | `{ rows?, cols? }`         | —       | `{ rows: 1 }` freezes the header row                                                                                       |
| `autoFilter`   | `boolean`                  | `false` | Filter over the header row and every written column                                                                        |
| `hidden`       | `boolean`                  | `false` | `state="hidden"` in the workbook part                                                                                      |
| `rowCount`     | `number`                   | —       | Data rows, excluding the header. A hint for up-front zip64 sizing; never written as `<dimension>`, so it need not be exact |

`ColumnOptions.width` is Excel's character-width unit, and Excel's schema caps it at 255: a wider value is written
as 255, and a negative, `NaN` or infinite width throws `WRITER_STATE` (`EC-COLS-WIDTH-OVER-255`). `addSheet` also
checks `freeze` (whole numbers that leave a scrolling cell inside the grid), `rowCount` (a whole number, 0 or more)
and every style id it is given, and a refused `addSheet` leaves the workbook as it was. `ColumnOptions.style` is the
column's default format (`<col style>`): it applies to cells the writer does not emit, such as ones a user types in
later. Cells you write keep the style passed to `writeRow`.

`SheetWriter` is:

| Member                      | Meaning                                                                          |
| --------------------------- | -------------------------------------------------------------------------------- |
| `name`                      | The name actually used, after sanitizing and de-duplication                      |
| `nextRow`                   | 1-based index of the row the next `writeRow` will produce                        |
| `writeRow(values, styles?)` | One row                                                                          |
| `writeRows(rows)`           | An `Iterable` or `AsyncIterable` of rows, written in order with back-pressure    |
| `merge(range)`              | Registers a merged range in A1 notation, checked at once; emitted at sheet close |
| `close()`                   | Resolves to a `SheetWriteSummary` (`{ name, rows, columns }`)                    |

### Sheet names

Excel's rules are enforced for you, so `addSheet` never fails on a name a user typed. `: \ / ? * [ ]` and control
characters (tab and line breaks included) become `_`,
leading and trailing apostrophes are stripped, the name is trimmed to 31 characters, the reserved name `History`
(any casing) becomes `History_`, an empty name becomes `Sheet<n>`, and a case-insensitive collision gets ` (2)`,
` (3)` … fitted inside the 31-character budget. Read `sheet.name` if you need to know what was used — the mapping
matters when you later look the sheet up by name.

### Sheet-name and cell-reference helpers

The rules above are exported, along with the A1 arithmetic the writer uses, for callers that key data by sheet
name before writing or build merge ranges from coordinates:

| Export                                            | What it does                                                                                                                                             |
| ------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `sanitizeSheetName(name, taken)`                  | Returns the name `addSheet` would pick. `taken` is a `Set<string>` of the **lower-cased** names already used; the result is added to it, lower-cased     |
| `isValidSheetName(name)`                          | `true` when the name needs no sanitizing. It does not check for collisions                                                                               |
| `MAX_SHEET_NAME_LENGTH`                           | `31`                                                                                                                                                     |
| `formatRef(row, col)`                             | 0-based row and column to an A1 reference: `formatRef(2, 1)` is `'B3'`                                                                                   |
| `formatRange(startRow, startCol, endRow, endCol)` | 0-based corners to a range, normalized so the top-left comes first: `formatRange(0, 0, 0, 2)` is `'A1:C1'`                                               |
| `parseRef(ref)`                                   | `'B3'` to `{ row: 2, col: 1 }` (a `CellRef`, 0-based); `$` anchors are tolerated; `null` when the text is not an upper-case A1 reference inside the grid |
| `parseRange(range)`                               | `'A1:C3'` to `{ start, end }` (a `CellRange`, normalized); a single reference parses as a one-cell range; `null` when malformed                          |
| `columnLetters(index)` / `columnIndexOf(letters)` | `0` to `'A'`, `26` to `'AA'`, and back; `columnIndexOf` returns `-1` for anything that is not upper-case column letters inside the grid                  |
| `MAX_ROWS` / `MAX_COLUMNS`                        | `1,048,576` / `16,384`                                                                                                                                   |

`formatRef`, `formatRange` and `columnLetters` throw `ROW_OUT_OF_RANGE` for a coordinate outside the grid, and
`sanitizeSheetName` throws `INVALID_SHEET_NAME` for a name that is not a string.

```ts
import { formatRange, sanitizeSheetName } from '@jetstreamapp/simple-excel';

const taken = new Set<string>();
const names = ['Accounts', 'accounts', 'Q1/Q2'].map(name => sanitizeSheetName(name, taken));
// ['Accounts', 'accounts (2)', 'Q1_Q2']

sheet.merge(formatRange(0, 0, 0, 2)); // 'A1:C1'
```

## Rows and values

```ts
await sheet.writeRow(['001xx000003DGb2AAG', 'Acme', 1234.5, true, new Date(), null]);

await sheet.writeRows(records.map(record => [record.Id, record.Name]));
await sheet.writeRows(asyncGeneratorOfRows());
```

Rows are strictly sequential: row _n_ is serialized and gone before row _n+1_ starts. There is no way to revisit a
written row, and there are no cell coordinates in the API — position in the array is the column. A row has to be an
array: an object such as `{ Id, Name }` throws `WRITER_STATE` rather than writing an empty row (`EC-ROW-NOT-ARRAY`),
so map records to arrays in header order first. A row array you reuse and change between calls is safe even when
you do not await each `writeRow`.

`CellInput` is `string | number | boolean | Date | null | undefined | bigint | CellError`:

| Input                | What is written                                                                                                                         |
| -------------------- | --------------------------------------------------------------------------------------------------------------------------------------- |
| `string`             | Inline string by default, a shared string under `strings: 'auto'` (see below); `_xHHHH_`-escaped, `xml:space="preserve"` when needed    |
| `number`             | `<v>` with the shortest round-trip form. `NaN` and `±Infinity` become the error cell `#NUM!`                                            |
| `boolean`            | `t="b"` with `1` / `0`                                                                                                                  |
| `Date`               | A serial under a date number format. An **invalid `Date` writes nothing** — no `<c>` at all. Before 1900 or after 9999: ISO text        |
| `null` / `undefined` | No `<c>` element, unless a non-default style was given for that cell (then an empty styled cell)                                        |
| `bigint`             | A number while `\|value\| <= 2^53`; a string beyond that, so no digits are silently lost                                                |
| `{ error: '#N/A' }`  | A real error cell (`t="e"`) for the `CellErrorCode` literals. Another `#` code (`#SPILL!`) is written as its text; anything else throws |

Anything else, such as a plain object, an array, a function or a symbol, throws `WRITER_STATE` with a message that
names the sheet, the cell and the type (`Sheet "Accounts" cell C7: a value of type object cannot be written…`).
SheetJS wrote an object as a blank cell and an array as its first element; the writer will not guess, so convert
such values (for example with `JSON.stringify`) before writing (`EC-CELL-UNSUPPORTED-TYPE`).

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
ids: `writeRow`, `columns` and `headerStyle` only accept ids that `registerStyle` returned (`EC-STYLE-ID-RANGE`).

`registerStyle` checks what Excel would otherwise repair or refuse, and throws `WRITER_STATE` naming the field: a
font size outside 1–409, a font name that is empty, longer than 31 characters or holds a control character, a
colour that is not `#RRGGBB`, a `fill` without a colour, an alignment or border value outside the lists above, a
number format code that is empty, longer than 255 characters or holds a control character, a numeric `numFmt` that is not a built-in or
registered id, and more than 64,000 styles in one workbook (`EC-STYLE-FIELD-RANGE`). A refused style registers
nothing and does not affect the workbook.

### Number formats

A `numFmt` string that matches a built-in code resolves to the built-in id and emits no `<numFmt>` element — so
`'0.00'` becomes id 2, `'@'` becomes id 49. Anything else is allocated a custom id from 164 upward. You can also
pass a built-in id directly (`numFmt: 14`).

Number formats are **not rendered** on read: a cell's value is the number, never the displayed text. See
[Dates and values](./dates-and-values.md#numbers).

### The default date format

A `Date` written without a date style is given one automatically, because a bare serial under the General format
shows as a five-digit number in Excel. The default code is **`yyyy-mm-dd hh:mm:ss`**, registered once per
workbook. If the style you pass for that cell has no number format of its own (a bold, filled row style, say), the
date cell keeps its font, fill, border and alignment and gains the default date format (`EC-DATE-STYLE-MERGE`). If
the style already carries a date or time number format, yours is used as it is:

```ts
const dateOnly = workbook.registerStyle({ numFmt: 'yyyy-mm-dd' });
await sheet.writeRow([new Date()], dateOnly); // your format wins
await sheet.writeRow([new Date()]); // yyyy-mm-dd hh:mm:ss
```

## Merges

```ts
sheet.merge('A1:C1');
```

A1 notation in either case (`'a1:c1'` works), registered while the sheet is open and emitted at close in the
order added. `merge()` checks the range at once and throws `WRITER_STATE` for a range that is not `A1:C1`-shaped,
a single cell (`'A1:A1'`), a range outside the grid, or one that overlaps or repeats a range already merged on the
sheet: Excel opens the repair dialog for each of these (`EC-MERGE-OVERLAP`). A refused merge does not affect the
rest of the workbook.

## Workbook-level rules

`workbook.close()` refuses two workbooks Excel would repair, with `WRITER_STATE`, and aborts the sink: one with no
sheet at all (`EC-WORKBOOK-NO-SHEETS`), and one whose sheets are all hidden (`EC-ALL-SHEETS-HIDDEN`). When the first
sheet is hidden, the first visible sheet is the one Excel opens on. `properties.created` must be a valid `Date`
between the years 1 and 9999, checked when the writer is created (`EC-DOCPROPS-CREATED-RANGE`). Control characters
in sheet names become `_` and are dropped from the title and creator, where XML cannot carry them; a font name or
number format code holding one is refused (`EC-XML-CONTROL-CHARS-METADATA`). Cell text keeps them, encoded as
`_xHHHH_`.

## Shared strings

Excel can store strings in a workbook-wide table (`xl/sharedStrings.xml`) and have cells point at indexes, or
inline them in the sheet. The table shrinks files with repeated values, and grows in memory with every unique one
— which is exactly how a streaming writer runs out of heap on a column of record ids.

The default is **inline** (ADR-001, revised): no table is built, memory does not grow with unique strings, the
compressed file is within a couple of percent of the table version on typical data (the platform deflater gets
about 20% more bytes, which costs roughly a tenth of the write time on repeat-heavy data), and it is the shape
every application reads faithfully — Numbers truncates shared strings at control characters and
mis-decodes protected `_xHHHH_` escapes in them, but reads inline strings correctly. It is also what SheetJS
writes by default, so files look the same as before to anyone migrating.

`strings: 'auto'` gives a **bounded hybrid**: strings are interned while the table is under budget, then the map
is frozen — existing entries still resolve, new strings go inline — and memory stops growing. Use it for very
low-cardinality data (picklists repeated across hundreds of thousands of rows) when file size matters more than
write speed.

| Option                | Values                           | Default    | Meaning                                                                               |
| --------------------- | -------------------------------- | ---------- | ------------------------------------------------------------------------------------- |
| `strings`             | `'inline' \| 'auto' \| 'shared'` | `'inline'` | `'auto'` builds a bounded table; `'shared'` interns everything (unbounded, for tests) |
| `sstBudget.maxUnique` | number                           | 65,536     | Stop interning once this many unique strings exist                                    |
| `sstBudget.maxChars`  | number                           | 16 Mi      | Stop interning once the interned text totals this many UTF-16 units                   |
| `sstBudget.maxLength` | number                           | 256        | Strings longer than this are always inline                                            |

`'shared'` switches the budget off entirely, so it will run out of memory on a large unique-heavy sheet. Use it
only when you are testing something about the table itself.

The result of `workbook.close()` reports what happened:

```ts
const result = await workbook.close();
result.sharedStrings; // { count, uniqueCount, frozen }
```

Under `'inline'` the counts are zero. Under `'auto'`, `frozen: true` means the budget was reached and the rest of the workbook was written inline. Mixed `s` and
`inlineStr` cells in one sheet are legal, and every reader in the compatibility matrix accepts them.

## Zip64

An entry larger than 4 GiB needs zip64 headers, and Excel only accepts them when they were declared **in the local
file header** — that is, before the first byte of the entry, when the writer cannot yet know how big it will be.
So the decision has to be made up front (ADR-002).

| `zip64`            | Behaviour                                                                                                                                                  |
| ------------------ | ---------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `'auto'` (default) | A worksheet gets zip64 only when its declared `rowCount` makes a > 4 GiB part plausible (about 40M cells, see below). Everything else stays a standard zip |
| `true`             | Every streamed part gets zip64 headers                                                                                                                     |
| `false`            | Never emitted                                                                                                                                              |

:::caution
`'auto'` is deliberately conservative because **SheetJS 0.20.3 cannot open a zip64 archive at all** ("Unsupported
ZIP file" even for a small one) and **Google Drive fails to convert one to a Google Sheet** (`EC-ZIP64-SMALL`).
Excel, LibreOffice, openpyxl and calamine all read them. Only turn `zip64: true` on when you know your consumers
can handle it.
:::

A sheet written without a `rowCount` stays 32-bit, and if a part does pass 4 GiB the writer fails with
`ENTRY_TOO_LARGE`, whose message says to enable zip64 (`zip64: true`) or split the data across sheets. Announcing
`rowCount` lets `'auto'` switch zip64 on for a sheet with a huge number of cells:

```ts
const sheet = workbook.addSheet('Huge', { header, rowCount: records.length });
```

`'auto'` estimates the part's size from its cell count (about 100 bytes a cell), so it cannot see long text. 140,000
rows of 32,000-character cells pass 4 GiB with an accurate `rowCount`, and the write fails cleanly with
`ENTRY_TOO_LARGE` at that point. When rows carry long text (email bodies, JSON blobs), estimate the uncompressed size
yourself and pass `zip64: true` when it could exceed about 3.5 GiB, accepting that SheetJS and Google Drive cannot
read that file.

## Compression and the deflater

| Option        | Values                 | Default     | Meaning                                                    |
| ------------- | ---------------------- | ----------- | ---------------------------------------------------------- |
| `compression` | `'deflate' \| 'store'` | `'deflate'` | `'store'` writes uncompressed entries: faster, much larger |
| `deflater`    | `DeflaterFactory`      | platform    | Swap in another compressor                                 |

The default compressor is `CompressionStream('deflate-raw')`, which has no compression level. Where
`CompressionStream` is unavailable or rejects `deflate-raw` the writer silently switches to `'store'` rather than
claiming method 8 over uncompressed bytes. In Node, `nodeDeflater(level)` gives you zlib and a level — level 1 is roughly 2–3× faster
than the default on xlsx-shaped XML:

```ts
import { createWorkbookWriter, nodeDeflater, toFile } from '@jetstreamapp/simple-excel/node';

const workbook = createWorkbookWriter(toFile('out.xlsx'), { deflater: nodeDeflater(1) });
```

:::caution
A stored entry is streamed like any other, so its sizes follow the data in a data descriptor. Excel, LibreOffice,
openpyxl and calamine read that, but **SheetJS 0.20.3 does not** ("Bad compressed size"). Avoid
`compression: 'store'` when SheetJS reads your files. Every current browser and Node 20.12+ compress, so the automatic
fallback only happens on older platforms.
:::

`createDeflater` is the default `DeflaterFactory`: the platform `CompressionStream('deflate-raw')`, or a
pass-through when the entry is stored or the platform cannot compress. A custom `deflater` has the same shape. The
writer calls it once per zip entry as `(onChunk, { method })`, where `method` is `'deflate' | 'store'`, and it
returns a `Deflater`: `push(chunk)`, `finish()`, `abort(reason?)` and the `bytesIn` / `bytesOut` counters. It must
hand compressed bytes to `onChunk` in order, awaiting each call, and `finish()` must resolve only after the last
`onChunk` has settled.

## Deterministic output

```ts
const workbook = createWorkbookWriter(sink, { deterministic: true });
```

Fixes the zip entry timestamps and the `docProps` dates, so the same rows produce byte-identical output. The
repository's golden-bytes suite depends on this; it is also what you want for content-addressed caching or for
diffing two exports.

:::caution
Determinism holds for one deflate implementation, not across engines. The compressed stream comes from the
platform's `CompressionStream('deflate-raw')`, and each engine tunes zlib differently: the same 20,000-row
workbook is 1,654,492 bytes in Chromium, 1,712,097 in Firefox and 1,686,065 in WebKit
(`npm run smoke:browsers`). The XML inside is identical everywhere — only the compressed container differs — so
compare hashes between runs on the same engine, or write with `compression: 'store'`.
:::

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
- Any failure while writing (a sink that rejects, a compressor that fails, a value the writer refuses mid-row)
  aborts the sink once, and every later call rejects with that same, original error (`EC-WRITER-FAILURE-STICKY`).
  Mistakes that write nothing, such as a refused `registerStyle`, `addSheet` or `merge`, are thrown to the caller
  and leave the workbook usable.

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

With `'truncate'`, the text is cut so that the suffix still fits inside the limit, never between the two halves of
an emoji (`EC-TRUNCATION-SURROGATE`). `onCellTruncated` is called **once**, after `workbook.close()` has finished
the file, with the workbook's total, and only when something was truncated; the same count is in the final result
as `truncatedCells`. An exception thrown from the callback rejects `close()` even though the file is complete. With `'throw'` a long cell raises `CELL_TOO_LONG`.

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
| `strings`          | `'auto' \| 'inline' \| 'shared'`        | `'inline'`             |
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
