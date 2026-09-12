---
id: streaming-and-memory
title: Streaming and memory
description: How the writer and reader stay flat, how back-pressure works, and which sink to use for a very large file.
---

# Streaming and memory

The design constraint is one sentence: **nothing that grows with the number of rows is ever held in memory.**

That rules out the usual xlsx architecture, where a workbook object is built up, serialized to one giant XML
string, and zipped. That approach fails at a predictable size — `RangeError: Invalid string length` around 512 MiB
of string, `Invalid array length` around 18M cells — and both SheetJS and `@office-kit/xlsx` hit it at roughly the
same point (`research/06-performance-baseline.md`).

## How the writer stays flat

Each `writeRow` appends cell XML to a small string builder. At 64 Ki characters the builder is joined, encoded to
UTF-8, pushed through the deflater and handed to the sink, and the builder is dropped. Bytes are leaving the
process while you are still producing rows.

What is resident at any moment:

| Structure                      | Size                                                                                                                     |
| ------------------------------ | ------------------------------------------------------------------------------------------------------------------------ |
| The row builder                | ≤ 64 Ki characters                                                                                                       |
| The deflater's internal window | a few tens of KB                                                                                                         |
| The zip central directory      | one record per part — tens of parts, a few KB                                                                            |
| The shared-string table        | none by default (`strings: 'inline'`); under `'auto'` bounded by `sstBudget`. See [Writing](./writing.md#shared-strings) |
| The style registry             | one entry per distinct `CellStyle` you registered                                                                        |

None of those is a function of row count. The shared-string table is the only one that could have been, which is
why it is off by default and budgeted when turned on.

## How the reader stays flat

A sheet is a zip entry; the reader opens it as a stream, inflates a chunk, pushes the decoded text through a
tokenizer and yields the rows that completed. The parser keeps one row at a time. Breaking out of the loop
cancels the inflate stream, so `head()` and an early `break` cost only what was already read.

The one structure proportional to the file is the **shared-string table**, which cells index into by number; it
has to be resident while a sheet is being read. It is parsed lazily — only when a sheet actually contains a
`t="s"` cell — and capped by `limits.maxSharedStringChars` (256 Mi UTF-16 units by default). A file written with
inline strings has no table at all and reads with a genuinely constant footprint.

What you do with the rows is your own memory budget. `for await (const row of sheet.rows())` is flat;
`await sheet.toObjects()` is not, because it returns an array of every record.

## Back-pressure

`ByteSink.write` returns a promise, and the writer awaits it. A sink that resolves slowly slows the producer
down instead of letting chunks queue up in memory. That is the whole mechanism, and it is why a sink must not
resolve `write` before it has actually taken the bytes.

```ts
const slowSink: ByteSink = {
  async write(chunk) {
    await sendSomewhere(chunk); // the writer waits here
  },
  async close() {},
  async abort() {},
};
```

The same applies on the way in: `fromWritableStream` awaits the stream writer's `ready` before each write, and
the Node sinks wait for `'drain'`, so a slow disk or a slow socket propagates all the way back to your row loop.

:::caution
A sink takes ownership of every chunk it is handed and the writer never copies. If your sink retains chunks
asynchronously and you are reusing a scratch buffer, copy first.
:::

## Choosing a sink

### Browser: download when finished

```ts
import { collectToBlob, createWorkbookWriter } from '@jetstreamapp/simple-excel';

const sink = collectToBlob();
const workbook = createWorkbookWriter(sink);
// ... write rows ...
await workbook.close();
const blob = await sink.result();
saveAs(blob, 'accounts.xlsx');
```

`collectToBlob` folds pending chunks into a sub-`Blob` every 32 MiB and drops the arrays. Chromium can page Blob
data out to disk, so the JS heap stays flat even for a multi-hundred-megabyte file — the final `Blob` is a list
of parts, not one allocation.

:::caution
Firefox and Safari keep Blob data in memory. `collectToBlob` still keeps the JS _heap_ flat there, but the
process footprint grows with the file. For files that might pass a few hundred megabytes on those browsers,
stream to disk instead. `sink.bytesWritten` lets you warn a user before that point.
:::

### Browser: stream straight to disk

On Chromium, the File System Access API gives you a real `WritableStream` to a file the user picked. Memory stays
constant no matter how large the export gets:

```ts
import { createWorkbookWriter, fromWritableStream } from '@jetstreamapp/simple-excel';

// must be called inside the user gesture, before generation starts
const handle = await window.showSaveFilePicker({ suggestedName: 'accounts.xlsx' });
const sink = fromWritableStream(await handle.createWritable());

const workbook = createWorkbookWriter(sink);
// ... write rows ...
await workbook.close(); // closes the file
```

The same adapter takes any `WritableStream<Uint8Array>` — an OPFS `createWritable()`, a `TransformStream` feeding
an upload, an Electron IPC bridge. `createWorkbookWriter` also accepts a `WritableStream` directly and wraps it
for you.

The other direction exists too: `toWritableStream(sink)` exposes a sink as a `WritableStream` for `pipeTo`
consumers.

### Node: a file or any writable

```ts
import { createWorkbookWriter } from '@jetstreamapp/simple-excel';
import { nodeDeflater, toFile, toWritable } from '@jetstreamapp/simple-excel/node';

const toDisk = createWorkbookWriter(toFile('accounts.xlsx'), { deflater: nodeDeflater(1) });
const toResponse = createWorkbookWriter(toWritable(httpResponse));
```

`toFile` streams through `fs.createWriteStream`; `toWritable` takes any Node `Writable` — a response, a socket, a
`PassThrough`. Both honour back-pressure. See [Node](./node.md).

### In memory

`collectToBytes({ maxBytes })` concatenates once at the end and throws `LIMIT_EXCEEDED` past its cap (1 GiB by
default). It is the right choice for tests, for small files, and for contexts with no `Blob` — but the cap is
there because a contiguous allocation is exactly what streaming exists to avoid. Set it to what you can actually
afford.

## Workers

The core entry touches no DOM API beyond `TextEncoder`, `TextDecoder`, the compression streams and a duck-typed
`Blob`, so it runs unchanged in a Web Worker, an MV3 extension service worker and an Electron renderer. Moving a
large export or import off the main thread is worth it: the row loop is CPU-bound and will otherwise block
rendering for seconds.

:::caution
Instantiate the worker with a literal `new Worker(new URL('./export.worker.ts', import.meta.url), { type: 'module' })`.
Bundlers pattern-match that exact shape, and an MV3 extension's `script-src 'self'` blocks `blob:` workers
outright, so an inlined worker will not run there (ADR-004).
:::

A `Blob` can be transferred back to the main thread by reference, so `collectToBlob` inside a worker costs no copy
on the way out.

## Reading a very large file

- Iterate with `for await`; do not call `toObjects()` on a file you have not sized.
- Use `head(n)` to inspect headers or template metadata before committing to a full read.
- `break` out of the loop as soon as you know enough — the rest of the sheet is never inflated.
- Lower `limits` for untrusted uploads; raise `maxInflatedBytes` deliberately for a large file you trust.
- Prefer a `Blob`/`File` or a `RandomAccessSource` over materializing an `ArrayBuffer`: the reader only pulls the
  ranges it needs.

## Writing a very large file

- Announce `rowCount` on `addSheet` when you know it. It enables `<dimension>` and lets the zip64 decision be
  made up front.
- Use `onProgress` (every 5,000 rows) for the UI, and an `AbortSignal` so a user can cancel.
- Leave `strings` on its inline default unless the data is very low-cardinality and file size matters more than
  write speed; then `strings: 'auto'` builds a bounded table. Excel re-shares strings the first time the user
  saves anyway.
- Use `nodeDeflater(1)` in Node when throughput matters more than a few percent of file size.

## The zip64 caveat

An xlsx part over 4 GiB needs zip64 headers, and Excel only accepts them when they were declared in the _local
file header_, before the entry's first byte. So the decision is made up front, from the `rowCount` you declared
(ADR-002).

:::caution
**SheetJS 0.20.3 cannot open a zip64 archive at all** — "Unsupported ZIP file", even for a small one — and
**Google Drive fails to convert one to a Google Sheet** on upload (`EC-ZIP64-SMALL`). Excel, LibreOffice, openpyxl
and calamine all read them. That is why `zip64: 'auto'` only turns it on for a sheet whose declared size makes a

> 4 GiB part plausible (about 40M cells; Google Sheets tops out at 10M cells, so such a file could not be imported
> there regardless). If any consumer of your files is SheetJS or Google Sheets, leave it on `'auto'` and do not set
> `zip64: true`.
> :::

A sheet written with an unknown row count stays 32-bit, and a part that does pass 4 GiB fails with
`ENTRY_TOO_LARGE` naming `zip64: true` as the fix. Excel's own 1,048,576-row limit means such a sheet would need
more than 4 KB of XML per row, so this is rare in practice.

## Numbers

The claim on this page is measurable, so here is the measurement. Apple M4, Node v24.18.0, 20 columns of
Salesforce-shaped data (about 1.7 KB of text per row), rows pulled from a generator so nothing but the writer is
resident, peak sampled every 25 ms. Run
[`2026-09-12-macbook-air-phase-e-combined`](https://github.com/jetstreamapp/simple-excel/tree/main/bench/results/2026-09-12-macbook-air-phase-e-combined),
with the streamed-source timings taken one size per process.

| Rows written        | Time   | Output | JS heap growth | Peak process RSS |
| ------------------- | ------ | ------ | -------------- | ---------------- |
| 100,000 (2M cells)  | 3.8 s  | 42 MB  | 45 MB          | 145 MB           |
| 900,000 (18M cells) | 35.4 s | 384 MB | 76 MB          | 215 MB           |
| 1,000,000           | 39.1 s | 427 MB | 77 MB          | 214 MB           |

Ten times the rows, the same heap. That is the whole design in one table. Reading is the same shape: streaming
the 1,000,000-row file back with `sheet.rows()` takes 10.3 s and grows the heap by 64 MB.

For contrast, on the same machine SheetJS writes 100k × 20 in 9.8 s with a 3,313 MB peak, and throws
`RangeError: Invalid string length` at both 1M × 20 and 18M cells. In a Chrome module worker writing to
`collectToBlob()`, simple-excel finishes 1,000,000 × 20 in 39.6 s for +254 MB of renderer RSS; SheetJS crashes
the renderer.

The gates all of this is measured against (`research/06-performance-baseline.md`) are, for 100k × 20 versus
SheetJS 0.20.3 with the same options: write time ≤ 1.0× (measured 0.28×), write peak memory ≤ 0.25× (0.03×),
read time ≤ 1.0× (0.26×), read peak memory ≤ 0.5× (0.48×), first byte at the sink under 100 ms (0.4 ms), and
1M × 20 completing in both Node and a Chrome module worker (it does). The absolute wall-clock targets in the
build plan — 100k × 20 written in 1.2 s, read in 1.5 s — are still **not** met on this dataset (2.77 s and
1.62 s); 06 records the shortfall and where the time goes.

:::caution
Reading back a file this large needs the inflate cap raised: 1,000,000 Salesforce-shaped rows inflate to 2.19 GB
of sheet XML, past the 1 GiB default, and `openWorkbook` refuses it with `ZIP_BOMB`. Pass
`limits: { maxInflatedBytes: 4 * 1024 * 1024 * 1024 }` for a file that size when you trust it.
:::
