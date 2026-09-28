---
id: node
title: Node
description: The /node entry — fromFile, toFile, toWritable and nodeDeflater.
---

# Node

Everything in the core entry works in Node. The `/node` entry adds the things that need `node:` modules: a
file-backed source, two file-backed sinks, and a zlib deflater with a compression level.

```ts
import { createWorkbookWriter, fromFile, nodeDeflater, openWorkbook, toFile, toWritable } from '@jetstreamapp/simple-excel/node';
```

The `/node` entry re-exports the whole core surface, so you never need both imports.

## `fromFile(path)`

```ts
function fromFile(path: string): Promise<FileSource>;
```

A `RandomAccessSource` over a file, using positional reads on a `FileHandle`. The reader pulls only the ranges it
needs — the central directory, the package parts, then each sheet as it is iterated — so opening a 2 GB workbook
does not read 2 GB.

```ts
const source = await fromFile('/tmp/upload.xlsx');
const workbook = await openWorkbook(source);

for await (const row of workbook.sheet(0).rows({ mode: 'object' })) {
  await handle(row);
}

await workbook.close(); // closes the file handle
```

`workbook.close()` closes the source. If you never open a workbook over it, call `source.close()` yourself.

## `toFile(path)`

```ts
function toFile(path: string): ByteSink;
```

Streams to a file through `fs.createWriteStream`, honouring back-pressure — a slow disk slows the row loop rather
than filling memory. `workbook.close()` ends the stream and resolves once the file is flushed.

```ts
const workbook = createWorkbookWriter(toFile('accounts.xlsx'));
const sheet = workbook.addSheet('Accounts', { header: ['Id', 'Name'] });
await sheet.writeRows(rows);
await sheet.close();
await workbook.close();
```

## `toWritable(writable)`

```ts
function toWritable(writable: NodeWritableLike): ByteSink;
```

The same sink over any Node `Writable`: an HTTP response, a socket, a `PassThrough`, a `gzip` stream. The
parameter is structurally typed (`write`, `end`, `destroy`, `once`, `off`), so the core types stay free of Node
types and you can pass a test double.

```ts
app.get('/export.xlsx', async (request, response) => {
  response.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  response.setHeader('Content-Disposition', 'attachment; filename="export.xlsx"');

  const workbook = createWorkbookWriter(toWritable(response), { deflater: nodeDeflater(1) });
  const sheet = workbook.addSheet('Export', { header });
  try {
    for await (const record of query()) {
      await sheet.writeRow(toRow(record));
    }
    await sheet.close();
    await workbook.close();
  } catch (error) {
    await workbook.abort(error);
    throw error;
  }
});
```

Streaming a response like this means the browser starts downloading immediately and the server never holds the
file. A stream error is sticky: the sink remembers it and every later call rejects with the original reason, so
you find out at the first `writeRow` rather than at the end.

## `nodeDeflater(level?)`

```ts
function nodeDeflater(level?: number): DeflaterFactory;
```

A deflater over `zlib.createDeflateRaw`, so you get a compression level — which the platform `CompressionStream`
does not expose (ADR-005). Levels are zlib's 1–9.

```ts
const fast = createWorkbookWriter(toFile('big.xlsx'), { deflater: nodeDeflater(1) });
const small = createWorkbookWriter(toFile('small.xlsx'), { deflater: nodeDeflater(9) });
```

Level 1 is roughly 2–3× faster than the default on xlsx-shaped XML, which is highly repetitive and compresses
well even at the low setting. For a large export that is usually the right trade.

`nodeDeflater` also honours `compression: 'store'`: with that option the factory returns a pass-through and the
level is ignored.

## When you need it

The core entry works in Node without this one — `collectToBytes()` and `fromWritableStream(Writable.toWeb(...))`
both do the job. Reach for `/node` when you want:

| Want                                          | Use                   |
| --------------------------------------------- | --------------------- |
| To read a file without loading it into memory | `fromFile`            |
| To write a file without buffering it          | `toFile`              |
| To stream into a response or socket           | `toWritable`          |
| A compression level                           | `nodeDeflater(level)` |

## Node version

`package.json` declares `engines.node >= 20.12`, because Node's `CompressionStream` and `DecompressionStream` only
accept the `deflate-raw` format from 20.12 (21.2 on the 21 line). On an older Node the writer detects that and
falls back to stored (uncompressed) parts, and `openWorkbook` fails with `UNSUPPORTED_ENVIRONMENT` because it cannot
inflate entries. `nodeDeflater()`, which goes through zlib, keeps written files compressed anywhere, and it is the
right default for server-side writes regardless, since it is the only way to choose a compression level.
