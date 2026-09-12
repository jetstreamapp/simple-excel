---
id: install
title: Install
description: Package, entry points, module formats and the platform features simple-excel needs.
---

# Install

```bash
npm install @jetstreamapp/simple-excel
```

## Entry points

| Import                            | Contents                                                                                                        |
| --------------------------------- | --------------------------------------------------------------------------------------------------------------- |
| `@jetstreamapp/simple-excel`      | Everything: `createWorkbookWriter`, `openWorkbook`, the sinks, `sniff`, `sourceFrom`, `XlsxError` and the types |
| `@jetstreamapp/simple-excel/node` | The same surface re-exported, plus `fromFile`, `toFile`, `toWritable` and `nodeDeflater`. See [Node](./node.md) |

Both entries ship ESM and CommonJS builds with TypeScript declarations. The core entry is browser-safe: it has no
`node:` imports and no DOM references beyond `TextEncoder`, `TextDecoder`, `CompressionStream`,
`DecompressionStream` and a duck-typed `Blob`, so it works unchanged in a Web Worker, an MV3 extension service
worker and an Electron renderer.

```ts
import { collectToBytes, createWorkbookWriter, isXlsxError, openWorkbook, sniff, sourceFrom } from '@jetstreamapp/simple-excel';
import type { CellInput, CellStyle, OpenOptions, Workbook, WorkbookWriterOptions } from '@jetstreamapp/simple-excel';
```

## Environment requirements

| Feature                              | Used for                              | Minimum                                           |
| ------------------------------------ | ------------------------------------- | ------------------------------------------------- |
| Web Streams (`WritableStream`)       | Stream sinks and the compression pump | Chrome 103+, Firefox 113+, Safari 16.4+, Node 20+ |
| `CompressionStream('deflate-raw')`   | Compressing zip entries on write      | Chrome 103+, Firefox 113+, Safari 16.4+, Node 20+ |
| `DecompressionStream('deflate-raw')` | Inflating zip entries on read         | same                                              |
| `TextEncoder` / `TextDecoder`        | UTF-8 on both sides                   | universal on the above                            |
| `Blob`                               | `collectToBlob()` only                | optional; `collectToBytes()` needs no `Blob`      |

`package.json` declares `engines.node >= 20`.

:::caution
Node has had `CompressionStream` since 18, but it only accepts the `deflate-raw` format from **21.2**. On an
earlier Node the writer falls back to stored (uncompressed) parts; pass `nodeDeflater()` from the `/node` entry to
keep files compressed — it uses zlib and works on every supported version.
It is a good default for server-side writes anyway, because it lets you pick a compression level.
:::

### When `CompressionStream` is missing

`hasNativeDeflate()` tells you whether the platform has it at all. If it does not, the writer stores every part
uncompressed (zip method 0) instead of failing. The file is a valid `.xlsx` that every reader opens; it is just
several times larger. In Node you can always get real compression by passing the zlib-backed deflater:

```ts
import { createWorkbookWriter, hasNativeDeflate } from '@jetstreamapp/simple-excel';
import { nodeDeflater, toFile } from '@jetstreamapp/simple-excel/node';

const workbook = createWorkbookWriter(toFile('out.xlsx'), {
  deflater: hasNativeDeflate() ? undefined : nodeDeflater(),
});
```

`collectToBlob()` throws `UNSUPPORTED_ENVIRONMENT` where there is no `Blob`; use `collectToBytes()` or a stream
sink there.

## Verifying the install

```ts
import { collectToBytes, createWorkbookWriter } from '@jetstreamapp/simple-excel';

const sink = collectToBytes();
const workbook = createWorkbookWriter(sink);
const sheet = workbook.addSheet('Sheet1', { header: ['A', 'B'] });
await sheet.writeRow([1, 2]);
await sheet.close();
await workbook.close();
const bytes = sink.result(); // a Uint8Array holding a complete .xlsx
```

`sink.result()` on `collectToBytes` is synchronous and only valid after `close()` has resolved; on
`collectToBlob` it returns a promise.
