---
id: intro
title: simple-excel
sidebar_label: Introduction
description: A streaming, dependency-free xlsx reader and writer for browsers and Node.
---

# simple-excel

`@jetstreamapp/simple-excel` reads and writes `.xlsx` files in browsers and in Node, with no runtime dependencies.

It writes a workbook by pushing rows into a sink as they arrive and reads one by pulling rows through an async
iterator. Nothing accumulates: there is no in-memory worksheet, no whole-sheet XML string and no unbounded shared
string table. A one-million-row export costs about as much memory as a one-thousand-row export.

:::info
The library is pre-release and the API may still shift before 1.0. `CHANGELOG.md` records what has landed.
:::

## What it is for

Exporting and importing tabular data — query results, record loads, report extracts — in files that have to open
cleanly in Excel and come back out carrying the same values.

```ts
import { collectToBlob, createWorkbookWriter } from '@jetstreamapp/simple-excel';

const sink = collectToBlob();
const workbook = createWorkbookWriter(sink);
const sheet = workbook.addSheet('Accounts', { header: ['Id', 'Name', 'Created'], freeze: { rows: 1 } });

await sheet.writeRow(['001xx000003DGb2AAG', 'Acme', new Date(2024, 2, 10)]);

await sheet.close();
await workbook.close();
const blob = await sink.result();
```

```ts
import { openWorkbook } from '@jetstreamapp/simple-excel';

const workbook = await openWorkbook(file);
for await (const row of workbook.sheet(0).rows({ mode: 'object' })) {
  console.log(row.Id, row.Name);
}
await workbook.close();
```

## When to pick it

Pick simple-excel when:

- the file is large, or you cannot predict how large it will be;
- it has to open in Excel without a repair prompt, and in Google Sheets, LibreOffice and Numbers too;
- you are running in a browser tab, a Web Worker or an extension service worker, where a multi-gigabyte heap spike
  ends the session;
- you want to know exactly what happens to a date, a 40,000-character cell or a control character, because the
  answer is written down;
- adding a dependency tree to a code path that touches untrusted uploads is not acceptable.

Pick something else when:

- you need to edit an existing workbook and preserve the parts you did not touch;
- you need formulas, charts, pivot tables, conditional formatting or data validation;
- you need to read `.xls`, `.xlsb` or `.ods` (simple-excel detects each and says what it is, but does not parse it);
- you need Excel's number formats rendered to display text.

[Comparison](./comparison.md) has the library-by-library reasoning.

## How the pieces fit

| Concept              | What it is                                                                                            |
| -------------------- | ----------------------------------------------------------------------------------------------------- |
| `ByteSink`           | Where written bytes go: a Blob, a byte array, a `WritableStream`, a file. See [Writing](./writing.md) |
| `WorkbookWriter`     | Owns the zip, the styles and the shared strings; hands out one `SheetWriter` at a time                |
| `SourceInput`        | Where read bytes come from: `ArrayBuffer`, `Uint8Array`, `Blob`/`File`, or a `RandomAccessSource`     |
| `Workbook` / `Sheet` | Lists sheets without touching sheet XML; streams rows on demand. See [Reading](./reading.md)          |
| `XlsxError`          | Every failure, carrying a stable `code`. See [Errors](./errors.md)                                    |

## Where to go next

- [Install](./install.md) — package, entry points, environment requirements
- [Writing](./writing.md) — sinks, sheets, rows, styles, options
- [Reading](./reading.md) — opening, listing sheets, array and object rows
- [Dates and values](./dates-and-values.md) — the value model and its edge cases
- [Errors](./errors.md) — every error code and when it fires
- [Streaming and memory](./streaming-and-memory.md) — how both sides stay flat, and what to do with very large files
- [Compatibility](./compatibility.md) — what was verified, against what
- [Node](./node.md) — the `/node` entry
- [Comparison](./comparison.md) — the longer version of the README table
- [Contributing](./contributing.md) — repo layout, commands, fixtures
