# simple-excel

> Streaming, dependency-free xlsx reader and writer for browsers and Node.

**Status: pre-release, under active development.** The API below is the target; `CHANGELOG.md` records what has
landed. The library is built against a committed corpus of real files and a compatibility oracle that includes
Excel itself, so every file it writes opens cleanly in Excel (Windows and Mac), Google Sheets, LibreOffice and
Numbers, and every file those applications write reads back correctly.

## Why another xlsx library

|                                                | simple-excel | SheetJS CE                        | ExcelJS             | @office-kit/xlsx            |
| ---------------------------------------------- | ------------ | --------------------------------- | ------------------- | --------------------------- |
| Streaming write (flat memory at 1M rows)       | yes          | no                                | yes                 | no (buffers until finalize) |
| Streaming read                                 | yes          | no                                | yes                 | yes                         |
| Runtime dependencies                           | 0            | 0                                 | many                | 3                           |
| Maintained on npm                              | yes          | no (fixes on the vendor CDN only) | inactive since 2023 | yes                         |
| Verified against Excel with a public corpus    | yes          | –                                 | –                   | –                           |
| Styles (fonts, fills, borders, number formats) | write        | Pro only                          | yes                 | yes                         |
| Formulas, charts, pivots, editing workbooks    | no           | partial                           | partial             | yes                         |

Pick simple-excel when you export or import tabular data and need it to be large, fast and correct. Pick a
full-featured library when you edit existing workbooks or need charts and formulas. The comparison numbers come from
`bench/results` and the compatibility verdicts from `research/05-compatibility-matrix.md`.

## Install

```sh
npm install @jetstreamapp/simple-excel
```

Evergreen browsers (Web Streams and `CompressionStream`) and Node 20+.

## Write

```ts
import { createWorkbookWriter, collectToBlob } from '@jetstreamapp/simple-excel';

const sink = collectToBlob();
const workbook = createWorkbookWriter(sink);
const sheet = workbook.addSheet('Accounts', { header: ['Id', 'Name', 'Created'], freeze: { rows: 1 } });
for await (const record of records) {
  await sheet.writeRow([record.Id, record.Name, record.CreatedDate]);
}
await sheet.close();
await workbook.close();
const blob = await sink.result();
```

## Read

```ts
import { openWorkbook } from '@jetstreamapp/simple-excel';

const workbook = await openWorkbook(file); // File | Blob | ArrayBuffer | Uint8Array
for (const { name, hidden } of workbook.sheets) {
  console.log(name, hidden);
}
for await (const row of workbook.sheet('Accounts').rows({ mode: 'object' })) {
  // values are string | number | boolean | Date | null
}
```

## Node

```ts
import { fromFile, toFile } from '@jetstreamapp/simple-excel/node';
```

## Documentation

`research/` holds the design: the format primer, the edge-case catalog, the compatibility matrix across eight
readers, the performance baseline and the architecture decisions. `docs/` is the user documentation site.

## License

MIT
