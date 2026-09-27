---
id: reading
title: Reading
description: openWorkbook, sources, sheet listing, array and object rows, head() and every read option.
---

# Reading

```ts
import { openWorkbook } from '@jetstreamapp/simple-excel';

const workbook = await openWorkbook(file);

for (const { name, index, kind, hidden } of workbook.sheets) {
  console.log(index, name, kind, hidden);
}

for await (const row of workbook.sheet('Accounts').rows({ mode: 'object' })) {
  console.log(row.Id, row.Name);
}

await workbook.close();
```

Opening a workbook reads the zip central directory and the small package parts (content types, relationships,
`workbook.xml`, `styles.xml`). It does **not** touch sheet XML, so listing sheets is cheap on a 200 MB file. The
shared-string table is parsed lazily — on the first `t="s"` cell of the first sheet that has one — and sheet
contents inflate only while you iterate. Dropping out of the loop cancels the underlying stream.

## `openWorkbook(input, options?)`

```ts
function openWorkbook(input: SourceInput, options?: OpenOptions): Promise<Workbook>;
function openWorkbook(input: SourceInput, options: OpenOptions & { errors: 'object' }): Promise<Workbook<CellValue | CellError>>;
```

Before anything else the input is sniffed, so a file that is not a modern xlsx fails with a classified error
naming what it actually is — `ENCRYPTED`, `LEGACY_XLS`, `XLSB`, `ODS` or `NOT_XLSX` — instead of a zip or XML
parse error three layers down. A `.docx` and a `.pptx` are named too. See [Errors](./errors.md).

### Sources

`SourceInput` is `ArrayBuffer | SharedArrayBuffer | ArrayBufferView | Uint8Array | Blob | RandomAccessSource`.

| Input                           | How it is read                                                                          |
| ------------------------------- | --------------------------------------------------------------------------------------- |
| `Uint8Array` or any typed array | Zero-copy subarray views over the same bytes; nothing is duplicated (`Buffer` included) |
| `DataView`                      | Same, over the bytes it views                                                           |
| `ArrayBuffer`                   | Same, over a view of the buffer                                                         |
| `SharedArrayBuffer`             | Each range is copied as it is read, because the decompressor refuses shared memory      |
| `Blob` / `File`                 | `slice().arrayBuffer()` per range, so the bytes never all have to be resident           |
| `RandomAccessSource`            | Anything with `{ size, read(offset, length) }` — including Node's `fromFile()`          |

The checks work across realms, so a buffer from an iframe, a worker or an Electron bridge is accepted. `Blob` is
duck-typed: any object with `size`, `slice` and `arrayBuffer` works, which matters in test environments and in
Electron. Anything else, a string included, is refused with `NOT_XLSX` (`EC-INPUT-TYPE`); a binary string from an
old API has to become bytes first.

`sourceFrom(input)` is exported if you want the `RandomAccessSource` yourself, and `RandomAccessSource` is the
extension point for a source that is not bytes in hand:

```ts
import { sourceFrom } from '@jetstreamapp/simple-excel';
import { fromFile } from '@jetstreamapp/simple-excel/node';

const fromBytes = sourceFrom(uint8Array);
const fromDisk = await fromFile('/tmp/upload.xlsx'); // already a RandomAccessSource
```

:::caution
The reader needs random access, because a zip's directory is at the end of the file. A pipe or an HTTP body
cannot be read incrementally — buffer it, save it to disk, or expose it as a `RandomAccessSource` backed by
ranged requests.
:::

## Options

```ts
const workbook = await openWorkbook(file, {
  dates: 'local',
  errors: 'string',
  limits: { maxInflatedBytes: 256 * 1024 * 1024 },
  signal: controller.signal,
  onWarning: warning => console.warn(warning.code, warning.sheet, warning.ref),
});
```

| Option      | Values                           | Default    | Meaning                                                                                                                   |
| ----------- | -------------------------------- | ---------- | ------------------------------------------------------------------------------------------------------------------------- |
| `dates`     | `'local' \| 'utc' \| 'serial'`   | `'local'`  | Which `Date` fields carry the wall clock, or `'serial'` for the raw number. See [Dates and values](./dates-and-values.md) |
| `errors`    | `'string' \| 'object' \| 'null'` | `'string'` | How an error cell comes back: its text (`'#N/A'`), a `{ error }` object, or `null`                                        |
| `limits`    | `ReadLimits`                     | see below  | Caps that make a hostile file fail fast instead of exhausting the host                                                    |
| `signal`    | `AbortSignal`                    | —          | Checked at open and between chunks; an aborted signal ends the read with `ABORTED`                                        |
| `onWarning` | `(warning: ReadWarning) => void` | —          | Told about problems the reader repaired instead of rejecting; see below                                                   |

`errors: 'object'` also changes the static type: the workbook becomes `Workbook<CellValue | CellError>` and row
values widen to include `{ error: '#N/A' }`.

### Warnings

Some damage Excel shrugs off, and so does the reader: the cell reads blank and the read carries on. `onWarning`
lets you notice. Each `ReadWarning` has a `code`, a `message`, the `sheet` and the `ref` of the first cell that hit
it, and each code is reported at most once per sheet however many cells hit it.

| `code`                   | What happened                                                                        |
| ------------------------ | ------------------------------------------------------------------------------------ |
| `SST_INDEX_OUT_OF_RANGE` | A shared-string cell points past the end of the table; it reads as `''`, as in Excel |
| `SHARED_STRINGS_MISSING` | A shared-string cell in a workbook that has no shared-string part; it reads as `''`  |

A sheet full of blank text cells with one of these warnings usually means a damaged file, not an empty one.

### Limits

| `ReadLimits`           | Default             | Guards                                               |
| ---------------------- | ------------------- | ---------------------------------------------------- |
| `maxInflatedBytes`     | 1 GiB per part      | Zip bombs; exceeded raises `ZIP_BOMB`                |
| `maxSharedStringChars` | 256 Mi UTF-16 units | A shared-string table designed to exhaust memory     |
| `maxEntries`           | 10,000              | Archives with an absurd number of parts              |
| `maxXmlDepth`          | 256                 | Deeply nested rich-text runs (`EC-XML-DEEP-NESTING`) |
| `maxTextLength`        | 64 Mi UTF-16 units  | One text node swallowing the heap                    |

Raise them for a legitimately huge file you trust; lower them for user uploads. Exceeding a limit raises
`LIMIT_EXCEEDED` (or `ZIP_BOMB` for the inflate caps), and the message names what was hit.

## Listing sheets

```ts
interface SheetInfo {
  readonly name: string;
  readonly index: number; // position in workbook order, from 0
  readonly kind: 'worksheet' | 'chartsheet';
  readonly hidden: boolean;
}
```

`workbook.sheets` is the workbook's own order. `workbook.date1904` says which date system the file uses (see
[Dates and values](./dates-and-values.md#the-1904-date-system)).

**Hidden sheets** are listed, not filtered — `hidden` is `true` for both `state="hidden"` and
`state="veryHidden"`. Deciding whether to show them is the application's call. Note that hiding is not a privacy
feature: Numbers makes hidden sheets visible on import and loses the state on export
(`EC-NUMBERS-HIDDEN-SHEETS-SHOWN`).

**Chartsheets** appear in the list with `kind: 'chartsheet'`. They hold a chart, not a grid, so iterating one
yields no rows. Filter them out when you are looking for data:

```ts
const dataSheets = workbook.sheets.filter(({ kind, hidden }) => kind === 'worksheet' && !hidden);
```

`workbook.sheet(nameOrIndex)` takes a name or a 0-based index. Names are matched exactly first and then
case-insensitively, because Excel itself compares sheet names without case. Anything unmatched raises
`SHEET_NOT_FOUND`, and the message lists the names the workbook does have.

## Rows

### Array mode

```ts
for await (const cells of sheet.rows()) {
  // cells: (string | number | boolean | Date | null)[]
}
```

Each row is a dense array, 0-based, with trailing empties trimmed and interior holes as `null`. Rows with no
cells are skipped unless `blankRows: true`, which yields them as `[]` so that positions still line up.

| `RowsOptions` | Type                | Default   | Meaning                                                                               |
| ------------- | ------------------- | --------- | ------------------------------------------------------------------------------------- |
| `startRow`    | number              | `1`       | 1-based first row to yield; earlier rows stream past without being materialized       |
| `maxRows`     | number              | —         | Stop after this many rows                                                             |
| `blankRows`   | boolean             | `false`   | Yield rows with no cells (array mode only)                                            |
| `maxColumns`  | number              | `16384`   | A cell past this column raises `LIMIT_EXCEEDED` rather than being dropped silently    |
| `formulas`    | `'value' \| 'text'` | `'value'` | `'value'` reads the cached result; `'text'` yields the formula text for formula cells |

### Object mode

```ts
for await (const row of sheet.rows({ mode: 'object' })) {
  // row: Record<string, string | number | boolean | Date | null>
}
```

Object mode takes one row as headers and keys every later row by it.

| `ObjectRowsOptions` | Type                   | Default                | Meaning                                                   |
| ------------------- | ---------------------- | ---------------------- | --------------------------------------------------------- |
| `headerRow`         | number                 | first row with a value | 1-based row holding the headers                           |
| `defval`            | `CellValue`            | `''`                   | Value for cells absent from a row                         |
| `headerNaming`      | `'sheetjs' \| 'index'` | `'sheetjs'`            | How headers become keys                                   |
| `dropEmptyHeaders`  | boolean                | `false`                | Drop columns whose header is empty instead of naming them |

Without `headerRow`, the header row is the first row at or after `startRow` (default 1) that holds any value, an
empty string or an error included. A sheet whose first row or two are blank therefore still finds its headers,
which is what `sheet_to_json` does when the sheet's range starts below row 1 (`EC-HEADER-ROW-FIRST-NONEMPTY`). It
does better than `sheet_to_json` when row 1 is formatted but empty: SheetJS then names every column `__EMPTY` and
reads the real header row as data. Pass `headerRow` to pin the row; an explicit `headerRow` on a blank row gives
`__EMPTY` columns, as before.

`startRow`, `maxRows`, `maxColumns` and `formulas` apply as well. The header row is consumed, not yielded: the
first object is the row after the header row, or `startRow` if that is later. `blankRows` does not apply — object
mode never yields empty records, which matches SheetJS's `blankrows: false`.

#### Header naming

`headerNaming: 'sheetjs'` reproduces SheetJS's `sheet_to_json` keys exactly, so a migration does not rename every
field in an application:

- A header cell with text becomes that text. Non-string headers are spelled the way Excel spells them (`TRUE` /
  `FALSE`, an ISO-style timestamp for a date, the code for an error cell such as `#N/A`, whatever `errors` says,
  `EC-HEADER-ERROR-CELL`). Numbers keep their JavaScript spelling (`123456789012`), where SheetJS shows Excel's
  General format (`1.23457E+11`).
- Every header becomes an own property of the record, `__proto__` and `constructor` included; the record's
  prototype never changes (`EC-HEADER-PROTO-KEY`).
- An **absent** header cell becomes `__EMPTY`, and each later absent header gets a counter: `__EMPTY_1`,
  `__EMPTY_2`, … (the first has no suffix). A header cell that holds an _empty string_ is a different thing: its
  key is the empty string.
- A **duplicate** header keeps the first occurrence as-is and suffixes the rest: `Name`, `Name_1`, `Name_2`, …
  skipping any suffix that is already taken.

`headerNaming: 'index'` ignores the header text and keys every column by its 0-based column index as a string
(`'0'`, `'1'`, `'2'`). Use it when the file has no usable header row but you still want objects.

`dropEmptyHeaders: true` drops every column whose header is empty — absent or empty text — instead of naming it.
That is usually what you want for a report export with trailing blank columns.

:::note
The column list comes from the widest of the header row, the widest data row and the sheet's declared
`<dimension>`. Excel, Google Sheets and Numbers all re-save with a dimension wider than the data, so a file from
one of them can grow trailing `__EMPTY` columns. This is deliberate: it is what `sheet_to_json` does, and object
mode is built to match it. The first column is the dimension's first column, so an all-blank column A that Excel
left out of the range does not become `__EMPTY` either (`EC-HEADER-DIMENSION-START-COLUMN`); a value that turns up
left of it is still kept, under a blank-header name.

`toObjects()` gives every record every key: a column that only appears in a later, wider row is added to the
earlier records with `defval`. Streaming `rows({ mode: 'object' })` cannot go back, so there a record only has the
keys named by the time it was read.
:::

#### `defval` and blank cells

A cell that is missing from the row and a cell holding an empty string are different things in the file, and
readers disagree about which they return (`EC-BLANK-VS-EMPTY-STRING`). Object mode settles it with `defval`:
every named column appears on every record, and a cell with no value gets `defval`, which defaults to `''`. Pass
`defval: null` if you need to tell "absent" from "empty text".

### `toObjects`

When the sheet is small enough to hold in memory, `toObjects` collects it in one call and hands back the headers
it used:

```ts
const { rows, headers, truncated } = await sheet.toObjects({ maxRows: 10_000 });
```

`truncated` is `true` when `maxRows` stopped the read before the sheet ended — which is how you tell "the file
has 10,000 rows" from "the file has more and you only asked for 10,000".

## `head()`: reading a prefix

```ts
const cells = await sheet.head(5); // Map<string, RawCell>, keyed 'A1', 'B1', …
const objectName = cells.get('B1')?.value;
const operation = cells.get('B2')?.value;
```

`head(rowCount)` returns the first rows as an A1-keyed map and then **stops inflating** — a preview of a
million-row sheet reads only its first chunks. It is the right tool for template headers, for sniffing which
columns a file has before committing to a full read, and for files whose interesting metadata sits in a few
scattered cells above the data.

```ts
interface RawCell {
  readonly value: CellValue;
  readonly error?: CellErrorCode; // present when the cell held an error; `value` is then the error text
  readonly formula?: string; // formula text, when the cell has one
}
```

Blank cells with no formula are simply absent from the map. `head()` always reports formula text and always gives
an error cell both its text and its code, whatever `errors` mode the workbook was opened with.

## Formulas

Formulas are read-only. A formula cell carries a cached result in the file, and that is what you get by default:

```ts
for await (const cells of sheet.rows()) {
  // the cached value of =SUM(A1:A9), not the text
}

for await (const cells of sheet.rows({ formulas: 'text' })) {
  // 'SUM(A1:A9)' for formula cells, values for the rest
}
```

Nothing is evaluated, and formulas are never written. Two caveats from the corpus: some writers never record a
cached value at all (openpyxl), and some applications replace formulas they cannot express with their cached
value on import (Numbers, `EC-NUMBERS-FORMULAS-REPLACED`).

## Cancelling a read

Breaking out of the loop cancels the inflate stream, so an early `break` costs only what was already read:

```ts
for await (const cells of sheet.rows()) {
  if (looksWrong(cells)) {
    break; // the rest of the sheet is never inflated
  }
}
```

For a read you need to cancel from elsewhere, pass a `signal` to `openWorkbook`. It is checked between chunks and
ends the iteration with `ABORTED`.

Call `workbook.close()` when you are done; it releases the source's handles (for a file-backed source, the file
descriptor). Reading any sheet of a closed workbook fails with `ABORTED`, whatever the source.

## Sniffing before you open

`sniff(head)` classifies the first bytes of an input without parsing it. Use it when you accept arbitrary uploads
and want to route CSV to your own parser rather than show an error:

```ts
import { SNIFF_BYTES, sniff } from '@jetstreamapp/simple-excel';

const head = new Uint8Array(await file.slice(0, SNIFF_BYTES).arrayBuffer());
switch (sniff(head)) {
  case 'zip':
    return openWorkbook(file);
  case 'text':
    return parseCsv(file);
  default:
    return showUnsupported();
}
```

`'text'` covers CSV in any common encoding: UTF-8, UTF-16 with a byte-order mark, and legacy single- and multi-byte
encodings such as Windows-1252 (what Excel for Windows saves CSV as) or Shift-JIS, recognised by the absence of NUL
and control bytes (`EC-SNIFF-LEGACY-TEXT`). Decoding it is up to your CSV path: try UTF-8 with
`new TextDecoder('utf-8', { fatal: true })` and fall back to `'windows-1252'`.

It returns `'zip' | 'cfb-encrypted' | 'cfb-legacy' | 'xml' | 'html' | 'text' | 'empty' | 'unknown'`. The
`SniffResult` type also lists `'xlsx' | 'xlsb' | 'ods'`, which `sniff` itself never returns: all three are zips,
and only the package contents tell them apart, so `openWorkbook` makes that distinction after opening the
archive. Pass at least `SNIFF_BYTES` (65,536 bytes): the encrypted-versus-legacy decision scans all of it for the
`EncryptedPackage` stream name.

## What the reader tolerates

Real files are not tidy. The reader deliberately accepts things the schema does not, because Excel does:

- `<row>` and `<c>` without `r` attributes — positions are inferred (`EC-MISSING-R-ATTRIBUTES`).
- A namespace prefix on every element, and Strict OOXML namespaces and relationship types
  (`EC-PREFIXED-ELEMENTS`, `EC-STRICT-NAMESPACES`).
- Relationship targets that are absolute (`/xl/worksheets/sheet1.xml`) or use backslashes
  (`EC-ABSOLUTE-REL-TARGETS`, `EC-BACKSLASH-REL-TARGETS`).
- A missing `<dimension>`, or one that disagrees with the cells — it is a hint, never the truth (`EC-NO-DIMENSION`).
- Part names whose case differs from the relationship that points at them (`xl/SharedStrings.xml`): OPC names are
  case-insensitive, so the exact name wins and a case-insensitive match is the fallback (`EC-PART-NAME-CASE`).
- A numeric `<v>` of only whitespace reads as blank (`EC-EMPTY-V-ELEMENT`); numeric text Excel never writes, such
  as `0x1A`, reads as its text rather than as a number (`EC-NUMBER-RADIX-PREFIX`).
- A workbook with no `xl/sharedStrings.xml` at all, every string inline (`EC-SST-ABSENT-INLINE-ONLY`).
- Extra parts with no extension, legacy VML comment parts, `[Content_Types].xml` as the last entry rather than
  the first (`EC-PART-NONSTANDARD-NAMES`, `EC-ZIP-CONTENT-TYPES-LAST`).
- Apache POI's spellings: `customWidth="true"` instead of `1`, six-digit colours, `<u val="none"/>` — which is
  what Salesforce report exports contain (`EC-POI-BOOLEAN-ATTRIBUTE-SPELLING`, `EC-POI-RGB-6-HEX`).

What it does not tolerate is on the [Errors](./errors.md) page: a DOCTYPE, duplicate entries, a CRC mismatch, an
unsupported compression method, a cell reference past column XFD (`EC-REF-BEYOND-XFD`), and anything past the
limits.
