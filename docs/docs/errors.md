---
id: errors
title: Errors
description: Every XlsxErrorCode, when it fires, and what the message is written to tell the user.
---

# Errors

Every failure the library raises is an `XlsxError` with a stable `code`. Branch on the code; show the message, or
your own copy.

```ts
import { isXlsxError } from '@jetstreamapp/simple-excel';

try {
  const workbook = await openWorkbook(file);
} catch (error) {
  if (isXlsxError(error)) {
    switch (error.code) {
      case 'ENCRYPTED':
        return showPasswordHelp(error.message);
      case 'LEGACY_XLS':
      case 'XLSB':
      case 'ODS':
        return showConvertHelp(error.message);
      default:
        return showGenericFailure(error.message);
    }
  }
  throw error;
}
```

```ts
class XlsxError extends Error {
  readonly code: XlsxErrorCode;
  readonly detail: Readonly<Record<string, unknown>> | undefined;
}

function isXlsxError(value: unknown): value is XlsxError;
```

`detail` carries whatever is useful for that code — the part name, the limit that was hit, the sniffed format,
the sheet names that do exist. It is for logs and for your own messaging, not a stable API.

`isXlsxError` is a duck-typed check (`instanceof`, or `name === 'XlsxError'`), so it still works across realms —
an error that crossed a worker boundary or came out of a bundled copy of the library.

## Messages are for end users

The messages name the fix, in the words of somebody who has a spreadsheet open, not a stack trace:

> This workbook is password-protected. Remove the password in Excel (File > Info > Protect Workbook) and save it
> again as .xlsx.

> This is an Excel Binary Workbook (.xlsb). Open it in Excel and save it as .xlsx.

> A cell holds 40000 characters, more than the 32767 Excel allows. Shorten it or write with
> `cellOverflow: 'truncate'`.

You can show them as they are. If you localize, branch on `code`.

## Input format errors

These fire when the bytes are not a modern xlsx. The format is sniffed before any parsing, so the user is told
what the file actually is instead of seeing a zip or XML failure.

| Code         | Fires when                                                                                                                                                                         | Message intent                                                 |
| ------------ | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | -------------------------------------------------------------- |
| `ENCRYPTED`  | A CFB container holding an `EncryptedPackage` stream — a password-protected workbook                                                                                               | Remove the password in Excel and save again as `.xlsx`         |
| `LEGACY_XLS` | A CFB container without encryption: a BIFF `.xls` workbook                                                                                                                         | It is a legacy Excel 97-2003 file; re-save it as `.xlsx`       |
| `XLSB`       | A zip whose main part is binary (`.xlsb`)                                                                                                                                          | It is an Excel Binary Workbook; re-save it as `.xlsx`          |
| `ODS`        | A zip with an OpenDocument `mimetype` entry                                                                                                                                        | It is an OpenDocument spreadsheet; open it and save as `.xlsx` |
| `NOT_XLSX`   | Anything else: CSV or other text, XML, HTML, `.docx`, `.pptx`, an empty file, a zip with no workbook part, or an input that is not bytes at all (a string, `null`, a plain object) | It is not a spreadsheet, and which of those it looks like      |

`NOT_XLSX` carries `detail.format` — `'text'`, `'xml'`, `'html'`, `'empty'`, `'docx'`, `'pptx'`, `'zip'`,
`'unknown'` — so an application that has its own CSV parser can route on it rather than showing an error at all.
[`sniff()`](./reading.md#sniffing-before-you-open) lets you make that decision before calling `openWorkbook`.

:::info[The `password-protected` contract]
The `ENCRYPTED` message **always contains the literal words `password-protected`**. That is a deliberate,
tested contract, because applications (Jetstream among them) match on it. It will not change without a major
version.
:::

### Why legacy formats are errors, not silent parses

SheetJS quietly parses `.xls`, `.ods`, SpreadsheetML 2003 and raw CSV bytes when they reach `XLSX.read` — which
is convenient right up to the point where a user uploads a `.xls`, gets different type coercion than everyone
else, and nobody can tell why. simple-excel supports `.xlsx` and `.xlsm` only (Strict OOXML is read-only) and
names everything else (ADR-007). Other libraries do not all get this right: office-kit fails on an `.xlsb` deep
inside its XML parser (its 0.11 release also reported every CFB file, legacy `.xls` included, as encrypted; 0.23
names both correctly).

## Container errors

| Code                  | Fires when                                                                                                                         |
| --------------------- | ---------------------------------------------------------------------------------------------------------------------------------- |
| `ZIP_TRUNCATED`       | The central directory or an entry runs past the end of the input, or a deflate stream ends early                                   |
| `ZIP_BOMB`            | A part declares, or inflates to, more than `limits.maxInflatedBytes`, or inflates more than 4 KB past the size its header declares |
| `ZIP_DUPLICATE_ENTRY` | Two central-directory entries normalize to the same name                                                                           |
| `ZIP_UNSUPPORTED`     | An entry uses a compression method other than stored or deflate, or is itself encrypted                                            |
| `ZIP_CRC_MISMATCH`    | An entry's bytes do not match the CRC-32 recorded for it                                                                           |

`ZIP_DUPLICATE_ENTRY` is a deliberate rejection rather than a guess. Readers disagree about duplicates — SheetJS
takes the first entry, office-kit, openpyxl and calamine take the last (`EC-ZIP-DUPLICATE-ENTRIES`) — so the same
file means different things depending on who opens it. That is exactly the shape of a smuggling bug, so it fails.

`ZIP_CRC_MISMATCH` is also stricter than some readers: SheetJS and office-kit do not check CRCs at all
(`EC-ZIP-CRC-MISMATCH`). A stream raises it after its last chunk, so a partially consumed iteration may have
yielded rows before the failure.

## XML errors

| Code            | Fires when                                                 |
| --------------- | ---------------------------------------------------------- |
| `XML_DOCTYPE`   | A `<!DOCTYPE` or an entity declaration appears in any part |
| `XML_MALFORMED` | Structurally broken XML in a part that had to be parsed    |

`XML_DOCTYPE` is not a parse failure — it is a refusal. Spreadsheet files never contain a DTD, and every entity
expansion attack (billion laughs, XXE) needs one. The tokenizer knows the five predefined entities and numeric
character references and nothing else, so there is no entity-expansion surface to guard in the first place
(`EC-XXE-DOCTYPE`).

## Limits

`LIMIT_EXCEEDED` covers every configurable cap, with the limit and the observed value in `detail`:

| Hit                                          | Configured by                  |
| -------------------------------------------- | ------------------------------ |
| Too many parts in the archive                | `limits.maxEntries`            |
| Shared-string table too large                | `limits.maxSharedStringChars`  |
| XML nested too deep                          | `limits.maxXmlDepth`           |
| One text node too long                       | `limits.maxTextLength`         |
| A cell past the column cap for this read     | `RowsOptions.maxColumns`       |
| A cell reference past column XFD             | — (`EC-REF-BEYOND-XFD`)        |
| More bytes than a memory sink allows         | `collectToBytes({ maxBytes })` |
| A declared size beyond what can be addressed | —                              |

Raising a limit is a deliberate act: the defaults are sized so a hostile upload fails fast rather than taking the
process with it. See [Reading](./reading.md#limits).

## Writer errors

| Code                 | Fires when                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               |
| -------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `WRITER_STATE`       | The writer was used out of order or given something it cannot write. Order: a second sheet opened while one is still open, a row written after `close()`, a sink written after it closed, `collectToBytes().result()` before `close()`. Workbook: `close()` with no sheet or with every sheet hidden, a `properties.created` that is invalid or outside the years 1–9999. Sheets: a `rowCount`, column width or `freeze` that is not a sensible whole number, a style id `registerStyle` did not return, a merge that is not `A1:C1`-shaped, a single cell, outside the grid, or overlapping another. Styles: a field outside Excel's ranges (see [Styles](./writing.md#styles)). Rows: a row that is not an array, a value of an unsupported type (the message names the sheet and cell), an error value that is not a `#` code (another `#` code such as `#SPILL!` is written as text) |
| `INVALID_SHEET_NAME` | A sheet name that cannot be sanitized into something Excel accepts (a non-string)                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                        |
| `ROW_OUT_OF_RANGE`   | A row past 1,048,576, a row wider than 16,384 columns, or more column definitions than a sheet can hold                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                  |
| `CELL_TOO_LONG`      | A cell longer than 32,767 characters while `cellOverflow: 'throw'`                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       |
| `ENTRY_TOO_LARGE`    | A zip entry or offset would overflow 32 bits while zip64 is off. The message says to enable zip64 (`zip64: true`) or split the data across sheets                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                        |

`WRITER_STATE` is almost always a sequencing mistake or a bad value in calling code, and it is deliberately loud: a
writer that carried on would produce a file Excel opens with its repair dialog. A check that fails before anything
is written (`registerStyle`, `addSheet`, `merge`) leaves the workbook usable. A failure once bytes are flowing
aborts the sink, and every later call rejects with the same error. Sheet names are sanitized rather than
rejected — see [Writing](./writing.md#sheet-names) — so `INVALID_SHEET_NAME` is rare.

## Lookup and lifecycle

| Code                      | Fires when                                                                                                                                         |
| ------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------- |
| `SHEET_NOT_FOUND`         | `workbook.sheet(x)` for a name or index that does not exist. The message lists the names that do                                                   |
| `ABORTED`                 | The caller's `AbortSignal` fired, `abort()` was called on the workbook writer or a sink, or a sheet was read after `workbook.close()`              |
| `UNSUPPORTED_ENVIRONMENT` | A required platform feature is missing: `collectToBlob()` with no `Blob`, reading with no `DecompressionStream('deflate-raw')` (Node before 20.12) |

`ABORTED` is what a cancelled export or read throws; treat it as an expected outcome, not a bug. Note that a
missing `CompressionStream` does _not_ raise `UNSUPPORTED_ENVIRONMENT` on write — the writer falls back to
storing parts uncompressed. Only reading genuinely requires `DecompressionStream`.

## The full list

```ts
type XlsxErrorCode =
  | 'NOT_XLSX'
  | 'ENCRYPTED'
  | 'LEGACY_XLS'
  | 'XLSB'
  | 'ODS'
  | 'ZIP_TRUNCATED'
  | 'ZIP_BOMB'
  | 'ZIP_DUPLICATE_ENTRY'
  | 'ZIP_UNSUPPORTED'
  | 'ZIP_CRC_MISMATCH'
  | 'XML_DOCTYPE'
  | 'XML_MALFORMED'
  | 'LIMIT_EXCEEDED'
  | 'SHEET_NOT_FOUND'
  | 'INVALID_SHEET_NAME'
  | 'ROW_OUT_OF_RANGE'
  | 'CELL_TOO_LONG'
  | 'ENTRY_TOO_LARGE'
  | 'WRITER_STATE'
  | 'ABORTED'
  | 'UNSUPPORTED_ENVIRONMENT';
```
