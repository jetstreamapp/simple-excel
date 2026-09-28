---
id: dates-and-values
title: Dates and values
description: The value model, date serials and their historical quirks, string escaping, number precision, booleans and error cells.
---

# Dates and values

Spreadsheet cells hold fewer types than you would expect, and the mapping to JavaScript has several traps that
predate everyone reading this. This page is the whole contract.

## The value model

A cell reads back as:

```ts
type CellValue = string | number | boolean | Date | null;
```

and accepts, on write:

```ts
type CellInput = CellValue | CellError | bigint | undefined;
```

| Write input                   | In the file                                            | Reads back as                                 |
| ----------------------------- | ------------------------------------------------------ | --------------------------------------------- |
| `string`                      | Shared or inline string, `_xHHHH_`-escaped             | `string`                                      |
| `number` (finite)             | `<v>` with the shortest round-trip form                | `number`                                      |
| `number` (`NaN`, `±Infinity`) | The error cell `#NUM!`                                 | `'#NUM!'` (or a `CellError`, per `errors`)    |
| `boolean`                     | `t="b"` with `1` / `0`                                 | `boolean`                                     |
| `Date` (valid)                | A serial under a date number format                    | `Date` (or a number with `dates: 'serial'`)   |
| `Date` (invalid)              | Nothing is written at all                              | absent                                        |
| `null` / `undefined`          | No `<c>` element, unless a non-default style was given | `null` in array mode, `defval` in object mode |
| `bigint`                      | A number while `\|value\| <= 2^53`, otherwise text     | `number` or `string`                          |
| `{ error: '#N/A' }`           | `t="e"` — a real error cell                            | `'#N/A'`, `{ error: '#N/A' }` or `null`       |

There is no "empty string versus blank" ambiguity to resolve on write: `''` writes an empty string cell, `null`
writes no cell. On read, object mode collapses both to `defval` unless you set `defval: null`
(`EC-BLANK-VS-EMPTY-STRING`).

## Dates

### Serials

Excel does not store dates. It stores a **number** — whole days since an epoch, with the time of day as the
fraction — and a **number format** on the cell that makes it render as a date. A cell is a date only because its
format says so (`EC-DATE-DETECTION-VIA-NUMFMT`), which is why a reader that skips the style table hands you
`45361.39583` instead of a `Date`.

simple-excel resolves the format for you: `cellXfs[s] → numFmtId → is this a date format`, using the built-in id
table (14–22, 45–47, plus the locale ranges) and, for custom codes, a scan for `y`, `m`, `d`, `h`, `s` or `AM/PM`
outside quoted and bracketed sections. `General` and `@` are never dates.

All of the serial arithmetic is integer arithmetic on the proleptic Gregorian calendar. A time zone never reaches
it.

### `dates` on write

```ts
const workbook = createWorkbookWriter(sink, { dates: 'local' });
```

A JS `Date` is an instant, but a spreadsheet cell holds a wall clock. `dates` says which fields of the `Date` are
the wall clock you meant:

| Value               | The serial is built from                                    |
| ------------------- | ----------------------------------------------------------- |
| `'local'` (default) | `getFullYear()`, `getMonth()`, `getDate()`, `getHours()`, … |
| `'utc'`             | `getUTCFullYear()`, `getUTCMonth()`, `getUTCHours()`, …     |

The default is "what you see in the debugger is what you see in Excel". Pick `'utc'` when your `Date`s were
constructed from UTC fields in the first place.

:::caution
Getting this wrong is silent and ugly. A library that takes UTC fields from a local-midnight `Date` puts the host
offset into the file as a time fraction — Excel shows `7:00 AM` for a date you meant as midnight — and turns
time-only values into negative serials, which Excel renders as `########`
(`EC-DATE-SERIAL-EPOCH-UTC-FIELDS`). That is a real observed behaviour of another library, not a hypothetical.
:::

### `dates` on read

```ts
const workbook = await openWorkbook(file, { dates: 'local' });
```

| Value               | A date cell becomes                                                           |
| ------------------- | ----------------------------------------------------------------------------- |
| `'local'` (default) | `new Date(y, m - 1, d, h, mi, s, ms)` from the serial's wall-clock components |
| `'utc'`             | `new Date(Date.UTC(...))` from the same components                            |
| `'serial'`          | The raw number. Never a `Date`                                                |

`'local'` matches what SheetJS's `sheet_to_json` produces, so switching libraries does not shift every timestamp
in an application. `'serial'` is the honest choice when you are going to reformat the value yourself and do not
want a `Date` inventing a time zone on the way.

### The 1900 leap-year bug

Excel believes 1900 was a leap year. It was not. Serial **60 is the fictional 1900-02-29**, so every serial below
61 is one day away from a correct calendar.

| Serial | Excel shows | simple-excel reads as |
| ------ | ----------- | --------------------- |
| 1      | 1900-01-01  | 1900-01-01            |
| 59     | 1900-02-28  | 1900-02-28            |
| 60     | 1900-02-29  | 1900-03-01            |
| 61     | 1900-03-01  | 1900-03-01            |

Serial 60 has no real date, so it maps to 1900-03-01 — the same value serial 61 gives — which is what most other
readers do. Writing is the mirror: 1900-01-01 through 1900-02-28 become serials 1–59, and only 1900-03-01 onward
lines up with the raw day count. Real dates are never shifted.

Several well-known writers get this off by one and write 1900-02-28 as serial 60, so Excel shows it as
1900-02-29 or 1900-03-01 (`EC-DATE-LEAP-WRITER-OFF-BY-ONE`). If your data lives in 1900 you should be suspicious
of every tool in the chain.

### Before 1900

Serial 0 is Excel's `1/0/1900` placeholder and negative serials display as `########`; there is no representation
for 1899 or earlier in the 1900 system. A `Date` before the epoch is therefore **written as ISO text**
(`1899-12-31T12:00:00`) rather than as a number that would display as garbage (`EC-DATE-PRE-1900`). The value
stays readable and obviously not a date cell.

On read, a **negative serial in a date-formatted cell comes back as the raw number**, exactly what Excel shows as
`########`. So does a serial past 9999-12-31, Excel's last date: a phone number typed into a date-formatted column
reads as the number, not as a year-10000 or invalid `Date` (`EC-DATE-SERIAL-OVER-9999`). The writer mirrors it:
a `Date` after 9999 is written as ISO text (`10000-01-01T00:00:00`), like one before 1900. SheetJS instead re-derives a time of day from it, which matters for files SheetJS itself wrote: its
time-only cells are negative serials (`EC-DATE-TIME-ONLY-NEGATIVE-SERIAL`), and they read back here as numbers.

`Date`s before roughly 1901 also carry local-mean-time offsets in JavaScript — Los Angeles is `-7:52:58`, not
`-8:00` — so local-field conversions of such dates are off by minutes and seconds
(`EC-DATE-HISTORICAL-TZ-OFFSET`). If you work with pre-1901 dates, use `dates: 'utc'` on both sides or carry them
as text.

### Time-only values

A duration or a clock time with no date is a serial below 1: `0.5` is noon. On read, those cells come back as a
`Date` on the marker day **1899-12-30**, because a `Date` has to have a date part:

```ts
// a cell holding 0.524268391204 (12:34:56.789)
value.getHours(); // 12
value.getFullYear(); // 1899 — the marker, not data
```

To write one, build a `Date` on that same marker day and the writer emits the fraction alone:

```ts
await sheet.writeRow([new Date(1899, 11, 30, 12, 34, 56, 789)]); // serial 0.5242…
```

:::caution
Readers disagree about the marker day — 1899-12-30 here and in office-kit, 1899-12-31 in SheetJS, a
`datetime.time` in openpyxl (`EC-TIME-ONLY-HAS-DATE-PART`). If you round-trip time-only values through another
tool, compare the time fields, never the whole `Date`.
:::

### The 1904 date system

Workbooks saved by old Mac Excel use a different epoch, shifted by 1,462 days. The flag lives in the workbook
part; the reader honours it everywhere and exposes it:

```ts
workbook.date1904; // boolean
```

To write one, pass `date1904: true`. There is no reason to unless you are matching an existing file
(`EC-DATE-1904`).

### The DST gap

Asking JavaScript for a local `Date` at a wall clock that does not exist — 02:30 on a US spring-forward Sunday —
silently gives you 03:30. That shift happens in `new Date(...)`, before any library sees the value
(`EC-DATE-DST-GAP`). Nothing downstream can undo it. If you are generating timestamps from wall-clock components,
either use `Date.UTC` and `dates: 'utc'`, or accept the hour.

## Strings

### `_xHHHH_` escapes

XML 1.0 cannot carry control characters, and real data has them: text pasted out of another system, long-text
fields from a CRM. The format's answer is `ST_Xstring`, which spells an illegal character as `_x0001_`.

simple-excel writes those escapes and decodes them on read. Two consequences worth knowing:

- **A literal `_x0041_` in your data** would be read back as `A` if written as-is, so the writer escapes the
  leading underscore: `_x005F_x0041_`. Overlapping runs are escaped too, which several writers get wrong
  (`EC-XML-ESCAPE-LITERAL`, `EC-XML-ESCAPE-OVERLAP`).
- **Only lowercase `_x` + four hex digits + `_` is an escape.** `_X0041_` with a capital X is literal text and
  stays literal. SheetJS decodes it anyway (`EC-XML-ESCAPE-CASE`), so a value that survives a round trip here may
  not survive one through SheetJS.

Leading and trailing whitespace is preserved with `xml:space="preserve"`; without it some readers trim it away.

### CRLF

A raw carriage return inside a text node is collapsed to a line feed by XML's own line-end normalization — the
`\r` is simply gone, in every reader. So the writer escapes CR as `_x000D_` (`EC-CRLF-NORMALIZED`), and the reader
decodes it without normalizing. `'a\r\nb'` written here reads back as `'a\r\nb'`.

Several other writers emit the raw CR and lose it, and SheetJS collapses `\r\n` to `\n` on read even when the
escape is present. If line endings matter to you, that is where they go missing.

### The 32,767-character limit

Excel refuses a cell longer than 32,767 characters; a file containing one opens with a repair prompt. The writer
applies a policy instead: truncate with a visible suffix (the default, `'...(truncated)'`, counted and reported)
or throw `CELL_TOO_LONG`. See [Writing](./writing.md#long-cells-the-32767-character-limit).

Reading is not affected — a longer cell produced by some other tool is returned whole.

### Rich text

Rich text is a sequence of styled runs inside one cell. On read the runs are flattened to plain text, which is
what every reader does and what a data pipeline wants. Writing rich text is not supported.

## Numbers

Numbers are written with JavaScript's shortest round-trip form (up to 17 significant digits), which Excel parses
exactly, with the exponent marker uppercased to match Excel's own output. `-0` is written as `0`, because the
format has no negative zero (`EC-NUM-NEGATIVE-ZERO`).

:::tip[Excel's display is not Excel's storage]
Excel **displays** and rounds to 15 significant digits: `9007199254740991` shows as `9.007199254741E+15`, and
editing that cell in Excel will commit the rounded value. The full-precision number is still in the file, and
this library returns what is stored (`EC-NUM-15-SIGNIFICANT-DIGITS`). If exact large integers matter, write them
as text — which `bigint` above `2^53` does for you.
:::

`NaN` and the infinities have no valid lexical form in the format. Rather than write something Excel will repair,
they become the error cell `#NUM!`. (Numbers.app does emit `<v>inf</v>`; Excel repairs those files and openpyxl
crashes on them — `EC-NUMBERS-INF-VALUE`.)

## Booleans

`true` and `false` are written as `t="b"` cells with `1` and `0`, and read back as booleans. A file that stored
`TRUE`/`FALSE` as text — a CSV import will do this — reads back as the strings, because that is what it is.

## Error cells

```ts
type CellErrorCode = '#NULL!' | '#DIV/0!' | '#VALUE!' | '#REF!' | '#NAME?' | '#NUM!' | '#N/A' | '#GETTING_DATA';
```

Write one by passing `{ error: '#N/A' }`; it becomes a genuine `t="e"` cell, not the text `#N/A`. Writers without
an error API write the text instead, and no reader can then tell the two apart (`EC-ERROR-CELL-AS-TEXT`).

On read, `OpenOptions.errors` decides the shape:

| `errors`             | An error cell becomes | Row type                                 |
| -------------------- | --------------------- | ---------------------------------------- |
| `'string'` (default) | `'#N/A'`              | `CellValue`                              |
| `'object'`           | `{ error: '#N/A' }`   | `CellValue \| CellError` (typed for you) |
| `'null'`             | `null`                | `CellValue`                              |

`'string'` is the default because it is what a user sees in Excel and what most pipelines want to write straight
into a CSV. Use `'object'` when you need to distinguish an error from a string that happens to read `#N/A`.

`head()` is the exception: it always reports both the text and the code on its `RawCell`, whatever the workbook
was opened with.

## Formula-looking text

Text beginning with `=` is written as text and never as a formula. Inferring one is both an injection surface
(`=cmd|calc`) and a way to make Excel repair the file — openpyxl and LibreOffice's CSV import both do it, and both
produce files Excel logs a repair for (`EC-FORMULA-LIKE-TEXT-WRITTEN-AS-FORMULA`). If you export user data to CSV
elsewhere in your application, that path needs the same hygiene; this library only controls the xlsx one.
