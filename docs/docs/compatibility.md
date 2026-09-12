---
id: compatibility
title: Compatibility
description: Which applications and readers the files were checked against, and how that checking is done.
---

# Compatibility

"It opens in Excel" is not a claim you can make from a unit test. This library is developed against a committed
corpus of real files and an oracle that runs actual applications over the output — Excel itself included. This
page is what has been checked, against what, and how.

The current results live in [`research/05-compatibility-matrix.md`](https://github.com/jetstreamapp/simple-excel/blob/main/research/05-compatibility-matrix.md),
which is regenerated from the newest oracle run rather than hand-maintained.

## The readers in the oracle

| Reader              | What it stands for                                                                                                                 |
| ------------------- | ---------------------------------------------------------------------------------------------------------------------------------- |
| `excel`             | Microsoft Excel 16 (read-only licence) driven by AppleScript: opens the file, and fails on a repair prompt or a recovery-log entry |
| `libreoffice`       | LibreOffice 26.2 headless — a re-save has to succeed                                                                               |
| `validator`         | Open XML SDK schema validation (`@xarsh/ooxml-validator`, Microsoft 365 profile)                                                   |
| `sheetjs`           | SheetJS CE 0.20.3 with `cellDates: true, cellText: false`                                                                          |
| `openpyxl`          | Python openpyxl 3.1 with `rich_text=True`                                                                                          |
| `calamine`          | Rust calamine via python-calamine with `skip_empty_area=False`                                                                     |
| `office-kit`        | `@office-kit/xlsx` 0.11 document API                                                                                               |
| `office-kit-stream` | `@office-kit/xlsx` streaming API (`loadWorkbookStream` + `iterRows`)                                                               |

Google Sheets, Apple Numbers and Excel on Windows are checked by hand — they cannot be automated locally — and
that checklist is part of the definition of done for 1.0.

## What the writer has passed

As of the `phase-b-writer` oracle run, the canonical workbook written by this library (in its inline-string,
shared-string and zip64 variants) is:

- **Open XML SDK validator: PASS** on all three. Most other writers fail it — ExcelJS and office-kit both emit
  `CT_Font` children out of schema order (`EC-STYLES-FONT-ELEMENT-ORDER`).
- **Excel: PASS** on all three — opened with no repair prompt and no recovery-log entry.
- **LibreOffice: PASS** on all three (headless re-save succeeds).
- **openpyxl, calamine, SheetJS**: read the standard variants with the residual differences that are each
  reader's own conventions, not the file's — SheetJS returns a serial rather than a `Date` for the fake
  1900-02-29, collapses CRLF and decodes `_X0041_`; calamine disagrees about blank versus empty string; openpyxl
  does not decode `_xHHHH_` escapes the way Excel does. Those are catalogued in
  [`research/04-edge-case-catalog.md`](https://github.com/jetstreamapp/simple-excel/blob/main/research/04-edge-case-catalog.md)
  and allowed per fixture, not silently ignored.
- **SheetJS on the zip64 variant: cannot open it** — see [the zip64 caveat](./streaming-and-memory.md#the-zip64-caveat).

Google Sheets, Numbers and Excel for Windows have **not** yet been recorded for this engine's output; they are
manual checks scheduled before 1.0.

## What the reader has been checked against

The corpus contains files produced by Excel 365 on macOS (including a Strict Open XML save and a 1904-date-system
workbook), Google Sheets, Apple Numbers 14.4, LibreOffice 26.2, Salesforce report exports (Apache POI), SheetJS,
ExcelJS, openpyxl, XlsxWriter, write-excel-file and `@office-kit/xlsx`, plus hand-built files for structural edge
cases and a set of deliberately hostile ones.

Every fixture has a recorded verdict, and the corpus suite fails if a fixture stops matching it. The hostile set
has to produce a _classified error_ — never a crash, never unbounded memory. Some of the things in there:

| Fixture family                        | What it is                                       |
| ------------------------------------- | ------------------------------------------------ |
| `hostile-truncated-central-directory` | Archive cut before its directory                 |
| `hostile-duplicate-sheet-entries`     | Two entries named `xl/worksheets/sheet1.xml`     |
| `hostile-zip-bomb-30mb-sheet`         | 30 MB of sheet XML deflated about 190:1          |
| `hostile-deeply-nested-rich-text`     | 50,000 nested `<r>` runs in one shared string    |
| `hostile-crc-mismatch`                | Entry bytes that do not match their recorded CRC |
| `hostile-biff8-xls-renamed`           | A real `.xls` under an `.xlsx` name              |
| `hostile-xlsb-renamed`                | A real `.xlsb` under an `.xlsx` name             |
| `hostile-not-a-zip`                   | SpreadsheetML 2003 XML with an `.xlsx` name      |

Several well-known readers fail these. SheetJS happily parses the renamed `.xls` and `.xlsb` as if nothing were
wrong; the Open XML SDK validator crashes with a stack overflow on the deeply nested rich text.

## How the checking works

### The fixture corpus

`fixtures/manifest.json` records a sha256 and the provenance of every file — which application and version wrote
it, or which script generated it. `npm run fixtures:check` fails if any file drifts from its hash. A fixture also
carries the policies that apply to it (for example, that a golden is compared with 32,767-character truncation
applied) and the diff categories that are expected for that generator.

### The typed dump and the ground truth

`fixtures/canonical/canonical.json` is the ground truth: the exact values a canonical workbook is supposed to
contain, including the awkward ones — control characters, `_x` escape literals, formula-looking text, a 17-digit
integer, the maximum double, dates in the 1900 leap window, a time-only value, a DST-gap datetime.

Every reader in the oracle dumps every fixture into the same typed shape, and `oracle/diff.mjs` compares it
against the ground truth, classifying each mismatch (`crlf-normalized`, `date-tz-offset-shift`,
`blank-vs-empty-string`, `escape-sequence-mangled`, …). Those categories are what the matrix's percentages are
made of, and each one maps to a catalog entry that explains whose fault it is.

### Golden bytes

With `deterministic: true` the writer's output is reproducible: fixed zip timestamps and fixed `docProps` dates.
The golden-bytes suite pins sha256 hashes of that output, so an accidental change to the XML or the container is
caught immediately. Updating a golden is deliberate — it requires an environment flag and a validator and oracle
pass first.

### The oracle run

`npm run oracle` drives the applications: Excel under AppleScript with repair detection, LibreOffice headless, a
Python virtual environment for openpyxl and calamine, the Open XML SDK validator, SheetJS and office-kit in
Node. It needs a Mac with Excel and LibreOffice installed, so it is a local pre-release gate rather than a CI
job; the results folder is committed with each run.

CI runs the subset that does not need those applications: the unit and corpus suites, the hostile suite, the
SheetJS parity suite, the validator, and the golden-byte pins.

## Two things worth knowing

:::note[The validator is a strictness gate, not a compatibility predictor]
Excel opens every golden in the corpus, including the ones the Open XML SDK validator rejects. The validator
catches real schema mistakes, which is why the writer is held to a clean result — but a validator failure
elsewhere does not mean Excel will refuse the file, and a validator pass does not mean it will accept it.
:::

:::note[Round-tripping through another application is not a repair]
Excel re-saving a workbook written by another library preserves that library's serials exactly — an ExcelJS
file's seven-hour date shift, an office-kit file's negative time serials
(`EC-EXCEL-RESAVE-PRESERVES-SHIFTED-SERIALS`). A bad writer's dates do not get fixed by opening and saving in
Excel. Whatever went into the file is what stays there.
:::
