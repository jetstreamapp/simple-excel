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
| `simple-excel`      | This library through its public API — the same typed dump the test suite builds, so the two cannot disagree                        |
| `excel`             | Microsoft Excel 16 (read-only licence) driven by AppleScript: opens the file, and fails on a repair prompt or a recovery-log entry |
| `libreoffice`       | LibreOffice 26.2 headless — a re-save has to succeed                                                                               |
| `validator`         | Open XML SDK schema validation (`@xarsh/ooxml-validator`, Microsoft 365 profile)                                                   |
| `sheetjs`           | SheetJS CE 0.20.3 with `cellDates: true, cellText: false`                                                                          |
| `openpyxl`          | Python openpyxl 3.1 with `rich_text=True`                                                                                          |
| `calamine`          | Rust calamine via python-calamine with `skip_empty_area=False`                                                                     |
| `office-kit`        | `@office-kit/xlsx` document API (0.11.0 in the recorded runs; the repo now pins 0.23)                                              |
| `office-kit-stream` | `@office-kit/xlsx` streaming API (`loadWorkbookStream` + `iterRows`)                                                               |

Google Sheets, Apple Numbers, Excel for Mac and Excel for the web (standing in for Excel on Windows) cannot be
automated locally, so they were checked by hand on 2026-09-12. All passed; the one expected failure is Drive
refusing to convert the zip64 variant (below).

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
- **SheetJS on the zip64 variant: cannot open it**, and **Google Drive fails to convert it** to a Google Sheet — see
  [the zip64 caveat](./streaming-and-memory.md#the-zip64-caveat). Neither affects files written with the default
  `zip64: 'auto'`.
- **office-kit**: reads all three at 94% of the ground-truth cells (`phase-d-office-kit-on-ours`). The gap is its
  own date handling — it puts a time-only value on 1899-12-31 rather than 1899-12-30 — plus the two writer
  conventions below. Its _streaming_ API additionally leaves `_xHHHH_` escapes undecoded when the strings are
  inline, so the inline variant reads at 88% there.
- **Two conventions of ours show up in every reader's dump**, and are deliberate: a pre-1900 date is written as
  ISO text rather than a serial ([ADR-003](https://github.com/jetstreamapp/simple-excel/blob/main/research/adr/ADR-003-date-semantics.md)),
  and a `Date` that falls in a DST gap serializes an hour later, because that instant does not exist locally.

**Google Sheets** imported `canonical.xlsx` and `canonical.inline.xlsx` cleanly and identically (manual check, 2026-09-12): every difference on screen was Sheets'
own behaviour — 15-significant-digit display of large numbers, its one-day offset for Excel serials below 61
(`EC-DATE-1900-LEAP-BUG`), seconds rounding under `hh:mm:ss` — with the stored values intact. The zip64 variant is
the exception: Drive refuses to convert it (`EC-ZIP64-SMALL`), which the default `zip64: 'auto'` never produces.
**Excel for Mac (365)** was also opened by hand (2026-09-12) with no repair prompt; its 15-digit display, its
zeroing of sub-precision subtractions and its `9.99999999999999E+307` formula cap all showed, with the stored values
intact in its re-save. **Apple Numbers 14.4** opened `canonical.xlsx` with one banner, "hidden sheets were made
visible" (Numbers has no hidden sheets); dates were right, time-only values sit on its 1899-12-31 marker day, and its
own limits apply — 15-significant-digit precision (`EC-NUMBERS-15-DIGIT-PRECISION`) and no representation of the max
double (`EC-NUMBERS-INF-VALUE`). **Excel for the web** (Microsoft's server-side engine, standing in for Excel on Windows, which the project cannot
run) re-downloaded both goldens with every value, escape, control character, hidden sheet and sheet name intact:
the re-downloads read back at 650/658, the same as the originals. Its only on-screen remark was the cosmetic
"text that looks like a formula" outline on the cells that deliberately hold `=`-prefixed text. The procedure and
results log live in
[`fixtures/golden/simple-excel/STEPS.md`](https://github.com/jetstreamapp/simple-excel/blob/main/fixtures/golden/simple-excel/STEPS.md).

## What the reader has passed

As of the `phase-d-simple-excel-*` oracle runs, this library reads every fixture in the corpus (the six Google
Sheets, Numbers and Excel-for-the-web re-exports of our goldens added since are covered by the corpus suite):

| Set                      | Result                                                                                                                                                                                                                       |
| ------------------------ | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 12 hostile files         | **12 of 12 rejected** with the classified error the manifest asks for — nothing crashed, hung or was accepted (the zip bomb is rejected against the read limits you configure; the default cap is a generous 1 GiB per part) |
| 15 structural edge cases | **13 PASS**; the two that differ are fixtures whose point is that a reader has to choose (a shared-string index out of range, which we read blank like Excel, and a 1904-epoch workbook)                                     |
| 22 goldens               | every remaining difference is a deviation the generator wrote into the file, each one allowed for that fixture by name in `test/corpus-policies.json`                                                                        |
| our own 3 goldens        | 650 of 658 cells; the 8 are the two writer conventions above                                                                                                                                                                 |
| 3 Jetstream assets       | both workbooks stream; the CSV asset is refused with `NOT_XLSX`, which is the designed behaviour for non-xlsx bytes                                                                                                          |

Two examples of what "the generator wrote it that way" means: a file from ExcelJS (and anything re-saved from it,
including by Excel itself) carries dates built from UTC fields, so 68 cells in it mean something other than the
ground truth in every reader; openpyxl refuses control characters at write time, so those cells are empty in the
file for everybody. The per-category breakdown for each fixture is in
[`research/05-compatibility-matrix.md`](https://github.com/jetstreamapp/simple-excel/blob/main/research/05-compatibility-matrix.md).

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

## Which browsers it has been run in

The browser is the primary target, so the whole surface is exercised in a real browser rather than only in Node.
`npm run smoke:browsers` runs one page — write 20,000 mixed rows to `collectToBlob()`, read them back, do the
same in a module worker, read an Excel- and a Sheets-authored file, reject a password-protected file and a zip
bomb, stream into a `WritableStream` — and fails on any difference. It is described in
[Contributing](./contributing.md#browser-smoke-test).

| Browser                         | Version               | Result                                                           |
| ------------------------------- | --------------------- | ---------------------------------------------------------------- |
| Chromium (Playwright, headless) | 153.0.8010.12         | 72/72 checks                                                     |
| Firefox (Playwright, headless)  | 155.0                 | 72/72 checks                                                     |
| WebKit (Playwright, headless)   | 26.6 (Safari 26.6 UA) | 72/72 checks                                                     |
| Safari (macOS, by hand)         | —                     | run `node test/browser/run.mjs --serve` and open the printed URL |

Run 2026-09-12 on macOS 26 (Apple silicon). Every check passed in all three: `CompressionStream('deflate-raw')`,
`DecompressionStream`, `WritableStream` and `Blob` are present everywhere, module workers load the library
unchanged, a `Blob` written in a worker transfers to the page by reference, and the reader returns the same
values for every trap cell — unicode, a CRLF kept as `\r\n`, a control character, an `_x0041_` literal, `-0`
(written and read as a plain zero, since Excel has no negative zero), `1e21`, `0.1 + 0.2`, `9007199254740991`,
a time-only `Date`, and a 40,000-character cell truncated to 32,767 with the `...(truncated)` suffix.

Two engine differences show up, neither of them a defect:

- **The compressed size differs per engine.** The same deterministic workbook is 1,654,492 bytes in Chromium,
  1,712,097 in Firefox and 1,686,065 in WebKit, because the bytes come from each engine's own
  `CompressionStream('deflate-raw')` tuning. The XML inside is identical. See
  [the deterministic-output caveat](./writing.md#deterministic-output).
- **`performance.measureUserAgentSpecificMemory()` is Chromium-only**, and only in a document (it is not exposed
  in a dedicated worker). Where it exists, writing 200,000 rows (a 15.7 MB file) in a worker leaves an
  agent-cluster peak of about 4 MB, 1.7 MB of it the worker — the flat-memory claim, measured in a browser
  rather than in Node.

To check real Safari, which cannot be automated here, run `node test/browser/run.mjs --serve` and open the
printed URL: the page shows the same table.

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
