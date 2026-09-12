# 10 - `@office-kit/xlsx` evaluation

Can `@office-kit/xlsx` (MIT, TypeScript, v0.11.0) replace SheetJS in Jetstream as-is, with upstream fixes,
or not at all? Scored against the requirements rubric in 07 using the fixtures (04), the oracle runs (05) and
the benchmark (06).

_Accurate as of 2026-09-12 (benchmark runs `2026-09-11-macbook-air-*`; user-produced goldens run `2026-09-12-user-goldens`)._

## Verdict

**Not adoptable as-is, and the blocking gap is the one Jetstream cares most about.** The zip/XML core, the
document API and the hostile-input handling are solid (every golden opens in Excel and LibreOffice; DOCTYPE,
truncation and deep nesting are rejected cleanly). But the "fixed-memory streaming" writer is streaming in API
shape only:

- `appendRow` accumulates the worksheet XML in memory and the zip entry is deflated and handed to the sink
  only at `finalize()` - even with `toWritable(fs.createWriteStream())` the output file stays at 0 bytes
  until the end (300k numeric rows: RSS 67 → 130 MB, file 0 → 8 MB at finalize).
- Every string is interned in an unbounded shared string table that `finalize()` serializes with one
  `Array.join` - `RangeError: Invalid string length` in `serializeSharedStrings` at 1M×20 and at the
  18M-cell shape, the same exception SheetJS throws, at roughly the same size.
- Consequently the memory gate fails on the Jetstream-shaped dataset: 100k×20 `mixed` peaks at 1,158 MB
  (0.40× SheetJS's 2,899 MB; gate ≤ 0.25×). Only the numeric dataset is flat (54 MB, 0.03×).

Fixing this means a streaming zip-entry writer, chunked shared-string serialization and an inline-string or
bounded-SST mode - the same three pieces that make up Phase B of the hand-rolled design (07 A2). Everything
else in this document is edge-case work of the kind the maintainer has been shipping same-day. The decision
is recorded in ADR-008: **dual track, time-boxed** - open the upstream issues below (the streaming claim is a
documented feature, so a fix is in the maintainer's interest) while prototyping the streaming writer core
per 07, and decide at the time-box whether office-kit + our patches, or our writer + office-kit's reader, or
a full hand-roll ships.

## 1. API mapping to Jetstream's contracts

| Jetstream need (01)                                             | office-kit API                                                                                                | Fit               | Notes                                                                                                                                                                    |
| --------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------- | ----------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| Write array-of-rows sheets, streaming, bounded memory           | `createWriteOnlyWorkbook(sink)` → `addWorksheet` → `appendRow(WriteOnlyRowItem[])` → `close`/`finalize`       | **Blocked**       | Buffers the sheet and the SST until `finalize()` (see verdict, W6/W7). One open sheet at a time; `setColumnWidth` must precede the first row.                            |
| Sinks: Blob (web), IPC/fs (desktop), buffer (extension/canvas)  | `toBlob()`, `toArrayBuffer()`, `toFile()`, `toWritable()`, `XlsxSink`                                         | Good shape        | Web Streams based; maps onto the existing `HistoryWriteStream`. Bytes arrive only at finalize today.                                                                     |
| Header bold, per-column number formats, dates                   | `WriteOnlyRowItem = value \| { value, style }` with `numberFormat`, `font`                                    | Good              | Style dedupe is internal (`allocateXfId`).                                                                                                                               |
| Column widths, merges, freeze, autofilter (Permission Manager)  | document API: `setColumnWidth`, `mergeCells('A1:C1')`, `setFreezePanes(ws, 'A3')`, `setAutoFilter`            | Good with caveats | `freezePanes(ws, rows, 0)` rejects `cols=0`; `mergeCells` object form is `{minRow,minCol,maxRow,maxCol}`.                                                                |
| 32,767-char truncation with a count                             | none                                                                                                          | Adapter           | Keep Jetstream's `truncateCellsForExcel`.                                                                                                                                |
| Sheet-name sanitizing/dedupe                                    | validation throws on invalid names                                                                            | Adapter           | Keep Jetstream's helpers.                                                                                                                                                |
| Read: list sheets cheaply, pick one, typed rows                 | `loadWorkbookStream` → `sheetNames`, `openWorksheet(name).iterRows()` (cells with `styleId`) / `iterValues()` | Adapter needed    | Dates come back as serials in both APIs; resolve `styles.cellXfs[styleId].numFmtId` → `isDateFormat` → `excelToDate` (≈20 lines, see `oracle/office-kit/read-dump.mjs`). |
| Read: `sheet_to_json` parity (`__EMPTY`, `defval`, `blankrows`) | none                                                                                                          | Adapter           | Object-mode conversion stays in Jetstream.                                                                                                                               |
| Read: raw `B1/B2/B3/A5` (multi-object template)                 | `getCellByCoord`, or `iterRows({ minRow: 1, maxRow: 5 })`                                                     | Good              | Band queries with `minRow > 1` materialize the whole inflated sheet (README).                                                                                            |
| Sniff non-xlsx bytes (CSV, .xls, .ods, encrypted)               | throws `OpenXmlIoError` / `OpenXmlNotImplementedError`                                                        | Adapter           | Developer-facing messages; `.xls` reported as "Encrypted". Sniff first (ADR-007).                                                                                        |
| Error text contract `password-protected`                        | "Encrypted xlsx is not supported. Decrypt with msoffcrypto-tool first."                                       | Adapter           | Map the error class to Jetstream's message.                                                                                                                              |
| Workers, MV3 CSP, no wasm                                       | pure JS; deps fflate + fast-xml-parser + saxes                                                                | Verify            | fflate's async API spawns `blob:` workers (blocked by the extension CSP); confirm only the sync/stream API is used.                                                      |

## 2. Reader gaps found by the corpus

| #   | Gap                                                                                                                                                                                                                                                  | Evidence                                                                                                                            | Severity for Jetstream                                                                                  |
| --- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------- |
| R1  | `_xHHHH_` escapes are **not decoded** on read (`a_x0001_b`, `crlf_x000d_`)                                                                                                                                                                           | `EC-XML-ESCAPE-UNDECODED`; office-kit reading the SheetJS and office-kit goldens; Excel decodes them                                | High: Salesforce long text with CR/control chars round-trips wrong; related to open upstream issue #131 |
| R2  | Rows/cells **without `r`**: `loadWorkbook` throws, `loadWorkbookStream` mis-positions every cell                                                                                                                                                     | `edge-missing-r-attributes` (SheetJS, calamine, LibreOffice, Excel all infer)                                                       | Medium                                                                                                  |
| R3  | **Strict OOXML** namespaces/relationship types not recognized                                                                                                                                                                                        | `edge-strict-namespaces` (Excel, LibreOffice, SheetJS open it)                                                                      | Low-Medium (Excel's "Strict Open XML Spreadsheet" save)                                                 |
| R4  | **Backslash** relationship targets not normalized                                                                                                                                                                                                    | `edge-backslash-rel-targets`                                                                                                        | Low (calamine fixed the same in 0.29)                                                                   |
| R5  | Max double `1.7976931348623157E+308` **dropped**                                                                                                                                                                                                     | `EC-NUM-MAX-DOUBLE`                                                                                                                 | Low                                                                                                     |
| R6  | Dates are raw serials in both APIs (`date1904` correct once resolved)                                                                                                                                                                                | `EC-DATE-DETECTION-VIA-NUMFMT`, `edge-date1904` PASS                                                                                | Adapter (documented)                                                                                    |
| R7  | Every CFB input reported as "Encrypted" (legacy `.xls` too); an `.xlsb` (binary parts inside a zip) fails deep inside the XML parser (`readTagExp returned undefined`) instead of being recognized                                                   | `hostile-biff8-xls-renamed`, `hostile-xlsb-renamed` (both saved by Excel 365)                                                       | Adapter (sniff first)                                                                                   |
| R8  | SST index out of range: document API throws, stream API silently drops cells; Excel opens the file                                                                                                                                                   | `edge-sst-index-out-of-range`                                                                                                       | Low; inconsistent between the two APIs                                                                  |
| R9  | Stream mode returns `''` for a CDATA inline string that document mode reads correctly                                                                                                                                                                | `edge-inline-strings-cdata` (`iterRows` B1)                                                                                         | Low                                                                                                     |
| R10 | Read speed 1.2-1.5× SheetJS at 100k rows (memory 0.29-0.39×)                                                                                                                                                                                         | 06 shapes run                                                                                                                       | Fails the ≤ 1.0× time gate by a margin that may be acceptable given the memory win                      |
| R11 | **Cannot open Salesforce report exports**: Apache POI writes `<u val="none"/>` (a legal `ST_UnderlineValues` member) and office-kit throws `expected one of [single, double, singleAccounting, doubleAccounting], got "none"` for the whole workbook | `golden-sfdc-report-details`, `golden-sfdc-report-formatted` (both APIs); SheetJS, openpyxl, calamine, LibreOffice, Excel open them | **Blocking**: these are the files Jetstream users upload                                                |

## 3. Writer gaps

| #   | Gap                                                                                                                                                       | Evidence                                                                                                                                    | Severity                                                                 |
| --- | --------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------ |
| W6  | **Worksheet bytes are buffered until `finalize()`**; no sink receives data during `appendRow`                                                             | 300k-row probe (`toFile` and `toWritable`): file 0 bytes until finalize, RSS grows with rows                                                | **Blocking** for 1M-row exports in a browser tab                         |
| W7  | **Unbounded shared string table joined into one string** at finalize                                                                                      | `RangeError: Invalid string length` in `serializeSharedStrings` at 1M×20 and 18M cells; 300k rows with two unique strings each → 439 MB RSS | **Blocking** (Salesforce Ids/names are unique per row)                   |
| W1  | Overlapping escape runs: literal `_x005F_x0041_` written as `_x005F_x005F_x0041_` (Excel decodes to `_x005FA`)                                            | `EC-XML-ESCAPE-OVERLAP`, Excel oracle F3                                                                                                    | Low                                                                      |
| W2  | Date conversion uses UTC fields; local-midnight Dates get the host offset as a time fraction, time-only values become negative serials (Excel `########`) | `EC-DATE-SERIAL-EPOCH-UTC-FIELDS`, Excel oracle M2/N2/O2                                                                                    | High unless adapted (ADR-003); a convention, not a bug, but undocumented |
| W3  | `setFormula` cached value cannot be an error; rich text needs the `{ kind: 'rich-text', runs }` shape (`setCellRichText` not exported)                    | golden sidecar                                                                                                                              | Low                                                                      |
| W4  | `styles.xml` `CT_Font` child order fails the Open XML SDK validator                                                                                       | validator run                                                                                                                               | Low (every surveyed writer; Excel tolerates)                             |
| W5  | Zip64 entries capped at 4 GiB                                                                                                                             | README                                                                                                                                      | Accepted for v1 (ADR-002)                                                |

## 4. Security and hostile inputs

| Input                                | office-kit                                                       | Assessment                                                  |
| ------------------------------------ | ---------------------------------------------------------------- | ----------------------------------------------------------- |
| DOCTYPE/XXE in sharedStrings         | rejected: "DTD declarations are not permitted in OOXML payloads" | Good (SheetJS accepts, leaves entities undecoded)           |
| Truncated archive                    | rejected: "no End-of-Central-Directory signature"                | Good                                                        |
| CRC mismatch                         | accepted (no CRC check)                                          | Acceptable; note                                            |
| Duplicate entry names                | accepted, last entry wins (SheetJS: first wins)                  | Should reject or warn                                       |
| 50k-deep nesting                     | rejected: "Maximum nested tags exceeded"                         | Good                                                        |
| 30 MB / 190:1 zip bomb               | accepted (6,000 rows parsed)                                     | Default `decompressionLimits` allow it; pass tighter limits |
| Encrypted (CFB)                      | rejected with a clear developer-facing message                   | Adapter maps to the UI contract                             |
| CSV / ODS / SpreadsheetML 2003 bytes | rejected with zip errors                                         | Sniff first (ADR-007)                                       |

## 5. Platform checks

- Runtime floor Node 22+ (Jetstream 24); browsers with Web Streams (all targets).
- Bundle: README claims ≤ 120 KB min+brotli (streaming subpath ~49 KB); the bench's Chrome bundle measures it.
- CSP: verify no code path uses fflate's async functions (`blob:` workers) and no `eval`/`Function` in
  fast-xml-parser/saxes.
- Determinism: `docProps` timestamps vary per run; a fixed-clock option is needed for golden-bytes tests.

## 6. Performance vs gates (from 06)

| Gate                                      | Result (office-kit vs SheetJS 0.20.3, Apple M4, Node 24)                       | Verdict                                    |
| ----------------------------------------- | ------------------------------------------------------------------------------ | ------------------------------------------ |
| write 100k×20 `mixed` time ≤ 1.0×         | 4.62 s vs 10.36 s = 0.45×                                                      | PASS                                       |
| write 100k×20 `mixed` peak memory ≤ 0.25× | 1,158 MB vs 2,899 MB = 0.40×                                                   | **FAIL** (W6/W7)                           |
| write 100k `strings-unique` memory        | 890 MB vs 3,408 MB = 0.26×                                                     | borderline (SST-bound)                     |
| write 100k `numeric` memory               | 54 MB vs 1,814 MB = 0.03×                                                      | PASS (no strings → no SST)                 |
| write 100k `wide` (107 cols, 10.7M cells) | 16.3 s / 1,728 MB; SheetJS died (SIGABRT)                                      | survives where SheetJS cannot, at 1.7 GB   |
| read-typed 100k×20 `mixed` time ≤ 1.0×    | 6.55 s vs 6.50 s = 1.01×                                                       | parity (marginal)                          |
| read-typed 100k×20 `mixed` memory ≤ 0.5×  | 522 MB vs 1,934 MB = 0.27×                                                     | PASS                                       |
| read-typed 100k `wide`                    | 63.5 s, **5,424 MB** footprint (SST resident); SheetJS timed out at 600 s      | both unusable at this shape                |
| write ceiling (`mixed`)                   | office-kit ok at 250k (11.7 s, 2,434 MB), fails at 500k; SheetJS fails at 250k | ceiling ≈ 2× SheetJS's, still far below 1M |
| 1M × 20 write (Node)                      | `RangeError: Invalid string length` (SST join)                                 | **FAIL**                                   |
| 18M cells write (Node)                    | `RangeError: Invalid string length` (SST join)                                 | **FAIL**                                   |
| 100k × 20 write in a Chrome module worker | 6.42 s, renderer RSS +1,618 MB (SheetJS 8.90 s, +2,875 MB)                     | works, but 1.6 GB in a tab                 |
| 1M × 20 write in a Chrome module worker   | renderer crash (office-kit reached 4.0 GB first; SheetJS too)                  | **FAIL**                                   |
| first byte < 100 ms                       | 2.8 ms Node / 10.6 ms Chrome (headers only; the sheet arrives at finalize)     | technically PASS, materially FAIL          |

Combined verdict from `bench/results/2026-09-11-macbook-air-combined`: **3 pass / 4 fail**. For comparison
ExcelJS's streaming writer passes the 100k memory gate (349 MB, 0.12×, 2.95 s) because it really streams
(archiver + data descriptors, first byte at 1.9 s) - but it is unmaintained, 4× larger, and reads slowly
(9.94 s, 1.5×). That contrast is the strongest argument that the missing piece is an engineering task, not a
research risk.

Two more friction items from the harness (`bench/engines/office-kit.mjs`):

- W8: a bare `Date` in `appendRow` is written as a serial **without a number format** (it displays as a
  number in Excel) unless wrapped as `{ value, style: { numberFormat } }` per cell.
- W9: `appendRow` never yields to the event loop and `toFile` only drains between turns, so a tight loop
  piles up deflated chunks; the adapter must yield every few hundred rows.

## 7. Project health

- 20 commits 2026-07-05 → 2026-09-03; 6 releases; one primary maintainer (baseballyama) plus two outside
  contributors; same-day merges for external PRs (#128, #129); a fix PR for a user-reported bug within hours
  (#131 → #132).
- ~2,100 vitest tests, 3-tier validator (OPC + XSD + semantic) in CI, openpyxl fixture corpus as a submodule,
  size-limit guards, Renovate, changesets, `SECURITY.md`, `THIRD_PARTY_NOTICES.md`.
- Risk: bus factor of one; mitigated by MIT + the fork-as-last-resort policy (ADR-006) and by this corpus and
  oracle as the upgrade gate.

## 8. Proposed upstream issues (each reproducible with a fixture in this repo)

0. **Reader rejects `<u val="none"/>`** (Apache POI / Salesforce report exports; fixtures `golden-sfdc-report-*`) -
   accept every `ST_UnderlineValues` member.
1. **Streaming writer buffers the worksheet and emits nothing to the sink until `finalize()`** (300k-row
   probe; README claims a fixed memory budget). Proposal: deflate and emit each worksheet entry incrementally
   with data descriptors; serialize `<si>` chunks incrementally.
2. **Shared string table is serialized with one `Array.join`** → `RangeError` at ~1M rows; add an
   inline-string mode or a bounded SST (`EC-*` and ADR-001 describe the policy).
3. Reader: decode `_xHHHH_` escapes (fixture `golden-canonical-sheetjs`, Excel oracle E2).
4. Reader: infer positions when `r` is absent (`edge-missing-r-attributes`).
5. Reader: accept Strict OOXML relationship types/namespaces (`edge-strict-namespaces`).
6. Reader: normalize backslashes in relationship targets (`edge-backslash-rel-targets`).
7. Writer: escape overlapping `_x` runs (`_x005F_x0041_` → `_x005F_x005F_x005F_x0041_`; Excel oracle F3).
8. Docs: state the UTC-field Date convention on `dateToExcel`/`excelToDate`/`appendRow`.
9. Reader: distinguish legacy `.xls` from encrypted packages in the CFB check.
10. Reader: keep `1.7976931348623157E+308`; stream-mode CDATA inline strings.
11. Export `setCellRichText`/`setCellFormula` from the worksheet subpath; `CT_Font` child order; fixed-clock option.

## 9. What "100% of our needs" would require

Items marked Adapter in §1 are Jetstream-side glue any engine needs. The blockers are W6 and W7 (scale), R11
(Salesforce report exports) and R1 (escapes); R2-R4 matter for files from other generators. If the maintainer accepts issues 1-3, adoption
needs no fork and the remaining items are adapter code covered by this corpus.
