# 00 - Feasibility verdict

Is a from-scratch, streaming, styles-capable xlsx engine that replaces SheetJS in Jetstream a realistic
goal, and should we build it or adopt `@office-kit/xlsx`?

_Accurate as of 2026-09-12 (branch `chore/xlsx-hand-roll`)._

_Superseded in part: the ADR-008 addendum chose to build the library; the result is `@jetstreamapp/simple-excel`,
released 0.1.0 (2026-09-12) and 0.1.1 (2026-09-26). The decision and status below are the feasibility-phase
snapshot; the counts are updated to 2026-09-27._

## Verdict: GO - realistic, and narrower than it looks

**Building it is realistic.** The expensive parts of a spreadsheet library are the ones Jetstream never
needs: round-trip editing of existing workbooks, formula evaluation, charts, pivots, number-format rendering.
Jetstream only (a) writes new workbooks from rows it already holds and (b) reads rows and a few fixed cells
from workbooks users upload (01). With that scope the engine is a streaming zip writer, an XML tokenizer, a
row writer/reader, a small styles registry and a date/escape layer - roughly 4,600-5,200 lines (07 §11), with
the platform providing deflate on every target (ADR-005). No wasm, no CSP change, no dependency.

**The requirements are now written down and testable.** The corpus (58 fixtures, 95 catalogued edge cases,
04), the compatibility oracle across eight readers including Excel itself (05) and the benchmark harness (06)
exist and are committed; they are the acceptance suite for any engine, ours or someone else's.

**The engine decision is dual-track (ADR-008).** `@office-kit/xlsx` has the right thesis and a responsive
maintainer, but its streaming writer buffers the whole worksheet and an unbounded shared string table, so it
fails at exactly the sizes that broke SheetJS, and its reader rejects Salesforce report exports outright (10). Those three missing pieces are Phase B of the hand-rolled
design. We open the upstream issues with our fixtures and build Phase A-B ourselves in parallel; at the
four-week mark we ship whichever combination passes the gates.

## What the research changed

| Assumption going in                                 | What the evidence says                                                                                                                                                                                |
| --------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| "Streaming xlsx writers are a solved problem in JS" | ExcelJS streams (and passes the memory gate) but is unmaintained; office-kit and write-excel-file buffer until close; SheetJS materializes everything. None is both maintained and streaming.         |
| "Dates are just serial numbers"                     | Three conventions collide (local fields, UTC fields, naive); SheetJS's own reader hands Jetstream local wall-clock Dates only through `sheet_to_json`; office-kit needs a boundary adapter (ADR-003). |
| "Escaping is `&amp;` and friends"                   | `_xHHHH_` is the real trap: control chars, CR, literal `_x0041_` and overlapping runs are all mishandled by at least one major library, and Excel is the arbiter (Excel oracle E2/F3/F7).             |
| "The validator tells us if Excel will open it"      | Every surveyed writer fails the Open XML SDK validator (font element order) and Excel opens all of them; Excel silently repairs formula-injected files and only leaves a recovery log.                |
| "Legacy formats can be dropped"                     | Confirmed, but the _sniffing_ must be explicit: today SheetJS silently parses CSV, ODS and BIFF8 bytes that reach it (ADR-007).                                                                       |

## Gates for "done" on the engine (whichever path)

1. Every golden and edge fixture reads identically to SheetJS-with-Jetstream-options after the adapter
   (05: PASS or documented policy).
2. Every hostile fixture is rejected with a classified error; `password-protected` reaches the UI.
3. Writer output passes the Excel oracle (no recovery log), LibreOffice, openpyxl and calamine.
4. Benchmark gates (06): write 100k×20 ≤ 1.0× time and ≤ 0.25× memory vs SheetJS; 1M×20 in a Chrome worker;
   18M cells in Node; read-typed ≤ 1.0× time, ≤ 0.5× memory.
5. Runs in a module worker under the web, desktop and MV3 CSPs without wasm or `blob:` scripts.

## Scope guard

Out of scope for v1 (07 §9): formulas, number-format rendering, charts, images, comments, data validation,
conditional formatting, workbook editing, encrypted files, `.xls`/`.xlsb`/`.ods`. Data-validation lists for
load templates and hyperlinks to records are the first v2 candidates.

## Where things stand

| Area                                                                                                                                                                                                       | Status                                                                                 |
| ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------- |
| Usage inventory (01), format primer (02), library landscape (03), reference architecture (07), migration plan (08), risks (09), ADRs                                                                       | written                                                                                |
| Fixtures: 28 goldens from 13 generators (incl. Excel 365 desktop, Excel for the web, Google Sheets, Numbers, Salesforce/POI), 15 edge, 12 hostile, 3 Jetstream assets; manifest with sha256 and provenance | committed                                                                              |
| Oracle: SheetJS, office-kit (document + stream), openpyxl, calamine, Open XML SDK validator, LibreOffice, Excel (silent-repair detection)                                                                  | committed with 15 result runs                                                          |
| Benchmark: sheetjs, exceljs, office-kit, write-excel-file; Node runs (baseline, scale, ceiling, shapes, combined) and a Chrome worker run committed                                                        | office-kit gate verdict 3 pass / 4 fail; renderer-memory API TODO in `bench/README.md` |
| office-kit evaluation (10) and ADR-008                                                                                                                                                                     | decided: hand-roll (ADR-008 addendum); shipped as `@jetstreamapp/simple-excel` 0.1.0   |
| Jetstream code changes                                                                                                                                                                                     | none (hardening candidates documented only, 08)                                        |
