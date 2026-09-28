---
id: contributing
title: Contributing
description: Repository layout, the commands, and how the fixture corpus, oracle and benchmark fit together.
---

# Contributing

The repository is [jetstreamapp/simple-excel](https://github.com/jetstreamapp/simple-excel). `AGENTS.md` at the
root is the short version of this page for anyone (or anything) editing the code; `research/11-build-plan.md` is
the contract every module is built to, and is worth reading before changing anything in `src/`.

## Layout

| Path                | What is in it                                                                                                                                                                          |
| ------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `src/`              | The library. Browser-first: no Node globals, no `node:` imports, no DOM beyond `TextEncoder`/`TextDecoder`, the compression streams and a duck-typed `Blob`. Zero runtime dependencies |
| `src/node/`         | The `/node` entry — the only place `node:` imports are allowed                                                                                                                         |
| `src/**/__tests__/` | Unit tests, next to the module they cover                                                                                                                                              |
| `test/`             | Integration suites: corpus reads, golden bytes, round trips, hostile inputs, SheetJS parity, memory                                                                                    |
| `fixtures/`         | The fixture corpus and `manifest.json` (sha256 and provenance for every file)                                                                                                          |
| `oracle/`           | The compatibility oracle: SheetJS, office-kit, openpyxl, calamine, the Open XML SDK validator, LibreOffice, Excel                                                                      |
| `bench/`            | The benchmark harness; this engine's adapters are `bench/engines/simple-excel.mjs` and `simple-excel-zlib.mjs`                                                                         |
| `research/`         | Format primer, library landscape, edge-case catalog, compatibility matrix, performance baseline, reference architecture, ADRs, build plan                                              |
| `docs/`             | This site — a Docusaurus project with its own `package-lock.json`                                                                                                                      |

Inside `src/`, the layering is strict: `zip/` and `compress/` know nothing about spreadsheets, `xml/` knows
nothing about SpreadsheetML, `sml/` owns the format rules, and `write/` and `read/` are thin facades over them.

## Commands

```bash
npm test                    # unit + integration (vitest)
npm run test:corpus         # corpus, golden-bytes, hostile and parity suites only
npm run typecheck
npm run lint
npm run format              # always run this after editing a source file
npm run build               # esbuild bundles (esm + cjs, core + node entry) and declarations
npm run fixtures:check      # fails on manifest sha256 drift
npm run smoke:browsers      # the same page in Chromium, Firefox and WebKit
npm run build && npm run bench -- --engines sheetjs,simple-excel --sizes 1k,10k,100k
npm run oracle -- --tag kind:golden --label <label>   # local only
npm run research:regenerate # re-renders research 04, 05 and 06 from the committed JSON
```

`npm run oracle` needs tooling that is not on CI: a Python virtual environment for openpyxl and calamine,
LibreOffice, and Microsoft Excel with Automation permission granted. It is a local pre-release gate; its results
folder is committed with each run.

The docs site is a separate npm project:

```bash
cd docs && npm ci && npm run build
```

## Rules that are easy to get wrong

- **The catalog is the arbiter.** Escaping, date and container rules live in `research/02-format-primer.md` and
  `research/04-edge-case-catalog.md`, and most entries were verified against Excel itself. When a test disagrees
  with your intuition, the catalog wins.
- **Streaming is the design, not an option.** A writer pushes bytes to its sink as rows arrive and never holds a
  worksheet or an unbounded shared-string table. A reader pulls rows through an async iterator.
- **Errors are classified.** Throw `XlsxError` with a `code`, never a bare `Error`. The encrypted-file message
  must contain the words `password-protected`.
- **Hot-path discipline.** No per-cell object allocation, no regex per cell, no string concatenation past the
  chunk boundary. Run `npm run bench` before and after any change to the writer or reader core.
- **Deterministic output.** With `deterministic: true` the bytes are reproducible; the golden-bytes suite depends
  on it.
- Every exported symbol needs an explicit type (`isolatedDeclarations`), and there is no `any` in `src/`.

## Tests

Unit tests sit beside their module and are named for the catalog entry they prove (`EC-...`), so a failure points
straight at the rule it broke. The integration suites in `test/` are manifest-driven: every fixture has a
recorded verdict, and a reader change that moves one fails the build.

Golden bytes are sha256 pins of the writer's deterministic output. Updating them is deliberate —
`UPDATE_GOLDENS=1`, and only after a validator and oracle pass.

## Adding a fixture

A fixture is a real file plus its provenance. Put the file under `fixtures/` in the folder that matches its kind
(`golden/`, `edge/`, `hostile/`, `jetstream/`) and register it:

```bash
node fixtures/register.mjs golden/my-app/canonical.xlsx \
  --id golden-canonical-my-app \
  --generator "MyApp 1.2.3" \
  --provenance "script:generators/canonical-my-app.mjs" \
  --license MIT \
  --tags kind:golden,generator:my-app \
  --expected canonical/canonical.json
```

| Flag               | Meaning                                                                              |
| ------------------ | ------------------------------------------------------------------------------------ |
| `--id`             | Stable identifier; re-registering the same id replaces the entry                     |
| `--generator`      | Name and version of whatever produced the file                                       |
| `--provenance`     | Where it came from: `script:generators/x.mjs`, an application and date, a bug report |
| `--license`        | SPDX licence of the file; default `MIT`                                              |
| `--tags`           | Comma-separated, e.g. `kind:hostile`, `generator:excel`, `policy:truncate-32767`     |
| `--expected`       | Ground-truth dump to compare reads against (usually `canonical/canonical.json`)      |
| `--expected-error` | For a hostile fixture: the `XlsxErrorCode` reading it must produce                   |
| `--notes`          | Anything a future reader needs to know                                               |

Size and sha256 are always recomputed from disk, so `npm run fixtures:check` will tell you if the file later
drifts.

Then run `npm run test:corpus`. A new fixture with an `--expected` dump has to read at its recorded verdict, and
a new `kind:hostile` fixture has to produce its `--expected-error` classified error rather than a crash.

If the fixture demonstrates a behaviour that is not already in the catalog, add an entry to
`fixtures/edge-cases.json` and run `npm run research:regenerate` — `research/04` is rendered from that file, not
edited by hand. The same goes for `research/05` (rendered from the newest oracle run of each label) and
`research/06` (rendered from the newest `*-combined` benchmark run).

## Browser smoke test

Node covers the format rules; it cannot tell you whether the library works in the browsers the package claims to
support. `npm run smoke:browsers` builds `dist/`, serves the repository root cross-origin isolated, and runs one
page — `test/browser/smoke.html` — in Playwright's Chromium, Firefox and WebKit builds, printing a table per
browser and exiting non-zero on any failed check.

```bash
npm run smoke:browsers                               # build + all three browsers
node test/browser/run.mjs --browsers firefox,webkit  # a subset, against the dist/ that is already built
node test/browser/run.mjs --serve                    # keep serving; open the URL by hand
npx playwright install firefox webkit                # Firefox and WebKit are not installed by default
```

The page runs eight steps and evaluates every expectation in the page, against the values it wrote itself, so the
runner only ever sees pass/fail plus a readable detail:

| Step                 | What it proves                                                                                                                                                                                |
| -------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Feature detection    | `hasNativeDeflate()`, `CompressionStream`, `DecompressionStream`, `WritableStream`, `Blob`                                                                                                    |
| Write (main thread)  | 20,000 mixed rows to `collectToBlob()` — bold header, frozen row, autofilter, a date number format, a hidden second sheet, `deterministic: true`                                              |
| Read back            | `openWorkbook(blob)` (the `slice().arrayBuffer()` source path), then every sentinel cell: unicode, CRLF, a control character, an `_x0041_` literal, `-0`, `1e21`, dates, truncation at 32,767 |
| Module worker        | the same write in a `type: 'module'` worker, the Blob transferred back and read on the page, plus one read of an `ArrayBuffer` transferred in                                                 |
| Real fixtures        | an Excel-authored and a Google-Sheets-authored golden, fetched and read                                                                                                                       |
| Hostile input        | `ENCRYPTED` for the password-protected file, `ZIP_BOMB` for the 30 MB sheet against an 8 MiB inflate cap                                                                                      |
| `fromWritableStream` | bytes streamed into a `WritableStream` equal a `collectToBytes()` run of the same workbook                                                                                                    |
| Memory               | 200,000 rows written in the worker while the page samples `performance.measureUserAgentSpecificMemory()` (Chromium only; recorded, never asserted)                                            |

The files are `test/browser/smoke.html`, `smoke.js` (the page), `smoke.worker.js`, `smoke-workbook.js` (the
dataset and the read-back checks, shared by both threads) and `run.mjs` (the static server and the Playwright
driver). They are plain ESM served over HTTP and import `/dist/esm/index.mjs` by URL, so there is no bundler in
the loop and the page is exactly what a browser would load.

Two things the runner has to do deliberately: Chromium is launched with `channel: 'chromium'` because
Playwright's default headless shell rejects `measureUserAgentSpecificMemory()` with a `SecurityError`, and the
server sends `Cross-Origin-Opener-Policy`/`Cross-Origin-Embedder-Policy` so the page is cross-origin isolated
(the same reason `bench/lib/chrome.mjs` does).

### Checking real Safari

Playwright's WebKit is not Safari. For the real thing:

```bash
node test/browser/run.mjs --serve
```

then open the printed URL in Safari. The page renders the same table it hands the runner, and
`window.__smokeResults` in the Web Inspector console holds the full object. Record the Safari version and the
outcome in the "Browsers" section of `research/05-compatibility-matrix.md`.

## Benchmarks

`npm run bench` runs each engine in a fresh Node process with `--expose-gc`, one warm-up plus three timed runs,
reporting median wall time, peak RSS above baseline, output size and first-byte latency. The datasets are seeded
and shaped like real exports: `mixed` (the Salesforce record shape), `wide` (107 columns), `strings-unique` (the
shared-string worst case) and `numeric` (the XML-generation floor).

Add an engine by writing an adapter in `bench/engines/`; this library's are `bench/engines/simple-excel.mjs`
(platform `CompressionStream`) and `bench/engines/simple-excel-zlib.mjs` (`nodeDeflater(1)`), which share
`bench/engines/_simple-excel.mjs` and import `dist/`, so `npm run build` has to have run.

## Commits and releases

Conventional commit messages. `CHANGELOG.md`'s `[Unreleased]` headings decide the version bump —
`### Breaking Changes` major, `### Added`/`### Deprecated` minor, `### Changed`/`### Removed`/`### Fixed`/
`### Security` patch — so keep it accurate as you go. Releases are cut from `main` by CI.

A pre-commit hook blocks commits that fail `npm run format:check` or `npm run lint`.
