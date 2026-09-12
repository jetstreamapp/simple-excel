# xlsx engine benchmark

Reproducible engine comparison behind the xlsx engine replacement. SheetJS CE (what Jetstream ships)
fails at scale — `RangeError: Invalid string length` at 958k × 10 cells, `Invalid array length` at
18M cells — so this harness measures every candidate on the same seeded datasets, in fresh processes,
against the exact SheetJS call Jetstream makes today.

```bash
npm run bench -- -- --help
npm run bench -- -- --engines sheetjs,exceljs,office-kit,write-excel-file --ops write,read-typed,read-raw --datasets mixed --sizes 1k,10k,100k --label baseline
npm run bench -- -- --datasets mixed --sizes 1m,18m-cells --engines sheetjs,office-kit --ops write --runs 1 --label scale
npm run bench -- -- --datasets strings-unique,numeric,wide --sizes 100k --engines sheetjs,office-kit --ops write,read-typed --label shapes
npm run bench -- -- --merge results/<a>,results/<b> --label combined   # one summary + gates over several runs
```

Everything here is pure (`node check-purity.mjs`): no `@jetstream/*` imports, so the
folder can lift into a standalone repository. Engines resolve through pnpm's walk-up: `xlsx` from the
repo root, the rest from `node_modules`.

## Layout

| Path                      | What                                                                                                      |
| ------------------------- | --------------------------------------------------------------------------------------------------------- |
| `run.mjs`                 | CLI. Builds the cell plan, forks one child per cell, writes `results/<date>-<host>-<label>/`.             |
| `lib/child.mjs`           | Runs ONE (engine, op, dataset, size) cell — or one fixture build — in a fresh process, prints JSON.       |
| `lib/dataset.mjs`         | Seeded xorshift PRNG + streaming row generators (`createRowIterator`, `materialize`, `resolveShape`).     |
| `lib/metrics.mjs`         | Ratios, gates, `summary.md` rendering.                                                                    |
| `lib/excel-limits.mjs`    | 32,767-char truncation (Jetstream's rule, applied by every writer) and the shared date format.            |
| `lib/engines.mjs`         | Engine registry (`engines/*.mjs`).                                                                        |
| `lib/chrome.mjs` + `web/` | Optional Chromium step (write in a module Worker, memory via `measureUserAgentSpecificMemory`).           |
| `engines/*.mjs`           | One adapter per engine, all exposing the same interface.                                                  |
| `results/`                | Committed results, one folder per run: `results.json` (machine info, options, every cell) + `summary.md`. |
| `../.generated/bench/`    | Gitignored scratch: read fixtures, write outputs, logs, browser bundle.                                   |

## Datasets (`lib/dataset.mjs`)

All rows come from a seeded xorshift32 PRNG (`--seed`, default 20260911): the same `{ dataset, rows, seed }`
yields byte-identical rows for every engine. Row 1 is always the header (column names).

| Dataset          | Columns | Shape                                                                                                                                                                                                                                                                                                        |
| ---------------- | ------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `mixed`          | 20      | Salesforce export: 18-char Ids, unicode names, integers + decimals, booleans, JS `Date` objects (`CreatedDate`, `CloseDate`), ISO datetime strings, multi-line `Description`, JSON subquery blob `Contacts__r` (1% of blobs > 32,767 chars), 10% nulls in every column except `Id`. ~1.7 KB of text per row. |
| `wide`           | 107     | 60 strings (5–30 random chars), 30 numbers (half integer, half 4-place decimals), 10 `Date` objects, 7 booleans. No nulls.                                                                                                                                                                                   |
| `strings-unique` | 20      | Globally unique 24-char strings (row/column prefix + random tail). Shared-string-table worst case: nothing dedupes.                                                                                                                                                                                          |
| `numeric`        | 20      | 10 integers (0–1e6) + 10 decimals (4 places). No strings at all.                                                                                                                                                                                                                                             |

Sizes: `1k`, `10k`, `100k`, `1m` rows. Presets that pin the dataset: `18m-cells` = mixed × 900,000 rows
(the historical failure shape) and `wide-100k` = wide × 100,000 (10.7M cells).

## Engines and the adapter interface

Each `engines/<name>.mjs` exports:

```js
export const name = 'my-engine';
export const version = '1.2.3';
export const supportsStreaming = true; // informational
export const readInput = 'path'; // or 'bytes' — what readTyped/readRaw receive
export async function write(rowIterable, { columns, rowCount, outPath }) {
  return { bytes, firstByteMs }; // firstByteMs optional (ms from write() start until the sink got its first chunk)
}
export async function readTyped(pathOrBytes) {
  return { rows, cells }; // dates as Date, numbers as numbers
}
export async function readRaw(pathOrBytes) {
  return { rows, cells }; // every value as a string
}
// export const readRaw = null;  -> op reported as "skipped"
// export const skipReason = '...'; -> engine reported as "skipped"
// export async function load() { return { ...interface } } -> resolve the interface lazily (see ours.mjs)
```

| Engine             | Version notes                                                                                                                                                                                                                                                                                                      |
| ------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `sheetjs`          | Exactly Jetstream's calls: `aoa_to_sheet(aoa, { dense: true })`, `XLSX.write(wb, { bookType: 'xlsx', bookSST: false, type: 'array', compression: rows > 10_000 })`; reads `XLSX.read(buf, { cellText: false, cellDates: true, type: 'array' })` + `sheet_to_json` with Jetstream's `dateNF`/`defval`/`rawNumbers`. |
| `exceljs`          | `stream.xlsx.WorkbookWriter` (`useStyles: false`, `useSharedStrings: false` → inline strings; dates still get the built-in `mm-dd-yy` style) and `WorkbookReader` (`styles: 'cache'` converts dates for read-typed).                                                                                               |
| `office-kit`       | `createWriteOnlyWorkbook(toFile(path))` + `appendRow` per row, `loadWorkbookStream(fromFile(path))` + `iterRows`/`iterValues`. See the friction notes at the top of `engines/office-kit.mjs`.                                                                                                                      |
| `write-excel-file` | Whole-array `writeXlsxFile(data, { filePath, dateFormat })`; reads via `read-excel-file/node` (typed only, no raw mode).                                                                                                                                                                                           |
| `ours`             | Stub: set `XLSX_ENGINE_OURS=/path/to/module.mjs` (same interface) to benchmark the in-house engine; skipped otherwise.                                                                                                                                                                                             |

Every writer applies `lib/excel-limits.mjs` truncation (`slice(0, 32_767 - '...(truncated)'.length) + '...(truncated)'`)
so all engines write the same cell values.

To add an engine: create `engines/<name>.mjs` with the interface above, install its dependency in
`package.json` (never the repo root), run a 1k smoke cell, then add a row to the table above.

## What a cell measures (`lib/child.mjs`)

Each cell runs in its own `node --expose-gc --max-old-space-size=12288` process (`--node-flags` to change):

1. Load the engine. For `write`, materialise the dataset in memory (`--source materialized`, default — rows
   are already in memory in Jetstream) or hand the engine a fresh generator (`--source stream`). For reads,
   locate the fixture; engines with `readInput: 'bytes'` get the file read into a `Uint8Array` up front.
2. `gc()`, settle, take the **pre-warm-up RSS** (engine loaded, input resident, nothing run yet).
3. One warm-up run (`--warmup`). A warm-up failure is the cell's error.
4. `gc()`, settle, take the **post-warm-up RSS**.
5. `--runs` timed runs (default 3; 1 for ≥ 1M rows), `gc()` between them.

| Metric                   | Definition                                                                                                                                                                                                                                                                                                                                                                                                                                     |
| ------------------------ | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `median`/`min`/`max`     | Wall time of the timed runs (`performance.now()`), including writing to disk for writers.                                                                                                                                                                                                                                                                                                                                                      |
| `RSS footprint`          | Peak RSS while the op ran − pre-warm-up RSS. Peak = max of: RSS polled every 25 ms from an unref'd `setInterval` during timed runs, RSS read right after each run, and `process.resourceUsage().maxRSS` (the process high-water mark, only if it rose after the pre-warm-up snapshot). The poller cannot fire while a synchronous engine (SheetJS) blocks the event loop; the high-water mark covers that case. This is the ratio/gate metric. |
| `post-warm-up delta`     | Same peak − post-warm-up RSS (the literal "baseline after warm-up + gc()" definition). Undercounts engines whose warm-up footprint stays resident, because V8 rarely returns pages to the OS; kept as a secondary column.                                                                                                                                                                                                                      |
| `peak RSS`               | Absolute high-water RSS of the child (input data included).                                                                                                                                                                                                                                                                                                                                                                                    |
| `first byte`             | ms from the start of `write()` until the sink received its first chunk (streaming writers only; `—` otherwise).                                                                                                                                                                                                                                                                                                                                |
| `output`                 | Bytes written.                                                                                                                                                                                                                                                                                                                                                                                                                                 |
| `rows read`/`cells read` | Data rows (header excluded) × columns seen by the reader.                                                                                                                                                                                                                                                                                                                                                                                      |
| `time/mem vs sheetjs`    | Candidate ÷ sheetjs for the same op/dataset/size. Lower is better.                                                                                                                                                                                                                                                                                                                                                                             |

Statuses: `ok`, `ERROR (phase) Name: message` (the exception, never rethrown), `TIMEOUT` (`--timeout`,
default 600 s; the child is SIGKILLed), `skipped` (engine/op unsupported), `fixture failed`. A V8 fatal
OOM (no result line, "heap out of memory" on stderr) is reported as `FatalHeapOutOfMemory`.

### Read fixtures

`read-typed`/`read-raw` cells read `.generated/bench/fixtures/<dataset>-<size>.xlsx`, produced once and
reused (`--rebuild-fixtures` to regenerate). Files up to 2.5M cells are written by **sheetjs** (the files
Jetstream produces today); larger ones (`1m`, `18m-cells`, `wide-100k`) by **office-kit** because
SheetJS cannot produce them. `results.json` records the producer per cell; `summary.md` names it above
each read table.

## Gates (`lib/metrics.mjs`)

Evaluated for every non-sheetjs engine present in the run (engines whose cells were all skipped are
ignored). A gate is `NOT-RUN` when its cell — or the sheetjs cell it compares against — is missing or
skipped; the verdict is `PASS` / `FAIL` / `INCOMPLETE` accordingly. Use `--merge` to evaluate gates across
the baseline + scale runs together.

| Gate                                    | Cell            |
| --------------------------------------- | --------------- |
| write RSS footprint ≤ 0.25× sheetjs     | mixed 100k      |
| write median time ≤ 1.0× sheetjs        | mixed 100k      |
| read-typed median time ≤ 1.0× sheetjs   | mixed 100k      |
| read-typed RSS footprint ≤ 0.5× sheetjs | mixed 100k      |
| write 1M × 20 succeeds                  | mixed 1m        |
| write 900k × 20 (18M cells) succeeds    | mixed 18m-cells |
| first byte reaches the sink in < 100 ms | mixed 100k      |

## Browser step (`lib/chrome.mjs`)

```bash
node bench/lib/chrome.mjs --engines office-kit,sheetjs --datasets mixed --sizes 100k,1m
```

Bundles `web/worker.js` (dataset generator + office-kit streaming + SheetJS) with esbuild into
`.generated/bench/web/worker.bundle.js`, serves `web/` on 127.0.0.1 with COOP/COEP headers, launches
Playwright's Chromium (`npx playwright install chromium` if missing) and runs each cell in a fresh
page + module Worker. Rows stream from the seeded generator (office-kit consumes it row by row; SheetJS
materialises the array-of-arrays itself, as part of its cost — a resident 1M-row mixed dataset is ~4.5 GB,
more than a renderer can hold). Output goes to a counting null sink. Memory is sampled every `--sample-ms` (100):
`performance.measureUserAgentSpecificMemory()` in the page (covers the worker; `worker peak` = the
`DedicatedWorkerGlobalScope` attribution), renderer process RSS via `ps`, and CDP `Performance.getMetrics`
JSHeapUsedSize (page isolate only). Renderer crashes (`page.on('crash')`) and worker `error` events are
recorded as failures. Results land in `results/<date>-<host>-<label>/chrome.{json,md}`.

Known limitation / TODO: in headless Chromium 151 `measureUserAgentSpecificMemory()` never resolved
inside the 6–9 s runs (it waits for a GC Chrome schedules on its own timetable, even when called on a
concurrent chain with a grace period), so the `agent-cluster peak` / `worker peak` columns are empty and
`renderer RSS delta` (process RSS via `ps`) is the number to read. Left to try: `--headed`, longer runs,
`--js-flags=--expose-gc` with a forced GC, or attaching a CDP session to the worker target for its own
`JSHeapUsedSize`.

## Latest results (2026-09-11, Apple M4 / 32 GB / Node 24.18)

| Folder                                    | What                                                                                             |
| ----------------------------------------- | ------------------------------------------------------------------------------------------------ |
| `results/2026-09-11-macbook-air-baseline` | mixed 1k/10k/100k × sheetjs, exceljs, office-kit, write-excel-file × write, read-typed, read-raw |
| `results/2026-09-11-macbook-air-scale`    | mixed 1m + 18m-cells writes (sheetjs, office-kit) — all four `RangeError: Invalid string length` |
| `results/2026-09-11-macbook-air-ceiling`  | mixed 250k/500k writes: office-kit ok at 250k, fails at 500k; sheetjs fails at 250k              |
| `results/2026-09-11-macbook-air-shapes`   | strings-unique / numeric / wide at 100k (sheetjs, office-kit; write + read-typed)                |
| `results/2026-09-11-macbook-air-combined` | `--merge` of the four above: one summary, gates evaluated across all of them                     |
| `results/2026-09-11-macbook-air-chrome`   | Chromium step (`chrome.md` / `chrome.json`): mixed 100k + 1m writes in a module Worker           |

Headline: office-kit at mixed 100k is 0.45× sheetjs write time with a 2.8 ms first byte and 0.40× the
RSS footprint (gate asks ≤ 0.25×), read-typed at parity (1.01×) with 0.27× the memory; it survives
`wide` 100k (10.7M cells) where sheetjs dies, but its shared-string table caps mixed data between 250k
and 500k rows, so the 1M and 18M-cell gates fail. Verdict per engine is in `combined/summary.md`.

## Caveats

- Numbers are from one machine with other processes running; `results.json` records the load average at
  start and end. Compare ratios across runs, not absolute ms.
- `post-warm-up delta` includes V8 heap slack: a streaming engine that fits in the pages the warm-up already
  reserved legitimately reports ~0 MB, which is why ratios use the pre-warm-up `RSS footprint` instead.
- `--source materialized` (default) keeps the whole dataset resident in the baseline (≈ 450 MB per 100k
  mixed rows), which is the Jetstream situation but means 900k rows start from a ~4 GB baseline.
- Timezones: SheetJS writes `Date` cells as **local wall-clock** serials; office-kit, exceljs and
  write-excel-file write **UTC wall-clock**. The files differ by the machine's UTC offset. A replacement
  engine needs Jetstream to shift dates (or the engine to accept a "local" epoch) to keep today's output.
- SheetJS output below 10k rows is uncompressed (Jetstream only enables compression above 10k), so its
  `output` size is ~5× the others at `1k`.
- `read-raw` has no formatted-text mode outside SheetJS: exceljs/office-kit stringify values (dates stay
  serial numbers).
- office-kit's write-only workbook keeps its shared-string table resident until `finalize()` and then
  serialises it as **one JS string** (`serializeSharedStrings` → `Array.join`). With Jetstream-shaped rows
  (~1.7 KB of text each) that string passes V8's ~2^29-char limit somewhere below 1M rows, so `1m` and
  `18m-cells` fail with `RangeError: Invalid string length` in `finalize()` — the same exception SheetJS
  throws, just later. There is no inline-string option in the streaming API (0.11.0); see the `ceiling`
  results for where the limit sits and `engines/office-kit.mjs` for the friction notes.
- exceljs's `WorkbookReader` (`row.values`) and office-kit's `iterRows` are sparse; the adapters rebuild
  dense rows so `cells read` is comparable to `sheet_to_json` with `defval`.
