# xlsx engine benchmark

Reproducible engine comparison behind the xlsx engine replacement. SheetJS CE (what Jetstream ships)
fails at scale — `RangeError: Invalid string length` at 958k × 10 cells, `Invalid array length` at
18M cells — so this harness measures every candidate on the same seeded datasets, in fresh processes,
against the exact SheetJS call Jetstream makes today.

```bash
npm run build                                # required once: the simple-excel adapters import dist/
npm run bench -- --help
npm run bench -- --engines sheetjs,exceljs,office-kit,write-excel-file,simple-excel,simple-excel-zlib --ops write,read-typed,read-raw --datasets mixed --sizes 1k,10k,100k --label baseline
npm run bench -- --datasets mixed --sizes 1m,18m-cells --engines sheetjs,simple-excel --ops write --runs 1 --label scale
npm run bench -- --datasets strings-unique,numeric,wide --sizes 100k --engines sheetjs,simple-excel --ops write,read-typed --label shapes
npm run bench -- --merge results/<a>,results/<b> --label combined   # one summary + gates over several runs
```

`npm run bench` deliberately does not build first (a `--merge` needs no build, and a stale rebuild in the
middle of a measurement session is worse than an explicit step): the `simple-excel` adapters import
`dist/esm/index.mjs` and `dist/esm/node.mjs`, and report `skipped: dist/ is missing` when they are absent.

The tooling stays free of `@jetstream/*` imports (`node scripts/check-purity.mjs`, run in CI). Engines resolve from
the root `node_modules`.

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
| `engines/*.mjs`           | One adapter per engine, all exposing the same interface (`_*.mjs` is shared adapter code, not an engine). |
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
// export async function load() { return { ...interface } } -> resolve the interface lazily (see simple-excel.mjs)
```

| Engine              | Version notes                                                                                                                                                                                                                                                                                                                  |
| ------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `sheetjs`           | Exactly Jetstream's calls: `aoa_to_sheet(aoa, { dense: true })`, `XLSX.write(wb, { bookType: 'xlsx', bookSST: false, type: 'array', compression: rows > 10_000 })`; reads `XLSX.read(buf, { cellText: false, cellDates: true, type: 'array' })` + `sheet_to_json` with Jetstream's `dateNF`/`defval`/`rawNumbers`.             |
| `exceljs`           | `stream.xlsx.WorkbookWriter` (`useStyles: false`, `useSharedStrings: false` → inline strings; dates still get the built-in `mm-dd-yy` style) and `WorkbookReader` (`styles: 'cache'` converts dates for read-typed).                                                                                                           |
| `office-kit`        | `createWriteOnlyWorkbook(toFile(path))` + `appendRow` per row, `loadWorkbookStream(fromFile(path))` + `iterRows`/`iterValues`. See the friction notes at the top of `engines/office-kit.mjs`.                                                                                                                                  |
| `write-excel-file`  | Whole-array `writeXlsxFile(data, { filePath, dateFormat })`; reads via `read-excel-file/node` (typed only, no raw mode).                                                                                                                                                                                                       |
| `simple-excel`      | `@jetstreamapp/simple-excel` from `dist/esm/`: `createWorkbookWriter(toFile(path))` streaming through the platform `CompressionStream('deflate-raw')` (the browser path), `writeRows` over the row iterable, `header` and `rowCount` declared; reads `openWorkbook(bytes)` + `sheet(0).toObjects()` (typed) or `rows()` (raw). |
| `simple-excel-zlib` | The same adapter with the node entry's `nodeDeflater(1)`. The pair isolates what the compressor costs.                                                                                                                                                                                                                         |

Every writer applies `lib/excel-limits.mjs` truncation (`slice(0, 32_767 - '...(truncated)'.length) + '...(truncated)'`)
so all engines write the same cell values. `simple-excel` is the exception in form only: its writer applies the
identical rule and suffix internally (`cellOverflow: 'truncate'`), so the adapter does not pre-scan rows and the
bytes still match — the cost stays inside the engine, where a caller would actually pay it.

To add an engine: create `engines/<name>.mjs` with the interface above, add its dependency to the root
`package.json` `devDependencies`, run a 1k smoke cell, then add a row to the table above.

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
node bench/lib/chrome.mjs --engines simple-excel,sheetjs --datasets mixed --sizes 100k,1m
```

Bundles `web/worker.js` (dataset generator + simple-excel from `dist/esm/index.mjs` + office-kit streaming +
SheetJS) with esbuild into `.generated/bench/web/worker.bundle.js`, serves `web/` on 127.0.0.1 with COOP/COEP
headers, launches Playwright's Chromium (`npx playwright install chromium` if missing) and runs each cell in a
fresh page + module Worker. Rows stream from the seeded generator (simple-excel and office-kit consume it row by
row; SheetJS materialises the array-of-arrays itself, as part of its cost — a resident 1M-row mixed dataset is
~4.5 GB, more than a renderer can hold). office-kit's output goes to a counting null sink; **simple-excel writes
to `collectToBlob()`**, the real browser download path, so its numbers include keeping the finished file. Memory is sampled every `--sample-ms` (100):
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

## Latest results (2026-09-12, Apple M4 / 32 GB / Node v24.18.0)

| Folder                                             | What                                                                                                                                                                        |
| -------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `results/2026-09-12-macbook-air-phase-e-optimized` | mixed 1k/10k/100k × simple-excel, simple-excel-zlib, sheetjs × write, read-typed, read-raw                                                                                  |
| `results/2026-09-12-macbook-air-phase-e-others`    | mixed 100k × exceljs, office-kit, write-excel-file, re-measured in the same session                                                                                         |
| `results/2026-09-12-macbook-air-phase-e-scale`     | mixed 1m + 18m-cells writes (sheetjs, simple-excel, simple-excel-zlib) — sheetjs `RangeError` on both                                                                       |
| `results/2026-09-12-macbook-air-phase-e-combined`  | `--merge` of the three above: one summary, gates across all of them. This is what `research/06` renders                                                                     |
| `results/2026-09-12-macbook-air-phase-e-chrome`    | Chromium 153 module worker: simple-excel + sheetjs, mixed 100k and 1m                                                                                                       |
| `results/2026-09-12-macbook-air-inline-default`    | mixed 100k × simple-excel, simple-excel-zlib in the shipped default (`strings: 'inline'`): write 3.08 s / 58 MB (1.73 s with `nodeDeflater(1)`), read-typed 1.64 s / 810 MB |

Headline, measured with the bounded shared-string table (now `strings: 'auto'`; the inline default is in
`inline-default`): simple-excel passes all seven gates on both deflate paths. At mixed 100k it writes in 2.77 s
(0.28× sheetjs) with a 107 MB RSS footprint (0.03×) and a 0.4 ms first byte, and reads typed in 1.62 s (0.26×) at
882 MB (0.48×); `nodeDeflater(1)` takes the write to 1.71 s for a 20% larger file. It writes 1M × 20 in 31.9 s and the
18M-cell shape in 27.0 s where sheetjs throws `RangeError: Invalid string length`, and completes 1M × 20 in a
Chrome worker (39.6 s, +254 MB renderer RSS) where sheetjs crashes the renderer. The absolute wall-clock targets
in `research/11-build-plan.md` §1 (100k written in 1.2 s, read in 1.5 s) are still not met on this dataset; see
`research/06-performance-baseline.md` for what the Phase E optimisation pass bought (write 1.65×, read 1.30×) and
where the rest of the time goes (the platform compressor on write, the XML tokenizer on read).

Earlier runs, kept for continuity: the Phase D set `results/2026-09-12-macbook-air-phase-d-*` (the same shapes
before the optimisation pass) and `results/2026-09-11-macbook-air-{baseline,scale,ceiling,shapes,combined,chrome}`
(the four candidate engines before simple-excel existed; the `ceiling` run is where office-kit's shared-string
table gives out, between 250k and 500k rows).

## Caveats

- Numbers are from one machine with other processes running; `results.json` records the load average at
  start and end. Compare ratios across runs, not absolute ms.
- `post-warm-up delta` includes V8 heap slack: a streaming engine that fits in the pages the warm-up already
  reserved legitimately reports ~0 MB, which is why ratios use the pre-warm-up `RSS footprint` instead.
- `--source materialized` (default) keeps the whole dataset resident in the baseline (≈ 450 MB per 100k
  mixed rows), which is the Jetstream situation but means 900k rows start from a ~4 GB baseline.
- Timezones: SheetJS writes `Date` cells as **local wall-clock** serials; office-kit, exceljs and
  write-excel-file write **UTC wall-clock**. The files differ by the machine's UTC offset. simple-excel
  defaults to `dates: 'local'`, so it matches SheetJS byte for byte here and needs no shifting; `dates: 'utc'`
  switches it to the other convention.
- SheetJS output below 10k rows is uncompressed (Jetstream only enables compression above 10k), so its
  `output` size is ~5× the others at `1k`.
- `read-raw` has no formatted-text mode outside SheetJS: exceljs/office-kit stringify values (dates stay
  serial numbers), and simple-excel streams typed array rows and counts them without stringifying. So
  `read-raw` is only strictly comparable within a column pair, and the honest comparison is `read-typed`,
  where every engine produces the same records.
- office-kit's write-only workbook keeps its shared-string table resident until `finalize()` and then
  serialises it as **one JS string** (`serializeSharedStrings` → `Array.join`). With Jetstream-shaped rows
  (~1.7 KB of text each) that string passes V8's ~2^29-char limit somewhere below 1M rows, so `1m` and
  `18m-cells` fail with `RangeError: Invalid string length` in `finalize()` — the same exception SheetJS
  throws, just later. There is no inline-string option in the streaming API (0.11.0); see the `ceiling`
  results for where the limit sits and `engines/office-kit.mjs` for the friction notes.
- exceljs's `WorkbookReader` (`row.values`) and office-kit's `iterRows` are sparse; the adapters rebuild
  dense rows so `cells read` is comparable to `sheet_to_json` with `defval`.
