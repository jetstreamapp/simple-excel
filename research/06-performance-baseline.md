# 06 - Performance baseline and targets

Measured engine performance on Jetstream-shaped data, the gates a replacement must clear, and the harness
that produces the numbers (`bench`, `npm run bench --`). The results section is rendered from the newest
`*-combined` run by `research/regenerate.sh 06`; the newer `inline-default` run is summarised by hand below.

_Accurate as of 2026-09-12; the results section is the Phase E run of `@jetstreamapp/simple-excel` itself, taken
after the hot-path optimisation pass (crc32, cell-text escaping, chunk encoding, attribute parsing)._

## Why these numbers matter

The three production failures are all "whole workbook in memory" failures (01): SheetJS built ~18M cells or
958k rows as one string/array and hit V8's limits at roughly 512 MiB of string or 3+ GB of heap. So the
gates are memory-first: an engine that is a little slower but streams with a flat heap wins; one that is
faster but still materializes loses.

## Datasets (seeded, `bench/lib/dataset.mjs`)

| Dataset          | Shape                                                                                                                                                | What it stresses                          |
| ---------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------- |
| `mixed`          | 20 cols: 18-char Ids, unicode names, decimals, booleans, JS Dates, ISO datetimes, long text with newlines, JSON blobs (1% > 32,767 chars), 10% nulls | The Salesforce record export / load shape |
| `wide`           | 107 cols: 60 strings, 30 numbers, 10 dates, 7 booleans                                                                                               | The 18M-cell incident shape (172k × 107)  |
| `strings-unique` | 20 cols of unique 24-char strings                                                                                                                    | Shared-string-table worst case            |
| `numeric`        | 20 numeric cols                                                                                                                                      | Best case / XML generation floor          |

Sizes 1k, 10k, 100k, 1M rows; `18m-cells` = 900k × 20 `mixed`. Reads consume files written by SheetJS (or
office-kit where SheetJS cannot produce them) — SheetJS writes inline strings with Jetstream's options, so the
read numbers are for a file with no shared-string table at all. Every cell runs in a fresh Node process
(`--expose-gc --max-old-space-size=12288`), 1 warm-up + 3 timed runs (1 for ≥ 1M), reporting median wall
time, peak RSS delta over the post-warm-up baseline, output bytes and first-byte latency for streaming writers.

## Gates (mixed 100k×20 unless stated; ratios vs SheetJS 0.20.3 with Jetstream's options)

Measured for `simple-excel` on 2026-09-12 (Apple M4, Node v24.18.0), runs
`2026-09-12-macbook-air-phase-e-{optimized,others,scale,chrome,combined}`. The `simple-excel-zlib` column is the
same engine with `nodeDeflater(1)`.

| Gate                           | Threshold                                                 | simple-excel                                                                  | simple-excel-zlib       |
| ------------------------------ | --------------------------------------------------------- | ----------------------------------------------------------------------------- | ----------------------- |
| write time                     | ≤ 1.0×                                                    | **PASS** 0.28× (2.77 s)                                                       | **PASS** 0.18× (1.71 s) |
| write peak memory              | ≤ 0.25×                                                   | **PASS** 0.03× (107 MB)                                                       | **PASS** 0.03× (100 MB) |
| read-typed time                | ≤ 1.0×                                                    | **PASS** 0.26× (1.62 s)                                                       | **PASS** 0.26× (1.60 s) |
| read-typed peak memory         | ≤ 0.5×                                                    | **PASS** 0.48× (882 MB)                                                       | **PASS** 0.48× (886 MB) |
| 1M × 20 write                  | succeeds in Node and in a Chrome module worker under 2 GB | **PASS** 31.9 s Node / 39.6 s Chrome, +254 MB renderer RSS                    | **PASS** 19.1 s Node    |
| 18M cells write                | succeeds in Node under 3 GB                               | **PASS** 27.0 s, 215 MB peak RSS                                              | **PASS** 15.6 s         |
| first byte (streaming writers) | < 100 ms at 100k rows                                     | **PASS** 0.4 ms                                                               | **PASS** 0.2 ms         |
| 100 MB xlsx read (1M rows)     | completes in Node under 1.5 GB                            | **PASS** 10.3 s, 245 MB peak RSS (needs `maxInflatedBytes` raised past 1 GiB) | same engine             |

Stretch: write memory ≤ 0.1× (**met**, 0.03×), time ≤ 0.5× (**met**, 0.28× / 0.18×).

These are `strings: 'auto'` numbers. The shipped default (inline strings) measures 3.08 s (0.32×) / 58 MB (0.02×)
on write and 1.64 s (0.27×) / 810 MB (0.44×) on read; see "Inline strings as the default" below.

The absolute targets in `11-build-plan.md` §1 are a separate set, and four of them are **not met** on this dataset:

| Target (`11-build-plan.md` §1) | Threshold      | Measured (mixed 100k×20)                    | Verdict  |
| ------------------------------ | -------------- | ------------------------------------------- | -------- |
| write 100k×20                  | ≤ 1.2 s        | 2.77 s (1.71 s with `nodeDeflater(1)`)      | **FAIL** |
| read-typed 100k×20             | ≤ 1.5 s        | 1.62 s                                      | **FAIL** |
| write throughput               | ≥ 100k rows/s  | 36.1k rows/s (58.5k with `nodeDeflater(1)`) | **FAIL** |
| read throughput                | ≥ 150k rows/s  | 61.7k rows/s typed, 68.0k rows/s array mode | **FAIL** |
| 1M × 20 write heap growth      | < 250 MB       | 77 MB (rows streamed, nothing materialised) | **PASS** |
| 18M cells                      | < 3 GB         | 76 MB heap growth, 215 MB peak RSS          | **PASS** |
| core bundle                    | ≤ 40 KB brotli | 29.0 KB (`npm run size`)                    | **PASS** |

A `mixed` row carries ~1.7 KB of text, so 100k rows is 170 MB of XML written and 213 MB read; the four time
targets were written without pinning a dataset. The optimisation pass below closed most of the gap - write is
1.65× faster and read 1.30× faster than Phase D - and what is left is the platform compressor (1.6 s of the
2.77 s write) and the per-byte cost of tokenizing 213 MB of XML. See "Reading the results" for the breakdown.

## Prior baselines (for continuity)

SheetJS on this machine (Apple M4, Node 24, 2026-08-23): write 10k×20 566 ms / 905 MB peak, 100k×20 7.1 s /
3.2 GB; read typed 10k 589 ms / 495 MB, 100k 7.5 s / 1.5 GB. excelize-wasm (killed): 4-8× slower writes,
10-17× slower reads.

## Results

<!-- generated:start (bench/results) -->

Source run: `2026-09-12-macbook-air-phase-e-combined`

Generated 2026-09-12T20:35:55.352Z. Results folder: `bench/results/2026-09-12-macbook-air-phase-e-combined`.

## Machine

- Apple M4 (10 cores), 32 GB RAM, darwin 25.5.0 arm64
- Node v24.18.0 (V8 13.6.233.17-node.50), child flags: `--expose-gc --max-old-space-size=12288`
- Load average at start 4.39 / 3.41 / 3.97, at end 4.60 / 4.54 / 4.32 (other processes may have been running)
- Options: runs=3, warmup=1, source=materialized, timeout=600s, seed=20260911

## Engines

| engine            | version                        | streaming | note |
| ----------------- | ------------------------------ | --------- | ---- |
| simple-excel      | 0.0.0                          | yes       |      |
| simple-excel-zlib | 0.0.0                          | yes       |      |
| sheetjs           | 0.20.3                         | no        |      |
| exceljs           | 4.4.0                          | yes       |      |
| office-kit        | 0.11.0                         | yes       |      |
| write-excel-file  | 2.3.10 (read-excel-file 5.8.8) | no        |      |

## Metrics

- `median`/`min`/`max`: wall time of the timed runs (after one warm-up run of the same cell).
- `RSS footprint`: highest RSS observed while the op ran (25 ms poller, post-run read, process high-water mark) minus the RSS before the warm-up (engine loaded, input data resident, gc() done). This is the ratio/gate metric.
- `post-warm-up delta`: the same peak minus the RSS after warm-up + gc(). Undercounts engines whose warm-up footprint stays resident (V8 rarely returns pages), kept for reference.
- `peak RSS`: the absolute high-water RSS of the child process (input data included).
- `first byte`: ms from the start of `write()` until the sink receives its first chunk (streaming writers only).
- Ratios are candidate ÷ sheetjs for the same op/dataset/size; lower is better.

## write — mixed

| size      | engine            | status                                            | median  | min     | max     | time vs sheetjs | RSS footprint | mem vs sheetjs | post-warm-up delta | peak RSS | output   | first byte |
| --------- | ----------------- | ------------------------------------------------- | ------- | ------- | ------- | --------------- | ------------- | -------------- | ------------------ | -------- | -------- | ---------- |
| 1k        | simple-excel      | ok                                                | 46.6 ms | 39.9 ms | 52.2 ms | 0.56x           | 16 MB         | 0.09x          | 8 MB               | 87 MB    | 450 KB   | 0.5 ms     |
| 1k        | simple-excel-zlib | ok                                                | 35.2 ms | 32.6 ms | 44.2 ms | 0.42x           | 16 MB         | 0.09x          | 9 MB               | 87 MB    | 529 KB   | 0.6 ms     |
| 1k        | sheetjs           | ok                                                | 83.6 ms | 81.5 ms | 105 ms  | —               | 171 MB        | —              | 136 MB             | 248 MB   | 2.5 MB   | —          |
| 10k       | simple-excel      | ok                                                | 285 ms  | 285 ms  | 291 ms  | 0.41x           | 44 MB         | 0.04x          | 17 MB              | 205 MB   | 4.2 MB   | 0.5 ms     |
| 10k       | simple-excel-zlib | ok                                                | 195 ms  | 189 ms  | 222 ms  | 0.28x           | 42 MB         | 0.04x          | 18 MB              | 203 MB   | 5.1 MB   | 0.3 ms     |
| 10k       | sheetjs           | ok                                                | 699 ms  | 685 ms  | 1.19 s  | —               | 1,084 MB      | —              | 901 MB             | 1,250 MB | 24.9 MB  | —          |
| 100k      | simple-excel      | ok                                                | 2.77 s  | 2.76 s  | 2.99 s  | 0.28x           | 107 MB        | 0.03x          | 44 MB              | 692 MB   | 42.4 MB  | 0.4 ms     |
| 100k      | simple-excel-zlib | ok                                                | 1.71 s  | 1.70 s  | 1.75 s  | 0.18x           | 100 MB        | 0.03x          | 42 MB              | 685 MB   | 50.9 MB  | 0.2 ms     |
| 100k      | sheetjs           | ok                                                | 9.75 s  | 8.50 s  | 10.24 s | —               | 3,313 MB      | —              | 3,289 MB           | 3,904 MB | 95.2 MB  | —          |
| 100k      | exceljs           | ok                                                | 2.49 s  | 2.47 s  | 2.53 s  | 0.26x           | 328 MB        | 0.10x          | 3 MB               | 923 MB   | 52.3 MB  | 1.61 s     |
| 100k      | office-kit        | ok                                                | 4.36 s  | 4.35 s  | 4.40 s  | 0.45x           | 1,241 MB      | 0.37x          | 501 MB             | 1,832 MB | 41.2 MB  | 3.1 ms     |
| 100k      | write-excel-file  | ok                                                | 3.21 s  | 3.20 s  | 3.22 s  | 0.33x           | 1,139 MB      | 0.34x          | 556 MB             | 1,731 MB | 48.2 MB  | —          |
| 1m        | simple-excel      | ok                                                | 31.95 s | 31.95 s | 31.95 s | —               | 328 MB        | —              | 70 MB              | 3,489 MB | 427.0 MB | 0.9 ms     |
| 1m        | simple-excel-zlib | ok                                                | 19.09 s | 19.09 s | 19.09 s | —               | 650 MB        | —              | 66 MB              | 3,837 MB | 511.3 MB | 0.4 ms     |
| 1m        | sheetjs           | ERROR (warm-up) RangeError: Invalid string length | —       | —       | —       | —               | —             | —              | —                  | —        | —        | —          |
| 18m-cells | simple-excel      | ok                                                | 27.02 s | 27.02 s | 27.02 s | —               | 327 MB        | —              | 46 MB              | 4,158 MB | 383.8 MB | 0.8 ms     |
| 18m-cells | simple-excel-zlib | ok                                                | 15.55 s | 15.55 s | 15.55 s | —               | 322 MB        | —              | 42 MB              | 4,152 MB | 459.7 MB | 0.4 ms     |
| 18m-cells | sheetjs           | ERROR (warm-up) RangeError: Invalid string length | —       | —       | —       | —               | —             | —              | —                  | —        | —        | —          |

## read-typed — mixed

Fixtures written by: sheetjs (see `fixture` in results.json per cell).

| size | engine            | status | median  | min     | max     | time vs sheetjs | RSS footprint | mem vs sheetjs | post-warm-up delta | peak RSS | rows read | cells read |
| ---- | ----------------- | ------ | ------- | ------- | ------- | --------------- | ------------- | -------------- | ------------------ | -------- | --------- | ---------- |
| 1k   | simple-excel      | ok     | 41.5 ms | 34.7 ms | 42.0 ms | 0.56x           | 56 MB         | 0.58x          | 31 MB              | 113 MB   | 1,000     | 20,000     |
| 1k   | simple-excel-zlib | ok     | 41.6 ms | 28.5 ms | 42.0 ms | 0.56x           | 56 MB         | 0.58x          | 31 MB              | 113 MB   | 1,000     | 20,000     |
| 1k   | sheetjs           | ok     | 74.5 ms | 73.5 ms | 75.3 ms | —               | 97 MB         | —              | 43 MB              | 166 MB   | 1,000     | 20,000     |
| 10k  | simple-excel      | ok     | 155 ms  | 153 ms  | 169 ms  | 0.27x           | 220 MB        | 0.56x          | 101 MB             | 322 MB   | 10,000    | 200,000    |
| 10k  | simple-excel-zlib | ok     | 155 ms  | 153 ms  | 164 ms  | 0.27x           | 219 MB        | 0.55x          | 100 MB             | 321 MB   | 10,000    | 200,000    |
| 10k  | sheetjs           | ok     | 585 ms  | 574 ms  | 592 ms  | —               | 394 MB        | —              | 74 MB              | 504 MB   | 10,000    | 200,000    |
| 100k | simple-excel      | ok     | 1.62 s  | 1.61 s  | 1.65 s  | 0.26x           | 882 MB        | 0.48x          | 14 MB              | 1,125 MB | 100,000   | 2,000,000  |
| 100k | simple-excel-zlib | ok     | 1.60 s  | 1.60 s  | 1.62 s  | 0.26x           | 886 MB        | 0.48x          | 17 MB              | 1,128 MB | 100,000   | 2,000,000  |
| 100k | sheetjs           | ok     | 6.17 s  | 6.13 s  | 6.22 s  | —               | 1,837 MB      | —              | 402 MB             | 2,092 MB | 100,000   | 2,000,000  |
| 100k | exceljs           | ok     | 9.53 s  | 9.51 s  | 9.67 s  | 1.54x           | 193 MB        | 0.10x          | 30 MB              | 268 MB   | 100,000   | 2,000,000  |
| 100k | office-kit        | ok     | 6.57 s  | 6.45 s  | 6.78 s  | 1.06x           | 449 MB        | 0.24x          | 105 MB             | 517 MB   | 100,000   | 2,000,000  |
| 100k | write-excel-file  | ok     | 6.80 s  | 6.67 s  | 6.86 s  | 1.10x           | 5,290 MB      | 2.88x          | 1,083 MB           | 5,360 MB | 100,000   | 2,000,000  |

## read-raw — mixed

Fixtures written by: sheetjs (see `fixture` in results.json per cell).

| size | engine            | status                                                | median  | min     | max     | time vs sheetjs | RSS footprint | mem vs sheetjs | post-warm-up delta | peak RSS | rows read | cells read |
| ---- | ----------------- | ----------------------------------------------------- | ------- | ------- | ------- | --------------- | ------------- | -------------- | ------------------ | -------- | --------- | ---------- |
| 1k   | simple-excel      | ok                                                    | 22.0 ms | 21.3 ms | 24.1 ms | 0.27x           | 36 MB         | 0.37x          | 15 MB              | 94 MB    | 1,000     | 20,000     |
| 1k   | simple-excel-zlib | ok                                                    | 35.8 ms | 30.7 ms | 43.1 ms | 0.45x           | 35 MB         | 0.35x          | 14 MB              | 92 MB    | 1,000     | 20,000     |
| 1k   | sheetjs           | ok                                                    | 80.3 ms | 78.7 ms | 81.4 ms | —               | 98 MB         | —              | 47 MB              | 163 MB   | 1,000     | 20,000     |
| 10k  | simple-excel      | ok                                                    | 157 ms  | 151 ms  | 160 ms  | 0.26x           | 103 MB        | 0.26x          | 56 MB              | 205 MB   | 10,000    | 200,000    |
| 10k  | simple-excel-zlib | ok                                                    | 158 ms  | 157 ms  | 159 ms  | 0.26x           | 101 MB        | 0.26x          | 55 MB              | 203 MB   | 10,000    | 200,000    |
| 10k  | sheetjs           | ok                                                    | 604 ms  | 602 ms  | 608 ms  | —               | 393 MB        | —              | 75 MB              | 502 MB   | 10,000    | 200,000    |
| 100k | simple-excel      | ok                                                    | 1.47 s  | 1.45 s  | 1.49 s  | 0.23x           | 76 MB         | 0.04x          | 32 MB              | 318 MB   | 100,000   | 2,000,000  |
| 100k | simple-excel-zlib | ok                                                    | 1.46 s  | 1.46 s  | 1.47 s  | 0.23x           | 76 MB         | 0.04x          | 32 MB              | 319 MB   | 100,000   | 2,000,000  |
| 100k | sheetjs           | ok                                                    | 6.29 s  | 6.29 s  | 6.34 s  | —               | 1,966 MB      | —              | 460 MB             | 2,221 MB | 100,000   | 2,000,000  |
| 100k | exceljs           | ok                                                    | 9.58 s  | 3.14 s  | 9.61 s  | 1.52x           | 200 MB        | 0.10x          | 12 MB              | 275 MB   | 100,000   | 2,000,000  |
| 100k | office-kit        | ok                                                    | 6.64 s  | 6.30 s  | 6.70 s  | 1.06x           | 533 MB        | 0.27x          | 186 MB             | 601 MB   | 100,000   | 2,000,000  |
| 100k | write-excel-file  | skipped: write-excel-file does not implement read-raw | —       | —       | —       | —               | —             | —              | —                  | —        | —         | —          |

## Gates

### simple-excel: PASS

| gate                                                       | status | detail                                  |
| ---------------------------------------------------------- | ------ | --------------------------------------- |
| write RSS footprint <= 0.25x sheetjs (mixed 100k)          | PASS   | 0.03x (107 MB vs 3,313 MB; limit 0.25x) |
| write median time <= 1.0x sheetjs (mixed 100k)             | PASS   | 0.28x (2.77 s vs 9.75 s; limit 1x)      |
| read-typed median time <= 1.0x sheetjs (mixed 100k)        | PASS   | 0.26x (1.62 s vs 6.17 s; limit 1x)      |
| read-typed RSS footprint <= 0.5x sheetjs (mixed 100k)      | PASS   | 0.48x (882 MB vs 1,837 MB; limit 0.5x)  |
| write 1M x 20 (mixed 1m) succeeds                          | PASS   | 31.95 s, 328 MB RSS footprint           |
| write 900k x 20 (mixed 18m-cells) succeeds                 | PASS   | 27.02 s, 327 MB RSS footprint           |
| first byte reaches the sink in < 100 ms (mixed 100k write) | PASS   | 0.4 ms (limit 100 ms)                   |

### simple-excel-zlib: PASS

| gate                                                       | status | detail                                  |
| ---------------------------------------------------------- | ------ | --------------------------------------- |
| write RSS footprint <= 0.25x sheetjs (mixed 100k)          | PASS   | 0.03x (100 MB vs 3,313 MB; limit 0.25x) |
| write median time <= 1.0x sheetjs (mixed 100k)             | PASS   | 0.18x (1.71 s vs 9.75 s; limit 1x)      |
| read-typed median time <= 1.0x sheetjs (mixed 100k)        | PASS   | 0.26x (1.60 s vs 6.17 s; limit 1x)      |
| read-typed RSS footprint <= 0.5x sheetjs (mixed 100k)      | PASS   | 0.48x (886 MB vs 1,837 MB; limit 0.5x)  |
| write 1M x 20 (mixed 1m) succeeds                          | PASS   | 19.09 s, 650 MB RSS footprint           |
| write 900k x 20 (mixed 18m-cells) succeeds                 | PASS   | 15.55 s, 322 MB RSS footprint           |
| first byte reaches the sink in < 100 ms (mixed 100k write) | PASS   | 0.2 ms (limit 100 ms)                   |

### exceljs: FAIL

| gate                                                       | status  | detail                                  |
| ---------------------------------------------------------- | ------- | --------------------------------------- |
| write RSS footprint <= 0.25x sheetjs (mixed 100k)          | PASS    | 0.10x (328 MB vs 3,313 MB; limit 0.25x) |
| write median time <= 1.0x sheetjs (mixed 100k)             | PASS    | 0.26x (2.49 s vs 9.75 s; limit 1x)      |
| read-typed median time <= 1.0x sheetjs (mixed 100k)        | FAIL    | 1.54x (9.53 s vs 6.17 s; limit 1x)      |
| read-typed RSS footprint <= 0.5x sheetjs (mixed 100k)      | PASS    | 0.10x (193 MB vs 1,837 MB; limit 0.5x)  |
| write 1M x 20 (mixed 1m) succeeds                          | NOT-RUN | cell not in this run                    |
| write 900k x 20 (mixed 18m-cells) succeeds                 | NOT-RUN | cell not in this run                    |
| first byte reaches the sink in < 100 ms (mixed 100k write) | FAIL    | 1.61 s (limit 100 ms)                   |

### office-kit: FAIL

| gate                                                       | status  | detail                                    |
| ---------------------------------------------------------- | ------- | ----------------------------------------- |
| write RSS footprint <= 0.25x sheetjs (mixed 100k)          | FAIL    | 0.37x (1,241 MB vs 3,313 MB; limit 0.25x) |
| write median time <= 1.0x sheetjs (mixed 100k)             | PASS    | 0.45x (4.36 s vs 9.75 s; limit 1x)        |
| read-typed median time <= 1.0x sheetjs (mixed 100k)        | FAIL    | 1.06x (6.57 s vs 6.17 s; limit 1x)        |
| read-typed RSS footprint <= 0.5x sheetjs (mixed 100k)      | PASS    | 0.24x (449 MB vs 1,837 MB; limit 0.5x)    |
| write 1M x 20 (mixed 1m) succeeds                          | NOT-RUN | cell not in this run                      |
| write 900k x 20 (mixed 18m-cells) succeeds                 | NOT-RUN | cell not in this run                      |
| first byte reaches the sink in < 100 ms (mixed 100k write) | PASS    | 3.1 ms (limit 100 ms)                     |

### write-excel-file: FAIL

| gate                                                       | status  | detail                                                |
| ---------------------------------------------------------- | ------- | ----------------------------------------------------- |
| write RSS footprint <= 0.25x sheetjs (mixed 100k)          | FAIL    | 0.34x (1,139 MB vs 3,313 MB; limit 0.25x)             |
| write median time <= 1.0x sheetjs (mixed 100k)             | PASS    | 0.33x (3.21 s vs 9.75 s; limit 1x)                    |
| read-typed median time <= 1.0x sheetjs (mixed 100k)        | FAIL    | 1.10x (6.80 s vs 6.17 s; limit 1x)                    |
| read-typed RSS footprint <= 0.5x sheetjs (mixed 100k)      | FAIL    | 2.88x (5,290 MB vs 1,837 MB; limit 0.5x)              |
| write 1M x 20 (mixed 1m) succeeds                          | NOT-RUN | cell not in this run                                  |
| write 900k x 20 (mixed 18m-cells) succeeds                 | NOT-RUN | cell not in this run                                  |
| first byte reaches the sink in < 100 ms (mixed 100k write) | FAIL    | engine reported no first-byte latency (not streaming) |

## Failures

- **sheetjs / write / mixed / 1m** — ERROR (warm-up) RangeError: Invalid string length

  ```
  RangeError: Invalid string length
      at Array.join (<anonymous>)
      at write_ws_xml_data (file:///Users/austinturner/dev/jetstream-simple-excel/node_modules/xlsx/xlsx.mjs:16417:11)
      at write_ws_xml (file:///Users/austinturner/dev/jetstream-simple-excel/node_modules/xlsx/xlsx.mjs:16458:11)
      at write_zip_xlsx (file:///Users/austinturner/dev/jetstream-simple-excel/node_modules/xlsx/xlsx.mjs:27056:25)
      at write_zip (file:///Users/austinturner/dev/jetstream-simple-excel/node_modules/xlsx/xlsx.mjs:27294:19)
  ```

- **sheetjs / write / mixed / 18m-cells** — ERROR (warm-up) RangeError: Invalid string length

  ```
  RangeError: Invalid string length
      at Array.join (<anonymous>)
      at write_ws_xml_data (file:///Users/austinturner/dev/jetstream-simple-excel/node_modules/xlsx/xlsx.mjs:16417:11)
      at write_ws_xml (file:///Users/austinturner/dev/jetstream-simple-excel/node_modules/xlsx/xlsx.mjs:16458:11)
      at write_zip_xlsx (file:///Users/austinturner/dev/jetstream-simple-excel/node_modules/xlsx/xlsx.mjs:27056:25)
      at write_zip (file:///Users/austinturner/dev/jetstream-simple-excel/node_modules/xlsx/xlsx.mjs:27294:19)
  ```

<!-- generated:end -->

## Reading the results (runs `2026-09-12-macbook-air-phase-e-{optimized,others,scale,chrome}`; `phase-e-combined` merges the Node runs)

- **simple-excel passes all seven harness gates on both deflate paths.** At `mixed` 100k×20 it writes in 2.77 s
  (0.28× SheetJS) with a 107 MB RSS footprint (0.03×) and a 0.4 ms first byte, and reads typed in 1.62 s (0.26×)
  at 882 MB (0.48×). Streaming array-mode reads (`read-raw`) cost 1.47 s and 76 MB (0.04×) — the 882 MB is the
  100,000 materialised records `toObjects()` returns, not the parser.
- **The scale gates are the point.** 1M × 20 writes in 31.9 s and the 18M-cell shape (900k × 20) in 27.0 s, where
  SheetJS throws `RangeError: Invalid string length` in `write_ws_xml_data` for both. Measured on their own terms
  with rows streamed from the generator (nothing materialised), both shapes grow the **JS heap by 76-77 MB** and
  peak RSS at 214-215 MB — the 250 MB heap gate and the 3 GB 18M-cell gate pass with an order of magnitude to
  spare. The 328 MB "RSS footprint" the harness reports for 1M is measured against a baseline that already holds
  the 4.2 GB materialised dataset, so it is V8 slack around the writer, not the writer.
- **`nodeDeflater(1)` is worth 38% of write wall time** and nothing else: 1.71 s vs 2.77 s at 100k, 19.1 s vs
  31.9 s at 1M, for a file 20% larger (50.9 MB vs 42.4 MB at 100k). Reads are unaffected. The platform
  `CompressionStream` is the default because it is the only one a browser has.
- **Chrome module worker** (`phase-e-chrome`, Chromium 153 headless, `collectToBlob()` sink): simple-excel writes
  100k × 20 in 3.88 s with +236 MB renderer RSS and 1M × 20 in 39.6 s with +254 MB — flat across a 10× row
  increase, holding the finished 427 MB Blob. SheetJS writes 100k in 7.64 s with +3,039 MB and **crashes the
  renderer** at 1M. `measureUserAgentSpecificMemory()` still rarely resolves inside a run, so renderer RSS from
  `ps` remains the number to read.
- **The absolute time targets in `11-build-plan.md` §1 are still not met on this dataset**, but the gap is now
  small enough to name: write 100k × 20 takes 2.77 s against a 1.2 s target (36.1k rows/s against 100k rows/s),
  read-typed 1.62 s against 1.5 s (61.7k rows/s against 150k rows/s). `mixed` carries ~1.7 KB of text per row, so
  100k rows is 170 MB of XML in and 213 MB out; the targets were set without that shape in mind. On a lighter
  shape the same engine is far quicker — 10k × 20 writes in 285 ms and reads in 155 ms.

### Inline strings as the default (run `2026-09-12-macbook-air-inline-default`)

After ADR-001's addendum made `strings: 'inline'` the writer default, the 100k × 20 `mixed` cells were re-measured
in the default configuration. The tables above (`phase-e-*`) were measured with the bounded shared-string table
that is now `strings: 'auto'`.

| op         | engine            | inline default (new) | `strings: 'auto'` (phase-e) |
| ---------- | ----------------- | -------------------- | --------------------------- |
| write      | simple-excel      | 3.08 s, 58 MB        | 2.77 s, 107 MB              |
| write      | simple-excel-zlib | 1.73 s, 74 MB        | 1.71 s, 100 MB              |
| read-typed | simple-excel      | 1.64 s, 810 MB       | 1.62 s, 882 MB              |
| read-raw   | simple-excel      | 1.48 s, 46 MB        | 1.47 s, 76 MB               |

Inline strings send about 20% more bytes into the compressor on this repeat-heavy data (30% of its strings
repeat), and the platform `CompressionStream` is the bottleneck, so the platform write is 11% slower; the zlib path
and both reads are unchanged within noise, and the writer's memory footprint drops by about 50 MB because there
is no table. The trade was made for fidelity in Apple Numbers (ADR-001 addendum). The 1M-row, 18M-cell and Chrome
figures were not re-measured; expect the same ~10% on the platform write path.

### What the Phase E optimisation pass changed

Five changes, each measured on its own with `bench` and a `--cpu-prof` run of the same workload. Byte-for-byte
identical output: `test/golden-bytes` stayed green throughout, so no golden or oracle re-run was needed.

| Change                                                                                                       | Effect                                                     |
| ------------------------------------------------------------------------------------------------------------ | ---------------------------------------------------------- |
| `crc32`: slicing-by-16 over a `DataView`, sixteen 256-entry tables in one flat `Int32Array`                  | 560 MB/s → 2.8 GB/s; -12% write, -19% read                 |
| `encodeCellText`: one precompiled pattern finds the clean case, and only a _lone_ surrogate is work          | -25% write; a clean 137-char cell no longer walks per char |
| `needsSpacePreserve`: `/[\t\n\r]/` instead of a per-character interior scan                                  | folded into the same -25%                                  |
| `encodeXmlChunk`: `encodeInto` a reused buffer instead of `TextEncoder.encode` per flush (`src/xml/utf8.ts`) | -11% write                                                 |
| Tokenizer: attributes parsed once into spans per tag, hot element names interned                             | -6% read-raw, -5% read-typed                               |

Two candidates were measured and **dropped**: a larger deflate flush chunk (256 Ki instead of 64 Ki chars) moves
the platform compressor by 2-3%, inside the noise and not worth regenerating the goldens for; and pipelining
`writer.write` on the `CompressionStream` behind `writer.ready` buys nothing, because Node runs the deflate on the
JS thread — a producer that burns 558 ms of CPU between chunks and a 211 ms compression simply add up (770 ms)
whether the writes are awaited or not. Replacing the worksheet writer's `string[]` + `join` with a per-row or
per-fragment `encodeInto` was also measured (84 ms vs 86 ms for 20k rows) and dropped as noise.

- **Where the write time goes now** (100k × 20, `--cpu-prof`, self time, compression off so the whole write is
  0.90 s): `TextEncoder.encodeInto` 11%, **`crc32` 11%**, the cell-text scan 8%, shared-string `intern` 7%, the
  builder's `join` 6%, GC 5%, `numberText` 3%, `needsSpacePreserve` 3%. With the platform compressor on top the
  write is 2.77 s, so **deflate is 1.9 s of it** — the single largest item by far, and not ours: Node's
  `CompressionStream` compresses this XML at about 135 MB/s whatever chunk size it is fed.
- **Where the read time goes now** (100k × 20 typed, 1.58 s): `decodeEntities` 11% (the SheetJS fixture escapes
  every `"` in a JSON blob as `&quot;`), `appendText` 11%, the tokenizer's `stepText`/`scan` 15%, GC 8%,
  `decodeIfEscaped` 5%, `parseAttributes` 5%, `TextDecoder` 5%, **`crc32` 4%** (was 15%), `attr` 2%.
- **What is left on the table.** The reader is now tokenizer-bound: roughly 60% of a typed read is in
  `src/xml/tokenizer.ts`, and every remaining idea there (fusing the `&` scan into the tag scan, a byte-level
  tokenizer that never materialises a chunk string) is a rewrite rather than a tweak. On the write side the only
  material lever left is the compressor itself — a hand-rolled deflate could beat the platform's level-6 default
  the way `nodeDeflater(1)` does, at the cost of the bundle budget and of owning a compressor.
- **Reading back a 1M-row file needs the inflate cap raised.** Our own 1M × 20 output inflates to 2.19 GB of sheet
  XML, over the 1 GiB default `limits.maxInflatedBytes`, so `openWorkbook` refuses it with `ZIP_BOMB` — the cap
  working as designed. With the cap at 4 GiB the same file streams back in 10.3 s (1,000,001 rows / 19.9M cells)
  for 64 MB of heap growth and a 245 MB peak RSS, well inside the 1.5 GB target.
- **A reader fix from the Phase D profiling stayed in.** `WorksheetParser.prepareSharedStrings` used to scan every
  chunk of the sheet twice for `t="s"` even when the package has no shared-strings part at all — 16% of the read
  on a SheetJS-written file, which uses inline strings only. The parse skips the scan when the part is absent
  (`WorksheetReadContext.hasSharedStrings`); a `t="s"` cell in such a file still resolves, through the deferred
  path.
- **The other engines** were re-measured in the same session at 100k (`phase-e-others`) and are unchanged in
  character: ExcelJS is still the closest writer (2.49 s) but its first byte is 1.61 s and it reads at 1.54×
  SheetJS; office-kit writes at 0.45× SheetJS but holds 1,241 MB at 100k and dies with the same `RangeError` past
  ~250k rows (2026-09-11 `ceiling` run); write-excel-file reads 100k at a 5,290 MB footprint.
- Load average was 3-10 during the runs (other processes were active); ratios are more trustworthy than absolute
  numbers. Bundle size, the last gate in `11-build-plan.md` §1, is checked by `npm run size`: the core browser
  entry is **29.0 KB brotli** (99.2 KB minified, 33.1 KB gzip) against a 40 KB budget; it was 24.9 KB in 0.1, before
  the writer's input validation (2026-09-27).
- The 0.2 fixes (input validation, header-row detection, failure latching) were measured before and after on
  mixed 100k × 20 (`bench/results/2026-09-27-macbook-air-audit-fixes-before` and `…-after`, median of three): write
  2.86 s → 2.93 s with the same 2.83 s best run (zlib 1.74 s → 1.72 s), typed read 1.69 s → 1.61 s (zlib
  1.60 s → 1.66 s), footprints unchanged. Every difference is inside run-to-run noise.
