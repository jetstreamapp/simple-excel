# 06 - Performance baseline and targets

Measured engine performance on Jetstream-shaped data, the gates a replacement must clear, and the harness
that produces the numbers (`bench`, `npm run bench --`). The results section is
rendered from the newest committed run by `research/regenerate.sh 06`.

_Accurate as of 2026-09-12._

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
office-kit where SheetJS cannot produce them). Every cell runs in a fresh Node process
(`--expose-gc --max-old-space-size=12288`), 1 warm-up + 3 timed runs (1 for ≥ 1M), reporting median wall
time, peak RSS delta over the post-warm-up baseline, output bytes and first-byte latency for streaming writers.

## Gates (mixed 100k×20 unless stated; ratios vs SheetJS 0.20.3 with Jetstream's options)

| Gate                           | Threshold                                                 |
| ------------------------------ | --------------------------------------------------------- |
| write time                     | ≤ 1.0×                                                    |
| write peak memory              | ≤ 0.25×                                                   |
| read-typed time                | ≤ 1.0×                                                    |
| read-typed peak memory         | ≤ 0.5×                                                    |
| 1M × 20 write                  | succeeds in Node and in a Chrome module worker under 2 GB |
| 18M cells write                | succeeds in Node under 3 GB                               |
| first byte (streaming writers) | < 100 ms at 100k rows                                     |
| 100 MB xlsx read (1M rows)     | completes in Node under 1.5 GB                            |

Stretch: write memory ≤ 0.1×, time ≤ 0.5×.

## Prior baselines (for continuity)

SheetJS on this machine (Apple M4, Node 24, 2026-08-23): write 10k×20 566 ms / 905 MB peak, 100k×20 7.1 s /
3.2 GB; read typed 10k 589 ms / 495 MB, 100k 7.5 s / 1.5 GB. excelize-wasm (killed): 4-8× slower writes,
10-17× slower reads.

## Results

<!-- generated:start (bench/results) -->

Source run: ``

Generated 2026-09-12T04:22:41.486Z. Results folder: `bench/results/2026-09-11-macbook-air-combined`.

## Machine

- Apple M4 (10 cores), 32 GB RAM, darwin 25.5.0 arm64
- Node v24.18.0 (V8 13.6.233.17-node.50), child flags: `--expose-gc --max-old-space-size=12288`
- Load average at start 3.24 / 3.38 / 3.30, at end 6.43 / 8.58 / 8.78 (other processes may have been running)
- Options: runs=3, warmup=1, source=materialized, timeout=600s, seed=20260911

## Engines

| engine           | version                        | streaming | note |
| ---------------- | ------------------------------ | --------- | ---- |
| sheetjs          | 0.20.3                         | no        |      |
| exceljs          | 4.4.0                          | yes       |      |
| office-kit       | 0.11.0                         | yes       |      |
| write-excel-file | 2.3.10 (read-excel-file 5.8.8) | no        |      |

## Metrics

- `median`/`min`/`max`: wall time of the timed runs (after one warm-up run of the same cell).
- `RSS footprint`: highest RSS observed while the op ran (25 ms poller, post-run read, process high-water mark) minus the RSS before the warm-up (engine loaded, input data resident, gc() done). This is the ratio/gate metric.
- `post-warm-up delta`: the same peak minus the RSS after warm-up + gc(). Undercounts engines whose warm-up footprint stays resident (V8 rarely returns pages), kept for reference.
- `peak RSS`: the absolute high-water RSS of the child process (input data included).
- `first byte`: ms from the start of `write()` until the sink receives its first chunk (streaming writers only).
- Ratios are candidate ÷ sheetjs for the same op/dataset/size; lower is better.

## write — mixed

| size      | engine           | status                                            | median  | min     | max     | time vs sheetjs | RSS footprint | mem vs sheetjs | post-warm-up delta | peak RSS | output   | first byte |
| --------- | ---------------- | ------------------------------------------------- | ------- | ------- | ------- | --------------- | ------------- | -------------- | ------------------ | -------- | -------- | ---------- |
| 1k        | sheetjs          | ok                                                | 96.2 ms | 69.7 ms | 99.6 ms | —               | 172 MB        | —              | 137 MB             | 248 MB   | 2.5 MB   | —          |
| 1k        | exceljs          | ok                                                | 38.2 ms | 36.0 ms | 39.0 ms | 0.40x           | 27 MB         | 0.16x          | 3 MB               | 117 MB   | 543 KB   | 25.2 ms    |
| 1k        | office-kit       | ok                                                | 56.7 ms | 56.3 ms | 59.9 ms | 0.59x           | 30 MB         | 0.17x          | 14 MB              | 111 MB   | 435 KB   | 5.4 ms     |
| 1k        | write-excel-file | ok                                                | 52.5 ms | 45.2 ms | 56.0 ms | 0.55x           | 36 MB         | 0.21x          | 25 MB              | 122 MB   | 503 KB   | —          |
| 10k       | sheetjs          | ok                                                | 807 ms  | 803 ms  | 1.10 s  | —               | 1,087 MB      | —              | 906 MB             | 1,258 MB | 24.9 MB  | —          |
| 10k       | exceljs          | ok                                                | 271 ms  | 270 ms  | 276 ms  | 0.34x           | 107 MB        | 0.10x          | 2 MB               | 284 MB   | 5.2 MB   | 174 ms     |
| 10k       | office-kit       | ok                                                | 436 ms  | 433 ms  | 446 ms  | 0.54x           | 183 MB        | 0.17x          | 65 MB              | 355 MB   | 4.1 MB   | 4.0 ms     |
| 10k       | write-excel-file | ok                                                | 325 ms  | 319 ms  | 344 ms  | 0.40x           | 188 MB        | 0.17x          | 49 MB              | 362 MB   | 4.8 MB   | —          |
| 100k      | sheetjs          | ok                                                | 10.36 s | 9.50 s  | 11.47 s | —               | 2,899 MB      | —              | 2,682 MB           | 3,490 MB | 95.2 MB  | —          |
| 100k      | exceljs          | ok                                                | 2.95 s  | 2.55 s  | 3.29 s  | 0.29x           | 349 MB        | 0.12x          | 31 MB              | 943 MB   | 52.3 MB  | 1.90 s     |
| 100k      | office-kit       | ok                                                | 4.62 s  | 4.51 s  | 4.67 s  | 0.45x           | 1,158 MB      | 0.40x          | 517 MB             | 1,749 MB | 41.2 MB  | 2.8 ms     |
| 100k      | write-excel-file | ok                                                | 3.68 s  | 3.40 s  | 3.78 s  | 0.36x           | 1,161 MB      | 0.40x          | 632 MB             | 1,754 MB | 48.2 MB  | —          |
| 1m        | sheetjs          | ERROR (warm-up) RangeError: Invalid string length | —       | —       | —       | —               | —             | —              | —                  | —        | —        | —          |
| 1m        | office-kit       | ERROR (warm-up) RangeError: Invalid string length | —       | —       | —       | —               | —             | —              | —                  | —        | —        | —          |
| 18m-cells | sheetjs          | ERROR (warm-up) RangeError: Invalid string length | —       | —       | —       | —               | —             | —              | —                  | —        | —        | —          |
| 18m-cells | office-kit       | ERROR (warm-up) RangeError: Invalid string length | —       | —       | —       | —               | —             | —              | —                  | —        | —        | —          |
| 250k      | sheetjs          | ERROR (warm-up) RangeError: Invalid string length | —       | —       | —       | —               | —             | —              | —                  | —        | —        | —          |
| 500k      | sheetjs          | ERROR (warm-up) RangeError: Invalid string length | —       | —       | —       | —               | —             | —              | —                  | —        | —        | —          |
| 250k      | office-kit       | ok                                                | 11.70 s | 11.70 s | 11.70 s | —               | 2,434 MB      | —              | 838 MB             | 3,628 MB | 102.9 MB | 4.6 ms     |
| 500k      | office-kit       | ERROR (warm-up) RangeError: Invalid string length | —       | —       | —       | —               | —             | —              | —                  | —        | —        | —          |

## write — strings-unique

| size | engine     | status | median | min    | max    | time vs sheetjs | RSS footprint | mem vs sheetjs | post-warm-up delta | peak RSS | output  | first byte |
| ---- | ---------- | ------ | ------ | ------ | ------ | --------------- | ------------- | -------------- | ------------------ | -------- | ------- | ---------- |
| 100k | sheetjs    | ok     | 4.43 s | 4.37 s | 6.36 s | —               | 3,408 MB      | —              | 2,782 MB           | 4,093 MB | 53.5 MB | —          |
| 100k | office-kit | ok     | 4.12 s | 4.03 s | 4.22 s | 0.93x           | 890 MB        | 0.26x          | 366 MB             | 1,577 MB | 37.1 MB | 2.4 ms     |

## write — numeric

| size | engine     | status | median | min    | max    | time vs sheetjs | RSS footprint | mem vs sheetjs | post-warm-up delta | peak RSS | output  | first byte |
| ---- | ---------- | ------ | ------ | ------ | ------ | --------------- | ------------- | -------------- | ------------------ | -------- | ------- | ---------- |
| 100k | sheetjs    | ok     | 3.09 s | 2.73 s | 3.27 s | —               | 1,814 MB      | —              | 1,295 MB           | 1,931 MB | 32.3 MB | —          |
| 100k | office-kit | ok     | 1.84 s | 1.84 s | 1.90 s | 0.60x           | 54 MB         | 0.03x          | 7 MB               | 173 MB   | 17.0 MB | 3.3 ms     |

## write — wide

| size | engine     | status                                                                                                      | median  | min     | max     | time vs sheetjs | RSS footprint | mem vs sheetjs | post-warm-up delta | peak RSS | output   | first byte |
| ---- | ---------- | ----------------------------------------------------------------------------------------------------------- | ------- | ------- | ------- | --------------- | ------------- | -------------- | ------------------ | -------- | -------- | ---------- |
| 100k | sheetjs    | ERROR (process) ChildProcessError: child exited with code null signal SIGABRT and no result; stderr tail: … | —       | —       | —       | —               | —             | —              | —                  | —        | —        | —          |
| 100k | office-kit | ok                                                                                                          | 16.30 s | 16.14 s | 16.49 s | —               | 1,728 MB      | —              | 475 MB             | 3,523 MB | 147.9 MB | 2.5 ms     |

## read-typed — mixed

Fixtures written by: sheetjs (see `fixture` in results.json per cell).

| size | engine           | status | median  | min     | max     | time vs sheetjs | RSS footprint | mem vs sheetjs | post-warm-up delta | peak RSS | rows read | cells read |
| ---- | ---------------- | ------ | ------- | ------- | ------- | --------------- | ------------- | -------------- | ------------------ | -------- | --------- | ---------- |
| 1k   | sheetjs          | ok     | 63.4 ms | 62.5 ms | 63.4 ms | —               | 101 MB        | —              | 44 MB              | 167 MB   | 1,000     | 20,000     |
| 1k   | exceljs          | ok     | 53.8 ms | 48.0 ms | 54.6 ms | 0.85x           | 42 MB         | 0.41x          | 19 MB              | 117 MB   | 1,000     | 20,000     |
| 1k   | office-kit       | ok     | 78.0 ms | 77.0 ms | 79.7 ms | 1.23x           | 179 MB        | 1.77x          | 63 MB              | 247 MB   | 1,000     | 20,000     |
| 1k   | write-excel-file | ok     | 79.0 ms | 71.8 ms | 85.4 ms | 1.25x           | 185 MB        | 1.83x          | 71 MB              | 257 MB   | 1,000     | 20,000     |
| 10k  | sheetjs          | ok     | 619 ms  | 603 ms  | 710 ms  | —               | 392 MB        | —              | 75 MB              | 505 MB   | 10,000    | 200,000    |
| 10k  | exceljs          | ok     | 963 ms  | 943 ms  | 998 ms  | 1.55x           | 124 MB        | 0.32x          | 71 MB              | 199 MB   | 10,000    | 200,000    |
| 10k  | office-kit       | ok     | 616 ms  | 604 ms  | 649 ms  | 1.00x           | 639 MB        | 1.63x          | 23 MB              | 708 MB   | 10,000    | 200,000    |
| 10k  | write-excel-file | ok     | 592 ms  | 587 ms  | 614 ms  | 0.96x           | 659 MB        | 1.68x          | 141 MB             | 730 MB   | 10,000    | 200,000    |
| 100k | sheetjs          | ok     | 6.50 s  | 6.44 s  | 6.71 s  | —               | 1,934 MB      | —              | 525 MB             | 2,185 MB | 100,000   | 2,000,000  |
| 100k | exceljs          | ok     | 9.94 s  | 9.76 s  | 10.35 s | 1.53x           | 209 MB        | 0.11x          | 25 MB              | 284 MB   | 100,000   | 2,000,000  |
| 100k | office-kit       | ok     | 6.55 s  | 6.49 s  | 6.67 s  | 1.01x           | 522 MB        | 0.27x          | 183 MB             | 591 MB   | 100,000   | 2,000,000  |
| 100k | write-excel-file | ok     | 8.93 s  | 8.26 s  | 11.15 s | 1.37x           | 4,787 MB      | 2.48x          | 1,946 MB           | 4,858 MB | 100,000   | 2,000,000  |

## read-typed — strings-unique

Fixtures written by: sheetjs (see `fixture` in results.json per cell).

| size | engine     | status | median | min    | max    | time vs sheetjs | RSS footprint | mem vs sheetjs | post-warm-up delta | peak RSS | rows read | cells read |
| ---- | ---------- | ------ | ------ | ------ | ------ | --------------- | ------------- | -------------- | ------------------ | -------- | --------- | ---------- |
| 100k | sheetjs    | ok     | 4.09 s | 3.92 s | 4.40 s | —               | 1,123 MB      | —              | 173 MB             | 1,290 MB | 100,000   | 2,000,000  |
| 100k | office-kit | ok     | 4.99 s | 4.87 s | 5.01 s | 1.22x           | 320 MB        | 0.29x          | 70 MB              | 388 MB   | 100,000   | 2,000,000  |

## read-typed — numeric

Fixtures written by: sheetjs (see `fixture` in results.json per cell).

| size | engine     | status | median | min    | max    | time vs sheetjs | RSS footprint | mem vs sheetjs | post-warm-up delta | peak RSS | rows read | cells read |
| ---- | ---------- | ------ | ------ | ------ | ------ | --------------- | ------------- | -------------- | ------------------ | -------- | --------- | ---------- |
| 100k | sheetjs    | ok     | 3.16 s | 3.14 s | 3.21 s | —               | 904 MB        | —              | 197 MB             | 1,029 MB | 100,000   | 2,000,000  |
| 100k | office-kit | ok     | 4.68 s | 4.28 s | 4.83 s | 1.48x           | 355 MB        | 0.39x          | 36 MB              | 424 MB   | 100,000   | 2,000,000  |

## read-typed — wide

Fixtures written by: office-kit (see `fixture` in results.json per cell).

| size | engine     | status             | median  | min     | max     | time vs sheetjs | RSS footprint | mem vs sheetjs | post-warm-up delta | peak RSS | rows read | cells read |
| ---- | ---------- | ------------------ | ------- | ------- | ------- | --------------- | ------------- | -------------- | ------------------ | -------- | --------- | ---------- |
| 100k | sheetjs    | TIMEOUT (600.00 s) | —       | —       | —       | —               | —             | —              | —                  | —        | —         | —          |
| 100k | office-kit | ok                 | 63.47 s | 57.68 s | 65.43 s | —               | 5,424 MB      | —              | 4,015 MB           | 5,492 MB | 100,000   | 10,700,000 |

## read-raw — mixed

Fixtures written by: sheetjs (see `fixture` in results.json per cell).

| size | engine           | status                                                | median  | min     | max     | time vs sheetjs | RSS footprint | mem vs sheetjs | post-warm-up delta | peak RSS | rows read | cells read |
| ---- | ---------------- | ----------------------------------------------------- | ------- | ------- | ------- | --------------- | ------------- | -------------- | ------------------ | -------- | --------- | ---------- |
| 1k   | sheetjs          | ok                                                    | 74.6 ms | 72.8 ms | 75.9 ms | —               | 98 MB         | —              | 47 MB              | 164 MB   | 1,000     | 20,000     |
| 1k   | exceljs          | ok                                                    | 122 ms  | 57.8 ms | 124 ms  | 1.64x           | 40 MB         | 0.41x          | 18 MB              | 115 MB   | 1,000     | 20,000     |
| 1k   | office-kit       | ok                                                    | 84.5 ms | 82.2 ms | 90.0 ms | 1.13x           | 180 MB        | 1.83x          | 60 MB              | 249 MB   | 1,000     | 20,000     |
| 1k   | write-excel-file | skipped: write-excel-file does not implement read-raw | —       | —       | —       | —               | —             | —              | —                  | —        | —         | —          |
| 10k  | sheetjs          | ok                                                    | 623 ms  | 614 ms  | 699 ms  | —               | 393 MB        | —              | 78 MB              | 503 MB   | 10,000    | 200,000    |
| 10k  | exceljs          | ok                                                    | 1.03 s  | 339 ms  | 1.13 s  | 1.66x           | 143 MB        | 0.36x          | 88 MB              | 218 MB   | 10,000    | 200,000    |
| 10k  | office-kit       | ok                                                    | 666 ms  | 623 ms  | 702 ms  | 1.07x           | 612 MB        | 1.56x          | 23 MB              | 680 MB   | 10,000    | 200,000    |
| 10k  | write-excel-file | skipped: write-excel-file does not implement read-raw | —       | —       | —       | —               | —             | —              | —                  | —        | —         | —          |
| 100k | sheetjs          | ok                                                    | 6.54 s  | 6.53 s  | 6.74 s  | —               | 2,059 MB      | —              | 970 MB             | 2,314 MB | 100,000   | 2,000,000  |
| 100k | exceljs          | ok                                                    | 10.09 s | 3.16 s  | 11.33 s | 1.54x           | 204 MB        | 0.10x          | 9 MB               | 279 MB   | 100,000   | 2,000,000  |
| 100k | office-kit       | ok                                                    | 6.76 s  | 6.43 s  | 6.78 s  | 1.03x           | 493 MB        | 0.24x          | 143 MB             | 562 MB   | 100,000   | 2,000,000  |
| 100k | write-excel-file | skipped: write-excel-file does not implement read-raw | —       | —       | —       | —               | —             | —              | —                  | —        | —         | —          |

## Gates

### exceljs: FAIL

| gate                                                       | status  | detail                                  |
| ---------------------------------------------------------- | ------- | --------------------------------------- |
| write RSS footprint <= 0.25x sheetjs (mixed 100k)          | PASS    | 0.12x (349 MB vs 2,899 MB; limit 0.25x) |
| write median time <= 1.0x sheetjs (mixed 100k)             | PASS    | 0.29x (2.95 s vs 10.36 s; limit 1x)     |
| read-typed median time <= 1.0x sheetjs (mixed 100k)        | FAIL    | 1.53x (9.94 s vs 6.50 s; limit 1x)      |
| read-typed RSS footprint <= 0.5x sheetjs (mixed 100k)      | PASS    | 0.11x (209 MB vs 1,934 MB; limit 0.5x)  |
| write 1M x 20 (mixed 1m) succeeds                          | NOT-RUN | cell not in this run                    |
| write 900k x 20 (mixed 18m-cells) succeeds                 | NOT-RUN | cell not in this run                    |
| first byte reaches the sink in < 100 ms (mixed 100k write) | FAIL    | 1.90 s (limit 100 ms)                   |

### office-kit: FAIL

| gate                                                       | status | detail                                            |
| ---------------------------------------------------------- | ------ | ------------------------------------------------- |
| write RSS footprint <= 0.25x sheetjs (mixed 100k)          | FAIL   | 0.40x (1,158 MB vs 2,899 MB; limit 0.25x)         |
| write median time <= 1.0x sheetjs (mixed 100k)             | PASS   | 0.45x (4.62 s vs 10.36 s; limit 1x)               |
| read-typed median time <= 1.0x sheetjs (mixed 100k)        | FAIL   | 1.01x (6.55 s vs 6.50 s; limit 1x)                |
| read-typed RSS footprint <= 0.5x sheetjs (mixed 100k)      | PASS   | 0.27x (522 MB vs 1,934 MB; limit 0.5x)            |
| write 1M x 20 (mixed 1m) succeeds                          | FAIL   | ERROR (warm-up) RangeError: Invalid string length |
| write 900k x 20 (mixed 18m-cells) succeeds                 | FAIL   | ERROR (warm-up) RangeError: Invalid string length |
| first byte reaches the sink in < 100 ms (mixed 100k write) | PASS   | 2.8 ms (limit 100 ms)                             |

### write-excel-file: FAIL

| gate                                                       | status  | detail                                                |
| ---------------------------------------------------------- | ------- | ----------------------------------------------------- |
| write RSS footprint <= 0.25x sheetjs (mixed 100k)          | FAIL    | 0.40x (1,161 MB vs 2,899 MB; limit 0.25x)             |
| write median time <= 1.0x sheetjs (mixed 100k)             | PASS    | 0.36x (3.68 s vs 10.36 s; limit 1x)                   |
| read-typed median time <= 1.0x sheetjs (mixed 100k)        | FAIL    | 1.37x (8.93 s vs 6.50 s; limit 1x)                    |
| read-typed RSS footprint <= 0.5x sheetjs (mixed 100k)      | FAIL    | 2.48x (4,787 MB vs 1,934 MB; limit 0.5x)              |
| write 1M x 20 (mixed 1m) succeeds                          | NOT-RUN | cell not in this run                                  |
| write 900k x 20 (mixed 18m-cells) succeeds                 | NOT-RUN | cell not in this run                                  |
| first byte reaches the sink in < 100 ms (mixed 100k write) | FAIL    | engine reported no first-byte latency (not streaming) |

## Failures

- **sheetjs / write / mixed / 1m** — ERROR (warm-up) RangeError: Invalid string length

  ```
  RangeError: Invalid string length
      at Array.join (<anonymous>)
      at write_ws_xml_data (file:///Users/austinturner/dev/worktrees/jetstream/workspace-5/node_modules/.pnpm/xlsx@https+++cdn.sheetjs.com+xlsx-0.20.3+xlsx-0.20.3.tgz/node_modules/xlsx/xlsx.mjs:16417:11)
      at write_ws_xml (file:///Users/austinturner/dev/worktrees/jetstream/workspace-5/node_modules/.pnpm/xlsx@https+++cdn.sheetjs.com+xlsx-0.20.3+xlsx-0.20.3.tgz/node_modules/xlsx/xlsx.mjs:16458:11)
      at write_zip_xlsx (file:///Users/austinturner/dev/worktrees/jetstream/workspace-5/node_modules/.pnpm/xlsx@https+++cdn.sheetjs.com+xlsx-0.20.3+xlsx-0.20.3.tgz/node_modules/xlsx/xlsx.mjs:27056:25)
      at write_zip (file:///Users/austinturner/dev/worktrees/jetstream/workspace-5/node_modules/.pnpm/xlsx@https+++cdn.sheetjs.com+xlsx-0.20.3+xlsx-0.20.3.tgz/node_modules/xlsx/xlsx.mjs:27294:19)
  ```

- **office-kit / write / mixed / 1m** — ERROR (warm-up) RangeError: Invalid string length

  ```
  RangeError: Invalid string length
      at Array.join (<anonymous>)
      at serializeSharedStrings (file:///Users/austinturner/dev/worktrees/jetstream/workspace-5/node_modules/.pnpm/@office-kit+xlsx@0.11.0/node_modules/@office-kit/xlsx/dist/workbook-CARZS94U.mjs:203:15)
      at sharedStringsToBytes (file:///Users/austinturner/dev/worktrees/jetstream/workspace-5/node_modules/.pnpm/@office-kit+xlsx@0.11.0/node_modules/@office-kit/xlsx/dist/workbook-CARZS94U.mjs:196:34)
      at finalizeImpl (file:///Users/austinturner/dev/worktrees/jetstream/workspace-5/node_modules/.pnpm/@office-kit+xlsx@0.11.0/node_modules/@office-kit/xlsx/dist/streaming.mjs:555:78)
      at async Object.finalize (file:///Users/austinturner/dev/worktrees/jetstream/workspace-5/node_modules/.pnpm/@office-kit+xlsx@0.11.0/node_modules/@office-kit/xlsx/dist/streaming.mjs:536:4)
  ```

- **sheetjs / write / mixed / 18m-cells** — ERROR (warm-up) RangeError: Invalid string length

  ```
  RangeError: Invalid string length
      at Array.join (<anonymous>)
      at write_ws_xml_data (file:///Users/austinturner/dev/worktrees/jetstream/workspace-5/node_modules/.pnpm/xlsx@https+++cdn.sheetjs.com+xlsx-0.20.3+xlsx-0.20.3.tgz/node_modules/xlsx/xlsx.mjs:16417:11)
      at write_ws_xml (file:///Users/austinturner/dev/worktrees/jetstream/workspace-5/node_modules/.pnpm/xlsx@https+++cdn.sheetjs.com+xlsx-0.20.3+xlsx-0.20.3.tgz/node_modules/xlsx/xlsx.mjs:16458:11)
      at write_zip_xlsx (file:///Users/austinturner/dev/worktrees/jetstream/workspace-5/node_modules/.pnpm/xlsx@https+++cdn.sheetjs.com+xlsx-0.20.3+xlsx-0.20.3.tgz/node_modules/xlsx/xlsx.mjs:27056:25)
      at write_zip (file:///Users/austinturner/dev/worktrees/jetstream/workspace-5/node_modules/.pnpm/xlsx@https+++cdn.sheetjs.com+xlsx-0.20.3+xlsx-0.20.3.tgz/node_modules/xlsx/xlsx.mjs:27294:19)
  ```

- **office-kit / write / mixed / 18m-cells** — ERROR (warm-up) RangeError: Invalid string length

  ```
  RangeError: Invalid string length
      at Array.join (<anonymous>)
      at serializeSharedStrings (file:///Users/austinturner/dev/worktrees/jetstream/workspace-5/node_modules/.pnpm/@office-kit+xlsx@0.11.0/node_modules/@office-kit/xlsx/dist/workbook-CARZS94U.mjs:203:15)
      at sharedStringsToBytes (file:///Users/austinturner/dev/worktrees/jetstream/workspace-5/node_modules/.pnpm/@office-kit+xlsx@0.11.0/node_modules/@office-kit/xlsx/dist/workbook-CARZS94U.mjs:196:34)
      at finalizeImpl (file:///Users/austinturner/dev/worktrees/jetstream/workspace-5/node_modules/.pnpm/@office-kit+xlsx@0.11.0/node_modules/@office-kit/xlsx/dist/streaming.mjs:555:78)
      at async Object.finalize (file:///Users/austinturner/dev/worktrees/jetstream/workspace-5/node_modules/.pnpm/@office-kit+xlsx@0.11.0/node_modules/@office-kit/xlsx/dist/streaming.mjs:536:4)
  ```

- **sheetjs / write / wide / 100k** — ERROR (process) ChildProcessError: child exited with code null signal SIGABRT and no result; stderr tail: …
- **sheetjs / read-typed / wide / 100k** — TIMEOUT (600.00 s)
- **sheetjs / write / mixed / 250k** — ERROR (warm-up) RangeError: Invalid string length

  ```
  RangeError: Invalid string length
      at Array.join (<anonymous>)
      at write_ws_xml_data (file:///Users/austinturner/dev/worktrees/jetstream/workspace-5/node_modules/.pnpm/xlsx@https+++cdn.sheetjs.com+xlsx-0.20.3+xlsx-0.20.3.tgz/node_modules/xlsx/xlsx.mjs:16417:11)
      at write_ws_xml (file:///Users/austinturner/dev/worktrees/jetstream/workspace-5/node_modules/.pnpm/xlsx@https+++cdn.sheetjs.com+xlsx-0.20.3+xlsx-0.20.3.tgz/node_modules/xlsx/xlsx.mjs:16458:11)
      at write_zip_xlsx (file:///Users/austinturner/dev/worktrees/jetstream/workspace-5/node_modules/.pnpm/xlsx@https+++cdn.sheetjs.com+xlsx-0.20.3+xlsx-0.20.3.tgz/node_modules/xlsx/xlsx.mjs:27056:25)
      at write_zip (file:///Users/austinturner/dev/worktrees/jetstream/workspace-5/node_modules/.pnpm/xlsx@https+++cdn.sheetjs.com+xlsx-0.20.3+xlsx-0.20.3.tgz/node_modules/xlsx/xlsx.mjs:27294:19)
  ```

- **office-kit / write / mixed / 500k** — ERROR (warm-up) RangeError: Invalid string length

  ```
  RangeError: Invalid string length
      at Array.join (<anonymous>)
      at serializeSharedStrings (file:///Users/austinturner/dev/worktrees/jetstream/workspace-5/node_modules/.pnpm/@office-kit+xlsx@0.11.0/node_modules/@office-kit/xlsx/dist/workbook-CARZS94U.mjs:203:15)
      at sharedStringsToBytes (file:///Users/austinturner/dev/worktrees/jetstream/workspace-5/node_modules/.pnpm/@office-kit+xlsx@0.11.0/node_modules/@office-kit/xlsx/dist/workbook-CARZS94U.mjs:196:34)
      at finalizeImpl (file:///Users/austinturner/dev/worktrees/jetstream/workspace-5/node_modules/.pnpm/@office-kit+xlsx@0.11.0/node_modules/@office-kit/xlsx/dist/streaming.mjs:555:78)
      at async Object.finalize (file:///Users/austinturner/dev/worktrees/jetstream/workspace-5/node_modules/.pnpm/@office-kit+xlsx@0.11.0/node_modules/@office-kit/xlsx/dist/streaming.mjs:536:4)
  ```

- **sheetjs / write / mixed / 500k** — ERROR (warm-up) RangeError: Invalid string length

  ```
  RangeError: Invalid string length
      at Array.join (<anonymous>)
      at write_ws_xml_data (file:///Users/austinturner/dev/worktrees/jetstream/workspace-5/node_modules/.pnpm/xlsx@https+++cdn.sheetjs.com+xlsx-0.20.3+xlsx-0.20.3.tgz/node_modules/xlsx/xlsx.mjs:16417:11)
      at write_ws_xml (file:///Users/austinturner/dev/worktrees/jetstream/workspace-5/node_modules/.pnpm/xlsx@https+++cdn.sheetjs.com+xlsx-0.20.3+xlsx-0.20.3.tgz/node_modules/xlsx/xlsx.mjs:16458:11)
      at write_zip_xlsx (file:///Users/austinturner/dev/worktrees/jetstream/workspace-5/node_modules/.pnpm/xlsx@https+++cdn.sheetjs.com+xlsx-0.20.3+xlsx-0.20.3.tgz/node_modules/xlsx/xlsx.mjs:27056:25)
      at write_zip (file:///Users/austinturner/dev/worktrees/jetstream/workspace-5/node_modules/.pnpm/xlsx@https+++cdn.sheetjs.com+xlsx-0.20.3+xlsx-0.20.3.tgz/node_modules/xlsx/xlsx.mjs:27294:19)
  ```

<!-- generated:end -->

## Reading the results (runs `2026-09-11-macbook-air-{baseline,scale,ceiling,shapes,chrome}`; `combined` merges the Node runs)

- **SheetJS** reproduces every production failure on demand: 100k×20 `mixed` takes 10.4 s and 2.9 GB, the
  `wide` shape (100k × 107 = 10.7M cells) dies with SIGABRT, and both 1M×20 and the 18M-cell shape throw
  `RangeError: Invalid string length` in `write_ws_xml_data`. Reading the `wide` file timed out at 600 s.
- **ExcelJS streaming** is the only writer that passes the 100k memory gate (349 MB, 0.12×) because it really
  streams to the sink (first byte at 1.9 s, output grows as rows are written); it is also the fastest writer
  (2.95 s, 0.29×). It is unmaintained and reads slowly (1.55× SheetJS at 10k).
- **office-kit** is 2.2× faster than SheetJS on writes and flat on numeric data (54 MB), but on
  Jetstream-shaped data it holds the worksheet XML and an unbounded shared string table until `finalize()`:
  1,158 MB at 100k×20 (0.40×, gate 0.25×) and the same `Invalid string length` as SheetJS at 1M×20 and 18M cells
  (in `serializeSharedStrings`). Reads use 0.29-0.39× the memory but 1.2-1.5× the time. Details and the
  300k-row sink probe are in 10.
- **write-excel-file** tracks office-kit closely on writes (it also buffers) and is write-only.
- **Ceiling run** (`mixed`, write): SheetJS already fails at 250k rows; office-kit succeeds at 250k (11.7 s,
  2.4 GB) and fails at 500k - roughly twice SheetJS's ceiling, still far below the 1M-row target.
- **Chrome module worker** (`2026-09-11-macbook-air-chrome`, Chromium 151 headless): office-kit writes 100k×20
  in 6.4 s with +1.6 GB renderer RSS (SheetJS 8.9 s, +2.9 GB); both crash the renderer at 1M rows.
  `performance.measureUserAgentSpecificMemory()` never resolved inside the runs, so renderer RSS from `ps`
  is the usable number (TODO in `bench/README.md`).
- Time is dominated by XML string building and shared-string bookkeeping, not deflate: the numeric dataset
  (no strings) writes in 1.84 s with office-kit and 3.09 s with SheetJS.
- Load average was 3-8 during the runs (other processes were active); ratios are more trustworthy than
  absolute numbers. The Chrome (worker) path is documented in `bench/README.md`; see the bench folder for
  whether a Chrome run has been committed.
