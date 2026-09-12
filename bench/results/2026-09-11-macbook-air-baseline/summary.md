# xlsx engine benchmark — baseline

Generated 2026-09-12T03:52:00.584Z. Results folder: `bench/results/2026-09-11-macbook-air-baseline`.

## Machine

- Apple M4 (10 cores), 32 GB RAM, darwin 25.5.0 arm64
- Node v24.18.0 (V8 13.6.233.17-node.50), child flags: `--expose-gc --max-old-space-size=12288`
- Load average at start 3.24 / 3.38 / 3.30, at end 8.38 / 6.13 / 4.53 (other processes may have been running)
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

| size | engine           | status | median  | min     | max     | time vs sheetjs | RSS footprint | mem vs sheetjs | post-warm-up delta | peak RSS | output  | first byte |
| ---- | ---------------- | ------ | ------- | ------- | ------- | --------------- | ------------- | -------------- | ------------------ | -------- | ------- | ---------- |
| 1k   | sheetjs          | ok     | 96.2 ms | 69.7 ms | 99.6 ms | —               | 172 MB        | —              | 137 MB             | 248 MB   | 2.5 MB  | —          |
| 1k   | exceljs          | ok     | 38.2 ms | 36.0 ms | 39.0 ms | 0.40x           | 27 MB         | 0.16x          | 3 MB               | 117 MB   | 543 KB  | 25.2 ms    |
| 1k   | office-kit       | ok     | 56.7 ms | 56.3 ms | 59.9 ms | 0.59x           | 30 MB         | 0.17x          | 14 MB              | 111 MB   | 435 KB  | 5.4 ms     |
| 1k   | write-excel-file | ok     | 52.5 ms | 45.2 ms | 56.0 ms | 0.55x           | 36 MB         | 0.21x          | 25 MB              | 122 MB   | 503 KB  | —          |
| 10k  | sheetjs          | ok     | 807 ms  | 803 ms  | 1.10 s  | —               | 1,087 MB      | —              | 906 MB             | 1,258 MB | 24.9 MB | —          |
| 10k  | exceljs          | ok     | 271 ms  | 270 ms  | 276 ms  | 0.34x           | 107 MB        | 0.10x          | 2 MB               | 284 MB   | 5.2 MB  | 174 ms     |
| 10k  | office-kit       | ok     | 436 ms  | 433 ms  | 446 ms  | 0.54x           | 183 MB        | 0.17x          | 65 MB              | 355 MB   | 4.1 MB  | 4.0 ms     |
| 10k  | write-excel-file | ok     | 325 ms  | 319 ms  | 344 ms  | 0.40x           | 188 MB        | 0.17x          | 49 MB              | 362 MB   | 4.8 MB  | —          |
| 100k | sheetjs          | ok     | 10.36 s | 9.50 s  | 11.47 s | —               | 2,899 MB      | —              | 2,682 MB           | 3,490 MB | 95.2 MB | —          |
| 100k | exceljs          | ok     | 2.95 s  | 2.55 s  | 3.29 s  | 0.29x           | 349 MB        | 0.12x          | 31 MB              | 943 MB   | 52.3 MB | 1.90 s     |
| 100k | office-kit       | ok     | 4.62 s  | 4.51 s  | 4.67 s  | 0.45x           | 1,158 MB      | 0.40x          | 517 MB             | 1,749 MB | 41.2 MB | 2.8 ms     |
| 100k | write-excel-file | ok     | 3.68 s  | 3.40 s  | 3.78 s  | 0.36x           | 1,161 MB      | 0.40x          | 632 MB             | 1,754 MB | 48.2 MB | —          |

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

| gate                                                       | status  | detail                                    |
| ---------------------------------------------------------- | ------- | ----------------------------------------- |
| write RSS footprint <= 0.25x sheetjs (mixed 100k)          | FAIL    | 0.40x (1,158 MB vs 2,899 MB; limit 0.25x) |
| write median time <= 1.0x sheetjs (mixed 100k)             | PASS    | 0.45x (4.62 s vs 10.36 s; limit 1x)       |
| read-typed median time <= 1.0x sheetjs (mixed 100k)        | FAIL    | 1.01x (6.55 s vs 6.50 s; limit 1x)        |
| read-typed RSS footprint <= 0.5x sheetjs (mixed 100k)      | PASS    | 0.27x (522 MB vs 1,934 MB; limit 0.5x)    |
| write 1M x 20 (mixed 1m) succeeds                          | NOT-RUN | cell not in this run                      |
| write 900k x 20 (mixed 18m-cells) succeeds                 | NOT-RUN | cell not in this run                      |
| first byte reaches the sink in < 100 ms (mixed 100k write) | PASS    | 2.8 ms (limit 100 ms)                     |

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
