# xlsx engine benchmark — phase-e-others

Generated 2026-09-12T20:31:43.399Z. Results folder: `bench/results/2026-09-12-macbook-air-phase-e-others`.

## Machine

- Apple M4 (10 cores), 32 GB RAM, darwin 25.5.0 arm64
- Node v24.18.0 (V8 13.6.233.17-node.50), child flags: `--expose-gc --max-old-space-size=12288`
- Load average at start 3.83 / 4.36 / 4.26, at end 3.33 / 3.91 / 4.09 (other processes may have been running)
- Options: runs=3, warmup=1, source=materialized, timeout=600s, seed=20260911

## Engines

| engine           | version                        | streaming | note |
| ---------------- | ------------------------------ | --------- | ---- |
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

| size | engine           | status | median | min    | max    | time vs sheetjs | RSS footprint | mem vs sheetjs | post-warm-up delta | peak RSS | output  | first byte |
| ---- | ---------------- | ------ | ------ | ------ | ------ | --------------- | ------------- | -------------- | ------------------ | -------- | ------- | ---------- |
| 100k | exceljs          | ok     | 2.49 s | 2.47 s | 2.53 s | —               | 328 MB        | —              | 3 MB               | 923 MB   | 52.3 MB | 1.61 s     |
| 100k | office-kit       | ok     | 4.36 s | 4.35 s | 4.40 s | —               | 1,241 MB      | —              | 501 MB             | 1,832 MB | 41.2 MB | 3.1 ms     |
| 100k | write-excel-file | ok     | 3.21 s | 3.20 s | 3.22 s | —               | 1,139 MB      | —              | 556 MB             | 1,731 MB | 48.2 MB | —          |

## read-typed — mixed

Fixtures written by: sheetjs (see `fixture` in results.json per cell).

| size | engine           | status | median | min    | max    | time vs sheetjs | RSS footprint | mem vs sheetjs | post-warm-up delta | peak RSS | rows read | cells read |
| ---- | ---------------- | ------ | ------ | ------ | ------ | --------------- | ------------- | -------------- | ------------------ | -------- | --------- | ---------- |
| 100k | exceljs          | ok     | 9.53 s | 9.51 s | 9.67 s | —               | 193 MB        | —              | 30 MB              | 268 MB   | 100,000   | 2,000,000  |
| 100k | office-kit       | ok     | 6.57 s | 6.45 s | 6.78 s | —               | 449 MB        | —              | 105 MB             | 517 MB   | 100,000   | 2,000,000  |
| 100k | write-excel-file | ok     | 6.80 s | 6.67 s | 6.86 s | —               | 5,290 MB      | —              | 1,083 MB           | 5,360 MB | 100,000   | 2,000,000  |

## read-raw — mixed

Fixtures written by: sheetjs (see `fixture` in results.json per cell).

| size | engine           | status                                                | median | min    | max    | time vs sheetjs | RSS footprint | mem vs sheetjs | post-warm-up delta | peak RSS | rows read | cells read |
| ---- | ---------------- | ----------------------------------------------------- | ------ | ------ | ------ | --------------- | ------------- | -------------- | ------------------ | -------- | --------- | ---------- |
| 100k | exceljs          | ok                                                    | 9.58 s | 3.14 s | 9.61 s | —               | 200 MB        | —              | 12 MB              | 275 MB   | 100,000   | 2,000,000  |
| 100k | office-kit       | ok                                                    | 6.64 s | 6.30 s | 6.70 s | —               | 533 MB        | —              | 186 MB             | 601 MB   | 100,000   | 2,000,000  |
| 100k | write-excel-file | skipped: write-excel-file does not implement read-raw | —      | —      | —      | —               | —             | —              | —                  | —        | —         | —          |

## Gates

### exceljs: FAIL

| gate                                                       | status  | detail                           |
| ---------------------------------------------------------- | ------- | -------------------------------- |
| write RSS footprint <= 0.25x sheetjs (mixed 100k)          | NOT-RUN | sheetjs baseline not in this run |
| write median time <= 1.0x sheetjs (mixed 100k)             | NOT-RUN | sheetjs baseline not in this run |
| read-typed median time <= 1.0x sheetjs (mixed 100k)        | NOT-RUN | sheetjs baseline not in this run |
| read-typed RSS footprint <= 0.5x sheetjs (mixed 100k)      | NOT-RUN | sheetjs baseline not in this run |
| write 1M x 20 (mixed 1m) succeeds                          | NOT-RUN | cell not in this run             |
| write 900k x 20 (mixed 18m-cells) succeeds                 | NOT-RUN | cell not in this run             |
| first byte reaches the sink in < 100 ms (mixed 100k write) | FAIL    | 1.61 s (limit 100 ms)            |

### office-kit: INCOMPLETE

| gate                                                       | status  | detail                           |
| ---------------------------------------------------------- | ------- | -------------------------------- |
| write RSS footprint <= 0.25x sheetjs (mixed 100k)          | NOT-RUN | sheetjs baseline not in this run |
| write median time <= 1.0x sheetjs (mixed 100k)             | NOT-RUN | sheetjs baseline not in this run |
| read-typed median time <= 1.0x sheetjs (mixed 100k)        | NOT-RUN | sheetjs baseline not in this run |
| read-typed RSS footprint <= 0.5x sheetjs (mixed 100k)      | NOT-RUN | sheetjs baseline not in this run |
| write 1M x 20 (mixed 1m) succeeds                          | NOT-RUN | cell not in this run             |
| write 900k x 20 (mixed 18m-cells) succeeds                 | NOT-RUN | cell not in this run             |
| first byte reaches the sink in < 100 ms (mixed 100k write) | PASS    | 3.1 ms (limit 100 ms)            |

### write-excel-file: FAIL

| gate                                                       | status  | detail                                                |
| ---------------------------------------------------------- | ------- | ----------------------------------------------------- |
| write RSS footprint <= 0.25x sheetjs (mixed 100k)          | NOT-RUN | sheetjs baseline not in this run                      |
| write median time <= 1.0x sheetjs (mixed 100k)             | NOT-RUN | sheetjs baseline not in this run                      |
| read-typed median time <= 1.0x sheetjs (mixed 100k)        | NOT-RUN | sheetjs baseline not in this run                      |
| read-typed RSS footprint <= 0.5x sheetjs (mixed 100k)      | NOT-RUN | sheetjs baseline not in this run                      |
| write 1M x 20 (mixed 1m) succeeds                          | NOT-RUN | cell not in this run                                  |
| write 900k x 20 (mixed 18m-cells) succeeds                 | NOT-RUN | cell not in this run                                  |
| first byte reaches the sink in < 100 ms (mixed 100k write) | FAIL    | engine reported no first-byte latency (not streaming) |
