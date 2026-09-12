# xlsx engine benchmark — inline-default

Generated 2026-09-12T22:44:37.775Z. Results folder: `bench/results/2026-09-12-macbook-air-inline-default`.

## Machine

- Apple M4 (10 cores), 32 GB RAM, darwin 25.5.0 arm64
- Node v24.18.0 (V8 13.6.233.17-node.50), child flags: `--expose-gc --max-old-space-size=12288`
- Load average at start 6.31 / 7.21 / 6.52, at end 5.47 / 6.82 / 6.42 (other processes may have been running)
- Options: runs=3, warmup=1, source=materialized, timeout=600s, seed=20260911

## Engines

| engine            | version | streaming | note |
| ----------------- | ------- | --------- | ---- |
| simple-excel      | 0.0.0   | yes       |      |
| simple-excel-zlib | 0.0.0   | yes       |      |

## Metrics

- `median`/`min`/`max`: wall time of the timed runs (after one warm-up run of the same cell).
- `RSS footprint`: highest RSS observed while the op ran (25 ms poller, post-run read, process high-water mark) minus the RSS before the warm-up (engine loaded, input data resident, gc() done). This is the ratio/gate metric.
- `post-warm-up delta`: the same peak minus the RSS after warm-up + gc(). Undercounts engines whose warm-up footprint stays resident (V8 rarely returns pages), kept for reference.
- `peak RSS`: the absolute high-water RSS of the child process (input data included).
- `first byte`: ms from the start of `write()` until the sink receives its first chunk (streaming writers only).
- Ratios are candidate ÷ sheetjs for the same op/dataset/size; lower is better.

## write — mixed

| size | engine            | status | median | min    | max    | time vs sheetjs | RSS footprint | mem vs sheetjs | post-warm-up delta | peak RSS | output  | first byte |
| ---- | ----------------- | ------ | ------ | ------ | ------ | --------------- | ------------- | -------------- | ------------------ | -------- | ------- | ---------- |
| 100k | simple-excel      | ok     | 3.08 s | 3.00 s | 3.31 s | —               | 58 MB         | —              | 48 MB              | 611 MB   | 43.0 MB | 0.5 ms     |
| 100k | simple-excel-zlib | ok     | 1.73 s | 1.67 s | 1.75 s | —               | 74 MB         | —              | 73 MB              | 660 MB   | 51.2 MB | 0.2 ms     |

## read-typed — mixed

Fixtures written by: sheetjs (see `fixture` in results.json per cell).

| size | engine            | status | median | min    | max    | time vs sheetjs | RSS footprint | mem vs sheetjs | post-warm-up delta | peak RSS | rows read | cells read |
| ---- | ----------------- | ------ | ------ | ------ | ------ | --------------- | ------------- | -------------- | ------------------ | -------- | --------- | ---------- |
| 100k | simple-excel      | ok     | 1.64 s | 1.62 s | 1.72 s | —               | 810 MB        | —              | 255 MB             | 1,053 MB | 100,000   | 2,000,000  |
| 100k | simple-excel-zlib | ok     | 1.69 s | 1.65 s | 1.84 s | —               | 813 MB        | —              | 40 MB              | 1,056 MB | 100,000   | 2,000,000  |

## read-raw — mixed

Fixtures written by: sheetjs (see `fixture` in results.json per cell).

| size | engine            | status | median | min    | max    | time vs sheetjs | RSS footprint | mem vs sheetjs | post-warm-up delta | peak RSS | rows read | cells read |
| ---- | ----------------- | ------ | ------ | ------ | ------ | --------------- | ------------- | -------------- | ------------------ | -------- | --------- | ---------- |
| 100k | simple-excel      | ok     | 1.48 s | 1.47 s | 1.48 s | —               | 46 MB         | —              | 95 MB              | 289 MB   | 100,000   | 2,000,000  |
| 100k | simple-excel-zlib | ok     | 1.52 s | 1.48 s | 1.58 s | —               | 44 MB         | —              | 95 MB              | 287 MB   | 100,000   | 2,000,000  |

## Gates

### simple-excel: INCOMPLETE

| gate                                                       | status  | detail                           |
| ---------------------------------------------------------- | ------- | -------------------------------- |
| write RSS footprint <= 0.25x sheetjs (mixed 100k)          | NOT-RUN | sheetjs baseline not in this run |
| write median time <= 1.0x sheetjs (mixed 100k)             | NOT-RUN | sheetjs baseline not in this run |
| read-typed median time <= 1.0x sheetjs (mixed 100k)        | NOT-RUN | sheetjs baseline not in this run |
| read-typed RSS footprint <= 0.5x sheetjs (mixed 100k)      | NOT-RUN | sheetjs baseline not in this run |
| write 1M x 20 (mixed 1m) succeeds                          | NOT-RUN | cell not in this run             |
| write 900k x 20 (mixed 18m-cells) succeeds                 | NOT-RUN | cell not in this run             |
| first byte reaches the sink in < 100 ms (mixed 100k write) | PASS    | 0.5 ms (limit 100 ms)            |

### simple-excel-zlib: INCOMPLETE

| gate                                                       | status  | detail                           |
| ---------------------------------------------------------- | ------- | -------------------------------- |
| write RSS footprint <= 0.25x sheetjs (mixed 100k)          | NOT-RUN | sheetjs baseline not in this run |
| write median time <= 1.0x sheetjs (mixed 100k)             | NOT-RUN | sheetjs baseline not in this run |
| read-typed median time <= 1.0x sheetjs (mixed 100k)        | NOT-RUN | sheetjs baseline not in this run |
| read-typed RSS footprint <= 0.5x sheetjs (mixed 100k)      | NOT-RUN | sheetjs baseline not in this run |
| write 1M x 20 (mixed 1m) succeeds                          | NOT-RUN | cell not in this run             |
| write 900k x 20 (mixed 18m-cells) succeeds                 | NOT-RUN | cell not in this run             |
| first byte reaches the sink in < 100 ms (mixed 100k write) | PASS    | 0.2 ms (limit 100 ms)            |
