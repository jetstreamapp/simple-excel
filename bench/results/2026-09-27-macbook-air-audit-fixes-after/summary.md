# xlsx engine benchmark — audit-fixes-after

Generated 2026-09-27T18:25:39.545Z. Results folder: `bench/results/2026-09-27-macbook-air-audit-fixes-after`.

## Machine

- Apple M4 (10 cores), 32 GB RAM, darwin 25.5.0 arm64
- Node v24.21.0 (V8 13.6.233.17-node.53), child flags: `--expose-gc --max-old-space-size=12288`
- Load average at start 2.68 / 3.47 / 5.26, at end 2.93 / 3.44 / 5.16 (other processes may have been running)
- Options: runs=3, warmup=1, source=materialized, timeout=600s, seed=20260911

## Engines

| engine            | version | streaming | note |
| ----------------- | ------- | --------- | ---- |
| simple-excel      | 0.1.1   | yes       |      |
| simple-excel-zlib | 0.1.1   | yes       |      |

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
| 100k | simple-excel      | ok     | 2.93 s | 2.83 s | 2.95 s | —               | 108 MB        | —              | 45 MB              | 695 MB   | 43.0 MB | 0.3 ms     |
| 100k | simple-excel-zlib | ok     | 1.72 s | 1.70 s | 1.73 s | —               | 98 MB         | —              | 41 MB              | 684 MB   | 51.2 MB | 0.3 ms     |

## read-typed — mixed

Fixtures written by: sheetjs (see `fixture` in results.json per cell).

| size | engine            | status | median | min    | max    | time vs sheetjs | RSS footprint | mem vs sheetjs | post-warm-up delta | peak RSS | rows read | cells read |
| ---- | ----------------- | ------ | ------ | ------ | ------ | --------------- | ------------- | -------------- | ------------------ | -------- | --------- | ---------- |
| 100k | simple-excel      | ok     | 1.61 s | 1.61 s | 1.63 s | —               | 910 MB        | —              | 44 MB              | 1,153 MB | 100,000   | 2,000,000  |
| 100k | simple-excel-zlib | ok     | 1.66 s | 1.61 s | 1.67 s | —               | 902 MB        | —              | 30 MB              | 1,145 MB | 100,000   | 2,000,000  |

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
| first byte reaches the sink in < 100 ms (mixed 100k write) | PASS    | 0.3 ms (limit 100 ms)            |

### simple-excel-zlib: INCOMPLETE

| gate                                                       | status  | detail                           |
| ---------------------------------------------------------- | ------- | -------------------------------- |
| write RSS footprint <= 0.25x sheetjs (mixed 100k)          | NOT-RUN | sheetjs baseline not in this run |
| write median time <= 1.0x sheetjs (mixed 100k)             | NOT-RUN | sheetjs baseline not in this run |
| read-typed median time <= 1.0x sheetjs (mixed 100k)        | NOT-RUN | sheetjs baseline not in this run |
| read-typed RSS footprint <= 0.5x sheetjs (mixed 100k)      | NOT-RUN | sheetjs baseline not in this run |
| write 1M x 20 (mixed 1m) succeeds                          | NOT-RUN | cell not in this run             |
| write 900k x 20 (mixed 18m-cells) succeeds                 | NOT-RUN | cell not in this run             |
| first byte reaches the sink in < 100 ms (mixed 100k write) | PASS    | 0.3 ms (limit 100 ms)            |
