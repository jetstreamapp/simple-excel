# xlsx engine benchmark — shapes

Generated 2026-09-12T04:00:59.261Z. Results folder: `bench/results/2026-09-11-macbook-air-shapes`.

## Machine

- Apple M4 (10 cores), 32 GB RAM, darwin 25.5.0 arm64
- Node v24.18.0 (V8 13.6.233.17-node.50), child flags: `--expose-gc --max-old-space-size=12288`
- Load average at start 5.39 / 6.31 / 4.94, at end 7.09 / 9.99 / 9.25 (other processes may have been running)
- Options: runs=3, warmup=1, source=materialized, timeout=600s, seed=20260911

## Engines

| engine     | version | streaming | note |
| ---------- | ------- | --------- | ---- |
| sheetjs    | 0.20.3  | no        |      |
| office-kit | 0.11.0  | yes       |      |

## Metrics

- `median`/`min`/`max`: wall time of the timed runs (after one warm-up run of the same cell).
- `RSS footprint`: highest RSS observed while the op ran (25 ms poller, post-run read, process high-water mark) minus the RSS before the warm-up (engine loaded, input data resident, gc() done). This is the ratio/gate metric.
- `post-warm-up delta`: the same peak minus the RSS after warm-up + gc(). Undercounts engines whose warm-up footprint stays resident (V8 rarely returns pages), kept for reference.
- `peak RSS`: the absolute high-water RSS of the child process (input data included).
- `first byte`: ms from the start of `write()` until the sink receives its first chunk (streaming writers only).
- Ratios are candidate ÷ sheetjs for the same op/dataset/size; lower is better.

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

## Gates

### office-kit: INCOMPLETE

| gate                                                       | status  | detail               |
| ---------------------------------------------------------- | ------- | -------------------- |
| write RSS footprint <= 0.25x sheetjs (mixed 100k)          | NOT-RUN | cell not in this run |
| write median time <= 1.0x sheetjs (mixed 100k)             | NOT-RUN | cell not in this run |
| read-typed median time <= 1.0x sheetjs (mixed 100k)        | NOT-RUN | cell not in this run |
| read-typed RSS footprint <= 0.5x sheetjs (mixed 100k)      | NOT-RUN | cell not in this run |
| write 1M x 20 (mixed 1m) succeeds                          | NOT-RUN | cell not in this run |
| write 900k x 20 (mixed 18m-cells) succeeds                 | NOT-RUN | cell not in this run |
| first byte reaches the sink in < 100 ms (mixed 100k write) | NOT-RUN | cell not in this run |

## Failures

- **sheetjs / write / wide / 100k** — ERROR (process) ChildProcessError: child exited with code null signal SIGABRT and no result; stderr tail: …
- **sheetjs / read-typed / wide / 100k** — TIMEOUT (600.00 s)
