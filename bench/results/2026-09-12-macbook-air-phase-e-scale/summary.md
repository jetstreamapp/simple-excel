# xlsx engine benchmark — phase-e-scale

Generated 2026-09-12T20:26:02.729Z. Results folder: `bench/results/2026-09-12-macbook-air-phase-e-scale`.

## Machine

- Apple M4 (10 cores), 32 GB RAM, darwin 25.5.0 arm64
- Node v24.18.0 (V8 13.6.233.17-node.50), child flags: `--expose-gc --max-old-space-size=12288`
- Load average at start 3.18 / 3.82 / 4.07, at end 4.60 / 4.54 / 4.32 (other processes may have been running)
- Options: runs=1, warmup=1, source=materialized, timeout=600s, seed=20260911

## Engines

| engine            | version | streaming | note |
| ----------------- | ------- | --------- | ---- |
| simple-excel      | 0.0.0   | yes       |      |
| simple-excel-zlib | 0.0.0   | yes       |      |
| sheetjs           | 0.20.3  | no        |      |

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
| 1m        | simple-excel      | ok                                                | 31.95 s | 31.95 s | 31.95 s | —               | 328 MB        | —              | 70 MB              | 3,489 MB | 427.0 MB | 0.9 ms     |
| 1m        | simple-excel-zlib | ok                                                | 19.09 s | 19.09 s | 19.09 s | —               | 650 MB        | —              | 66 MB              | 3,837 MB | 511.3 MB | 0.4 ms     |
| 1m        | sheetjs           | ERROR (warm-up) RangeError: Invalid string length | —       | —       | —       | —               | —             | —              | —                  | —        | —        | —          |
| 18m-cells | simple-excel      | ok                                                | 27.02 s | 27.02 s | 27.02 s | —               | 327 MB        | —              | 46 MB              | 4,158 MB | 383.8 MB | 0.8 ms     |
| 18m-cells | simple-excel-zlib | ok                                                | 15.55 s | 15.55 s | 15.55 s | —               | 322 MB        | —              | 42 MB              | 4,152 MB | 459.7 MB | 0.4 ms     |
| 18m-cells | sheetjs           | ERROR (warm-up) RangeError: Invalid string length | —       | —       | —       | —               | —             | —              | —                  | —        | —        | —          |

## Gates

### simple-excel: INCOMPLETE

| gate                                                       | status  | detail                        |
| ---------------------------------------------------------- | ------- | ----------------------------- |
| write RSS footprint <= 0.25x sheetjs (mixed 100k)          | NOT-RUN | cell not in this run          |
| write median time <= 1.0x sheetjs (mixed 100k)             | NOT-RUN | cell not in this run          |
| read-typed median time <= 1.0x sheetjs (mixed 100k)        | NOT-RUN | cell not in this run          |
| read-typed RSS footprint <= 0.5x sheetjs (mixed 100k)      | NOT-RUN | cell not in this run          |
| write 1M x 20 (mixed 1m) succeeds                          | PASS    | 31.95 s, 328 MB RSS footprint |
| write 900k x 20 (mixed 18m-cells) succeeds                 | PASS    | 27.02 s, 327 MB RSS footprint |
| first byte reaches the sink in < 100 ms (mixed 100k write) | NOT-RUN | cell not in this run          |

### simple-excel-zlib: INCOMPLETE

| gate                                                       | status  | detail                        |
| ---------------------------------------------------------- | ------- | ----------------------------- |
| write RSS footprint <= 0.25x sheetjs (mixed 100k)          | NOT-RUN | cell not in this run          |
| write median time <= 1.0x sheetjs (mixed 100k)             | NOT-RUN | cell not in this run          |
| read-typed median time <= 1.0x sheetjs (mixed 100k)        | NOT-RUN | cell not in this run          |
| read-typed RSS footprint <= 0.5x sheetjs (mixed 100k)      | NOT-RUN | cell not in this run          |
| write 1M x 20 (mixed 1m) succeeds                          | PASS    | 19.09 s, 650 MB RSS footprint |
| write 900k x 20 (mixed 18m-cells) succeeds                 | PASS    | 15.55 s, 322 MB RSS footprint |
| first byte reaches the sink in < 100 ms (mixed 100k write) | NOT-RUN | cell not in this run          |

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
