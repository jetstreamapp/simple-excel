# xlsx engine benchmark — phase-d-scale

Generated 2026-09-12T18:48:09.831Z. Results folder: `bench/results/2026-09-12-macbook-air-phase-d-scale`.

## Machine

- Apple M4 (10 cores), 32 GB RAM, darwin 25.5.0 arm64
- Node v24.18.0 (V8 13.6.233.17-node.50), child flags: `--expose-gc --max-old-space-size=12288`
- Load average at start 4.49 / 4.72 / 5.85, at end 3.33 / 4.03 / 5.12 (other processes may have been running)
- Options: runs=1, warmup=1, source=materialized, timeout=600s, seed=20260911

## Engines

| engine            | version | streaming | note |
| ----------------- | ------- | --------- | ---- |
| sheetjs           | 0.20.3  | no        |      |
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

| size      | engine            | status                                            | median  | min     | max     | time vs sheetjs | RSS footprint | mem vs sheetjs | post-warm-up delta | peak RSS | output   | first byte |
| --------- | ----------------- | ------------------------------------------------- | ------- | ------- | ------- | --------------- | ------------- | -------------- | ------------------ | -------- | -------- | ---------- |
| 1m        | sheetjs           | ERROR (warm-up) RangeError: Invalid string length | —       | —       | —       | —               | —             | —              | —                  | —        | —        | —          |
| 1m        | simple-excel      | ok                                                | 45.79 s | 45.79 s | 45.79 s | —               | 352 MB        | —              | 43 MB              | 4,599 MB | 427.0 MB | 0.4 ms     |
| 1m        | simple-excel-zlib | ok                                                | 32.01 s | 32.01 s | 32.01 s | —               | 182 MB        | —              | 480 MB             | 4,430 MB | 511.3 MB | 0.4 ms     |
| 18m-cells | sheetjs           | ERROR (warm-up) RangeError: Invalid string length | —       | —       | —       | —               | —             | —              | —                  | —        | —        | —          |
| 18m-cells | simple-excel      | ok                                                | 40.84 s | 40.84 s | 40.84 s | —               | 325 MB        | —              | 44 MB              | 4,157 MB | 383.8 MB | 0.7 ms     |
| 18m-cells | simple-excel-zlib | ok                                                | 29.55 s | 29.55 s | 29.55 s | —               | 126 MB        | —              | 414 MB             | 3,959 MB | 459.7 MB | 0.5 ms     |

## Gates

### simple-excel: INCOMPLETE

| gate                                                       | status  | detail                        |
| ---------------------------------------------------------- | ------- | ----------------------------- |
| write RSS footprint <= 0.25x sheetjs (mixed 100k)          | NOT-RUN | cell not in this run          |
| write median time <= 1.0x sheetjs (mixed 100k)             | NOT-RUN | cell not in this run          |
| read-typed median time <= 1.0x sheetjs (mixed 100k)        | NOT-RUN | cell not in this run          |
| read-typed RSS footprint <= 0.5x sheetjs (mixed 100k)      | NOT-RUN | cell not in this run          |
| write 1M x 20 (mixed 1m) succeeds                          | PASS    | 45.79 s, 352 MB RSS footprint |
| write 900k x 20 (mixed 18m-cells) succeeds                 | PASS    | 40.84 s, 325 MB RSS footprint |
| first byte reaches the sink in < 100 ms (mixed 100k write) | NOT-RUN | cell not in this run          |

### simple-excel-zlib: INCOMPLETE

| gate                                                       | status  | detail                        |
| ---------------------------------------------------------- | ------- | ----------------------------- |
| write RSS footprint <= 0.25x sheetjs (mixed 100k)          | NOT-RUN | cell not in this run          |
| write median time <= 1.0x sheetjs (mixed 100k)             | NOT-RUN | cell not in this run          |
| read-typed median time <= 1.0x sheetjs (mixed 100k)        | NOT-RUN | cell not in this run          |
| read-typed RSS footprint <= 0.5x sheetjs (mixed 100k)      | NOT-RUN | cell not in this run          |
| write 1M x 20 (mixed 1m) succeeds                          | PASS    | 32.01 s, 182 MB RSS footprint |
| write 900k x 20 (mixed 18m-cells) succeeds                 | PASS    | 29.55 s, 126 MB RSS footprint |
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
