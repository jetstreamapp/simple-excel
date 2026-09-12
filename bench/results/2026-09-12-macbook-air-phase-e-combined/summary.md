# xlsx engine benchmark — phase-e-combined

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
