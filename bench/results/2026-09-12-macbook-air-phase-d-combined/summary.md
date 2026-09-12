# xlsx engine benchmark — phase-d-combined

Generated 2026-09-12T19:16:57.376Z. Results folder: `bench/results/2026-09-12-macbook-air-phase-d-combined`.

## Machine

- Apple M4 (10 cores), 32 GB RAM, darwin 25.5.0 arm64
- Node v24.18.0 (V8 13.6.233.17-node.50), child flags: `--expose-gc --max-old-space-size=12288`
- Load average at start 3.93 / 6.86 / 7.17, at end 3.33 / 4.03 / 5.12 (other processes may have been running)
- Options: runs=3, warmup=1, source=materialized, timeout=600s, seed=20260911

## Engines

| engine            | version                        | streaming | note |
| ----------------- | ------------------------------ | --------- | ---- |
| sheetjs           | 0.20.3                         | no        |      |
| exceljs           | 4.4.0                          | yes       |      |
| office-kit        | 0.11.0                         | yes       |      |
| write-excel-file  | 2.3.10 (read-excel-file 5.8.8) | no        |      |
| simple-excel      | 0.0.0                          | yes       |      |
| simple-excel-zlib | 0.0.0                          | yes       |      |

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
| 1k        | sheetjs           | ok                                                | 86.4 ms | 85.8 ms | 112 ms  | —               | 171 MB        | —              | 136 MB             | 248 MB   | 2.5 MB   | —          |
| 1k        | exceljs           | ok                                                | 48.9 ms | 38.8 ms | 50.1 ms | 0.57x           | 24 MB         | 0.14x          | 1 MB               | 115 MB   | 543 KB   | 33.7 ms    |
| 1k        | office-kit        | ok                                                | 68.7 ms | 60.6 ms | 70.7 ms | 0.80x           | 39 MB         | 0.23x          | 28 MB              | 123 MB   | 435 KB   | 8.0 ms     |
| 1k        | write-excel-file  | ok                                                | 57.7 ms | 49.5 ms | 63.9 ms | 0.67x           | 35 MB         | 0.21x          | 23 MB              | 121 MB   | 503 KB   | —          |
| 1k        | simple-excel      | ok                                                | 63.8 ms | 56.7 ms | 107 ms  | 0.74x           | 16 MB         | 0.09x          | 7 MB               | 87 MB    | 450 KB   | 0.4 ms     |
| 1k        | simple-excel-zlib | ok                                                | 56.1 ms | 43.3 ms | 72.8 ms | 0.65x           | 14 MB         | 0.08x          | 7 MB               | 85 MB    | 529 KB   | 0.4 ms     |
| 10k       | sheetjs           | ok                                                | 894 ms  | 774 ms  | 1.25 s  | —               | 1,086 MB      | —              | 910 MB             | 1,258 MB | 24.9 MB  | —          |
| 10k       | exceljs           | ok                                                | 313 ms  | 292 ms  | 345 ms  | 0.35x           | 107 MB        | 0.10x          | 4 MB               | 285 MB   | 5.2 MB   | 192 ms     |
| 10k       | office-kit        | ok                                                | 463 ms  | 436 ms  | 487 ms  | 0.52x           | 228 MB        | 0.21x          | 110 MB             | 401 MB   | 4.1 MB   | 4.2 ms     |
| 10k       | write-excel-file  | ok                                                | 337 ms  | 321 ms  | 414 ms  | 0.38x           | 188 MB        | 0.17x          | 49 MB              | 362 MB   | 4.8 MB   | —          |
| 10k       | simple-excel      | ok                                                | 457 ms  | 453 ms  | 475 ms  | 0.51x           | 43 MB         | 0.04x          | 15 MB              | 204 MB   | 4.2 MB   | 0.4 ms     |
| 10k       | simple-excel-zlib | ok                                                | 344 ms  | 333 ms  | 358 ms  | 0.38x           | 41 MB         | 0.04x          | 17 MB              | 202 MB   | 5.1 MB   | 0.3 ms     |
| 100k      | sheetjs           | ok                                                | 9.43 s  | 8.95 s  | 10.10 s | —               | 2,896 MB      | —              | 2,951 MB           | 3,488 MB | 95.2 MB  | —          |
| 100k      | exceljs           | ok                                                | 2.52 s  | 2.50 s  | 2.58 s  | 0.27x           | 317 MB        | 0.11x          | 3 MB               | 912 MB   | 52.3 MB  | 1.62 s     |
| 100k      | office-kit        | ok                                                | 4.70 s  | 4.50 s  | 4.72 s  | 0.50x           | 981 MB        | 0.34x          | 361 MB             | 1,572 MB | 41.2 MB  | 2.9 ms     |
| 100k      | write-excel-file  | ok                                                | 3.62 s  | 3.50 s  | 3.69 s  | 0.38x           | 1,117 MB      | 0.39x          | 665 MB             | 1,709 MB | 48.2 MB  | —          |
| 100k      | simple-excel      | ok                                                | 4.56 s  | 4.55 s  | 4.56 s  | 0.48x           | 108 MB        | 0.04x          | 45 MB              | 693 MB   | 42.4 MB  | 0.4 ms     |
| 100k      | simple-excel-zlib | ok                                                | 3.32 s  | 3.15 s  | 3.35 s  | 0.35x           | 109 MB        | 0.04x          | 51 MB              | 695 MB   | 50.9 MB  | 0.3 ms     |
| 1m        | sheetjs           | ERROR (warm-up) RangeError: Invalid string length | —       | —       | —       | —               | —             | —              | —                  | —        | —        | —          |
| 1m        | simple-excel      | ok                                                | 45.79 s | 45.79 s | 45.79 s | —               | 352 MB        | —              | 43 MB              | 4,599 MB | 427.0 MB | 0.4 ms     |
| 1m        | simple-excel-zlib | ok                                                | 32.01 s | 32.01 s | 32.01 s | —               | 182 MB        | —              | 480 MB             | 4,430 MB | 511.3 MB | 0.4 ms     |
| 18m-cells | sheetjs           | ERROR (warm-up) RangeError: Invalid string length | —       | —       | —       | —               | —             | —              | —                  | —        | —        | —          |
| 18m-cells | simple-excel      | ok                                                | 40.84 s | 40.84 s | 40.84 s | —               | 325 MB        | —              | 44 MB              | 4,157 MB | 383.8 MB | 0.7 ms     |
| 18m-cells | simple-excel-zlib | ok                                                | 29.55 s | 29.55 s | 29.55 s | —               | 126 MB        | —              | 414 MB             | 3,959 MB | 459.7 MB | 0.5 ms     |

## read-typed — mixed

Fixtures written by: sheetjs (see `fixture` in results.json per cell).

| size | engine            | status | median  | min     | max     | time vs sheetjs | RSS footprint | mem vs sheetjs | post-warm-up delta | peak RSS | rows read | cells read |
| ---- | ----------------- | ------ | ------- | ------- | ------- | --------------- | ------------- | -------------- | ------------------ | -------- | --------- | ---------- |
| 1k   | sheetjs           | ok     | 72.2 ms | 64.8 ms | 74.6 ms | —               | 97 MB         | —              | 43 MB              | 166 MB   | 1,000     | 20,000     |
| 1k   | exceljs           | ok     | 63.8 ms | 58.5 ms | 73.8 ms | 0.88x           | 43 MB         | 0.44x          | 19 MB              | 118 MB   | 1,000     | 20,000     |
| 1k   | office-kit        | ok     | 80.4 ms | 80.4 ms | 86.6 ms | 1.11x           | 178 MB        | 1.83x          | 62 MB              | 247 MB   | 1,000     | 20,000     |
| 1k   | write-excel-file  | ok     | 90.8 ms | 87.0 ms | 91.9 ms | 1.26x           | 186 MB        | 1.91x          | 71 MB              | 256 MB   | 1,000     | 20,000     |
| 1k   | simple-excel      | ok     | 42.3 ms | 36.6 ms | 47.3 ms | 0.59x           | 57 MB         | 0.58x          | 32 MB              | 114 MB   | 1,000     | 20,000     |
| 1k   | simple-excel-zlib | ok     | 40.0 ms | 36.7 ms | 41.1 ms | 0.55x           | 55 MB         | 0.57x          | 31 MB              | 112 MB   | 1,000     | 20,000     |
| 10k  | sheetjs           | ok     | 590 ms  | 581 ms  | 638 ms  | —               | 399 MB        | —              | 79 MB              | 510 MB   | 10,000    | 200,000    |
| 10k  | exceljs           | ok     | 954 ms  | 921 ms  | 955 ms  | 1.62x           | 124 MB        | 0.31x          | 70 MB              | 199 MB   | 10,000    | 200,000    |
| 10k  | office-kit        | ok     | 626 ms  | 611 ms  | 640 ms  | 1.06x           | 612 MB        | 1.53x          | 23 MB              | 680 MB   | 10,000    | 200,000    |
| 10k  | write-excel-file  | ok     | 629 ms  | 612 ms  | 680 ms  | 1.07x           | 699 MB        | 1.75x          | 147 MB             | 770 MB   | 10,000    | 200,000    |
| 10k  | simple-excel      | ok     | 204 ms  | 203 ms  | 210 ms  | 0.35x           | 216 MB        | 0.54x          | 98 MB              | 318 MB   | 10,000    | 200,000    |
| 10k  | simple-excel-zlib | ok     | 209 ms  | 208 ms  | 213 ms  | 0.35x           | 218 MB        | 0.55x          | 99 MB              | 320 MB   | 10,000    | 200,000    |
| 100k | sheetjs           | ok     | 6.44 s  | 6.36 s  | 6.81 s  | —               | 2,037 MB      | —              | 434 MB             | 2,291 MB | 100,000   | 2,000,000  |
| 100k | exceljs           | ok     | 11.07 s | 10.55 s | 11.23 s | 1.72x           | 178 MB        | 0.09x          | 61 MB              | 254 MB   | 100,000   | 2,000,000  |
| 100k | office-kit        | ok     | 6.70 s  | 6.67 s  | 6.81 s  | 1.04x           | 452 MB        | 0.22x          | 111 MB             | 521 MB   | 100,000   | 2,000,000  |
| 100k | write-excel-file  | ok     | 9.04 s  | 9.02 s  | 10.40 s | 1.40x           | 4,484 MB      | 2.20x          | 1,548 MB           | 4,554 MB | 100,000   | 2,000,000  |
| 100k | simple-excel      | ok     | 2.10 s  | 2.09 s  | 2.13 s  | 0.33x           | 881 MB        | 0.43x          | 13 MB              | 1,124 MB | 100,000   | 2,000,000  |
| 100k | simple-excel-zlib | ok     | 2.15 s  | 2.09 s  | 2.18 s  | 0.33x           | 880 MB        | 0.43x          | 11 MB              | 1,123 MB | 100,000   | 2,000,000  |

## read-raw — mixed

Fixtures written by: sheetjs (see `fixture` in results.json per cell).

| size | engine            | status                                                | median  | min     | max     | time vs sheetjs | RSS footprint | mem vs sheetjs | post-warm-up delta | peak RSS | rows read | cells read |
| ---- | ----------------- | ----------------------------------------------------- | ------- | ------- | ------- | --------------- | ------------- | -------------- | ------------------ | -------- | --------- | ---------- |
| 1k   | sheetjs           | ok                                                    | 68.5 ms | 66.8 ms | 73.8 ms | —               | 98 MB         | —              | 48 MB              | 163 MB   | 1,000     | 20,000     |
| 1k   | exceljs           | ok                                                    | 111 ms  | 65.7 ms | 170 ms  | 1.61x           | 45 MB         | 0.45x          | 22 MB              | 120 MB   | 1,000     | 20,000     |
| 1k   | office-kit        | ok                                                    | 83.6 ms | 81.5 ms | 91.4 ms | 1.22x           | 178 MB        | 1.82x          | 61 MB              | 247 MB   | 1,000     | 20,000     |
| 1k   | write-excel-file  | skipped: write-excel-file does not implement read-raw | —       | —       | —       | —               | —             | —              | —                  | —        | —         | —          |
| 1k   | simple-excel      | ok                                                    | 49.7 ms | 39.0 ms | 50.6 ms | 0.73x           | 35 MB         | 0.36x          | 15 MB              | 92 MB    | 1,000     | 20,000     |
| 1k   | simple-excel-zlib | ok                                                    | 51.3 ms | 47.4 ms | 58.7 ms | 0.75x           | 35 MB         | 0.36x          | 16 MB              | 92 MB    | 1,000     | 20,000     |
| 10k  | sheetjs           | ok                                                    | 609 ms  | 606 ms  | 613 ms  | —               | 393 MB        | —              | 75 MB              | 503 MB   | 10,000    | 200,000    |
| 10k  | exceljs           | ok                                                    | 948 ms  | 315 ms  | 986 ms  | 1.56x           | 148 MB        | 0.38x          | 95 MB              | 222 MB   | 10,000    | 200,000    |
| 10k  | office-kit        | ok                                                    | 624 ms  | 615 ms  | 635 ms  | 1.03x           | 632 MB        | 1.61x          | 22 MB              | 701 MB   | 10,000    | 200,000    |
| 10k  | write-excel-file  | skipped: write-excel-file does not implement read-raw | —       | —       | —       | —               | —             | —              | —                  | —        | —         | —          |
| 10k  | simple-excel      | ok                                                    | 209 ms  | 202 ms  | 219 ms  | 0.34x           | 98 MB         | 0.25x          | 54 MB              | 200 MB   | 10,000    | 200,000    |
| 10k  | simple-excel-zlib | ok                                                    | 215 ms  | 202 ms  | 235 ms  | 0.35x           | 98 MB         | 0.25x          | 55 MB              | 200 MB   | 10,000    | 200,000    |
| 100k | sheetjs           | ok                                                    | 6.57 s  | 6.44 s  | 6.67 s  | —               | 2,073 MB      | —              | 472 MB             | 2,323 MB | 100,000   | 2,000,000  |
| 100k | exceljs           | ok                                                    | 9.83 s  | 3.28 s  | 10.08 s | 1.49x           | 215 MB        | 0.10x          | 23 MB              | 290 MB   | 100,000   | 2,000,000  |
| 100k | office-kit        | ok                                                    | 6.54 s  | 6.48 s  | 6.63 s  | 1.00x           | 454 MB        | 0.22x          | 105 MB             | 522 MB   | 100,000   | 2,000,000  |
| 100k | write-excel-file  | skipped: write-excel-file does not implement read-raw | —       | —       | —       | —               | —             | —              | —                  | —        | —         | —          |
| 100k | simple-excel      | ok                                                    | 1.97 s  | 1.96 s  | 2.01 s  | 0.30x           | 75 MB         | 0.04x          | 32 MB              | 318 MB   | 100,000   | 2,000,000  |
| 100k | simple-excel-zlib | ok                                                    | 2.06 s  | 1.96 s  | 2.09 s  | 0.31x           | 77 MB         | 0.04x          | 34 MB              | 320 MB   | 100,000   | 2,000,000  |

## Gates

### exceljs: FAIL

| gate                                                       | status  | detail                                  |
| ---------------------------------------------------------- | ------- | --------------------------------------- |
| write RSS footprint <= 0.25x sheetjs (mixed 100k)          | PASS    | 0.11x (317 MB vs 2,896 MB; limit 0.25x) |
| write median time <= 1.0x sheetjs (mixed 100k)             | PASS    | 0.27x (2.52 s vs 9.43 s; limit 1x)      |
| read-typed median time <= 1.0x sheetjs (mixed 100k)        | FAIL    | 1.72x (11.07 s vs 6.44 s; limit 1x)     |
| read-typed RSS footprint <= 0.5x sheetjs (mixed 100k)      | PASS    | 0.09x (178 MB vs 2,037 MB; limit 0.5x)  |
| write 1M x 20 (mixed 1m) succeeds                          | NOT-RUN | cell not in this run                    |
| write 900k x 20 (mixed 18m-cells) succeeds                 | NOT-RUN | cell not in this run                    |
| first byte reaches the sink in < 100 ms (mixed 100k write) | FAIL    | 1.62 s (limit 100 ms)                   |

### office-kit: FAIL

| gate                                                       | status  | detail                                  |
| ---------------------------------------------------------- | ------- | --------------------------------------- |
| write RSS footprint <= 0.25x sheetjs (mixed 100k)          | FAIL    | 0.34x (981 MB vs 2,896 MB; limit 0.25x) |
| write median time <= 1.0x sheetjs (mixed 100k)             | PASS    | 0.50x (4.70 s vs 9.43 s; limit 1x)      |
| read-typed median time <= 1.0x sheetjs (mixed 100k)        | FAIL    | 1.04x (6.70 s vs 6.44 s; limit 1x)      |
| read-typed RSS footprint <= 0.5x sheetjs (mixed 100k)      | PASS    | 0.22x (452 MB vs 2,037 MB; limit 0.5x)  |
| write 1M x 20 (mixed 1m) succeeds                          | NOT-RUN | cell not in this run                    |
| write 900k x 20 (mixed 18m-cells) succeeds                 | NOT-RUN | cell not in this run                    |
| first byte reaches the sink in < 100 ms (mixed 100k write) | PASS    | 2.9 ms (limit 100 ms)                   |

### write-excel-file: FAIL

| gate                                                       | status  | detail                                                |
| ---------------------------------------------------------- | ------- | ----------------------------------------------------- |
| write RSS footprint <= 0.25x sheetjs (mixed 100k)          | FAIL    | 0.39x (1,117 MB vs 2,896 MB; limit 0.25x)             |
| write median time <= 1.0x sheetjs (mixed 100k)             | PASS    | 0.38x (3.62 s vs 9.43 s; limit 1x)                    |
| read-typed median time <= 1.0x sheetjs (mixed 100k)        | FAIL    | 1.40x (9.04 s vs 6.44 s; limit 1x)                    |
| read-typed RSS footprint <= 0.5x sheetjs (mixed 100k)      | FAIL    | 2.20x (4,484 MB vs 2,037 MB; limit 0.5x)              |
| write 1M x 20 (mixed 1m) succeeds                          | NOT-RUN | cell not in this run                                  |
| write 900k x 20 (mixed 18m-cells) succeeds                 | NOT-RUN | cell not in this run                                  |
| first byte reaches the sink in < 100 ms (mixed 100k write) | FAIL    | engine reported no first-byte latency (not streaming) |

### simple-excel: PASS

| gate                                                       | status | detail                                  |
| ---------------------------------------------------------- | ------ | --------------------------------------- |
| write RSS footprint <= 0.25x sheetjs (mixed 100k)          | PASS   | 0.04x (108 MB vs 2,896 MB; limit 0.25x) |
| write median time <= 1.0x sheetjs (mixed 100k)             | PASS   | 0.48x (4.56 s vs 9.43 s; limit 1x)      |
| read-typed median time <= 1.0x sheetjs (mixed 100k)        | PASS   | 0.33x (2.10 s vs 6.44 s; limit 1x)      |
| read-typed RSS footprint <= 0.5x sheetjs (mixed 100k)      | PASS   | 0.43x (881 MB vs 2,037 MB; limit 0.5x)  |
| write 1M x 20 (mixed 1m) succeeds                          | PASS   | 45.79 s, 352 MB RSS footprint           |
| write 900k x 20 (mixed 18m-cells) succeeds                 | PASS   | 40.84 s, 325 MB RSS footprint           |
| first byte reaches the sink in < 100 ms (mixed 100k write) | PASS   | 0.4 ms (limit 100 ms)                   |

### simple-excel-zlib: PASS

| gate                                                       | status | detail                                  |
| ---------------------------------------------------------- | ------ | --------------------------------------- |
| write RSS footprint <= 0.25x sheetjs (mixed 100k)          | PASS   | 0.04x (109 MB vs 2,896 MB; limit 0.25x) |
| write median time <= 1.0x sheetjs (mixed 100k)             | PASS   | 0.35x (3.32 s vs 9.43 s; limit 1x)      |
| read-typed median time <= 1.0x sheetjs (mixed 100k)        | PASS   | 0.33x (2.15 s vs 6.44 s; limit 1x)      |
| read-typed RSS footprint <= 0.5x sheetjs (mixed 100k)      | PASS   | 0.43x (880 MB vs 2,037 MB; limit 0.5x)  |
| write 1M x 20 (mixed 1m) succeeds                          | PASS   | 32.01 s, 182 MB RSS footprint           |
| write 900k x 20 (mixed 18m-cells) succeeds                 | PASS   | 29.55 s, 126 MB RSS footprint           |
| first byte reaches the sink in < 100 ms (mixed 100k write) | PASS   | 0.3 ms (limit 100 ms)                   |

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
