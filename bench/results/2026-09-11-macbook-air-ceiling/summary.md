# xlsx engine benchmark — ceiling

Generated 2026-09-12T04:20:17.388Z. Results folder: `bench/results/2026-09-11-macbook-air-ceiling`.

## Machine

- Apple M4 (10 cores), 32 GB RAM, darwin 25.5.0 arm64
- Node v24.18.0 (V8 13.6.233.17-node.50), child flags: `--expose-gc --max-old-space-size=12288`
- Load average at start 5.81 / 9.15 / 8.98, at end 6.43 / 8.58 / 8.78 (other processes may have been running)
- Options: runs=1, warmup=1, source=materialized, timeout=600s, seed=20260911

## Engines

| engine     | version | streaming | note |
| ---------- | ------- | --------- | ---- |
| office-kit | 0.11.0  | yes       |      |
| sheetjs    | 0.20.3  | no        |      |

## Metrics

- `median`/`min`/`max`: wall time of the timed runs (after one warm-up run of the same cell).
- `RSS footprint`: highest RSS observed while the op ran (25 ms poller, post-run read, process high-water mark) minus the RSS before the warm-up (engine loaded, input data resident, gc() done). This is the ratio/gate metric.
- `post-warm-up delta`: the same peak minus the RSS after warm-up + gc(). Undercounts engines whose warm-up footprint stays resident (V8 rarely returns pages), kept for reference.
- `peak RSS`: the absolute high-water RSS of the child process (input data included).
- `first byte`: ms from the start of `write()` until the sink receives its first chunk (streaming writers only).
- Ratios are candidate ÷ sheetjs for the same op/dataset/size; lower is better.

## write — mixed

| size | engine     | status                                            | median  | min     | max     | time vs sheetjs | RSS footprint | mem vs sheetjs | post-warm-up delta | peak RSS | output   | first byte |
| ---- | ---------- | ------------------------------------------------- | ------- | ------- | ------- | --------------- | ------------- | -------------- | ------------------ | -------- | -------- | ---------- |
| 250k | office-kit | ok                                                | 11.70 s | 11.70 s | 11.70 s | —               | 2,434 MB      | —              | 838 MB             | 3,628 MB | 102.9 MB | 4.6 ms     |
| 500k | office-kit | ERROR (warm-up) RangeError: Invalid string length | —       | —       | —       | —               | —             | —              | —                  | —        | —        | —          |
| 250k | sheetjs    | ERROR (warm-up) RangeError: Invalid string length | —       | —       | —       | —               | —             | —              | —                  | —        | —        | —          |
| 500k | sheetjs    | ERROR (warm-up) RangeError: Invalid string length | —       | —       | —       | —               | —             | —              | —                  | —        | —        | —          |

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
