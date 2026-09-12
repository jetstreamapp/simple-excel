# xlsx engine browser benchmark — phase-e-chrome

Generated 2026-09-12T20:40:55.590Z. Chromium 153.0.8010.12 (Playwright 1.63.0), headless.
Apple M4, 32 GB RAM. One fresh page + module Worker per run. simple-excel writes to `collectToBlob()` (the real download path, so the finished file is held); office-kit discards bytes through a counting sink; SheetJS returns one ArrayBuffer.

| engine       | dataset | rows      | run | status                                                               | write   | first byte | output   | agent-cluster peak | worker peak | renderer RSS delta | renderer RSS peak | page heap peak |
| ------------ | ------- | --------- | --- | -------------------------------------------------------------------- | ------- | ---------- | -------- | ------------------ | ----------- | ------------------ | ----------------- | -------------- |
| simple-excel | mixed   | 100,000   | 1   | ok                                                                   | 3.88 s  | 4.5 ms     | 42.4 MB  | —                  | —           | 235.5 MB           | 409.2 MB          | 1.4 MB         |
| sheetjs      | mixed   | 100,000   | 1   | ok                                                                   | 7.64 s  | —          | 95.2 MB  | —                  | —           | 3038.6 MB          | 3212.9 MB         | 1.4 MB         |
| simple-excel | mixed   | 1,000,000 | 1   | ok                                                                   | 39.64 s | 0.9 ms     | 427.0 MB | 1010 KB            | —           | 253.7 MB           | 429.0 MB          | 1.4 MB         |
| sheetjs      | mixed   | 1,000,000 | 1   | ERROR RendererCrash: page crashed (renderer out of memory or killed) | —       | —          | —        | —                  | —           | 3087.5 MB          | 3263.2 MB         | 1.4 MB         |

Memory columns: `agent-cluster peak` / `worker peak` come from `performance.measureUserAgentSpecificMemory()` measured in the page (worker attribution = `DedicatedWorkerGlobalScope` entries; each measurement resolves at the next GC Chrome schedules, so short runs may get only one or none — `measurements` in chrome.json says how many landed);
`renderer RSS` is the `--type=renderer` process RSS sampled with `ps`; `page heap` is CDP `Performance.getMetrics` JSHeapUsedSize for the page isolate only (the worker has its own isolate).
