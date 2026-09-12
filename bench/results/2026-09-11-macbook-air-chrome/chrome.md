# xlsx engine browser benchmark — chrome

Generated 2026-09-12T04:30:05.601Z. Chromium 151.0.7922.34 (Playwright 1.62.1), headless.
Apple M4, 32 GB RAM. One fresh page + module Worker per run; output discarded through a counting sink.

| engine     | dataset | rows      | run | status                                                               | write  | first byte | output  | agent-cluster peak | worker peak | renderer RSS delta | renderer RSS peak | page heap peak |
| ---------- | ------- | --------- | --- | -------------------------------------------------------------------- | ------ | ---------- | ------- | ------------------ | ----------- | ------------------ | ----------------- | -------------- |
| office-kit | mixed   | 100,000   | 1   | ok                                                                   | 6.42 s | 10.6 ms    | 41.2 MB | —                  | —           | 1618.4 MB          | 1753.6 MB         | 1.4 MB         |
| sheetjs    | mixed   | 100,000   | 1   | ok                                                                   | 8.90 s | —          | 95.2 MB | —                  | —           | 2874.5 MB          | 3009.6 MB         | 1.4 MB         |
| office-kit | mixed   | 1,000,000 | 1   | ERROR RendererCrash: page crashed (renderer out of memory or killed) | —      | —          | —       | —                  | —           | 3881.8 MB          | 4016.8 MB         | 1.4 MB         |
| sheetjs    | mixed   | 1,000,000 | 1   | ERROR RendererCrash: page crashed (renderer out of memory or killed) | —      | —          | —       | —                  | —           | 2550.8 MB          | 2686.4 MB         | 1.4 MB         |

Memory columns: `agent-cluster peak` / `worker peak` come from `performance.measureUserAgentSpecificMemory()` measured in the page (worker attribution = `DedicatedWorkerGlobalScope` entries; each measurement resolves at the next GC Chrome schedules, so short runs may get only one or none — `measurements` in chrome.json says how many landed);
`renderer RSS` is the `--type=renderer` process RSS sampled with `ps`; `page heap` is CDP `Performance.getMetrics` JSHeapUsedSize for the page isolate only (the worker has its own isolate).
