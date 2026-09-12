#!/usr/bin/env node
/**
 * Browser (Chromium) write benchmark: bundles `web/worker.js` with esbuild, serves `web/` with COOP/COEP
 * headers (cross-origin isolation unlocks `performance.measureUserAgentSpecificMemory`), drives one
 * fresh page + module Worker per cell through Playwright, and samples memory three ways while the
 * write runs: in-page `measureUserAgentSpecificMemory` (covers the worker), CDP `Performance.getMetrics`
 * JSHeapUsedSize (page isolate only) and the renderer process RSS via `ps`.
 *
 *   node bench/lib/chrome.mjs --engines simple-excel,sheetjs --datasets mixed --sizes 100k,1m
 *
 * Requires Playwright's Chromium (`npx playwright install chromium`) and, for simple-excel, a built
 * `dist/` (`npm run build`) because the worker imports the bundle.
 */
import { execFileSync } from 'node:child_process';
import { createServer } from 'node:http';
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import os from 'node:os';
import { extname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { DATASET_NAMES, DEFAULT_SEED, SIZE_NAMES, resolveShape } from './dataset.mjs';
import { formatBytes, formatMs } from './metrics.mjs';
import { createResultsDir, machineInfo } from './results-dir.mjs';

const BENCH_DIR = fileURLToPath(new URL('..', import.meta.url));
const WEB_DIR = join(BENCH_DIR, 'web');
const BUNDLE_DIR = resolve(BENCH_DIR, '../.generated/bench/web');
const BUNDLE_PATH = join(BUNDLE_DIR, 'worker.bundle.js');

const HELP = `
xlsx engine browser (Chromium) write benchmark

Usage: node bench/lib/chrome.mjs [options]

Options:
  --engines <list>   simple-excel,office-kit,sheetjs (default: all three; only these are bundled for the browser)
  --datasets <list>  ${DATASET_NAMES.join(',')} (default: mixed)
  --sizes <list>     ${SIZE_NAMES.join(',')} (default: 100k,1m)
  --runs <n>         runs per cell, each in a fresh page + worker (default: 1)
  --sample-ms <n>    memory sampling interval (default: 100)
  --timeout <sec>    per-run timeout (default: 600)
  --label <name>     results folder suffix (default: chrome)
  --headed           show the browser window
  --skip-bundle      reuse the existing worker bundle
  --help             this text
`;

function parseArgs(argv) {
  const options = {
    engines: ['simple-excel', 'office-kit', 'sheetjs'],
    datasets: ['mixed'],
    sizes: ['100k', '1m'],
    runs: 1,
    sampleMs: 100,
    timeoutSeconds: 600,
    label: 'chrome',
    headed: false,
    skipBundle: false,
    help: false,
  };
  const list = value =>
    value
      .split(',')
      .map(item => item.trim())
      .filter(Boolean);
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    const next = () => {
      const value = argv[++i];
      if (value === undefined) {
        throw new Error(`${arg} requires a value`);
      }
      return value;
    };
    switch (arg) {
      case '--engines':
        options.engines = list(next());
        break;
      case '--datasets':
        options.datasets = list(next());
        break;
      case '--sizes':
        options.sizes = list(next());
        break;
      case '--runs':
        options.runs = Number(next());
        break;
      case '--sample-ms':
        options.sampleMs = Number(next());
        break;
      case '--timeout':
        options.timeoutSeconds = Number(next());
        break;
      case '--label':
        options.label = next();
        break;
      case '--headed':
        options.headed = true;
        break;
      case '--skip-bundle':
        options.skipBundle = true;
        break;
      case '--help':
      case '-h':
        options.help = true;
        break;
      default:
        throw new Error(`Unknown option ${arg} (see --help)`);
    }
  }
  return options;
}

async function bundleWorker() {
  const esbuild = await import('esbuild');
  mkdirSync(BUNDLE_DIR, { recursive: true });
  const result = await esbuild.build({
    entryPoints: [join(WEB_DIR, 'worker.js')],
    bundle: true,
    format: 'esm',
    platform: 'browser',
    target: 'es2022',
    outfile: BUNDLE_PATH,
    logLevel: 'warning',
    metafile: true,
  });
  const bytes = Object.values(result.metafile.outputs)[0]?.bytes ?? 0;
  console.log(`[chrome] bundled worker -> ${BUNDLE_PATH} (${formatBytes(bytes)})`);
}

const CONTENT_TYPES = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.map': 'application/json' };

function startServer() {
  const routes = { '/': join(WEB_DIR, 'index.html'), '/index.html': join(WEB_DIR, 'index.html'), '/worker.bundle.js': BUNDLE_PATH };
  const server = createServer((request, response) => {
    const path = new URL(request.url, 'http://localhost').pathname;
    const file = routes[path];
    if (path === '/favicon.ico') {
      response.writeHead(204).end();
      return;
    }
    if (!file) {
      response.writeHead(404).end('not found');
      return;
    }
    response.writeHead(200, {
      'Content-Type': CONTENT_TYPES[extname(file)] ?? 'application/octet-stream',
      'Cross-Origin-Opener-Policy': 'same-origin',
      'Cross-Origin-Embedder-Policy': 'require-corp',
      'Cache-Control': 'no-store',
    });
    response.end(readFileSync(file));
  });
  return new Promise(resolvePromise => {
    server.listen(0, '127.0.0.1', () => resolvePromise({ server, url: `http://127.0.0.1:${server.address().port}/` }));
  });
}

/** RSS (bytes) of every `--type=renderer` child of the browser process, keyed by pid. macOS/Linux `ps`. */
function rendererRss(browserPid) {
  try {
    const output = execFileSync('ps', ['-axo', 'pid=,ppid=,rss=,command='], { encoding: 'utf8' });
    const result = {};
    for (const line of output.split('\n')) {
      const match = line.match(/^\s*(\d+)\s+(\d+)\s+(\d+)\s+(.*)$/);
      if (match && Number(match[2]) === browserPid && match[4].includes('--type=renderer')) {
        result[match[1]] = Number(match[3]) * 1024;
      }
    }
    return result;
  } catch {
    return {};
  }
}

function startProcessSampler({ cdpSession, browserPid, sampleMs }) {
  let peakHeap = 0;
  let peakRss = 0;
  const state = { stopped: false };
  (async () => {
    while (!state.stopped) {
      const startedAt = Date.now();
      try {
        const { metrics } = await cdpSession.send('Performance.getMetrics');
        const heap = metrics.find(metric => metric.name === 'JSHeapUsedSize')?.value ?? 0;
        peakHeap = Math.max(peakHeap, heap);
      } catch {
        // session gone (page crashed) — keep sampling RSS
      }
      const rss = Math.max(0, ...Object.values(rendererRss(browserPid)));
      peakRss = Math.max(peakRss, rss);
      const elapsed = Date.now() - startedAt;
      if (elapsed < sampleMs) {
        await new Promise(resolvePromise => setTimeout(resolvePromise, sampleMs - elapsed));
      }
    }
  })();
  return {
    stop() {
      state.stopped = true;
      return { peakPageHeapBytes: peakHeap, peakRendererRssBytes: peakRss };
    },
  };
}

async function runCell(browser, browserPid, url, cell, options) {
  const context = await browser.newContext();
  const page = await context.newPage();
  let crashed = false;
  page.on('crash', () => {
    crashed = true;
  });
  page.on('console', message => {
    if (message.type() === 'error') {
      console.error(`[chrome]   console.error: ${message.text()}`);
    }
  });
  await page.goto(url);
  await page.waitForFunction(() => window.benchReady === true);
  const cdpSession = await context.newCDPSession(page);
  await cdpSession.send('Performance.enable');
  const baselineRss = Math.max(0, ...Object.values(rendererRss(browserPid)));
  const sampler = startProcessSampler({ cdpSession, browserPid, sampleMs: options.sampleMs });
  const startedAt = Date.now();
  let outcome;
  try {
    outcome = await Promise.race([
      page.evaluate(args => window.bench.run(args), {
        engine: cell.engine,
        dataset: cell.dataset,
        rows: cell.rows,
        seed: DEFAULT_SEED,
        sampleMs: options.sampleMs,
      }),
      new Promise((_, reject) => setTimeout(() => reject(new Error('timeout')), options.timeoutSeconds * 1000)),
    ]);
  } catch (error) {
    outcome = {
      ok: false,
      error: crashed
        ? { name: 'RendererCrash', message: 'page crashed (renderer out of memory or killed)' }
        : error.message === 'timeout'
          ? { name: 'Timeout', message: `exceeded ${options.timeoutSeconds}s` }
          : { name: error.name, message: error.message },
    };
  }
  const process = sampler.stop();
  await context.close().catch(() => {});
  return {
    ...cell,
    ...outcome,
    status: outcome.ok ? 'ok' : outcome.error?.name === 'Timeout' ? 'timeout' : 'error',
    wallMs: Date.now() - startedAt,
    process: {
      ...process,
      baselineRendererRssBytes: baselineRss,
      rendererRssDeltaBytes: Math.max(0, process.peakRendererRssBytes - baselineRss),
    },
  };
}

function renderMarkdown(results) {
  const lines = [`# xlsx engine browser benchmark — ${results.label}`, ''];
  lines.push(
    `Generated ${results.createdAt}. Chromium ${results.browser.version} (Playwright ${results.browser.playwright}), ${results.browser.headless ? 'headless' : 'headed'}.`,
  );
  lines.push(
    `${results.machine.cpuModel}, ${results.machine.totalMemGB} GB RAM. One fresh page + module Worker per run. simple-excel writes to \`collectToBlob()\` (the real download path, so the finished file is held); office-kit discards bytes through a counting sink; SheetJS returns one ArrayBuffer.`,
    '',
  );
  lines.push(
    '| engine | dataset | rows | run | status | write | first byte | output | agent-cluster peak | worker peak | renderer RSS delta | renderer RSS peak | page heap peak |',
  );
  lines.push('| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |');
  for (const cell of results.cells) {
    const status = cell.status === 'ok' ? 'ok' : `${cell.status.toUpperCase()} ${cell.error?.name ?? ''}: ${cell.error?.message ?? ''}`;
    lines.push(
      `| ${cell.engine} | ${cell.dataset} | ${cell.rows.toLocaleString('en-US')} | ${cell.run} | ${status} | ${formatMs(cell.ms)} | ${formatMs(cell.firstByteMs)} | ${formatBytes(cell.bytes)} | ${formatBytes(cell.memory?.peakBytes)} | ${formatBytes(cell.memory?.peakWorkerBytes)} | ${formatBytes(cell.process?.rendererRssDeltaBytes)} | ${formatBytes(cell.process?.peakRendererRssBytes)} | ${formatBytes(cell.process?.peakPageHeapBytes)} |`,
    );
  }
  lines.push(
    '',
    'Memory columns: `agent-cluster peak` / `worker peak` come from `performance.measureUserAgentSpecificMemory()` measured in the page (worker attribution = `DedicatedWorkerGlobalScope` entries; each measurement resolves at the next GC Chrome schedules, so short runs may get only one or none — `measurements` in chrome.json says how many landed);',
  );
  lines.push(
    '`renderer RSS` is the `--type=renderer` process RSS sampled with `ps`; `page heap` is CDP `Performance.getMetrics` JSHeapUsedSize for the page isolate only (the worker has its own isolate).',
    '',
  );
  return lines.join('\n');
}

async function main() {
  const options = parseArgs(process.argv.slice(2));
  if (options.help) {
    console.log(HELP.trim());
    return;
  }
  const { chromium } = await import('@playwright/test');
  if (!options.skipBundle) {
    await bundleWorker();
  }
  const shapes = new Map();
  for (const size of options.sizes) {
    for (const dataset of options.datasets) {
      const shape = resolveShape(dataset, size);
      shapes.set(`${shape.dataset}:${shape.size}`, shape);
    }
  }
  const { server, url } = await startServer();
  const browserServer = await chromium.launchServer({ headless: !options.headed, channel: 'chromium' });
  const browserPid = browserServer.process().pid;
  const browser = await chromium.connect(browserServer.wsEndpoint());
  const resultsDir = createResultsDir(options.label);
  const results = {
    label: options.label,
    createdAt: new Date().toISOString(),
    resultsDir: resultsDir.relative,
    machine: machineInfo(),
    browser: {
      version: browser.version(),
      playwright: JSON.parse(readFileSync(fileURLToPath(import.meta.resolve('@playwright/test/package.json')), 'utf8')).version,
      headless: !options.headed,
    },
    options,
    cells: [],
  };
  console.log(`[chrome] ${shapes.size * options.engines.length * options.runs} runs -> ${resultsDir.relative} (served at ${url})`);
  try {
    for (const shape of shapes.values()) {
      for (const engine of options.engines) {
        for (let run = 1; run <= options.runs; run++) {
          console.log(`[chrome] ${engine} / ${shape.dataset} / ${shape.size} run ${run}/${options.runs}`);
          const cell = await runCell(
            browser,
            browserPid,
            url,
            { engine, dataset: shape.dataset, size: shape.size, rows: shape.rows, columns: shape.columns.length, cells: shape.cells, run },
            options,
          );
          results.cells.push(cell);
          console.log(
            `[chrome]   -> ${cell.status === 'ok' ? `${formatMs(cell.ms)}, ${formatBytes(cell.bytes)}, agent-cluster peak ${formatBytes(cell.memory?.peakBytes)}, renderer RSS delta ${formatBytes(cell.process?.rendererRssDeltaBytes)}` : `${cell.status.toUpperCase()} ${cell.error?.name}: ${cell.error?.message}`}`,
          );
          results.loadAvgEnd = os.loadavg();
          writeFileSync(join(resultsDir.absolute, 'chrome.json'), JSON.stringify(results, null, 2));
          writeFileSync(join(resultsDir.absolute, 'chrome.md'), renderMarkdown(results));
        }
      }
    }
  } finally {
    await browser.close().catch(() => {});
    await browserServer.close().catch(() => {});
    server.close();
  }
  console.log(`[chrome] done: ${join(resultsDir.absolute, 'chrome.md')}`);
}

main().catch(error => {
  console.error(`[chrome] ${error.stack ?? error.message}`);
  process.exit(1);
});
