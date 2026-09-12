#!/usr/bin/env node
/**
 * xlsx engine benchmark CLI. Forks `lib/child.mjs` once per (engine, op, dataset, size) cell so every
 * measurement starts from a fresh process, then writes `results/<date>-<host>-<label>/results.json`
 * and `summary.md`. See README.md for the full description.
 */
import { spawn } from 'node:child_process';
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import os from 'node:os';
import { join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { DATASET_NAMES, DEFAULT_SEED, SIZE_NAMES, resolveShape } from './lib/dataset.mjs';
import { RESULT_MARKER } from './lib/child.mjs';
import { OPS, listEngineNames, loadEngine } from './lib/engines.mjs';
import { renderSummaryMarkdown } from './lib/metrics.mjs';
import { createResultsDir, machineInfo } from './lib/results-dir.mjs';

const BENCH_DIR = fileURLToPath(new URL('.', import.meta.url));
const GENERATED_DIR = resolve(BENCH_DIR, '../.generated/bench');
const FIXTURES_DIR = join(GENERATED_DIR, 'fixtures');
const OUTPUT_DIR = join(GENERATED_DIR, 'out');
const CHILD_PATH = join(BENCH_DIR, 'lib/child.mjs');

const DEFAULT_NODE_FLAGS = ['--expose-gc', '--max-old-space-size=12288'];
/** Fixtures up to this many cells are produced by SheetJS (what Jetstream ships); larger ones by office-kit. */
const SHEETJS_FIXTURE_MAX_CELLS = 2_500_000;
const LARGE_FIXTURE_PRODUCER = 'office-kit';

const HELP = `
xlsx engine benchmark

Usage: node bench/run.mjs [options]
       npm run bench -- -- [options]

Options:
  --engines <list>   comma-separated engines (default: all in engines/: ${listEngineNames().join(',')})
  --ops <list>       comma-separated ops: ${OPS.join(',')} (default: all)
  --datasets <list>  comma-separated datasets: ${DATASET_NAMES.join(',')} (default: mixed)
  --sizes <list>     comma-separated sizes: ${SIZE_NAMES.join(',')} (default: 1k,10k)
                     18m-cells and wide-100k are presets that pin the dataset
  --runs <n>         timed runs per cell (default: 3, forced to 1 for >= 1M rows unless given)
  --warmup <n>       warm-up runs per cell (default: 1)
  --source <mode>    write input: materialized (rows in memory, like Jetstream) | stream (default: materialized)
  --label <name>     results folder suffix (default: adhoc)
  --timeout <sec>    per-cell timeout in seconds (default: 600)
  --seed <n>         dataset PRNG seed (default: ${DEFAULT_SEED})
  --node-flags <s>   space-separated flags for the child (default: "${DEFAULT_NODE_FLAGS.join(' ')}")
  --rebuild-fixtures regenerate read fixtures even if present
  --merge <dirs>     comma-separated results folders to merge into one summary (no benchmarks run)
  --help             this text

Read fixtures are written to .generated/bench/fixtures/ by sheetjs
(<= ${SHEETJS_FIXTURE_MAX_CELLS.toLocaleString('en-US')} cells) or ${LARGE_FIXTURE_PRODUCER} (larger), and reused across runs.
`;

function parseArgs(argv) {
  const options = {
    engines: listEngineNames(),
    ops: OPS,
    datasets: ['mixed'],
    sizes: ['1k', '10k'],
    runs: null,
    warmup: 1,
    source: 'materialized',
    label: 'adhoc',
    timeoutSeconds: 600,
    seed: DEFAULT_SEED,
    nodeFlags: DEFAULT_NODE_FLAGS,
    rebuildFixtures: false,
    merge: null,
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
      case '--ops':
        options.ops = list(next());
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
      case '--warmup':
        options.warmup = Number(next());
        break;
      case '--source':
        options.source = next();
        break;
      case '--label':
        options.label = next();
        break;
      case '--timeout':
        options.timeoutSeconds = Number(next());
        break;
      case '--seed':
        options.seed = Number(next());
        break;
      case '--node-flags':
        options.nodeFlags = next().split(/\s+/).filter(Boolean);
        break;
      case '--rebuild-fixtures':
        options.rebuildFixtures = true;
        break;
      case '--merge':
        options.merge = list(next());
        break;
      case '--help':
      case '-h':
        options.help = true;
        break;
      default:
        throw new Error(`Unknown option ${arg} (see --help)`);
    }
  }
  if (!['materialized', 'stream'].includes(options.source)) {
    throw new Error(`--source must be materialized or stream`);
  }
  return options;
}

/** Forks child.mjs for one spec, enforcing the per-cell timeout; never throws. */
function runChild(spec, { nodeFlags, timeoutSeconds, onLog }) {
  return new Promise(resolvePromise => {
    const startedAt = performance.now();
    const child = spawn(process.execPath, [...nodeFlags, CHILD_PATH, JSON.stringify(spec)], { stdio: ['ignore', 'pipe', 'pipe'] });
    let stdout = '';
    let stderrTail = '';
    let timedOut = false;
    const timer = setTimeout(() => {
      timedOut = true;
      child.kill('SIGKILL');
    }, timeoutSeconds * 1000);
    child.stdout.on('data', chunk => {
      stdout += chunk;
    });
    child.stderr.on('data', chunk => {
      const text = String(chunk);
      stderrTail = (stderrTail + text).slice(-32_000);
      onLog?.(text);
    });
    child.on('close', (code, signal) => {
      clearTimeout(timer);
      const durationMs = performance.now() - startedAt;
      if (timedOut) {
        resolvePromise({ status: 'timeout', timeoutMs: timeoutSeconds * 1000, durationMs });
        return;
      }
      const line = stdout.split('\n').find(candidate => candidate.startsWith(RESULT_MARKER));
      if (line) {
        resolvePromise({ ...JSON.parse(line.slice(RESULT_MARKER.length)), exitCode: code });
        return;
      }
      const fatalOom = /heap out of memory|Allocation failed|FATAL ERROR|OOM/i.test(stderrTail);
      const fatalLine = stderrTail.split('\n').find(candidate => candidate.startsWith('FATAL ERROR'));
      resolvePromise({
        status: 'error',
        phase: 'process',
        error: {
          name: fatalOom ? 'FatalHeapOutOfMemory' : 'ChildProcessError',
          message: `child exited with code ${code} signal ${signal} and no result${fatalLine ? `; ${fatalLine.trim()}` : `; stderr tail: ${stderrTail.trim().slice(-600)}`}`,
        },
        durationMs,
      });
    });
  });
}

function fixturePaths(shape) {
  const base = join(FIXTURES_DIR, `${shape.dataset}-${shape.size}`);
  return { xlsx: `${base}.xlsx`, meta: `${base}.json` };
}

function buildCellPlan(options) {
  const shapes = new Map();
  for (const size of options.sizes) {
    for (const dataset of options.datasets) {
      const shape = resolveShape(dataset, size);
      shapes.set(`${shape.dataset}:${shape.size}`, shape);
    }
  }
  const plan = [];
  for (const op of options.ops) {
    for (const shape of shapes.values()) {
      for (const engine of options.engines) {
        plan.push({ engine, op, shape });
      }
    }
  }
  return plan;
}

async function ensureFixture(shape, options, log) {
  const paths = fixturePaths(shape);
  if (!options.rebuildFixtures && existsSync(paths.xlsx) && existsSync(paths.meta)) {
    const meta = JSON.parse(readFileSync(paths.meta, 'utf8'));
    if (meta.seed === options.seed && meta.rows === shape.rows) {
      return { ok: true, meta, path: paths.xlsx };
    }
  }
  const producer = shape.cells <= SHEETJS_FIXTURE_MAX_CELLS ? 'sheetjs' : LARGE_FIXTURE_PRODUCER;
  log(`fixture ${shape.dataset}/${shape.size} via ${producer} ...`);
  const result = await runChild(
    { task: 'fixture', engine: producer, dataset: shape.dataset, size: shape.size, seed: options.seed, fixturePath: paths.xlsx },
    { nodeFlags: options.nodeFlags, timeoutSeconds: options.timeoutSeconds, onLog: text => process.stderr.write(text) },
  );
  if (result.status !== 'ok') {
    const reason = result.error ? `${result.error.name}: ${result.error.message}` : (result.reason ?? result.status);
    log(`fixture ${shape.dataset}/${shape.size} FAILED: ${reason}`);
    return { ok: false, reason: `${producer} could not write it — ${reason}` };
  }
  writeFileSync(paths.meta, JSON.stringify(result, null, 2));
  log(
    `fixture ${shape.dataset}/${shape.size} written (${(result.bytes / (1024 * 1024)).toFixed(1)} MB in ${(result.durationMs / 1000).toFixed(1)} s)`,
  );
  return { ok: true, meta: result, path: paths.xlsx };
}

async function describeEngines(names) {
  const engines = {};
  for (const name of names) {
    try {
      const engine = await loadEngine(name);
      engines[name] = { version: engine.version, supportsStreaming: engine.supportsStreaming, skipReason: engine.skipReason };
    } catch (error) {
      engines[name] = { version: 'load failed', supportsStreaming: false, skipReason: error.message };
    }
  }
  return engines;
}

function writeOutputs(results, resultsDir) {
  writeFileSync(join(resultsDir, 'results.json'), JSON.stringify(results, null, 2));
  writeFileSync(join(resultsDir, 'summary.md'), renderSummaryMarkdown(results));
}

async function runBenchmarks(options) {
  const { absolute: resultsDir, relative: resultsDirRelative } = createResultsDir(options.label);
  const log = message => console.log(`[bench] ${message}`);
  const plan = buildCellPlan(options);
  const results = {
    label: options.label,
    createdAt: new Date().toISOString(),
    resultsDir: resultsDirRelative,
    machine: machineInfo(),
    engines: await describeEngines(options.engines),
    options: { ...options, nodeFlags: options.nodeFlags, runs: options.runs ?? 3 },
    cells: [],
  };
  log(`${plan.length} cells -> ${results.resultsDir}`);
  mkdirSync(OUTPUT_DIR, { recursive: true });
  const fixtureCache = new Map();

  for (const [index, { engine, op, shape }] of plan.entries()) {
    const runs = options.runs ?? (shape.rows >= 1_000_000 ? 1 : 3);
    const cellLabel = `${engine} / ${op} / ${shape.dataset} / ${shape.size} (${shape.rows.toLocaleString('en-US')} rows x ${shape.columns.length} cols)`;
    log(`[${index + 1}/${plan.length}] ${cellLabel}`);
    const spec = {
      task: 'bench',
      engine,
      op,
      dataset: shape.dataset,
      size: shape.size,
      runs,
      warmup: options.warmup,
      source: options.source,
      seed: options.seed,
      outPath: join(OUTPUT_DIR, `${engine}-${shape.dataset}-${shape.size}.xlsx`),
    };
    let fixture;
    if (op !== 'write') {
      const key = `${shape.dataset}:${shape.size}`;
      if (!fixtureCache.has(key)) {
        fixtureCache.set(key, await ensureFixture(shape, options, log));
      }
      fixture = fixtureCache.get(key);
      if (!fixture.ok) {
        results.cells.push({
          engine,
          op,
          dataset: shape.dataset,
          size: shape.size,
          rows: shape.rows,
          columns: shape.columns.length,
          cells: shape.cells,
          status: 'fixture-failed',
          reason: fixture.reason,
        });
        writeOutputs(results, resultsDir);
        continue;
      }
      spec.fixturePath = fixture.path;
    }
    const cell = await runChild(spec, {
      nodeFlags: options.nodeFlags,
      timeoutSeconds: options.timeoutSeconds,
      onLog: text => process.stderr.write(text.replace(/^/gm, '        ')),
    });
    const record = {
      engine,
      op,
      dataset: shape.dataset,
      size: shape.size,
      rows: shape.rows,
      columns: shape.columns.length,
      cells: shape.cells,
      runs,
      ...cell,
      fixture: fixture
        ? { path: fixture.path, producer: fixture.meta.producer, producerVersion: fixture.meta.producerVersion, bytes: fixture.meta.bytes }
        : undefined,
    };
    results.cells.push(record);
    const summary =
      record.status === 'ok'
        ? `${(record.timing.medianMs / 1000).toFixed(2)} s median, ${Math.round(record.memory.footprintMB)} MB RSS footprint${record.firstByteMs != null ? `, first byte ${record.firstByteMs.toFixed(1)} ms` : ''}`
        : record.status === 'error'
          ? `ERROR ${record.error?.name}: ${record.error?.message}`
          : record.status;
    log(`    -> ${summary}`);
    results.loadAvgEnd = os.loadavg();
    writeOutputs(results, resultsDir);
  }
  results.loadAvgEnd = os.loadavg();
  results.finishedAt = new Date().toISOString();
  writeOutputs(results, resultsDir);
  log(`done: ${join(resultsDir, 'summary.md')}`);
}

/** Merges several results folders into one (later folders win on duplicate cells) and renders a combined summary. */
function mergeResults(dirs, label) {
  const sources = dirs.map(dir => JSON.parse(readFileSync(join(resolve(dir), 'results.json'), 'utf8')));
  const merged = new Map();
  for (const source of sources) {
    for (const cell of source.cells) {
      merged.set(`${cell.engine}|${cell.op}|${cell.dataset}|${cell.size}`, { ...cell, sourceLabel: source.label });
    }
  }
  const { absolute: resultsDir, relative: resultsDirRelative } = createResultsDir(label);
  const results = {
    label,
    createdAt: new Date().toISOString(),
    resultsDir: resultsDirRelative,
    mergedFrom: sources.map(source => source.resultsDir),
    machine: sources[0].machine,
    engines: Object.assign({}, ...sources.map(source => source.engines)),
    options: sources[0].options,
    loadAvgEnd: sources.at(-1).loadAvgEnd,
    cells: [...merged.values()],
  };
  writeOutputs(results, resultsDir);
  console.log(`[bench] merged ${sources.length} folders -> ${join(resultsDir, 'summary.md')}`);
}

async function main() {
  const options = parseArgs(process.argv.slice(2));
  if (options.help) {
    console.log(HELP.trim());
    return;
  }
  if (options.merge) {
    mergeResults(options.merge, options.label === 'adhoc' ? 'merged' : options.label);
    return;
  }
  await runBenchmarks(options);
}

main().catch(error => {
  console.error(`[bench] ${error.message}`);
  process.exit(1);
});
