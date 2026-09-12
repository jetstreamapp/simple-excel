/**
 * Runs exactly ONE benchmark cell (engine × op × dataset × size) — or one fixture build — in this
 * process and prints a single JSON result line prefixed with RESULT_MARKER on stdout. `run.mjs`
 * forks one of these per cell so engines never share a heap, module cache or GC history.
 *
 * Invoked as: node --expose-gc --max-old-space-size=<mb> child.mjs '<json spec>'
 *
 * Memory model: RSS is polled every SAMPLE_INTERVAL_MS from an unref'd interval while a timed run is in
 * flight. Synchronous engines (SheetJS) block the event loop so the poller cannot fire; the process
 * high-water mark (`process.resourceUsage().maxRSS`) and the RSS read immediately after each run fill
 * that gap. Peak = max of all three, delta = peak − baseline (baseline taken after the warm-up + gc()).
 * Because maxRSS is a lifetime high-water mark it also covers the warm-up run, which executes the same op.
 */
import { existsSync, mkdirSync, readFileSync, statSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { setTimeout as sleep } from 'node:timers/promises';
import { fileURLToPath } from 'node:url';
import { createRowIterator, materialize, resolveShape } from './dataset.mjs';
import { getOpImplementation, loadEngine } from './engines.mjs';

export const RESULT_MARKER = '@@BENCH_RESULT@@';
const SAMPLE_INTERVAL_MS = 25;
const MB = 1024 * 1024;

const toMB = bytes => Math.round((bytes / MB) * 10) / 10;

function serializeError(error) {
  if (!(error instanceof Error)) {
    return { name: 'Error', message: String(error) };
  }
  return {
    name: error.name,
    message: error.message,
    stack: typeof error.stack === 'string' ? error.stack.split('\n').slice(0, 6).join('\n') : undefined,
  };
}

function emit(result) {
  process.stdout.write(`${RESULT_MARKER}${JSON.stringify(result)}\n`);
}

function log(message) {
  process.stderr.write(`${message}\n`);
}

function startRssSampler() {
  let peak = 0;
  let samples = 0;
  const timer = setInterval(() => {
    samples++;
    const rss = process.memoryUsage().rss;
    if (rss > peak) {
      peak = rss;
    }
  }, SAMPLE_INTERVAL_MS);
  timer.unref();
  return {
    stop() {
      clearInterval(timer);
      return { peak, samples };
    },
  };
}

async function settle() {
  globalThis.gc?.();
  await sleep(100);
  globalThis.gc?.();
  await sleep(50);
}

function median(values) {
  const sorted = [...values].sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  return sorted.length % 2 === 0 ? (sorted[mid - 1] + sorted[mid]) / 2 : sorted[mid];
}

/** Builds the callable that performs one full run of the op, plus what it needs prepared up front. */
async function prepareRunner(engine, op, spec, shape) {
  const implementation = getOpImplementation(engine, op);
  if (!implementation) {
    return { skipReason: `${engine.name} does not implement ${op}` };
  }
  if (op === 'write') {
    mkdirSync(dirname(spec.outPath), { recursive: true });
    const generatorOptions = { dataset: shape.dataset, rows: shape.rows, seed: spec.seed };
    const rows = spec.source === 'stream' ? null : materialize(generatorOptions);
    const writeOptions = { columns: shape.columns, rowCount: shape.rows, outPath: spec.outPath };
    return {
      run: () => implementation(rows ?? createRowIterator(generatorOptions), writeOptions),
    };
  }
  if (!spec.fixturePath || !existsSync(spec.fixturePath)) {
    return { skipReason: `fixture missing: ${spec.fixturePath}` };
  }
  const input = engine.readInput === 'bytes' ? new Uint8Array(readFileSync(spec.fixturePath)) : spec.fixturePath;
  return { run: () => implementation(input) };
}

async function runBench(spec) {
  const startedAt = performance.now();
  const shape = resolveShape(spec.dataset, spec.size);
  const base = {
    engine: spec.engine,
    op: spec.op,
    dataset: shape.dataset,
    size: shape.size,
    rows: shape.rows,
    columns: shape.columns.length,
    cells: shape.cells,
    source: spec.op === 'write' ? spec.source : undefined,
    fixturePath: spec.op === 'write' ? undefined : spec.fixturePath,
  };

  const engine = await loadEngine(spec.engine);
  if (engine.skipReason) {
    return { ...base, status: 'skipped', reason: engine.skipReason, durationMs: performance.now() - startedAt };
  }

  const runner = await prepareRunner(engine, spec.op, spec, shape);
  if (runner.skipReason) {
    return { ...base, status: 'skipped', reason: runner.skipReason, durationMs: performance.now() - startedAt };
  }

  // Pre-warm-up snapshot: engine loaded, input resident, nothing run yet. The op's footprint is
  // measured against this, because RSS rarely shrinks back after the warm-up (V8 keeps its pages).
  await settle();
  const preWarmupRss = process.memoryUsage().rss;
  const highWaterBeforeWarmup = process.resourceUsage().maxRSS * 1024;

  for (let i = 0; i < spec.warmup; i++) {
    log(`  warm-up ${i + 1}/${spec.warmup}`);
    try {
      await runner.run();
    } catch (error) {
      return { ...base, status: 'error', phase: 'warm-up', error: serializeError(error), durationMs: performance.now() - startedAt };
    }
  }

  await settle();
  const baselineRss = process.memoryUsage().rss;
  const runsMs = [];
  const firstBytesMs = [];
  const postRunRss = [];
  let sampledPeak = 0;
  let sampleCount = 0;
  let lastResult = null;

  for (let i = 0; i < spec.runs; i++) {
    log(`  run ${i + 1}/${spec.runs}`);
    const sampler = startRssSampler();
    const t0 = performance.now();
    try {
      lastResult = await runner.run();
    } catch (error) {
      sampler.stop();
      return {
        ...base,
        status: 'error',
        phase: `run ${i + 1}`,
        error: serializeError(error),
        timing: runsMs.length > 0 ? { runsMs } : undefined,
        durationMs: performance.now() - startedAt,
      };
    }
    runsMs.push(performance.now() - t0);
    postRunRss.push(process.memoryUsage().rss);
    const { peak, samples } = sampler.stop();
    sampledPeak = Math.max(sampledPeak, peak);
    sampleCount += samples;
    if (typeof lastResult?.firstByteMs === 'number') {
      firstBytesMs.push(lastResult.firstByteMs);
    }
    await settle();
  }

  const highWaterRss = process.resourceUsage().maxRSS * 1024;
  // The lifetime high-water mark only describes the op if it rose after the pre-warm-up snapshot
  // (otherwise it was set while materialising the input and would overstate the engine).
  const opPeakRss = Math.max(sampledPeak, ...postRunRss, highWaterRss > highWaterBeforeWarmup ? highWaterRss : 0);
  const peakRss = Math.max(opPeakRss, highWaterRss);
  return {
    ...base,
    status: 'ok',
    timing: {
      medianMs: median(runsMs),
      minMs: Math.min(...runsMs),
      maxMs: Math.max(...runsMs),
      runsMs,
    },
    memory: {
      preWarmupRssMB: toMB(preWarmupRss),
      baselineRssMB: toMB(baselineRss),
      peakRssMB: toMB(peakRss),
      footprintMB: toMB(Math.max(0, opPeakRss - preWarmupRss)),
      peakRssDeltaMB: toMB(Math.max(0, peakRss - baselineRss)),
      sampledPeakRssMB: sampledPeak > 0 ? toMB(sampledPeak) : null,
      postRunRssMB: toMB(Math.max(...postRunRss)),
      highWaterRssMB: toMB(highWaterRss),
      highWaterBeforeWarmupMB: toMB(highWaterBeforeWarmup),
      samples: sampleCount,
    },
    bytes: spec.op === 'write' ? (lastResult?.bytes ?? null) : statSync(spec.fixturePath).size,
    firstByteMs: firstBytesMs.length > 0 ? median(firstBytesMs) : null,
    rowsRead: spec.op === 'write' ? undefined : (lastResult?.rows ?? null),
    cellsRead: spec.op === 'write' ? undefined : (lastResult?.cells ?? null),
    durationMs: performance.now() - startedAt,
  };
}

/** Writes a read fixture with the requested engine; no timing, streams rows to keep the producer's footprint low. */
async function runFixture(spec) {
  const startedAt = performance.now();
  const shape = resolveShape(spec.dataset, spec.size);
  const engine = await loadEngine(spec.engine);
  if (engine.skipReason) {
    return { status: 'skipped', reason: engine.skipReason };
  }
  mkdirSync(dirname(spec.fixturePath), { recursive: true });
  try {
    const result = await engine.write(createRowIterator({ dataset: shape.dataset, rows: shape.rows, seed: spec.seed }), {
      columns: shape.columns,
      rowCount: shape.rows,
      outPath: spec.fixturePath,
    });
    return {
      status: 'ok',
      producer: engine.name,
      producerVersion: engine.version,
      dataset: shape.dataset,
      size: shape.size,
      rows: shape.rows,
      columns: shape.columns.length,
      cells: shape.cells,
      seed: spec.seed,
      bytes: result.bytes,
      truncatedOversizedCells: true,
      createdAt: new Date().toISOString(),
      durationMs: performance.now() - startedAt,
    };
  } catch (error) {
    return { status: 'error', producer: engine.name, error: serializeError(error), durationMs: performance.now() - startedAt };
  }
}

async function main() {
  const spec = JSON.parse(process.argv[2] ?? '{}');
  try {
    const result = spec.task === 'fixture' ? await runFixture(spec) : await runBench(spec);
    emit(result);
  } catch (error) {
    emit({ status: 'error', phase: 'setup', error: serializeError(error) });
  }
}

// run.mjs imports RESULT_MARKER from this module, so only act as the child when executed directly
if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  main();
}
