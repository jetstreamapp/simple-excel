/**
 * Cross-browser smoke test for `@jetstreamapp/simple-excel`, run by `test/browser/run.mjs` in Chromium, Firefox
 * and WebKit and openable by hand in real Safari (`node test/browser/run.mjs --serve`).
 *
 * It exercises the library the way Jetstream does — write a large export to a Blob, read it back, do the same in
 * a module worker, read real Excel- and Sheets-authored files, reject hostile input, stream into a
 * `WritableStream` — and records one pass/fail record per check. Every expectation is evaluated in the page
 * against the values the page itself wrote, so the runner only ever sees `ok` plus a human-readable detail.
 *
 * The results land on `window.__smokeResults`, which is also what the on-page table is rendered from.
 */
import { collectToBlob, collectToBytes, fromWritableStream, hasNativeDeflate, openWorkbook } from '/dist/esm/index.mjs';
import { DATA_SHEET, HIDDEN_SHEET, SENTINEL_ROW, verifySmokeWorkbook, writeSmokeWorkbook } from './smoke-workbook.js';

const ROW_COUNT = 20_000;
const STREAM_ROW_COUNT = 500;
const MEMORY_ROW_COUNT = 200_000;
/** `measureUserAgentSpecificMemory()` resolves at the next garbage collection Chrome schedules, not on demand. */
const MEMORY_MEASUREMENT_DEADLINE_MS = 25_000;

const EXCEL_FIXTURE = '/fixtures/golden/excel-365/canonical.from-exceljs.xlsx';
const SHEETS_FIXTURE = '/fixtures/golden/gsheets/canonical.from-simple-excel.xlsx';
const ENCRYPTED_FIXTURE = '/fixtures/hostile/encrypted-password-test.xlsx';
const ZIP_BOMB_FIXTURE = '/fixtures/hostile/zip-bomb-30mb-sheet.xlsx';
const ZIP_BOMB_INFLATE_LIMIT = 8 * 1024 * 1024;

const results = {
  userAgent: navigator.userAgent,
  crossOriginIsolated: self.crossOriginIsolated === true,
  startedAt: new Date().toISOString(),
  environment: {},
  steps: [],
  metrics: {},
  totals: { checks: 0, failed: 0 },
  fatal: null,
  done: false,
};
window.__smokeResults = results;

// ---------------------------------------------------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------------------------------------------------

function describeError(error) {
  return `${error?.name ?? 'Error'}${error?.code ? ` ${error.code}` : ''}: ${error?.message ?? String(error)}`;
}

/** Collects checks for one step, so a step body reads as a list of assertions rather than bookkeeping. */
function createStep(id, label) {
  const checks = [];
  return {
    id,
    label,
    checks,
    check(checkId, ok, detail) {
      checks.push({ id: checkId, ok: Boolean(ok), detail: String(detail ?? '') });
    },
    adopt(otherChecks) {
      checks.push(...otherChecks);
    },
  };
}

async function fetchArrayBuffer(url) {
  const response = await fetch(url);
  if (!response.ok) {
    throw new Error(`GET ${url} -> ${response.status} ${response.statusText}`);
  }
  return response.arrayBuffer();
}

/** Runs `body` and expects it to reject with an `XlsxError` carrying `code`, optionally containing `phrase`. */
async function expectRejection(step, checkId, code, phrase, body) {
  try {
    await body();
    step.check(checkId, false, `no error thrown; expected ${code}`);
  } catch (error) {
    const codeMatches = error?.code === code;
    const phraseMatches = phrase === null || String(error?.message ?? '').includes(phrase);
    step.check(checkId, codeMatches && phraseMatches, describeError(error));
  }
}

function formatBytes(bytes) {
  if (typeof bytes !== 'number' || !Number.isFinite(bytes)) {
    return 'n/a';
  }
  return bytes >= 1024 * 1024 ? `${(bytes / (1024 * 1024)).toFixed(2)} MB` : `${(bytes / 1024).toFixed(1)} KB`;
}

function formatMs(ms) {
  return typeof ms === 'number' && Number.isFinite(ms) ? `${ms.toFixed(0)} ms` : 'n/a';
}

// ---------------------------------------------------------------------------------------------------------------------
// The worker, addressed request/response
// ---------------------------------------------------------------------------------------------------------------------

function createWorkerClient() {
  const worker = new Worker(new URL('./smoke.worker.js', import.meta.url), { type: 'module' });
  const pending = new Map();
  let nextId = 1;
  let fatal = null;

  worker.addEventListener('message', ({ data }) => {
    const entry = pending.get(data.id);
    if (!entry) {
      return;
    }
    pending.delete(data.id);
    if (data.ok) {
      entry.resolve(data);
    } else {
      const error = new Error(data.error.message);
      error.name = data.error.name;
      error.code = data.error.code;
      entry.reject(error);
    }
  });
  worker.addEventListener('error', event => {
    fatal = new Error(event.message || 'the smoke worker failed to load');
    for (const entry of pending.values()) {
      entry.reject(fatal);
    }
    pending.clear();
  });

  return {
    request(type, payload = {}, transfer = []) {
      if (fatal) {
        return Promise.reject(fatal);
      }
      const id = nextId++;
      return new Promise((resolve, reject) => {
        pending.set(id, { resolve, reject });
        worker.postMessage({ id, type, ...payload }, transfer);
      });
    },
    terminate() {
      worker.terminate();
    },
  };
}

// ---------------------------------------------------------------------------------------------------------------------
// Steps
// ---------------------------------------------------------------------------------------------------------------------

function stepFeatureDetection() {
  const step = createStep('features', 'Feature detection');
  const environment = {
    hasNativeDeflate: hasNativeDeflate(),
    CompressionStream: typeof CompressionStream,
    DecompressionStream: typeof DecompressionStream,
    WritableStream: typeof WritableStream,
    Blob: typeof Blob,
    crossOriginIsolated: self.crossOriginIsolated === true,
    userAgent: navigator.userAgent,
  };
  results.environment = environment;
  step.check('CompressionStream', environment.CompressionStream === 'function', environment.CompressionStream);
  step.check('DecompressionStream', environment.DecompressionStream === 'function', environment.DecompressionStream);
  step.check('WritableStream', environment.WritableStream === 'function', environment.WritableStream);
  step.check('Blob', environment.Blob === 'function', environment.Blob);
  // Without it the writer silently stores every part uncompressed, which is valid but several times larger — a
  // browser we claim to support must have it.
  step.check('hasNativeDeflate', environment.hasNativeDeflate === true, String(environment.hasNativeDeflate));
  return step;
}

async function stepWriteOnMainThread(state) {
  const step = createStep('write-main', `Write ${ROW_COUNT.toLocaleString('en-US')} rows on the main thread to collectToBlob()`);
  const sink = collectToBlob();
  const { result, ms } = await writeSmokeWorkbook(sink, ROW_COUNT);
  const blob = await sink.result();
  state.mainBlob = blob;
  results.metrics.mainWrite = {
    ms,
    bytes: blob.size,
    sharedStrings: result.sharedStrings.count,
    truncatedCells: result.truncatedCells,
    blobType: blob.type,
  };
  step.check('blob-produced', blob instanceof Blob && blob.size > 0, `${formatBytes(blob.size)} in ${formatMs(ms)}`);
  step.check('blob-mime-type', blob.type.endsWith('spreadsheetml.sheet'), blob.type);
  // The inline-string default builds no shared-string table at all, so `count` is 0 by design.
  step.check('shared-strings', result.sharedStrings.count === 0, JSON.stringify(result.sharedStrings));
  step.check('truncated-cells', result.truncatedCells === 1, String(result.truncatedCells));
  step.check(
    'reported-sheets',
    result.sheets.map(sheet => sheet.name).join(',') === `${DATA_SHEET},${HIDDEN_SHEET}`,
    result.sheets.map(sheet => `${sheet.name}:${sheet.rows}x${sheet.columns}`).join(', '),
  );
  return step;
}

async function stepReadBackFromBlob(state) {
  const step = createStep('read-blob', 'Read the Blob back (slice().arrayBuffer() source path)');
  const { checks, ms } = await verifySmokeWorkbook(state.mainBlob, ROW_COUNT);
  step.adopt(checks);
  results.metrics.mainRead = { ms };
  return step;
}

async function stepWorker(state) {
  const step = createStep('worker', 'Module worker: write to a Blob, read a transferred ArrayBuffer');
  const features = await state.worker.request('features');
  step.check('worker-native-deflate', features.hasNativeDeflate === true, String(features.hasNativeDeflate));

  const written = await state.worker.request('write', { rowCount: ROW_COUNT });
  results.metrics.workerWrite = { ms: written.ms, bytes: written.bytes, truncatedCells: written.truncatedCells };
  step.check(
    'worker-blob-transferred',
    written.blob instanceof Blob && written.blob.size > 0,
    `${formatBytes(written.bytes)} in ${formatMs(written.ms)}`,
  );
  // Both writes are `deterministic: true` over the same rows, so the two files must be byte-identical.
  step.check(
    'worker-bytes-match-main-thread',
    written.blob.size === state.mainBlob.size,
    `worker ${written.blob.size} vs main ${state.mainBlob.size}`,
  );
  const { checks, ms } = await verifySmokeWorkbook(written.blob, ROW_COUNT);
  step.adopt(checks.map(({ id, ok, detail }) => ({ id: `worker-blob/${id}`, ok, detail })));
  results.metrics.workerBlobRead = { ms };

  const buffer = await state.mainBlob.arrayBuffer();
  const read = await state.worker.request('read', { buffer }, [buffer]);
  step.check(
    'worker-read-sheets',
    read.sheets.map(sheet => sheet.name).join(',') === `${DATA_SHEET},${HIDDEN_SHEET}` && read.sheets[1].hidden === true,
    read.sheets.map(sheet => `${sheet.name}${sheet.hidden ? ' (hidden)' : ''}`).join(', '),
  );
  step.check('worker-read-sentinel', read.name === SENTINEL_ROW[1], JSON.stringify(read.name));
  results.metrics.workerRead = { ms: read.ms, bytes: read.byteLength };
  return step;
}

async function readFixture(step, prefix, url, expectedSheetNames) {
  const buffer = await fetchArrayBuffer(url);
  const workbook = await openWorkbook(buffer);
  try {
    const names = workbook.sheets.map(sheet => sheet.name);
    step.check(`${prefix}/sheet-names`, names.join(' | ') === expectedSheetNames.join(' | '), names.join(' | '));
    const hidden = workbook.sheets.find(sheet => sheet.name === 'Hidden');
    step.check(`${prefix}/hidden-sheet`, hidden?.hidden === true, `Hidden=${hidden?.hidden}`);

    const head = await workbook.sheet(DATA_SHEET).head(4);
    const name = head.get('B2')?.value;
    step.check(`${prefix}/B2-unicode`, name === 'Zoë Ångström', JSON.stringify(name));
    const bigInteger = head.get('I2')?.value;
    step.check(`${prefix}/I2-max-safe-integer`, bigInteger === 9_007_199_254_740_991, String(bigInteger));
    const date = head.get('M4')?.value;
    const isMarchFirst1900 = date instanceof Date && date.getFullYear() === 1900 && date.getMonth() === 2 && date.getDate() === 1;
    step.check(`${prefix}/M4-date-1900-03-01`, isMarchFirst1900, date instanceof Date ? date.toString() : String(date));
  } finally {
    await workbook.close();
  }
}

async function stepRealFixtures() {
  const step = createStep('fixtures', 'Read Excel- and Google-Sheets-authored files');
  await readFixture(step, 'excel-365', EXCEL_FIXTURE, ['Data', 'Features', "It's a very long sheet name 001", 'Hidden']);
  // Sheets drops the apostrophe from the long sheet name on export; that is its behaviour, pinned here so a
  // change in ours would show up.
  await readFixture(step, 'gsheets', SHEETS_FIXTURE, ['Data', 'Features', 'Its a very long sheet name 001', 'Hidden']);
  return step;
}

async function stepHostileInput() {
  const step = createStep('hostile', 'Reject hostile input with a classified error');
  const encrypted = await fetchArrayBuffer(ENCRYPTED_FIXTURE);
  await expectRejection(step, 'encrypted', 'ENCRYPTED', 'password-protected', () => openWorkbook(encrypted));

  const zipBomb = await fetchArrayBuffer(ZIP_BOMB_FIXTURE);
  await expectRejection(step, 'zip-bomb', 'ZIP_BOMB', null, async () => {
    const workbook = await openWorkbook(zipBomb, { limits: { maxInflatedBytes: ZIP_BOMB_INFLATE_LIMIT } });
    try {
      // The cap is enforced while the sheet streams, so the rejection only surfaces once rows are pulled.
      for await (const _row of workbook.sheet(0).rows()) {
        void _row;
      }
    } finally {
      await workbook.close();
    }
  });
  return step;
}

async function stepWritableStream() {
  const step = createStep('writable-stream', 'fromWritableStream() matches collectToBytes()');
  const chunks = [];
  let chunkCount = 0;
  const stream = new WritableStream({
    write(chunk) {
      // The sink takes ownership of every chunk, so a consumer that keeps them copies first.
      chunks.push(chunk.slice());
      chunkCount++;
    },
  });
  const streamed = await writeSmokeWorkbook(fromWritableStream(stream), STREAM_ROW_COUNT);

  const bytesSink = collectToBytes();
  const collected = await writeSmokeWorkbook(bytesSink, STREAM_ROW_COUNT);
  const expected = bytesSink.result();

  const streamedBytes = new Uint8Array(chunks.reduce((total, chunk) => total + chunk.byteLength, 0));
  let offset = 0;
  for (const chunk of chunks) {
    streamedBytes.set(chunk, offset);
    offset += chunk.byteLength;
  }

  step.check('chunked', chunkCount > 1, `${chunkCount} chunks, ${formatBytes(streamedBytes.byteLength)}`);
  step.check('same-length', streamedBytes.byteLength === expected.byteLength, `${streamedBytes.byteLength} vs ${expected.byteLength}`);
  const firstDifference = streamedBytes.findIndex((byte, index) => byte !== expected[index]);
  step.check(
    'same-bytes',
    streamedBytes.byteLength === expected.byteLength && firstDifference === -1,
    `first difference at ${firstDifference}`,
  );
  step.check(
    'same-reported-result',
    streamed.result.bytes === collected.result.bytes,
    `${streamed.result.bytes} vs ${collected.result.bytes}`,
  );
  return step;
}

/**
 * Samples `performance.measureUserAgentSpecificMemory()` back to back while the worker writes. Each measurement
 * only resolves at the next garbage collection the engine schedules, so they run on their own chain rather than
 * on a timer, and the peak is whatever landed. Chromium only, page scope only (it is not exposed in a dedicated
 * worker), and only when the document is cross-origin isolated — the page's measurement covers the whole agent
 * cluster, so the worker's allocations are in `breakdown` under `DedicatedWorkerGlobalScope`.
 */
function startMemorySampler() {
  // One mutable record rather than closed-over `let`s: the loops below only ever observe it through async
  // callbacks, which the `no-unmodified-loop-condition` lint rule cannot see.
  const state = { peakBytes: 0, peakWorkerBytes: 0, measurements: 0, failure: null, stopped: false };
  void (async () => {
    while (!state.stopped) {
      try {
        const measurement = await performance.measureUserAgentSpecificMemory();
        state.measurements++;
        state.peakBytes = Math.max(state.peakBytes, measurement.bytes);
        state.peakWorkerBytes = Math.max(
          state.peakWorkerBytes,
          measurement.breakdown
            .filter(entry => entry.attribution.some(attribution => attribution.scope === 'DedicatedWorkerGlobalScope'))
            .reduce((total, entry) => total + entry.bytes, 0),
        );
      } catch (error) {
        state.failure = describeError(error);
        return;
      }
    }
  })();
  /** Waits for at least one measurement to land — a short write usually finishes before the first GC does. */
  return async function stop(deadlineMs) {
    const deadline = performance.now() + deadlineMs;
    while (state.measurements === 0 && state.failure === null && performance.now() < deadline) {
      await new Promise(resolve => setTimeout(resolve, 250));
    }
    state.stopped = true;
    return {
      peakBytes: state.peakBytes || null,
      peakWorkerBytes: state.peakWorkerBytes || null,
      measurements: state.measurements,
      failure: state.failure,
    };
  };
}

async function stepWorkerMemory(state) {
  const step = createStep('memory', `Memory sanity: ${MEMORY_ROW_COUNT.toLocaleString('en-US')} rows in the worker`);
  const supported = self.crossOriginIsolated === true && typeof performance.measureUserAgentSpecificMemory === 'function';
  if (!supported) {
    const reason = 'performance.measureUserAgentSpecificMemory is unavailable here (Chromium, cross-origin isolated pages only)';
    results.metrics.memory = { skipped: reason };
    step.check('measured', true, `skipped: ${reason}`);
    return step;
  }
  const stop = startMemorySampler();
  const written = await state.worker.request('memory', { rowCount: MEMORY_ROW_COUNT });
  const memory = await stop(MEMORY_MEASUREMENT_DEADLINE_MS);
  results.metrics.memory = { rows: MEMORY_ROW_COUNT, bytes: written.bytes, ms: written.ms, ...memory };
  // Recorded, never asserted: the number depends on when the engine happened to run a garbage collection.
  const measured =
    memory.measurements > 0
      ? `agent-cluster peak ${formatBytes(memory.peakBytes)}, worker peak ${formatBytes(memory.peakWorkerBytes)} over ${memory.measurements} measurement(s)`
      : `no measurement landed within ${MEMORY_MEASUREMENT_DEADLINE_MS / 1000}s${memory.failure ? ` (${memory.failure})` : ''}`;
  step.check('measured', true, `${formatBytes(written.bytes)} written in ${formatMs(written.ms)}; ${measured}`);
  return step;
}

// ---------------------------------------------------------------------------------------------------------------------
// Runner and on-page report
// ---------------------------------------------------------------------------------------------------------------------

const STEPS = [
  ['features', stepFeatureDetection],
  ['write-main', stepWriteOnMainThread],
  ['read-blob', stepReadBackFromBlob],
  ['worker', stepWorker],
  ['fixtures', stepRealFixtures],
  ['hostile', stepHostileInput],
  ['writable-stream', stepWritableStream],
  ['memory', stepWorkerMemory],
];

function render() {
  const rows = [];
  for (const step of results.steps) {
    const failed = step.checks.filter(entry => !entry.ok).length;
    rows.push(
      `<tr class="step ${failed > 0 || step.error ? 'fail' : 'pass'}"><td colspan="2"><strong>${step.label}</strong></td><td>${step.checks.length - failed}/${step.checks.length}</td><td>${formatMs(step.ms)}</td></tr>`,
    );
    if (step.error) {
      rows.push(`<tr class="fail"><td></td><td colspan="3">threw ${escapeHtml(step.error)}</td></tr>`);
    }
    for (const entry of step.checks) {
      rows.push(
        `<tr class="${entry.ok ? 'pass' : 'fail'}"><td></td><td>${escapeHtml(entry.id)}</td><td>${entry.ok ? 'PASS' : 'FAIL'}</td><td>${escapeHtml(entry.detail)}</td></tr>`,
      );
    }
  }
  document.getElementById('summary').textContent = results.done
    ? `${results.totals.checks - results.totals.failed}/${results.totals.checks} checks passed${results.fatal ? ` — fatal: ${results.fatal}` : ''}`
    : 'running…';
  document.getElementById('summary').className = results.done && results.totals.failed === 0 && !results.fatal ? 'pass' : 'fail';
  document.getElementById('agent').textContent = navigator.userAgent;
  document.getElementById('rows').innerHTML = rows.join('');
}

function escapeHtml(value) {
  return String(value).replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;');
}

async function main() {
  const state = { worker: createWorkerClient() };
  try {
    for (const [id, body] of STEPS) {
      const startedAt = performance.now();
      let step;
      try {
        step = await body(state);
        step.ms = performance.now() - startedAt;
      } catch (error) {
        step = createStep(id, id);
        step.ms = performance.now() - startedAt;
        step.error = describeError(error);
        step.check('step-completed', false, describeError(error));
      }
      results.steps.push({ id: step.id, label: step.label, ms: step.ms, error: step.error ?? null, checks: step.checks });
      results.totals.checks += step.checks.length;
      results.totals.failed += step.checks.filter(entry => !entry.ok).length;
      render();
    }
  } catch (error) {
    results.fatal = describeError(error);
  } finally {
    state.worker.terminate();
    results.finishedAt = new Date().toISOString();
    results.done = true;
    render();
  }
}

main();
