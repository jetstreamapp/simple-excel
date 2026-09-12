/**
 * The module worker half of the browser smoke test. Everything Jetstream would move off the main thread happens
 * here: writing an export to `collectToBlob()` (the Blob is transferred back by reference), reading a workbook
 * out of an `ArrayBuffer` that was transferred in, and a best-effort memory measurement around a large write.
 *
 * Protocol (page -> worker): { id, type: 'features' | 'write' | 'read' | 'memory', ... }
 * Protocol (worker -> page): { id, ok: true, ...payload } | { id, ok: false, error: { name, code, message } }
 */
import { collectToBlob, hasNativeDeflate, openWorkbook } from '/dist/esm/index.mjs';
import { DATA_SHEET, writeSmokeWorkbook } from './smoke-workbook.js';

/** Holds the memory step's finished workbook alive while the page takes its measurement. */
const retainedForMeasurement = [];

function toErrorPayload(error) {
  return { name: error?.name ?? 'Error', code: error?.code ?? null, message: error?.message ?? String(error) };
}

async function writeToBlob(rowCount) {
  const sink = collectToBlob();
  const { result, ms } = await writeSmokeWorkbook(sink, rowCount);
  const blob = await sink.result();
  return { blob, result, ms };
}

const HANDLERS = {
  features() {
    return { hasNativeDeflate: hasNativeDeflate(), crossOriginIsolated: self.crossOriginIsolated === true };
  },

  async write({ rowCount }) {
    const { blob, result, ms } = await writeToBlob(rowCount);
    return {
      payload: {
        blob,
        bytes: blob.size,
        ms,
        sharedStrings: result.sharedStrings,
        truncatedCells: result.truncatedCells,
        hasNativeDeflate: hasNativeDeflate(),
      },
    };
  },

  /** Reads a workbook out of an ArrayBuffer the page transferred in, so the bytes never crossed as a copy. */
  async read({ buffer }) {
    const startedAt = performance.now();
    const workbook = await openWorkbook(buffer);
    try {
      const sheets = workbook.sheets.map(sheet => ({ name: sheet.name, hidden: sheet.hidden }));
      const { rows } = await workbook.sheet(DATA_SHEET).toObjects({ defval: null, maxRows: 2 });
      return { payload: { sheets, name: rows[0]?.Name ?? null, byteLength: buffer.byteLength, ms: performance.now() - startedAt } };
    } finally {
      await workbook.close();
    }
  },

  /**
   * The large write for the memory step. The measurement itself is taken in the page:
   * `performance.measureUserAgentSpecificMemory()` is not exposed in a dedicated worker, and the page's
   * measurement covers the whole agent cluster (this worker included) anyway.
   */
  async memory({ rowCount }) {
    const { blob, ms } = await writeToBlob(rowCount);
    // Keep the finished file referenced: the page measures after the write returns, and a Blob that has already
    // been collected would not be in the number. The worker is terminated right after this step.
    retainedForMeasurement.push(blob);
    return { payload: { bytes: blob.size, ms, rowCount } };
  },
};

self.addEventListener('message', async event => {
  const { id, type, ...request } = event.data;
  const handler = HANDLERS[type];
  try {
    if (!handler) {
      throw new Error(`Unknown smoke worker request "${type}"`);
    }
    const outcome = await handler(request);
    self.postMessage({ id, ok: true, ...(outcome?.payload ?? outcome) });
  } catch (error) {
    self.postMessage({ id, ok: false, error: toErrorPayload(error) });
  }
});
