/**
 * Browser-side write benchmark. Bundled by `lib/chrome.mjs` (esbuild) into
 * `.generated/bench/web/worker.bundle.js` and run as a module Worker by `index.html`.
 *
 * Protocol (page -> worker): { type: 'run', runId, engine, dataset, rows, seed }
 * Protocol (worker -> page): { type: 'phase', runId, phase: 'write-start', at }
 *                            { type: 'progress', runId, rows }
 *                            { type: 'result', runId, ok, ms, bytes, firstByteMs, rows, cells }
 *                            { type: 'result', runId, ok: false, error: { name, message } }
 *
 * Rows are pulled lazily from the seeded generator (a materialised 1M-row mixed dataset is ~4.5 GB,
 * more than a renderer can hold, and would fail both engines before they run). office-kit streams the
 * generator; SheetJS has to materialise the array-of-arrays itself, which is part of its cost. Output
 * goes to a counting null sink (bytes discarded) so the numbers are the engine's own cost without a
 * Blob/OPFS destination on top. Same row mapping as the Node adapters (truncation, date style).
 */
import { createWriteOnlyWorkbook } from '@office-kit/xlsx/streaming';
import * as XLSX from 'xlsx';
import { createRowIterator, getDataset } from '../lib/dataset.mjs';
import { DATE_NUMBER_FORMAT, truncateRow } from '../lib/excel-limits.mjs';

const DATE_STYLE = { numberFormat: DATE_NUMBER_FORMAT };
const PROGRESS_EVERY_ROWS = 25_000;

function createCountingNullSink(onFirstByte) {
  let bytes = 0;
  return {
    bytes: () => bytes,
    toBytes() {
      return {
        write(chunk) {
          if (bytes === 0) {
            onFirstByte();
          }
          bytes += chunk.byteLength;
        },
        async finish() {
          return new Uint8Array(0);
        },
        abort() {},
      };
    },
  };
}

function toWriteOnlyRow(row) {
  const truncated = truncateRow(row);
  let out = truncated;
  for (let i = 0; i < truncated.length; i++) {
    if (truncated[i] instanceof Date) {
      if (out === truncated) {
        out = truncated.slice();
      }
      out[i] = { value: truncated[i], style: DATE_STYLE };
    }
  }
  return out;
}

const ENGINES = {
  'office-kit': async (rowIterable, columns, report) => {
    const startedAt = performance.now();
    let firstByteMs = null;
    const sink = createCountingNullSink(() => {
      firstByteMs = performance.now() - startedAt;
    });
    const workbook = await createWriteOnlyWorkbook(sink);
    const worksheet = await workbook.addWorksheet('Sheet1');
    await worksheet.appendRow(columns);
    let rowIndex = 0;
    for (const row of rowIterable) {
      await worksheet.appendRow(toWriteOnlyRow(row));
      if (++rowIndex % PROGRESS_EVERY_ROWS === 0) {
        report(rowIndex);
      }
    }
    await worksheet.close();
    await workbook.finalize();
    return { bytes: sink.bytes(), firstByteMs };
  },
  sheetjs: async (rowIterable, columns, report, rowCount) => {
    const aoa = [columns];
    for (const row of rowIterable) {
      aoa.push(truncateRow(row));
      if ((aoa.length - 1) % PROGRESS_EVERY_ROWS === 0) {
        report(aoa.length - 1);
      }
    }
    const worksheet = XLSX.utils.aoa_to_sheet(aoa, { dense: true });
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    const bytes = XLSX.write(workbook, { bookType: 'xlsx', bookSST: false, type: 'array', compression: rowCount > 10_000 });
    return { bytes: bytes.byteLength, firstByteMs: null };
  },
};

self.onmessage = async event => {
  const { type, runId, engine, dataset, rows, seed } = event.data;
  if (type !== 'run') {
    return;
  }
  const post = message => self.postMessage({ runId, ...message });
  try {
    const implementation = ENGINES[engine];
    if (!implementation) {
      throw new Error(`Unknown browser engine "${engine}" (known: ${Object.keys(ENGINES).join(', ')})`);
    }
    const { columns } = getDataset(dataset);
    post({ type: 'phase', phase: 'write-start', at: performance.now() });
    const writeStart = performance.now();
    const result = await implementation(
      createRowIterator({ dataset, rows, seed }),
      columns,
      done => post({ type: 'progress', rows: done }),
      rows,
    );
    const ms = performance.now() - writeStart;
    post({ type: 'result', ok: true, ms, bytes: result.bytes, firstByteMs: result.firstByteMs, rows, cells: rows * columns.length });
  } catch (error) {
    post({ type: 'result', ok: false, error: { name: error?.name ?? 'Error', message: error?.message ?? String(error) } });
  }
};
