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
 * more than a renderer can hold, and would fail every engine before it runs). simple-excel and
 * office-kit stream the generator; SheetJS has to materialise the array-of-arrays itself, which is part
 * of its cost. Same row mapping as the Node adapters (truncation, date style).
 *
 * The sinks differ deliberately. office-kit writes to a counting null sink (bytes discarded), which is
 * the engine's cost with no destination on top; simple-excel writes to `collectToBlob()`, the real
 * browser download path, where chunks are folded into sub-Blobs Chromium can page out to disk. Our
 * numbers therefore include keeping the finished file, which is the pessimistic side of that comparison.
 *
 * Importing `../../dist/esm/index.mjs` means `npm run build` has to have run before the bundle step.
 */
import { createWriteOnlyWorkbook } from '@office-kit/xlsx/streaming';
import * as XLSX from 'xlsx';
import { collectToBlob, createWorkbookWriter } from '../../dist/esm/index.mjs';
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

/** Wraps a ByteSink so the harness sees when the first compressed chunk reaches the destination. */
function probeSink(sink, onFirstByte) {
  let sawFirstChunk = false;
  return {
    write(chunk) {
      if (!sawFirstChunk) {
        sawFirstChunk = true;
        onFirstByte();
      }
      return sink.write(chunk);
    },
    close: () => sink.close(),
    abort: reason => sink.abort(reason),
  };
}

/** Feeds the generator through in `PROGRESS_EVERY_ROWS` batches so the page can report progress. */
function* reportingRows(rowIterable, report) {
  let rowIndex = 0;
  for (const row of rowIterable) {
    yield row;
    if (++rowIndex % PROGRESS_EVERY_ROWS === 0) {
      report(rowIndex);
    }
  }
}

const ENGINES = {
  'simple-excel': async (rowIterable, columns, report, rowCount) => {
    const startedAt = performance.now();
    let firstByteMs = null;
    const blobSink = collectToBlob();
    const sink = probeSink(blobSink, () => {
      firstByteMs = performance.now() - startedAt;
    });
    const workbook = createWorkbookWriter(sink);
    const sheet = workbook.addSheet('Sheet1', { header: columns, rowCount });
    await sheet.writeRows(reportingRows(rowIterable, report));
    await sheet.close();
    await workbook.close();
    // resolving the Blob keeps the finished file alive to the end of the run, so its cost is measured
    const blob = await blobSink.result();
    return { bytes: blob.size, firstByteMs };
  },
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
