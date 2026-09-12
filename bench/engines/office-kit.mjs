import { statSync } from 'node:fs';
import { setImmediate as yieldToEventLoop } from 'node:timers/promises';
import { fromFile, toFile } from '@office-kit/xlsx/node';
import { createWriteOnlyWorkbook, loadWorkbookStream } from '@office-kit/xlsx/streaming';
import { builtinFormatCode, isDateFormat } from '@office-kit/xlsx/styles';
import { excelToDate } from '@office-kit/xlsx/utils';
import { DATE_NUMBER_FORMAT, truncateRow } from '../lib/excel-limits.mjs';
import { packageVersion } from '../lib/package-version.mjs';

/**
 * `@office-kit/xlsx` streaming write-only workbook + read-only streaming iterator.
 *
 * API friction observed while writing this adapter (0.11.0):
 * - `appendRow` serialises a Date as a bare serial number; without a `{ value, style: { numberFormat } }`
 *   wrapper Excel shows `42370.5`. Wrapping costs a `stableStringify` of the style per cell (deduped
 *   through the xf pool, but the hashing is per call).
 * - Strings always go through the shared string table (no inline-string option). The whole SST stays
 *   resident until `finalize()`, so peak memory scales with distinct strings (see `strings-unique`).
 *   Worse, `finalize()` serialises the SST as ONE JS string (`serializeSharedStrings` → `Array.join`):
 *   with Jetstream-shaped rows (~1.7 KB of text each) it throws `RangeError: Invalid string length`
 *   (V8's ~2^29-char cap) before reaching 1M rows — the worksheet streams, the SST does not.
 * - `appendRow` is async but never actually yields, and the `toFile` sink only drains the fs stream when
 *   the event loop turns. A tight loop therefore piles deflated chunks up in memory; the docs recommend
 *   yielding (`setImmediate`) — this adapter does so every YIELD_EVERY_ROWS rows.
 * - The streaming reader never converts dates: `iterRows` yields numbers plus a `styleId`, and the caller
 *   resolves `workbook.styles.cellXfs[styleId].numFmtId` → format code → `isDateFormat` → `excelToDate`.
 * - `iterRows` is sparse (absent cells are skipped), so dense rows must be rebuilt by the caller.
 */
export const name = 'office-kit';
export const version = packageVersion('@office-kit/xlsx/streaming', '@office-kit/xlsx');
export const supportsStreaming = true;
export const readInput = 'path';

const YIELD_EVERY_ROWS = 256;
const DATE_STYLE = { numberFormat: DATE_NUMBER_FORMAT };

/** Wraps a sink so the harness can see when the first deflated chunk reaches the destination. */
function probeSink(sink, onFirstByte) {
  return {
    toBytes() {
      const writer = sink.toBytes();
      let sawFirst = false;
      return {
        write(chunk) {
          if (!sawFirst) {
            sawFirst = true;
            onFirstByte();
          }
          writer.write(chunk);
        },
        finish: () => writer.finish(),
        abort: cause => writer.abort?.(cause),
      };
    },
  };
}

function toWriteOnlyRow(row) {
  const truncated = truncateRow(row);
  let out = truncated;
  for (let i = 0; i < truncated.length; i++) {
    const value = truncated[i];
    if (value instanceof Date) {
      if (out === truncated) {
        out = truncated.slice();
      }
      out[i] = { value, style: DATE_STYLE };
    }
  }
  return out;
}

export async function write(rowIterable, { columns, outPath }) {
  const startedAt = performance.now();
  let firstByteMs = null;
  const sink = probeSink(toFile(outPath), () => {
    firstByteMs = performance.now() - startedAt;
  });
  const workbook = await createWriteOnlyWorkbook(sink);
  try {
    const worksheet = await workbook.addWorksheet('Sheet1');
    await worksheet.appendRow(columns);
    let rowIndex = 0;
    for (const row of rowIterable) {
      await worksheet.appendRow(toWriteOnlyRow(row));
      if (++rowIndex % YIELD_EVERY_ROWS === 0) {
        await yieldToEventLoop();
      }
    }
    await worksheet.close();
    await workbook.finalize();
  } catch (error) {
    workbook.abort(error);
    throw error;
  }
  return { bytes: statSync(outPath).size, firstByteMs };
}

/** Resolves whether a cellXfs index carries a date number format; memoised per workbook. */
function createDateStyleResolver(workbook) {
  const cache = new Map();
  return styleId => {
    let isDate = cache.get(styleId);
    if (isDate === undefined) {
      const xf = workbook.styles.cellXfs[styleId];
      const code = xf ? (workbook.styles.numFmts.get(xf.numFmtId) ?? builtinFormatCode(xf.numFmtId)) : undefined;
      isDate = isDateFormat(code);
      cache.set(styleId, isDate);
    }
    return isDate;
  };
}

export async function readTyped(path) {
  const workbook = await loadWorkbookStream(fromFile(path));
  const worksheet = workbook.openWorksheet(workbook.sheetNames[0]);
  const isDateStyle = createDateStyleResolver(workbook);
  const epoch = { epoch: workbook.date1904 ? 'mac' : 'windows' };
  let rows = 0;
  let columnCount = 0;
  for await (const cells of worksheet.iterRows()) {
    if (rows === 0) {
      columnCount = cells.length;
    }
    const dense = Array.from({ length: columnCount }, () => null);
    for (const cell of cells) {
      let value = cell.value;
      if (typeof value === 'number' && cell.styleId !== 0 && isDateStyle(cell.styleId)) {
        value = excelToDate(value, epoch);
      }
      dense[cell.col - 1] = value;
    }
    rows++;
  }
  await workbook.close();
  const dataRows = Math.max(0, rows - 1);
  return { rows: dataRows, cells: dataRows * columnCount };
}

/** No formatted-text mode exists; raw = every value stringified (dates stay serial numbers). */
export async function readRaw(path) {
  const workbook = await loadWorkbookStream(fromFile(path));
  const worksheet = workbook.openWorksheet(workbook.sheetNames[0]);
  let rows = 0;
  let columnCount = 0;
  for await (const values of worksheet.iterValues()) {
    if (rows === 0) {
      columnCount = values.length;
    }
    for (let i = 0; i < values.length; i++) {
      const value = values[i];
      values[i] = value === null || value === undefined ? '' : typeof value === 'string' ? value : String(value);
    }
    rows++;
  }
  await workbook.close();
  const dataRows = Math.max(0, rows - 1);
  return { rows: dataRows, cells: dataRows * columnCount };
}
