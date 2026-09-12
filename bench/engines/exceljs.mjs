import { createRequire } from 'node:module';
import { createWriteStream, statSync } from 'node:fs';
import { Writable } from 'node:stream';
import ExcelJS from 'exceljs';
import { truncateRow } from '../lib/excel-limits.mjs';

/**
 * ExcelJS streaming writer (`WorkbookWriter`) and streaming reader (`WorkbookReader`).
 * Writer notes: `useStyles: false` still emits a built-in `mm-dd-yy` style for Date cells (the mock
 * stylesheet special-cases dates); `useSharedStrings: false` writes inline strings.
 */
export const name = 'exceljs';
export const version = createRequire(import.meta.url)('exceljs/package.json').version;
export const supportsStreaming = true;
export const readInput = 'path';

/** Writable that records when the first zip chunk arrives, then forwards everything to the file. */
function createProbedFileStream(outPath, onFirstByte) {
  const file = createWriteStream(outPath);
  let sawFirst = false;
  const probe = new Writable({
    write(chunk, encoding, callback) {
      if (!sawFirst) {
        sawFirst = true;
        onFirstByte();
      }
      if (!file.write(chunk, encoding)) {
        file.once('drain', callback);
      } else {
        callback();
      }
    },
    final(callback) {
      file.end(callback);
    },
  });
  return probe;
}

export async function write(rowIterable, { columns, outPath }) {
  const startedAt = performance.now();
  let firstByteMs = null;
  const stream = createProbedFileStream(outPath, () => {
    firstByteMs = performance.now() - startedAt;
  });
  const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({ stream, useStyles: false, useSharedStrings: false });
  const worksheet = workbook.addWorksheet('Sheet1');
  worksheet.addRow(columns).commit();
  for (const row of rowIterable) {
    worksheet.addRow(truncateRow(row)).commit();
  }
  await worksheet.commit();
  await workbook.commit();
  return { bytes: statSync(outPath).size, firstByteMs };
}

async function readRows(path, styles, mapValue) {
  const reader = new ExcelJS.stream.xlsx.WorkbookReader(path, {
    worksheets: 'emit',
    sharedStrings: 'cache',
    styles,
    hyperlinks: 'ignore',
    entries: 'ignore',
  });
  let rows = 0;
  let columnCount = 0;
  for await (const worksheet of reader) {
    for await (const row of worksheet) {
      // row.values is 1-based (index 0 is always empty)
      const values = row.values;
      if (rows === 0) {
        columnCount = values.length - 1;
      }
      for (let i = 1; i < values.length; i++) {
        values[i] = mapValue(values[i]);
      }
      rows++;
    }
    break;
  }
  const dataRows = Math.max(0, rows - 1);
  return { rows: dataRows, cells: dataRows * columnCount };
}

/** With `styles: 'cache'` ExcelJS converts date-formatted numbers to Date objects itself. */
export async function readTyped(path) {
  return readRows(path, 'cache', value => (value === undefined ? null : value));
}

/** No formatted-text mode exists; raw = styles ignored (dates stay serial numbers) and every value stringified. */
export async function readRaw(path) {
  return readRows(path, 'ignore', value => (value === undefined || value === null ? '' : String(value)));
}
