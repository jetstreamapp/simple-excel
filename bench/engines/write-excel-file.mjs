import { createRequire } from 'node:module';
import { statSync } from 'node:fs';
import readXlsxFile from 'read-excel-file/node';
import writeXlsxFile from 'write-excel-file/node';
import { DATE_NUMBER_FORMAT, truncateCell } from '../lib/excel-limits.mjs';

/**
 * `write-excel-file` (writer) paired with its sibling `read-excel-file` (reader). Both are whole-file,
 * in-memory APIs: the writer needs every row as `{ value, type }` cell objects up front, the reader
 * returns a fully materialised `rows` array. Neither has a "raw strings" mode.
 */
export const name = 'write-excel-file';
const require = createRequire(import.meta.url);
export const version = `${require('write-excel-file/package.json').version} (read-excel-file ${require('read-excel-file/package.json').version})`;
export const supportsStreaming = false;
export const readInput = 'path';

function toCell(value) {
  if (value === null || value === undefined) {
    return null;
  }
  if (value instanceof Date) {
    return { value, type: Date };
  }
  switch (typeof value) {
    case 'number':
      return { value, type: Number };
    case 'boolean':
      return { value, type: Boolean };
    default:
      return { value: truncateCell(String(value)), type: String };
  }
}

export async function write(rowIterable, { columns, outPath }) {
  const data = [columns.map(column => ({ value: column, type: String }))];
  for (const row of rowIterable) {
    data.push(row.map(toCell));
  }
  await writeXlsxFile(data, { filePath: outPath, dateFormat: DATE_NUMBER_FORMAT });
  return { bytes: statSync(outPath).size };
}

/** Built-in date formats come back as Date objects; custom ones would need the `dateFormat` option. */
export async function readTyped(path) {
  const rows = await readXlsxFile(path);
  const columnCount = rows.length > 0 ? rows[0].length : 0;
  const dataRows = Math.max(0, rows.length - 1);
  return { rows: dataRows, cells: dataRows * columnCount };
}

/** read-excel-file has no formatted/raw-string mode. */
export const readRaw = null;
