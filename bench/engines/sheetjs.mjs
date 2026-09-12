import { writeFileSync } from 'node:fs';
import * as XLSX from 'xlsx';
import { truncateRow } from '../lib/excel-limits.mjs';

/**
 * SheetJS CE exactly as Jetstream calls it today (`prepareExcelFile` / `parseFile`):
 * dense array-of-arrays sheet, no shared strings, compression only above 10k rows, cells over
 * 32,767 chars truncated first. Fully synchronous and fully in-memory — the baseline every other
 * engine is compared against.
 */
export const name = 'sheetjs';
export const version = XLSX.version;
export const supportsStreaming = false;
/** SheetJS reads from an ArrayBuffer in Jetstream (file input / fetch), so the harness hands it bytes. */
export const readInput = 'bytes';

export async function write(rowIterable, { columns, rowCount, outPath }) {
  const aoa = [columns];
  for (const row of rowIterable) {
    aoa.push(truncateRow(row));
  }
  const worksheet = XLSX.utils.aoa_to_sheet(aoa, { dense: true });
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
  // type:'array' yields an ArrayBuffer (what Jetstream wraps in a Blob for download)
  const bytes = XLSX.write(workbook, { bookType: 'xlsx', bookSST: false, type: 'array', compression: rowCount > 10_000 });
  writeFileSync(outPath, new Uint8Array(bytes));
  return { bytes: bytes.byteLength };
}

export async function readTyped(bytes) {
  const workbook = XLSX.read(bytes, { cellText: false, cellDates: true, type: 'array' });
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const records = XLSX.utils.sheet_to_json(worksheet, {
    dateNF: 'yyyy"-"mm"-"dd"T"hh:mm:ss',
    defval: '',
    blankrows: false,
    rawNumbers: true,
  });
  const columnCount = records.length > 0 ? Object.keys(records[0]).length : 0;
  return { rows: records.length, cells: records.length * columnCount };
}

export async function readRaw(bytes) {
  const workbook = XLSX.read(bytes, { type: 'array' });
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false, defval: '', blankrows: false });
  const columnCount = rows.length > 0 ? rows[0].length : 0;
  const dataRows = Math.max(0, rows.length - 1);
  return { rows: dataRows, cells: dataRows * columnCount };
}
