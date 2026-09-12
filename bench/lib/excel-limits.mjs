/**
 * Jetstream truncates oversized strings before handing rows to any xlsx writer (Salesforce long
 * text areas go up to 131,072 chars and subquery records are JSON-stringified into one cell).
 * Every engine adapter applies the same rule so the bytes they are asked to write are identical.
 */
export const EXCEL_MAX_CELL_CHARS = 32_767;
export const TRUNCATED_SUFFIX = '...(truncated)';
const KEEP_CHARS = EXCEL_MAX_CELL_CHARS - TRUNCATED_SUFFIX.length;

export function truncateCell(value) {
  if (typeof value === 'string' && value.length > EXCEL_MAX_CELL_CHARS) {
    return value.slice(0, KEEP_CHARS) + TRUNCATED_SUFFIX;
  }
  return value;
}

/**
 * Returns the row with oversized strings truncated. Only allocates a copy when something actually
 * changed, mirroring Jetstream's `truncateCellsForExcel`.
 */
export function truncateRow(row) {
  let copy = null;
  for (let i = 0; i < row.length; i++) {
    const value = row[i];
    if (typeof value === 'string' && value.length > EXCEL_MAX_CELL_CHARS) {
      copy ??= row.slice();
      copy[i] = truncateCell(value);
    }
  }
  return copy ?? row;
}

/** Number format applied to Date cells by engines that need an explicit one (Excel shows a serial otherwise). */
export const DATE_NUMBER_FORMAT = 'yyyy-mm-dd hh:mm:ss';
