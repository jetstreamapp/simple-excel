import { XlsxError } from '../errors';

export const MAX_ROWS: number = 1_048_576;
export const MAX_COLUMNS: number = 16_384;

const CODE_UPPER_A = 65;
const CODE_UPPER_Z = 90;
const CODE_ZERO = 48;
const CODE_NINE = 57;
const CODE_DOLLAR = 36;
/** `XFD` is the widest column there is, so more than three letters is always out of range. */
const MAX_COLUMN_LETTERS = 3;

/** Grown on demand: the writer asks for a column's letters once per cell and never builds the same string twice. */
const columnLettersCache: string[] = [];

function buildColumnLetters(columnIndex: number): string {
  let letters = '';
  let remaining = columnIndex;
  while (remaining >= 0) {
    letters = String.fromCharCode(CODE_UPPER_A + (remaining % 26)) + letters;
    remaining = Math.floor(remaining / 26) - 1;
  }
  return letters;
}

/** 0-based column index to letters (`0 -> 'A'`, `26 -> 'AA'`); cached, so repeated calls allocate nothing. */
export function columnLetters(columnIndex: number): string {
  const cached = columnLettersCache[columnIndex];
  if (cached !== undefined) {
    return cached;
  }
  if (!Number.isInteger(columnIndex) || columnIndex < 0 || columnIndex >= MAX_COLUMNS) {
    throw new XlsxError('ROW_OUT_OF_RANGE', `Column ${columnIndex} is outside the ${MAX_COLUMNS} columns a sheet can hold.`, {
      columnIndex,
    });
  }
  const letters = buildColumnLetters(columnIndex);
  columnLettersCache[columnIndex] = letters;
  return letters;
}

/** Letters to 0-based column index (`'A' -> 0`); -1 when not a valid column. */
export function columnIndexOf(letters: string): number {
  const { length } = letters;
  if (length === 0 || length > MAX_COLUMN_LETTERS) {
    return -1;
  }
  let oneBasedColumn = 0;
  for (let i = 0; i < length; i++) {
    const code = letters.charCodeAt(i);
    if (code < CODE_UPPER_A || code > CODE_UPPER_Z) {
      return -1;
    }
    oneBasedColumn = oneBasedColumn * 26 + (code - CODE_UPPER_A + 1);
  }
  return oneBasedColumn > MAX_COLUMNS ? -1 : oneBasedColumn - 1;
}

/** 0-based row and column to an A1 reference. */
export function formatRef(rowIndex: number, columnIndex: number): string {
  if (!Number.isInteger(rowIndex) || rowIndex < 0 || rowIndex >= MAX_ROWS) {
    throw new XlsxError('ROW_OUT_OF_RANGE', `Row ${rowIndex + 1} is outside the ${MAX_ROWS} rows a sheet can hold.`, { rowIndex });
  }
  return `${columnLetters(columnIndex)}${rowIndex + 1}`;
}

export interface CellRef {
  /** 0-based */
  readonly row: number;
  /** 0-based */
  readonly col: number;
}

/** Parse an A1 reference (`$` anchors tolerated). Returns null when the text is not a reference. */
export function parseRef(ref: string): CellRef | null {
  let position = 0;
  if (ref.charCodeAt(position) === CODE_DOLLAR) {
    position++;
  }
  let oneBasedColumn = 0;
  let letterCount = 0;
  while (position < ref.length) {
    const code = ref.charCodeAt(position);
    if (code < CODE_UPPER_A || code > CODE_UPPER_Z) {
      break;
    }
    oneBasedColumn = oneBasedColumn * 26 + (code - CODE_UPPER_A + 1);
    letterCount++;
    position++;
  }
  if (letterCount === 0 || letterCount > MAX_COLUMN_LETTERS || oneBasedColumn > MAX_COLUMNS) {
    return null;
  }
  if (ref.charCodeAt(position) === CODE_DOLLAR) {
    position++;
  }
  let oneBasedRow = 0;
  let digitCount = 0;
  while (position < ref.length) {
    const code = ref.charCodeAt(position);
    if (code < CODE_ZERO || code > CODE_NINE) {
      break;
    }
    oneBasedRow = oneBasedRow * 10 + (code - CODE_ZERO);
    digitCount++;
    position++;
  }
  if (digitCount === 0 || position !== ref.length || oneBasedRow < 1 || oneBasedRow > MAX_ROWS) {
    return null;
  }
  return { row: oneBasedRow - 1, col: oneBasedColumn - 1 };
}

/**
 * Column index of an A1 reference by scanning char codes; -1 when malformed. Hot path for the worksheet reader, which
 * is why it repeats `parseRef`'s letter scan rather than share a result object with it. Only the first character of
 * the row part is checked: the reader already knows the row from the enclosing `<row>` element.
 */
export function columnIndexOfRef(ref: string): number {
  let position = 0;
  if (ref.charCodeAt(position) === CODE_DOLLAR) {
    position++;
  }
  let oneBasedColumn = 0;
  let letterCount = 0;
  while (position < ref.length) {
    const code = ref.charCodeAt(position);
    if (code < CODE_UPPER_A || code > CODE_UPPER_Z) {
      break;
    }
    oneBasedColumn = oneBasedColumn * 26 + (code - CODE_UPPER_A + 1);
    letterCount++;
    position++;
  }
  if (letterCount === 0 || letterCount > MAX_COLUMN_LETTERS || oneBasedColumn > MAX_COLUMNS) {
    return -1;
  }
  if (ref.charCodeAt(position) === CODE_DOLLAR) {
    position++;
  }
  // `charCodeAt` past the end is NaN, which fails this test - a reference with no row part is malformed.
  const rowCode = ref.charCodeAt(position);
  if (!(rowCode >= CODE_ZERO && rowCode <= CODE_NINE)) {
    return -1;
  }
  return oneBasedColumn - 1;
}

export interface CellRange {
  readonly start: CellRef;
  readonly end: CellRef;
}

/** Parse `A1:C3` (or a single cell as a 1x1 range). Returns null when malformed. */
export function parseRange(range: string): CellRange | null {
  const separator = range.indexOf(':');
  if (separator === -1) {
    const single = parseRef(range);
    return single === null ? null : { start: single, end: single };
  }
  const first = parseRef(range.slice(0, separator));
  const second = parseRef(range.slice(separator + 1));
  if (first === null || second === null) {
    return null;
  }
  return {
    start: { row: Math.min(first.row, second.row), col: Math.min(first.col, second.col) },
    end: { row: Math.max(first.row, second.row), col: Math.max(first.col, second.col) },
  };
}

/** Format a range from 0-based corners (`formatRange(0, 0, 0, 2)` -> `'A1:C1'`); corners are normalized. */
export function formatRange(startRow: number, startColumn: number, endRow: number, endColumn: number): string {
  const top = Math.min(startRow, endRow);
  const bottom = Math.max(startRow, endRow);
  const left = Math.min(startColumn, endColumn);
  const right = Math.max(startColumn, endColumn);
  return `${formatRef(top, left)}:${formatRef(bottom, right)}`;
}
