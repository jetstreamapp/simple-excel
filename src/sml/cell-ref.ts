import { notImplemented } from '../internal/not-implemented';

export const MAX_ROWS: number = 1_048_576;
export const MAX_COLUMNS: number = 16_384;

/** 0-based column index to letters (`0 -> 'A'`, `26 -> 'AA'`); cached, so repeated calls allocate nothing. */
export function columnLetters(columnIndex: number): string {
  void columnIndex;
  throw notImplemented('sml/cell-ref');
}

/** Letters to 0-based column index (`'A' -> 0`); -1 when not a valid column. */
export function columnIndexOf(letters: string): number {
  void letters;
  throw notImplemented('sml/cell-ref');
}

/** 0-based row and column to an A1 reference. */
export function formatRef(rowIndex: number, columnIndex: number): string {
  void rowIndex;
  void columnIndex;
  throw notImplemented('sml/cell-ref');
}

export interface CellRef {
  /** 0-based */
  readonly row: number;
  /** 0-based */
  readonly col: number;
}

/** Parse an A1 reference (`$` anchors tolerated). Returns null when the text is not a reference. */
export function parseRef(ref: string): CellRef | null {
  void ref;
  throw notImplemented('sml/cell-ref');
}

/** Column index of an A1 reference by scanning char codes; -1 when malformed. Hot path for the worksheet reader. */
export function columnIndexOfRef(ref: string): number {
  void ref;
  throw notImplemented('sml/cell-ref');
}

export interface CellRange {
  readonly start: CellRef;
  readonly end: CellRef;
}

/** Parse `A1:C3` (or a single cell as a 1x1 range). Returns null when malformed. */
export function parseRange(range: string): CellRange | null {
  void range;
  throw notImplemented('sml/cell-ref');
}
