import { notImplemented } from '../internal/not-implemented';
import type { CellValue, OpenOptions, RawCell, RowsOptions } from '../types';

export interface WorksheetReadContext {
  /** Lazily loaded shared strings; called at most once per sheet read, only when a `t="s"` cell appears. */
  readonly sharedStrings: () => Promise<readonly string[]>;
  readonly isDateByXf: Uint8Array;
  readonly date1904: boolean;
  readonly dates: NonNullable<OpenOptions['dates']>;
  readonly errors: NonNullable<OpenOptions['errors']>;
  readonly maxXmlDepth: number;
  readonly maxTextLength: number;
}

export interface SheetRow {
  /** 1-based row number from the file (or inferred when `r` is missing). */
  readonly index: number;
  /** Dense, 0-based; trailing empties trimmed; holes are null. */
  readonly cells: CellValue[];
}

/**
 * Stream typed rows out of a worksheet part: UTF-8 decode -> tokenizer -> cell state machine. Handles missing `r`
 * attributes (row and column inferred), every `t` value (n, s, str, inlineStr, b, e, d), `<v/>`, `<is>` runs,
 * formulas (cached value, or text with `formulas: 'text'`), Strict ISO dates, and decodes `_xHHHH_` only when the
 * text contains `_x`. Yields in batches per decoded chunk; breaking out cancels the underlying stream.
 */
export function readWorksheetRows(
  chunks: AsyncIterable<Uint8Array>,
  context: WorksheetReadContext,
  options: RowsOptions,
): AsyncIterable<SheetRow> {
  void chunks;
  void context;
  void options;
  throw notImplemented('sml/worksheet-reader');
}

/** First `rowCount` rows as A1-keyed raw cells (the `head()` fast path); stops reading after the last wanted row. */
export function readWorksheetHead(
  chunks: AsyncIterable<Uint8Array>,
  context: WorksheetReadContext,
  rowCount: number,
): Promise<Map<string, RawCell>> {
  void chunks;
  void context;
  void rowCount;
  throw notImplemented('sml/worksheet-reader');
}
