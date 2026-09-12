import { notImplemented } from '../internal/not-implemented';
import type { CellInput, SheetOptions, StyleId } from '../types';
import type { ZipEntryWriter } from '../zip/zip-writer';
import type { SharedStringWriter } from './shared-strings';
import type { StyleRegistry } from './styles';

export interface WorksheetWriterContext {
  readonly entry: ZipEntryWriter;
  readonly styles: StyleRegistry;
  /** Null when strings are written inline only. */
  readonly sharedStrings: SharedStringWriter | null;
  readonly date1904: boolean;
  readonly dates: 'local' | 'utc';
  readonly cellOverflow: 'truncate' | 'throw';
  readonly truncationSuffix: string;
  readonly onCellTruncated: (count: number) => void;
}

export interface WorksheetWriteSummary {
  readonly rows: number;
  readonly columns: number;
  readonly truncatedCells: number;
  /** A1 range of all written cells (`A1:T31`), or null when the sheet is empty. */
  readonly dimension: string | null;
}

/**
 * Row-at-a-time worksheet XML in the fixed schema order (sheetPr? dimension? sheetViews? sheetFormatPr? cols?
 * sheetData autoFilter? mergeCells?). Cells append to a string builder that is encoded and pushed to the zip entry
 * every 64 Ki characters; `writeRow` therefore only awaits at flush boundaries. Rows and columns are monotonic.
 */
export class WorksheetWriter {
  constructor(context: WorksheetWriterContext, options: SheetOptions) {
    void context;
    void options;
    throw notImplemented('sml/worksheet-writer');
  }

  /** 1-based row number the next `writeRow` produces. */
  get nextRow(): number {
    throw notImplemented('sml/worksheet-writer');
  }

  writeRow(values: readonly CellInput[], styles?: StyleId | readonly (StyleId | undefined)[]): Promise<void> {
    void values;
    void styles;
    throw notImplemented('sml/worksheet-writer');
  }

  merge(range: string): void {
    void range;
    throw notImplemented('sml/worksheet-writer');
  }

  /** Close `</sheetData>`, write autoFilter/mergeCells, flush, close the zip entry. */
  close(): Promise<WorksheetWriteSummary> {
    throw notImplemented('sml/worksheet-writer');
  }
}
