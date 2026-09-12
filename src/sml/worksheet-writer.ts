import { XlsxError } from '../errors';
import type { CellError, CellInput, SheetOptions, StyleId } from '../types';
import { encodeCellText, needsSpacePreserve } from '../xml/escape';
import { encodeXmlChunk } from '../xml/utf8';
import type { ZipEntryWriter } from '../zip/zip-writer';
import { columnLetters, formatRef, MAX_COLUMNS, MAX_ROWS, parseRange } from './cell-ref';
import { componentsFromDate, type DateComponents, serialFromComponents } from './date';
import type { SharedStringWriter } from './shared-strings';
import type { StyleRegistry } from './styles';

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const SPREADSHEET_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const OFFICE_REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** The builder is joined, encoded and pushed to the zip entry once it holds this many UTF-16 units. */
const FLUSH_THRESHOLD_CHARS = 64 * 1024;
/** Excel's hard per-cell limit; a longer string is a repair-dialog error (EC-CELL-32767-LIMIT). */
const MAX_CELL_CHARS = 32_767;
/** 2^53: past this a bigint no longer survives the trip through a double, so it is written as text. */
const MAX_EXACT_BIGINT = 9_007_199_254_740_992n;

/** Excel's own defaults, in inches. All six attributes are required by the schema. */
const PAGE_MARGINS = '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>';

/** Handed back by `writeRow` when nothing had to be flushed, so a row costs no promise allocation. */
const NOTHING_TO_FLUSH: Promise<void> = Promise.resolve();

export interface WorksheetWriterContext {
  readonly entry: ZipEntryWriter;
  readonly styles: StyleRegistry;
  /** Null when strings are written inline only. */
  readonly sharedStrings: SharedStringWriter | null;
  readonly date1904: boolean;
  readonly dates: 'local' | 'utc';
  readonly cellOverflow: 'truncate' | 'throw';
  readonly truncationSuffix: string;
  /** Called once at close with the number of cells this sheet truncated, and only when that is more than zero. */
  readonly onCellTruncated: (count: number) => void;
  /** Excel marks exactly one sheet as the selected tab; the facade passes true for the first visible sheet. */
  readonly tabSelected?: boolean;
}

export interface WorksheetWriteSummary {
  readonly rows: number;
  readonly columns: number;
  readonly truncatedCells: number;
  /** A1 range of all written cells (`A1:T31`), or null when the sheet is empty. */
  readonly dimension: string | null;
}

/**
 * The A1 range covering `rows` rows and `columns` columns from the top-left cell. `<dimension>`, the sheet summary
 * and the autofilter (header row through the last written row) are all this same shape.
 */
export function sheetRange(rows: number, columns: number): string {
  return `A1:${columnLetters(columns - 1)}${rows}`;
}

/**
 * Row-at-a-time worksheet XML in the fixed schema order (sheetPr? dimension? sheetViews? sheetFormatPr? cols?
 * sheetData autoFilter? mergeCells?). Cells append to a string builder that is encoded and pushed to the zip entry
 * every 64 Ki characters; `writeRow` therefore only awaits at flush boundaries. Rows and columns are monotonic.
 *
 * The first failure is latched: the XML is left half-written at that point, so every later call reports the same
 * error rather than producing a file Excel would refuse. Callers recover by aborting the workbook.
 */
export class WorksheetWriter {
  private readonly entry: ZipEntryWriter;
  private readonly styles: StyleRegistry;
  private readonly sharedStrings: SharedStringWriter | null;
  private readonly date1904: boolean;
  private readonly dates: 'local' | 'utc';
  private readonly cellOverflow: 'truncate' | 'throw';
  private readonly truncationSuffix: string;
  private readonly onCellTruncated: (count: number) => void;
  private readonly autoFilter: boolean;
  /**
   * Columns the caller announced up front, used for `<dimension>`, which is written before any row. The wider of the
   * header and the column definitions wins: a dimension that is too wide costs nothing, while one that is too narrow
   * makes readers that trust it (SheetJS's `!ref`, for one) drop the cells outside it.
   */
  private readonly declaredColumns: number;

  private readonly pending: string[] = [];
  private readonly merges: string[] = [];
  private pendingChars = 0;
  private rowsWritten = 0;
  private columnsWritten = 0;
  private rowColumns = 0;
  private truncatedCells = 0;
  private closed = false;
  private failure: unknown;

  constructor(context: WorksheetWriterContext, options: SheetOptions) {
    this.entry = context.entry;
    this.styles = context.styles;
    this.sharedStrings = context.sharedStrings;
    this.date1904 = context.date1904;
    this.dates = context.dates;
    this.cellOverflow = context.cellOverflow;
    this.truncationSuffix = context.truncationSuffix;
    this.onCellTruncated = context.onCellTruncated;
    this.autoFilter = options.autoFilter === true;
    this.declaredColumns = Math.max(options.header?.length ?? 0, options.columns?.length ?? 0);

    this.append(`${XML_DECLARATION}<worksheet xmlns="${SPREADSHEET_NS}" xmlns:r="${OFFICE_REL_NS}">`);
    const dimension = declaredDimension(options, this.declaredColumns);
    if (dimension !== null) {
      this.append(`<dimension ref="${dimension}"/>`);
    }
    this.append(sheetViewsXml(options.freeze, context.tabSelected === true));
    this.append('<sheetFormatPr defaultRowHeight="15"/>');
    this.append(colsXml(options.columns));
    this.append('<sheetData>');
  }

  /** 1-based row number the next `writeRow` produces. */
  get nextRow(): number {
    return this.rowsWritten + 1;
  }

  writeRow(values: readonly CellInput[], styles?: StyleId | readonly (StyleId | undefined)[]): Promise<void> {
    if (this.failure !== undefined) {
      return Promise.reject(this.failure);
    }
    if (this.closed) {
      return Promise.reject(new XlsxError('WRITER_STATE', 'This sheet is closed; rows can no longer be written to it.'));
    }
    try {
      this.appendRow(values, styles);
    } catch (reason) {
      this.failure = reason;
      return Promise.reject(reason);
    }
    return this.pendingChars >= FLUSH_THRESHOLD_CHARS ? this.flush() : NOTHING_TO_FLUSH;
  }

  /** Register a merged range in A1 notation. Emitted at close, in the order the ranges were added. */
  merge(range: string): void {
    if (this.failure !== undefined) {
      throw this.failure;
    }
    if (this.closed) {
      throw new XlsxError('WRITER_STATE', 'This sheet is closed; merged ranges can no longer be added to it.', { range });
    }
    const parsed = parseRange(range);
    if (parsed === null) {
      throw new XlsxError('WRITER_STATE', `"${range}" is not a range like A1:C1.`, { range });
    }
    // A 1x1 merge is one of the things Excel opens the repair dialog for (primer section 12.8).
    if (parsed.start.row === parsed.end.row && parsed.start.col === parsed.end.col) {
      throw new XlsxError('WRITER_STATE', `"${range}" merges a single cell, which Excel rejects. Merge two or more cells.`, { range });
    }
    this.merges.push(`${formatRef(parsed.start.row, parsed.start.col)}:${formatRef(parsed.end.row, parsed.end.col)}`);
  }

  /** Close `</sheetData>`, write autoFilter/mergeCells, flush, close the zip entry. */
  async close(): Promise<WorksheetWriteSummary> {
    if (this.failure !== undefined) {
      throw this.failure;
    }
    if (this.closed) {
      throw new XlsxError('WRITER_STATE', 'This sheet is already closed.');
    }
    this.closed = true;

    this.append('</sheetData>');
    if (this.autoFilter && this.rowsWritten > 0 && this.columnsWritten > 0) {
      this.append(`<autoFilter ref="${sheetRange(this.rowsWritten, this.columnsWritten)}"/>`);
    }
    if (this.merges.length > 0) {
      this.append(`<mergeCells count="${this.merges.length}">`);
      for (const range of this.merges) {
        this.append(`<mergeCell ref="${range}"/>`);
      }
      this.append('</mergeCells>');
    }
    this.append(PAGE_MARGINS);
    this.append('</worksheet>');
    await this.flush();
    await this.entry.close();
    if (this.truncatedCells > 0) {
      // Reported once per sheet, with this sheet's count; the facade turns that into the workbook running total.
      this.onCellTruncated(this.truncatedCells);
    }

    return {
      rows: this.rowsWritten,
      columns: this.columnsWritten,
      truncatedCells: this.truncatedCells,
      dimension: this.rowsWritten === 0 || this.columnsWritten === 0 ? null : sheetRange(this.rowsWritten, this.columnsWritten),
    };
  }

  private appendRow(values: readonly CellInput[], styles: StyleId | readonly (StyleId | undefined)[] | undefined): void {
    const rowNumber = this.rowsWritten + 1;
    if (rowNumber > MAX_ROWS) {
      throw new XlsxError('ROW_OUT_OF_RANGE', `Row ${rowNumber} is past the ${MAX_ROWS} rows a sheet can hold.`, { row: rowNumber });
    }
    if (values.length > MAX_COLUMNS) {
      throw new XlsxError(
        'ROW_OUT_OF_RANGE',
        `A row of ${values.length} values is wider than the ${MAX_COLUMNS} columns a sheet can hold.`,
        {
          columns: values.length,
        },
      );
    }

    const rowNumberText = String(rowNumber);
    const sharedStyle = typeof styles === 'number' ? styles : undefined;
    const styleList = typeof styles === 'object' ? styles : undefined;
    this.rowColumns = 0;
    this.append(`<row r="${rowNumberText}">`);
    for (let i = 0; i < values.length; i++) {
      this.appendCell(columnLetters(i) + rowNumberText, values[i], styleList === undefined ? sharedStyle : styleList[i], i);
    }
    this.append('</row>');

    this.rowsWritten = rowNumber;
    if (this.rowColumns > this.columnsWritten) {
      this.columnsWritten = this.rowColumns;
    }
  }

  private appendCell(ref: string, value: CellInput, styleId: StyleId | undefined, columnIndex: number): void {
    const kind = typeof value;
    if (kind === 'string') {
      this.appendString(ref, value as string, styleId, columnIndex);
      return;
    }
    if (kind === 'number') {
      const numberValue = value as number;
      const style = styleAttribute(styleId);
      // NaN and the infinities have no xsd:double lexical form Excel accepts; an error cell is the honest stand-in.
      this.emit(
        Number.isFinite(numberValue)
          ? `<c r="${ref}"${style}><v>${numberText(numberValue)}</v></c>`
          : `<c r="${ref}"${style} t="e"><v>#NUM!</v></c>`,
        columnIndex,
      );
      return;
    }
    if (value === undefined || value === null) {
      // A blank cell only earns a `<c>` when it carries a style; otherwise it costs bytes and says nothing.
      if (styleId !== undefined && styleId !== 0) {
        this.emit(`<c r="${ref}" s="${styleId}"/>`, columnIndex);
      }
      return;
    }
    if (kind === 'boolean') {
      this.emit(`<c r="${ref}"${styleAttribute(styleId)} t="b"><v>${value === true ? 1 : 0}</v></c>`, columnIndex);
      return;
    }
    if (kind === 'bigint') {
      const bigintValue = value as bigint;
      if (bigintValue <= MAX_EXACT_BIGINT && bigintValue >= -MAX_EXACT_BIGINT) {
        this.emit(`<c r="${ref}"${styleAttribute(styleId)}><v>${numberText(Number(bigintValue))}</v></c>`, columnIndex);
      } else {
        this.appendString(ref, bigintValue.toString(), styleId, columnIndex);
      }
      return;
    }
    if (value instanceof Date) {
      this.appendDate(ref, value, styleId, columnIndex);
      return;
    }
    const errorCode = (value as CellError).error;
    if (typeof errorCode === 'string') {
      this.emit(`<c r="${ref}"${styleAttribute(styleId)} t="e"><v>${encodeCellText(errorCode)}</v></c>`, columnIndex);
      return;
    }
    throw new XlsxError('WRITER_STATE', `A cell value of type ${kind} cannot be written to a spreadsheet.`, { ref });
  }

  private appendDate(ref: string, value: Date, styleId: StyleId | undefined, columnIndex: number): void {
    const components = componentsFromDate(value, this.dates);
    if (components === null) {
      // An Invalid Date carries no wall clock at all; writing anything would be inventing one.
      return;
    }
    const serial = serialFromComponents(components, this.date1904);
    if (serial === null) {
      // Pre-epoch instants are not Excel dates (EC-DATE-PRE-1900); ISO text keeps the value readable.
      this.appendString(ref, isoText(components), styleId, columnIndex);
      return;
    }
    // A serial under a general-purpose style would show as a bare number, so a date always gets a date format:
    // the caller's style when it already has one, the workbook's default date style otherwise.
    const dateStyle = styleId !== undefined && this.styles.isDateStyle(styleId) ? styleId : this.styles.defaultDateStyle;
    this.emit(`<c r="${ref}" s="${dateStyle}"><v>${numberText(serial)}</v></c>`, columnIndex);
  }

  private appendString(ref: string, text: string, styleId: StyleId | undefined, columnIndex: number): void {
    const limited = this.limitLength(text);
    const index = this.sharedStrings === null ? -1 : this.sharedStrings.intern(limited);
    const style = styleAttribute(styleId);
    if (index >= 0) {
      this.emit(`<c r="${ref}"${style} t="s"><v>${index}</v></c>`, columnIndex);
      return;
    }
    const space = needsSpacePreserve(limited) ? ' xml:space="preserve"' : '';
    this.emit(`<c r="${ref}"${style} t="inlineStr"><is><t${space}>${encodeCellText(limited)}</t></is></c>`, columnIndex);
  }

  /** Apply the over-limit policy (EC-CELL-32767-LIMIT); truncation is counted and reported in the summary. */
  private limitLength(text: string): string {
    if (text.length <= MAX_CELL_CHARS) {
      return text;
    }
    if (this.cellOverflow === 'throw') {
      throw new XlsxError(
        'CELL_TOO_LONG',
        `A cell holds ${text.length} characters, more than the ${MAX_CELL_CHARS} Excel allows. Shorten it or write with cellOverflow: 'truncate'.`,
        { length: text.length, limit: MAX_CELL_CHARS },
      );
    }
    this.truncatedCells++;
    const keep = Math.max(0, MAX_CELL_CHARS - this.truncationSuffix.length);
    // The slice guards a suffix longer than the limit itself; the visible marker is what matters, not the tail.
    return (text.slice(0, keep) + this.truncationSuffix).slice(0, MAX_CELL_CHARS);
  }

  private emit(fragment: string, columnIndex: number): void {
    this.append(fragment);
    this.rowColumns = columnIndex + 1;
  }

  private append(fragment: string): void {
    if (fragment === '') {
      return;
    }
    this.pending.push(fragment);
    this.pendingChars += fragment.length;
  }

  private async flush(): Promise<void> {
    if (this.pendingChars === 0) {
      return;
    }
    const text = this.pending.join('');
    this.pending.length = 0;
    this.pendingChars = 0;
    try {
      await this.entry.write(encodeXmlChunk(text));
    } catch (reason) {
      this.failure = reason;
      throw reason;
    }
  }
}

function styleAttribute(styleId: StyleId | undefined): string {
  // xf 0 is the default a cell without `s` already gets, so naming it only makes the file bigger.
  return styleId === undefined || styleId === 0 ? '' : ` s="${styleId}"`;
}

/**
 * `String(n)` is the shortest round-trip form and a valid xsd:double, except that JS spells the exponent with a
 * lowercase `e` where Excel writes `E` (primer section 5). `-0` stringifies to `0`, which is what Excel stores.
 */
function numberText(value: number): string {
  const text = String(value);
  const exponent = text.indexOf('e');
  return exponent === -1 ? text : `${text.slice(0, exponent)}E${text.slice(exponent + 1)}`;
}

function pad(value: number, width: number): string {
  return String(value).padStart(width, '0');
}

function isoText(components: DateComponents): string {
  const { year, month, day, hour, minute, second } = components;
  return `${pad(year, 4)}-${pad(month, 2)}-${pad(day, 2)}T${pad(hour, 2)}:${pad(minute, 2)}:${pad(second, 2)}`;
}

/** `<dimension>` is only written when both extents are known up front; readers derive the range otherwise. */
function declaredDimension(options: SheetOptions, declaredColumns: number): string | null {
  if (options.rowCount === undefined || declaredColumns === 0) {
    return null;
  }
  const rows = options.rowCount + (options.header === undefined ? 0 : 1);
  return rows < 1 ? null : sheetRange(rows, declaredColumns);
}

function sheetViewsXml(freeze: SheetOptions['freeze'], tabSelected: boolean): string {
  const frozenRows = freeze?.rows ?? 0;
  const frozenColumns = freeze?.cols ?? 0;
  let pane = '';
  if (frozenRows > 0 || frozenColumns > 0) {
    // `topLeftCell` is the first cell that still scrolls, and the active pane is the one it lives in.
    const topLeftCell = formatRef(frozenRows, frozenColumns);
    const activePane = frozenRows > 0 ? (frozenColumns > 0 ? 'bottomRight' : 'bottomLeft') : 'topRight';
    const xSplit = frozenColumns > 0 ? ` xSplit="${frozenColumns}"` : '';
    const ySplit = frozenRows > 0 ? ` ySplit="${frozenRows}"` : '';
    pane =
      `<pane${xSplit}${ySplit} topLeftCell="${topLeftCell}" activePane="${activePane}" state="frozen"/>` +
      `<selection pane="${activePane}" activeCell="${topLeftCell}" sqref="${topLeftCell}"/>`;
  }
  const selected = tabSelected ? ' tabSelected="1"' : '';
  const view = pane === '' ? `<sheetView workbookViewId="0"${selected}/>` : `<sheetView workbookViewId="0"${selected}>${pane}</sheetView>`;
  return `<sheetViews>${view}</sheetViews>`;
}

/** One `<col>` per entry that actually sets something; Excel only honours `width` alongside `customWidth="1"`. */
function colsXml(columns: SheetOptions['columns']): string {
  if (columns === undefined || columns.length === 0) {
    return '';
  }
  if (columns.length > MAX_COLUMNS) {
    throw new XlsxError('ROW_OUT_OF_RANGE', `${columns.length} column definitions exceed the ${MAX_COLUMNS} columns a sheet can hold.`, {
      columns: columns.length,
    });
  }
  const parts: string[] = [];
  for (const [index, column] of columns.entries()) {
    if (column.width === undefined && column.hidden !== true && column.style === undefined) {
      continue;
    }
    const position = index + 1;
    const width = column.width === undefined ? '' : ` width="${column.width}" customWidth="1"`;
    const hidden = column.hidden === true ? ' hidden="1"' : '';
    const style = column.style === undefined || column.style === 0 ? '' : ` style="${column.style}"`;
    parts.push(`<col min="${position}" max="${position}"${width}${hidden}${style}/>`);
  }
  return parts.length === 0 ? '' : `<cols>${parts.join('')}</cols>`;
}
