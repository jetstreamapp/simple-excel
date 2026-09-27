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
/** Excel's widest column, in character units; `<col width>` above it fails the schema (EC-COLS-WIDTH-OVER-255). */
const MAX_COLUMN_WIDTH = 255;

const SURROGATE_HIGH_FIRST = 0xd800;
const SURROGATE_HIGH_LAST = 0xdbff;

/** The error literals `t="e"` may carry; anything else is a repair-dialog error (primer section 12.6). */
const CELL_ERROR_CODES: ReadonlySet<string> = new Set([
  '#NULL!',
  '#DIV/0!',
  '#VALUE!',
  '#REF!',
  '#NAME?',
  '#NUM!',
  '#N/A',
  '#GETTING_DATA',
]);

/** Excel's own defaults, in inches. All six attributes are required by the schema. */
const PAGE_MARGINS = '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>';

/** Rows per bucket of the merge overlap index. */
const MERGE_ROW_BLOCK_SHIFT = 4;
/** A merge spanning more row blocks than this is kept on a list every overlap check scans, rather than indexed. */
const MERGE_MAX_INDEXED_BLOCKS = 8;

/**
 * Handed back by `writeRow` when nothing had to be flushed, so a row costs no promise allocation. Exported so the
 * facade can tell "buffered" apart from a flush or a refusal without attaching a handler to every row.
 */
export const NOTHING_TO_FLUSH: Promise<void> = Promise.resolve();

export interface WorksheetWriterContext {
  readonly entry: ZipEntryWriter;
  readonly styles: StyleRegistry;
  /** Null when strings are written inline only. */
  readonly sharedStrings: SharedStringWriter | null;
  readonly date1904: boolean;
  readonly dates: 'local' | 'utc';
  readonly cellOverflow: 'truncate' | 'throw';
  readonly truncationSuffix: string;
  /** The sheet's name as written to the workbook; only used to make error messages point at the right sheet. */
  readonly sheetName: string;
  /** Excel marks exactly one sheet as the selected tab; the facade passes true for the first visible sheet. */
  readonly tabSelected?: boolean;
  /**
   * The sheet's merged ranges. The facade creates them up front so `merge()` validates synchronously even before
   * the zip entry has opened; a worksheet writer used on its own creates its own.
   */
  readonly merges?: MergedRanges;
}

export interface WorksheetWriteSummary {
  readonly rows: number;
  readonly columns: number;
  readonly truncatedCells: number;
  /** A1 range of all written cells (`A1:T31`), or null when the sheet is empty. */
  readonly dimension: string | null;
}

/**
 * The A1 range covering `rows` rows and `columns` columns from the top-left cell. The sheet summary and the
 * autofilter (header row through the last written row) are both this shape.
 */
export function sheetRange(rows: number, columns: number): string {
  return `A1:${columnLetters(columns - 1)}${rows}`;
}

/** A value's type as an error message names it: `array` and `null` are called out, class instances by class. */
function describeType(value: unknown): string {
  if (Array.isArray(value)) {
    return 'array';
  }
  if (value === null) {
    return 'null';
  }
  if (typeof value === 'object') {
    const className = (value as { constructor?: { name?: unknown } }).constructor?.name;
    return typeof className === 'string' && className !== 'Object' && className !== '' ? `object (${className})` : 'object';
  }
  return typeof value;
}

/** True when `styleId` is an integer in `[0, styleCount)`: one conversion and two comparisons. */
function isStyleIdInRange(styleId: StyleId, styleCount: number): boolean {
  return styleId >>> 0 === styleId && styleId < styleCount;
}

function styleIdError(where: string, styleId: unknown, styleCount: number): XlsxError {
  return new XlsxError(
    'WRITER_STATE',
    `${where}: ${typeof styleId === 'number' ? String(styleId) : `a ${describeType(styleId)}`} is not a registered style id ` +
      `(this workbook has ids 0 to ${styleCount - 1}). Use an id registerStyle returned.`,
    { styleId, styleCount },
  );
}

/** Throws `WRITER_STATE` unless `styleId` is a style the registry has handed out (EC-STYLE-ID-RANGE). */
function assertStyleId(styleId: unknown, styleCount: number, where: string): void {
  if (typeof styleId !== 'number' || !isStyleIdInRange(styleId, styleCount)) {
    throw styleIdError(where, styleId, styleCount);
  }
}

/**
 * Check a sheet's options before anything is queued, so a bad option fails `addSheet` itself rather than a later
 * `writeRow`: `rowCount` a non-negative safe integer, `header` an array, `headerStyle` and every column style a
 * registered id (EC-STYLE-ID-RANGE), every column width a finite non-negative number (EC-COLS-WIDTH-OVER-255).
 */
export function validateSheetOptions(options: SheetOptions, styleCount: number): void {
  if (typeof options !== 'object' || options === null) {
    throw new XlsxError('WRITER_STATE', 'Sheet options must be an object.', { options });
  }
  const { rowCount, header, headerStyle, columns } = options;
  if (rowCount !== undefined && !(Number.isSafeInteger(rowCount) && rowCount >= 0)) {
    throw new XlsxError('WRITER_STATE', `rowCount must be a whole number of rows, 0 or more; got ${String(rowCount)}.`, { rowCount });
  }
  if (header !== undefined && !Array.isArray(header)) {
    throw new XlsxError('WRITER_STATE', `header must be an array of cell values; got a ${describeType(header)}.`);
  }
  if (headerStyle !== undefined && headerStyle !== false) {
    assertStyleId(headerStyle, styleCount, 'headerStyle');
  }
  if (options.freeze !== undefined) {
    validateFreeze(options.freeze);
  }
  if (columns === undefined) {
    return;
  }
  if (!Array.isArray(columns)) {
    throw new XlsxError('WRITER_STATE', `columns must be an array of column options; got a ${describeType(columns)}.`);
  }
  if (columns.length > MAX_COLUMNS) {
    throw new XlsxError('ROW_OUT_OF_RANGE', `${columns.length} column definitions exceed the ${MAX_COLUMNS} columns a sheet can hold.`, {
      columns: columns.length,
    });
  }
  for (const [index, column] of columns.entries()) {
    if (typeof column !== 'object' || column === null) {
      throw new XlsxError('WRITER_STATE', `columns[${index}] must be an object; got a ${describeType(column)}.`);
    }
    const { width, style } = column;
    if (width !== undefined && !(typeof width === 'number' && Number.isFinite(width) && width >= 0)) {
      throw new XlsxError(
        'WRITER_STATE',
        `columns[${index}].width must be a finite number of characters, 0 or more; got ${String(width)}.`,
        { column: index, width },
      );
    }
    if (style !== undefined) {
      assertStyleId(style, styleCount, `columns[${index}].style`);
    }
  }
}

/** Frozen rows and columns must leave at least one scrolling cell inside the grid (`topLeftCell` is written as A1). */
function validateFreeze(freeze: NonNullable<SheetOptions['freeze']>): void {
  if (typeof freeze !== 'object' || freeze === null) {
    throw new XlsxError('WRITER_STATE', `freeze must be an object like { rows: 1 }; got a ${describeType(freeze)}.`);
  }
  const limits = [
    ['rows', freeze.rows, MAX_ROWS],
    ['cols', freeze.cols, MAX_COLUMNS],
  ] as const;
  for (const [field, value, limit] of limits) {
    if (value !== undefined && !(Number.isSafeInteger(value) && value >= 0 && value < limit)) {
      throw new XlsxError('WRITER_STATE', `freeze.${field} must be a whole number from 0 to ${limit - 1}; got ${String(value)}.`, {
        field: `freeze.${field}`,
        value,
      });
    }
  }
}

/**
 * The merged ranges of one sheet, validated as they are added (EC-MERGE-OVERLAP): A1 shape (either case), at least
 * two cells, inside the grid, and not overlapping or repeating a range already registered - each of which Excel
 * answers with the repair dialog (primer section 12.8). Overlap checks go through a row-block index, so a sheet with
 * thousands of merges does not pay a quadratic scan.
 */
export class MergedRanges {
  private readonly sheetName: string;
  /** Normalized `A1:C3` text, in the order the ranges were added. */
  private readonly refs: string[] = [];
  private readonly tops: number[] = [];
  private readonly lefts: number[] = [];
  private readonly bottoms: number[] = [];
  private readonly rights: number[] = [];
  /** Indexes of the merges touching each 16-row block. */
  private readonly byRowBlock = new Map<number, number[]>();
  /** Merges too tall to index by block; every check scans them. */
  private readonly tall: number[] = [];

  constructor(sheetName: string) {
    this.sheetName = sheetName;
  }

  /** The ranges in the order they were added, normalized to `A1:C3`. */
  get ranges(): readonly string[] {
    return this.refs;
  }

  /** Validate and register one range; throws `WRITER_STATE` without registering anything when it is refused. */
  add(range: string): void {
    const parsed = typeof range === 'string' ? parseRange(range.toUpperCase()) : null;
    if (parsed === null) {
      throw new XlsxError(
        'WRITER_STATE',
        `Sheet "${this.sheetName}": ${typeof range === 'string' ? `"${range}"` : `a ${describeType(range)}`} is not a range like A1:C1 inside A1:XFD1048576.`,
        { range },
      );
    }
    const { start, end } = parsed;
    const normalized = `${formatRef(start.row, start.col)}:${formatRef(end.row, end.col)}`;
    if (start.row === end.row && start.col === end.col) {
      throw new XlsxError(
        'WRITER_STATE',
        `Sheet "${this.sheetName}": "${range}" merges a single cell, which Excel rejects. Merge two or more cells.`,
        { range },
      );
    }
    const conflict = this.findOverlap(start.row, start.col, end.row, end.col);
    if (conflict >= 0) {
      const existing = this.refs[conflict] ?? '';
      throw new XlsxError(
        'WRITER_STATE',
        existing === normalized
          ? `Sheet "${this.sheetName}": "${range}" is already merged. Register each merged range once.`
          : `Sheet "${this.sheetName}": "${range}" overlaps the merged range "${existing}", which Excel rejects. Merged ranges must not overlap.`,
        { range, overlaps: existing },
      );
    }

    const index = this.refs.length;
    this.refs.push(normalized);
    this.tops.push(start.row);
    this.lefts.push(start.col);
    this.bottoms.push(end.row);
    this.rights.push(end.col);
    const firstBlock = start.row >> MERGE_ROW_BLOCK_SHIFT;
    const lastBlock = end.row >> MERGE_ROW_BLOCK_SHIFT;
    if (lastBlock - firstBlock >= MERGE_MAX_INDEXED_BLOCKS) {
      this.tall.push(index);
      return;
    }
    for (let block = firstBlock; block <= lastBlock; block++) {
      const bucket = this.byRowBlock.get(block);
      if (bucket === undefined) {
        this.byRowBlock.set(block, [index]);
      } else {
        bucket.push(index);
      }
    }
  }

  /** Index of a registered merge sharing at least one cell with the rectangle, or -1. */
  private findOverlap(top: number, left: number, bottom: number, right: number): number {
    for (const index of this.tall) {
      if (this.overlaps(index, top, left, bottom, right)) {
        return index;
      }
    }
    const firstBlock = top >> MERGE_ROW_BLOCK_SHIFT;
    const lastBlock = bottom >> MERGE_ROW_BLOCK_SHIFT;
    if (lastBlock - firstBlock >= MERGE_MAX_INDEXED_BLOCKS) {
      for (let index = 0; index < this.refs.length; index++) {
        if (this.overlaps(index, top, left, bottom, right)) {
          return index;
        }
      }
      return -1;
    }
    for (let block = firstBlock; block <= lastBlock; block++) {
      for (const index of this.byRowBlock.get(block) ?? []) {
        if (this.overlaps(index, top, left, bottom, right)) {
          return index;
        }
      }
    }
    return -1;
  }

  private overlaps(index: number, top: number, left: number, bottom: number, right: number): boolean {
    return (
      (this.tops[index] ?? 0) <= bottom &&
      (this.bottoms[index] ?? 0) >= top &&
      (this.lefts[index] ?? 0) <= right &&
      (this.rights[index] ?? 0) >= left
    );
  }
}

/**
 * Row-at-a-time worksheet XML in the fixed schema order (sheetPr? dimension? sheetViews? sheetFormatPr? cols?
 * sheetData autoFilter? mergeCells?). Cells append to a string builder that is encoded and pushed to the zip entry
 * every 64 Ki characters; `writeRow` therefore only awaits at flush boundaries. Rows and columns are monotonic.
 *
 * No `<dimension>` is written (EC-DIMENSION-FROM-HINT): it precedes `sheetData`, so it could only come from a hint,
 * and readers that trust a wrong one (SheetJS, openpyxl in read-only mode) silently drop rows and columns outside it,
 * while a missing one is tolerated everywhere (primer section 12.15).
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
  private readonly sheetName: string;
  private readonly autoFilter: boolean;
  private readonly merges: MergedRanges;

  private readonly pending: string[] = [];
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
    this.sheetName = context.sheetName;
    this.autoFilter = options.autoFilter === true;
    this.merges = context.merges ?? new MergedRanges(context.sheetName);

    this.append(`${XML_DECLARATION}<worksheet xmlns="${SPREADSHEET_NS}" xmlns:r="${OFFICE_REL_NS}">`);
    this.append(sheetViewsXml(options.freeze, context.tabSelected === true));
    this.append('<sheetFormatPr defaultRowHeight="15"/>');
    this.append(colsXml(options.columns));
    this.append('<sheetData>');
  }

  /** 1-based row number the next `writeRow` produces. */
  get nextRow(): number {
    return this.rowsWritten + 1;
  }

  /**
   * Append one row. Resolves to `NOTHING_TO_FLUSH` while the row only went into the buffer; a refused row (bad value,
   * unregistered style, over-long text under `cellOverflow: 'throw'`) rejects and latches the failure.
   */
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

  /** Register a merged range in A1 notation; validated now (EC-MERGE-OVERLAP), emitted at close in order. */
  merge(range: string): void {
    if (this.failure !== undefined) {
      throw this.failure;
    }
    if (this.closed) {
      throw new XlsxError('WRITER_STATE', 'This sheet is closed; merged ranges can no longer be added to it.', { range });
    }
    this.merges.add(range);
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
    const merges = this.merges.ranges;
    if (merges.length > 0) {
      this.append(`<mergeCells count="${merges.length}">`);
      for (const range of merges) {
        this.append(`<mergeCell ref="${range}"/>`);
      }
      this.append('</mergeCells>');
    }
    this.append(PAGE_MARGINS);
    this.append('</worksheet>');
    await this.flush();
    await this.entry.close();

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
    // Snapshotted per row: a style registered mid-row (a derived date style) is never one the caller passed.
    const styleCount = this.styles.count;
    let sharedStyle: StyleId | undefined;
    let styleList: readonly (StyleId | undefined)[] | undefined;
    if (typeof styles === 'number') {
      if (!isStyleIdInRange(styles, styleCount)) {
        throw styleIdError(`Sheet "${this.sheetName}" row ${rowNumberText}`, styles, styleCount);
      }
      sharedStyle = styles;
    } else if (Array.isArray(styles)) {
      styleList = styles;
    } else if (styles !== undefined) {
      throw new XlsxError(
        'WRITER_STATE',
        `Sheet "${this.sheetName}" row ${rowNumberText}: styles must be a style id or an array of them; got a ${describeType(styles)}.`,
      );
    }

    this.rowColumns = 0;
    this.append(`<row r="${rowNumberText}">`);
    for (let i = 0; i < values.length; i++) {
      const ref = columnLetters(i) + rowNumberText;
      let styleId = sharedStyle;
      if (styleList !== undefined) {
        styleId = styleList[i];
        if (styleId !== undefined && !isStyleIdInRange(styleId, styleCount)) {
          throw styleIdError(`Sheet "${this.sheetName}" cell ${ref}`, styleId, styleCount);
        }
      }
      this.appendCell(ref, values[i], styleId, i);
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
    const errorCode = kind === 'object' && !Array.isArray(value) ? (value as CellError).error : undefined;
    if (typeof errorCode === 'string') {
      if (CELL_ERROR_CODES.has(errorCode)) {
        this.emit(`<c r="${ref}"${styleAttribute(styleId)} t="e"><v>${errorCode}</v></c>`, columnIndex);
        return;
      }
      // `t="e"` with anything but a literal from the standard's list is a repair-dialog error (primer section 12.6,
      // EC-ERROR-CODE-UNKNOWN). A newer Excel error such as `#SPILL!` or `#CALC!`, which the reader can hand back from
      // a file, is written as its text: it displays the same and a read-then-write round trip keeps working.
      if (errorCode.startsWith('#')) {
        this.appendString(ref, errorCode, styleId, columnIndex);
        return;
      }
      throw new XlsxError(
        'WRITER_STATE',
        `Sheet "${this.sheetName}" cell ${ref}: "${errorCode}" is not an Excel error value. Use one of ${[...CELL_ERROR_CODES].join(' ')}.`,
        { ref, error: errorCode },
      );
    }
    // A Date from another realm (an iframe, a vm context, an Electron bridge) fails `instanceof`; this branch is only
    // reached on the way to an error, so the tag check costs nothing on the normal path.
    if (Object.prototype.toString.call(value) === '[object Date]') {
      this.appendDate(ref, value as unknown as Date, styleId, columnIndex);
      return;
    }
    throw new XlsxError(
      'WRITER_STATE',
      `Sheet "${this.sheetName}" cell ${ref}: a value of type ${describeType(value)} cannot be written. ` +
        'Convert it to a string, number, boolean, Date or null first.',
      { ref, type: describeType(value) },
    );
  }

  private appendDate(ref: string, value: Date, styleId: StyleId | undefined, columnIndex: number): void {
    const components = componentsFromDate(value, this.dates);
    if (components === null) {
      // An Invalid Date carries no wall clock at all; writing anything would be inventing one.
      return;
    }
    const serial = serialFromComponents(components, this.date1904);
    if (serial === null) {
      // Pre-epoch and post-9999 instants are not Excel dates (EC-DATE-PRE-1900, EC-DATE-SERIAL-OVER-9999); ISO text
      // keeps the value readable.
      this.appendString(ref, isoText(components), styleId, columnIndex);
      return;
    }
    // A serial under a general-purpose style would show as a bare number, so a date always gets a date format: the
    // caller's style when it already has one, otherwise the caller's style with the default date format merged in
    // (EC-DATE-STYLE-MERGE), so bold, fills, borders and alignment survive on date cells.
    const dateStyle = styleId === undefined ? this.styles.defaultDateStyle : this.styles.dateStyleFor(styleId);
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
    const suffix = this.truncationSuffix;
    if (suffix.length >= MAX_CELL_CHARS) {
      // A suffix at least as long as the limit is all that fits; the visible marker is what matters, not the tail.
      return cutBeforeLimit(suffix, MAX_CELL_CHARS);
    }
    return cutBeforeLimit(text, MAX_CELL_CHARS - suffix.length) + suffix;
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
      // A row refused while this flush was in flight is the first failure; the aborted write only follows from it.
      if (this.failure === undefined) {
        this.failure = reason;
      }
      throw this.failure;
    }
  }
}

/**
 * The first `limit` UTF-16 units of `text`, one fewer when the last one kept would be the high half of a surrogate
 * pair (EC-TRUNCATION-SURROGATE): a split pair becomes U+FFFD in the file, right before the truncation marker.
 */
function cutBeforeLimit(text: string, limit: number): string {
  const lastKept = text.charCodeAt(limit - 1);
  return text.slice(0, lastKept >= SURROGATE_HIGH_FIRST && lastKept <= SURROGATE_HIGH_LAST ? limit - 1 : limit);
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

/**
 * ISO 8601 text for a date Excel cannot hold as a serial. Years 0-9999 keep the four-digit form; a negative year gets
 * a leading minus over four digits (`-0050`) and a year past 9999 is written out in full (`10000`), the ISO 8601
 * expanded forms.
 */
function isoText(components: DateComponents): string {
  const { year, month, day, hour, minute, second } = components;
  const yearText = year < 0 ? `-${pad(-year, 4)}` : pad(year, 4);
  return `${yearText}-${pad(month, 2)}-${pad(day, 2)}T${pad(hour, 2)}:${pad(minute, 2)}:${pad(second, 2)}`;
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

/**
 * One `<col>` per entry that actually sets something; Excel only honours `width` alongside `customWidth="1"`. Widths
 * past Excel's 255 are clamped (EC-COLS-WIDTH-OVER-255); the options were validated by `validateSheetOptions`.
 */
function colsXml(columns: SheetOptions['columns']): string {
  if (columns === undefined || columns.length === 0) {
    return '';
  }
  const parts: string[] = [];
  for (const [index, column] of columns.entries()) {
    if (column.width === undefined && column.hidden !== true && column.style === undefined) {
      continue;
    }
    const position = index + 1;
    const width = column.width === undefined ? '' : ` width="${Math.min(column.width, MAX_COLUMN_WIDTH)}" customWidth="1"`;
    const hidden = column.hidden === true ? ' hidden="1"' : '';
    const style = column.style === undefined || column.style === 0 ? '' : ` style="${column.style}"`;
    parts.push(`<col min="${position}" max="${position}"${width}${hidden}${style}/>`);
  }
  return parts.length === 0 ? '' : `<cols>${parts.join('')}</cols>`;
}
