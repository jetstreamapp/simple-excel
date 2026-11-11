/**
 * The reader facade: sniff the bytes, open the container, resolve the package parts, and hand out sheet readers
 * that stream rows on demand. Nothing here parses sheet XML - `sml/worksheet-reader` does - and nothing is read
 * eagerly beyond the workbook, its relationships and the styles part (which is streamed).
 */
import { classifyError, isXlsxError, XlsxError } from '../errors';
import { isStringLengthError } from '../internal/string-length';
import { columnLetters, formatRef, MAX_COLUMNS, parseRange } from '../sml/cell-ref';
import {
  DEFAULT_MAX_PACKAGE_ENTRIES,
  parseContentTypes,
  parseRels,
  parseWorkbook,
  relTypeIs,
  type ContentTypes,
  type Relationship,
} from '../sml/package-parts';
import { parseSharedStrings } from '../sml/shared-strings-reader';
import { formatGeneralNumber } from '../sml/numfmt';
import { parseStyles, type ParsedStyles } from '../sml/styles';
import { readWorksheetHead, readWorksheetRows, type WorksheetReadContext, type WorksheetRowsOptions } from '../sml/worksheet-reader';
import type {
  AnyModeRowsOptions,
  ArrayRowsOptions,
  CellValue,
  ObjectModeRowsOptions,
  ObjectRowsOptions,
  ObjectsResult,
  OpenOptions,
  RawCell,
  ReadLimits,
  ReadValue,
  ReadWarning,
  ReadWarningCode,
  RowsOptions,
  Sheet,
  SheetInfo,
  SniffResult,
  Workbook,
} from '../types';
import { guardSource, sourceFrom, type SourceInput } from '../zip/source';
import { DEFAULT_MAX_PART_BYTES, ZipReader } from '../zip/zip-reader';
import { SNIFF_BYTES, sniff } from './sniff';

const DEFAULT_MAX_INFLATED_BYTES = 1024 * 1024 * 1024;
const DEFAULT_MAX_SHARED_STRING_CHARS = 256 * 1024 * 1024;
const DEFAULT_MAX_XML_DEPTH = 256;
/** Excel holds 32,767 characters in a cell and Salesforce's longest text field 131,072; 1 Mi leaves room for both. */
const DEFAULT_MAX_TEXT_LENGTH = 1024 * 1024;
const DEFAULT_MAX_ROW_CHARS = 64 * 1024 * 1024;
/**
 * Records x keys `toObjects()` may hold at once. A materialised cell costs 41 to 57 bytes in Node 24 (records with
 * more than a dozen keys are dictionary objects; Chrome's pointer compression roughly halves that), so 100M is 4 to
 * 6 GB at most: it never refuses a legitimate wide load (200k rows x 300 columns is 60M) and stops the pathological
 * one, a 16,384-column header over 65k rows (1.07G cells), a tenth of the way in.
 */
const DEFAULT_MAX_OBJECT_CELLS = 100_000_000;

/** The UI matches on the literal `password-protected`; changing it breaks Jetstream's error copy (ADR-007). */
const ENCRYPTED_MESSAGE =
  'This workbook is password-protected. Remove the password in Excel (File > Info > Protect Workbook) and save it again as .xlsx.';

const ODS_MIMETYPE_PREFIX = 'application/vnd.oasis.opendocument';
const ODS_MIMETYPE_ENTRY = 'mimetype';
const CONTENT_TYPES_PART = '[Content_Types].xml';
/** The package root, whose relationships live in `_rels/.rels`. */
const PACKAGE_ROOT_PART = '';
const DEFAULT_WORKBOOK_PART = 'xl/workbook.xml';

/** Main-document content types, lower-cased: which application wrote a package (Strict shares the xlsx ones). */
const SPREADSHEET_MAIN_TYPES: ReadonlySet<string> = new Set([
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml',
  'application/vnd.ms-excel.sheet.macroenabled.main+xml',
  'application/vnd.ms-excel.template.macroenabled.main+xml',
  'application/vnd.ms-excel.addin.macroenabled.main+xml',
]);
const XLSB_MAIN_TYPE = 'application/vnd.ms-excel.sheet.binary.macroenabled.main';
const WORDPROCESSING_MAIN_TYPES: ReadonlySet<string> = new Set([
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
  'application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml',
  'application/vnd.ms-word.document.macroenabled.main+xml',
  'application/vnd.ms-word.template.macroenabled.main+xml',
]);
const PRESENTATION_MAIN_TYPES: ReadonlySet<string> = new Set([
  'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',
  'application/vnd.openxmlformats-officedocument.presentationml.template.main+xml',
  'application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml',
  'application/vnd.ms-powerpoint.presentation.macroenabled.main+xml',
  'application/vnd.ms-powerpoint.slideshow.macroenabled.main+xml',
]);

/**
 * `<dimension ref="A1:T31"/>`, prefixed or not. It sits in the worksheet prologue, ahead of `<sheetData>`.
 *
 * The attribute run excludes `<` as well as `>`: a raw `<` cannot appear inside a start tag, and excluding it stops a
 * hostile prologue of repeated `<dimension` from making every start position rescan the rest of the chunk, which
 * would be quadratic. Each start now scans only up to the next `<`, so the match is linear in the prologue.
 */
const DIMENSION_PATTERN = /<(?:[\w.-]+:)?dimension[^<>]*\sref="([^"]+)"/;
/** How far into a worksheet part to look for `<dimension>` before giving up on it. */
const DIMENSION_SCAN_CHARS = 64 * 1024;

interface ResolvedLimits {
  readonly maxEntries: number;
  readonly maxInflatedBytes: number;
  readonly maxPartBytes: number;
  readonly maxSharedStringChars: number;
  readonly maxXmlDepth: number;
  readonly maxTextLength: number;
  readonly maxRowChars: number;
  readonly maxObjectCells: number;
}

function resolveLimits(limits: ReadLimits | undefined): ResolvedLimits {
  return {
    maxEntries: limits?.maxEntries ?? DEFAULT_MAX_PACKAGE_ENTRIES,
    maxInflatedBytes: limits?.maxInflatedBytes ?? DEFAULT_MAX_INFLATED_BYTES,
    maxPartBytes: limits?.maxPartBytes ?? DEFAULT_MAX_PART_BYTES,
    maxSharedStringChars: limits?.maxSharedStringChars ?? DEFAULT_MAX_SHARED_STRING_CHARS,
    maxXmlDepth: limits?.maxXmlDepth ?? DEFAULT_MAX_XML_DEPTH,
    maxTextLength: limits?.maxTextLength ?? DEFAULT_MAX_TEXT_LENGTH,
    maxRowChars: limits?.maxRowChars ?? DEFAULT_MAX_ROW_CHARS,
    maxObjectCells: limits?.maxObjectCells ?? DEFAULT_MAX_OBJECT_CELLS,
  };
}

function notXlsx(format: string, message: string): XlsxError {
  return new XlsxError('NOT_XLSX', message, { format });
}

function throwIfAborted(signal: AbortSignal | undefined): void {
  if (signal?.aborted === true) {
    throw new XlsxError('ABORTED', 'Reading this spreadsheet was cancelled.');
  }
}

/** Everything that is not a zip is a dead end; each message names the fix rather than the format we wanted. */
function rejectSniffedFormat(format: SniffResult): never {
  switch (format) {
    case 'cfb-encrypted':
      throw new XlsxError('ENCRYPTED', ENCRYPTED_MESSAGE);
    case 'cfb-legacy':
      throw new XlsxError('LEGACY_XLS', 'This is a legacy Excel 97-2003 (.xls) workbook. Open it in Excel and save it as .xlsx.');
    case 'empty':
      throw notXlsx('empty', 'The file is empty.');
    case 'xml':
      throw notXlsx('xml', 'This is not an .xlsx workbook (it looks like an XML document). Open it in Excel and save it as .xlsx.');
    case 'html':
      throw notXlsx('html', 'This is not an .xlsx workbook (it looks like an HTML page). Save it as .xlsx or import it as CSV.');
    case 'text':
      throw notXlsx('text', 'This is not an .xlsx workbook (it looks like csv/text). Save it as .xlsx or import it as CSV.');
    default:
      throw notXlsx('unknown', 'This is not an .xlsx workbook. Open it in Excel and save it as .xlsx.');
  }
}

/**
 * Parsers throw `XlsxError` for everything they anticipate. A string past the engine's limit means the caller raised
 * a limit further than one string can go (EC-CELL-TEXT-RUN-SPLIT); anything else means the part is damaged.
 */
function damagedPart(partName: string, cause: unknown): XlsxError {
  if (isStringLengthError(cause)) {
    return new XlsxError(
      'LIMIT_EXCEEDED',
      `The part ${partName} holds more text than one string can, so it was not read.`,
      { part: partName },
      { cause },
    );
  }
  return new XlsxError('XML_MALFORMED', `The part ${partName} could not be read; this file may be damaged.`, { part: partName }, { cause });
}

function parsePart<T>(partName: string, parse: () => T): T {
  try {
    return parse();
  } catch (error) {
    throw isXlsxError(error) ? error : damagedPart(partName, error);
  }
}

async function parsePartAsync<T>(partName: string, parse: () => Promise<T>): Promise<T> {
  try {
    return await parse();
  } catch (error) {
    throw isXlsxError(error) ? error : damagedPart(partName, error);
  }
}

/** `xl/workbook.xml` -> `xl/`; a part at the package root -> `''`. */
function directoryOf(partName: string): string {
  const lastSlash = partName.lastIndexOf('/');
  return lastSlash < 0 ? '' : partName.slice(0, lastSlash + 1);
}

/** The relationship part of a part: `xl/workbook.xml` -> `xl/_rels/workbook.xml.rels`, the root -> `_rels/.rels`. */
function relsPartFor(partName: string): string {
  return `${directoryOf(partName)}_rels/${partName.slice(partName.lastIndexOf('/') + 1)}.rels`;
}

async function readRels(zip: ZipReader, partName: string, limits: ResolvedLimits): Promise<Relationship[]> {
  const relsPart = relsPartFor(partName);
  if (!zip.has(relsPart)) {
    return [];
  }
  const xml = await zip.readText(relsPart, limits.maxPartBytes);
  return parsePart(relsPart, () => parseRels(xml, directoryOf(partName), limits.maxEntries));
}

function relationshipTarget(rels: readonly Relationship[], suffix: string): string | undefined {
  return rels.find(rel => !rel.external && rel.target !== '' && relTypeIs(rel.type, suffix))?.target;
}

/**
 * An OpenDocument package announces itself with a `mimetype` entry, which is the only thing about it a
 * SpreadsheetML reader can recognize: it has no `[Content_Types].xml` at all.
 */
async function rejectOpenDocument(zip: ZipReader, limits: ResolvedLimits): Promise<void> {
  if (!zip.has(ODS_MIMETYPE_ENTRY)) {
    return;
  }
  const mimetype = (await zip.readText(ODS_MIMETYPE_ENTRY, limits.maxPartBytes)).trim();
  if (mimetype.startsWith(ODS_MIMETYPE_PREFIX)) {
    throw new XlsxError('ODS', 'This is an OpenDocument (.ods) spreadsheet. Open it in Excel or LibreOffice and save it as .xlsx.', {
      mimetype,
    });
  }
}

function isMainContentType(contentType: string): boolean {
  const normalized = contentType.toLowerCase();
  return (
    SPREADSHEET_MAIN_TYPES.has(normalized) ||
    WORDPROCESSING_MAIN_TYPES.has(normalized) ||
    PRESENTATION_MAIN_TYPES.has(normalized) ||
    normalized === XLSB_MAIN_TYPE
  );
}

/**
 * The package's main part: the `officeDocument` relationship names it, and its content type says which application
 * wrote the package. The content-type overrides are the fallback for a package with no root relationships, and
 * `xl/workbook.xml` the last resort for one with neither.
 */
function locateMainPart(zip: ZipReader, rootRels: readonly Relationship[], contentTypes: ContentTypes | undefined): string | undefined {
  const candidates: string[] = [];
  const officeDocument = relationshipTarget(rootRels, 'officeDocument');
  if (officeDocument !== undefined) {
    candidates.push(officeDocument);
  }
  for (const [partName, contentType] of contentTypes?.overrides ?? []) {
    if (isMainContentType(contentType)) {
      candidates.push(partName);
    }
  }
  candidates.push(DEFAULT_WORKBOOK_PART);
  return candidates.find(candidate => zip.has(candidate));
}

/** Reject the OOXML packages that are not spreadsheets, by the content type of the part the workbook rel points at. */
function rejectForeignMainPart(mainPart: string, contentTypes: ContentTypes | undefined): void {
  const contentType = contentTypes?.typeOf(mainPart)?.toLowerCase();
  if (contentType === undefined) {
    return;
  }
  if (contentType === XLSB_MAIN_TYPE) {
    throw new XlsxError('XLSB', 'This is an Excel Binary Workbook (.xlsb). Open it in Excel and save it as .xlsx.', { part: mainPart });
  }
  if (WORDPROCESSING_MAIN_TYPES.has(contentType)) {
    throw notXlsx('docx', 'This is a Word document (.docx), not a spreadsheet. Put your data in a workbook and save it as .xlsx.');
  }
  if (PRESENTATION_MAIN_TYPES.has(contentType)) {
    throw notXlsx(
      'pptx',
      'This is a PowerPoint presentation (.pptx), not a spreadsheet. Put your data in a workbook and save it as .xlsx.',
    );
  }
}

// ---------------------------------------------------------------------------------------------------------------------
// Object mode: the `sheet_to_json` contract
// ---------------------------------------------------------------------------------------------------------------------

const EMPTY_HEADER_BASE = '__EMPTY';
/** Stands in for "this call names no columns from a header row", so widening a row allocates nothing. */
const NO_HEADER_CELLS: readonly ReadValue[] = [];
/** The one key a plain-object assignment does not store: it runs the `Object.prototype.__proto__` setter instead. */
const PROTOTYPE_KEY = '__proto__';

/**
 * SheetJS names a blank header `__EMPTY`, then `__EMPTY_1`, `__EMPTY_2`, and disambiguates a repeated header with
 * the same `_1`, `_2` suffixes, skipping any suffix already taken. Reproduced exactly: Jetstream's field mapping
 * keys off these names, so they are part of the contract (01 B2). A `Map`, never a plain object: header text is user
 * text, and `__proto__`, `constructor` or `toString` must count like any other name (EC-HEADER-PROTO-KEY).
 */
class HeaderNamer {
  private readonly used = new Map<string, number>();

  unique(base: string): string {
    const seen = this.used.get(base) ?? 0;
    if (seen === 0) {
      this.used.set(base, 1);
      return base;
    }
    let counter = seen;
    let candidate = `${base}_${counter++}`;
    while (this.used.has(candidate)) {
      candidate = `${base}_${counter++}`;
    }
    this.used.set(base, counter);
    this.used.set(candidate, 1);
    return candidate;
  }
}

function padTwo(value: number): string {
  return value < 10 ? `0${value}` : String(value);
}

/**
 * The text SheetJS puts in a header is the cell's formatted text. We have no number formatter, so the value's own
 * spelling stands in: identical for strings (every header in practice) and Excel's spelling for the other types, a
 * number in the General format whatever format the cell has (EC-HEADER-NUMBER-GENERAL). An
 * error cell names its column with the error text whatever `errors` says (EC-HEADER-ERROR-CELL); the row reader
 * hands the header row's errors over as text for that reason, and a `CellError` is read the same way.
 */
function headerTextOf(value: ReadValue, utcDates: boolean): string {
  if (typeof value === 'string') {
    return value;
  }
  if (typeof value === 'boolean') {
    return value ? 'TRUE' : 'FALSE';
  }
  if (value instanceof Date) {
    // The wall clock the cell shows lives in the UTC fields under `dates: 'utc'` and in the local ones otherwise, so
    // the key does not move with the server's time zone.
    if (utcDates) {
      const date = `${value.getUTCFullYear()}-${padTwo(value.getUTCMonth() + 1)}-${padTwo(value.getUTCDate())}`;
      return `${date}T${padTwo(value.getUTCHours())}:${padTwo(value.getUTCMinutes())}:${padTwo(value.getUTCSeconds())}`;
    }
    const date = `${value.getFullYear()}-${padTwo(value.getMonth() + 1)}-${padTwo(value.getDate())}`;
    return `${date}T${padTwo(value.getHours())}:${padTwo(value.getMinutes())}:${padTwo(value.getSeconds())}`;
  }
  if (typeof value === 'object' && value !== null) {
    return value.error;
  }
  // General's spelling, as Excel shows the cell and SheetJS names the key: no float noise (EC-HEADER-NUMBER-GENERAL).
  return typeof value === 'number' ? formatGeneralNumber(value) : String(value);
}

/** The first index below `limit` whose cell holds a value, or `limit` when none does. */
function firstValueIndex(cells: readonly ReadValue[], limit: number): number {
  const end = Math.min(limit, cells.length);
  for (let column = 0; column < end; column++) {
    const value = cells[column];
    if (value !== undefined && value !== null) {
      return column;
    }
  }
  return limit;
}

/**
 * One past the last header cell that holds a value, or 0 when the row holds none. A value is anything the header-row
 * search counts (EC-HEADER-ROW-FIRST-NONEMPTY), an empty string or an error included: a column whose header cell
 * says `''` is named `''`, as in SheetJS.
 */
function headerEndOf(headerCells: readonly ReadValue[]): number {
  for (let column = headerCells.length - 1; column >= 0; column--) {
    const value = headerCells[column];
    if (value !== undefined && value !== null) {
      return column + 1;
    }
  }
  return 0;
}

interface HeaderColumn {
  readonly name: string;
  /** 0-based column the header reads from. */
  readonly column: number;
  /** The name is `__proto__`, which only `Object.defineProperty` can store as an own property. */
  readonly definesOwnProperty: boolean;
  /** How many records were built before this column had a name; `toObjects` backfills those. */
  readonly firstRecord: number;
}

/** Store `value` under `column`'s name as an own, enumerable data property of a plain object. */
function setField(record: Record<string, ReadValue>, column: HeaderColumn, value: ReadValue): void {
  if (column.definesOwnProperty) {
    Object.defineProperty(record, column.name, { value, writable: true, enumerable: true, configurable: true });
    return;
  }
  record[column.name] = value;
}

/** What a `toObjects` call needs back from the row stream it drains. */
interface ObjectReadState {
  headers: string[];
  truncated: boolean;
  /** Set as soon as the read starts, so `toObjects` can count the keys it holds and backfill once the stream ends. */
  builder: ObjectRowBuilder | undefined;
}

/** 1-based header row "the first row with a value", as opposed to a row the caller named. */
const DETECT_HEADER_ROW = 0;
/** No header row at all: the caller named the columns with `headers`. */
const NO_HEADER_ROW = -1;

/** The values a read left out because they sit right of the last header (EC-DATA-BEYOND-HEADER). */
class DataBeyondHeader {
  /** Which columns held one, by 0-based index; allocated on the first such value, 16 KiB at most. */
  private readonly seen = new Uint8Array(MAX_COLUMNS);
  private firstRef = '';
  private firstColumn = MAX_COLUMNS;
  private lastColumn = -1;
  private columnCount = 0;
  private cellCount = 0;

  /** Record the values of `cells` from `headerEnd` on. `rowIndex` is 1-based. */
  note(rowIndex: number, cells: readonly ReadValue[], headerEnd: number): void {
    for (let column = headerEnd; column < cells.length; column++) {
      const value = cells[column];
      if (value === undefined || value === null) {
        continue;
      }
      if (this.cellCount++ === 0) {
        this.firstRef = formatRef(rowIndex - 1, column);
      }
      if (this.seen[column] === 0) {
        this.seen[column] = 1;
        this.columnCount++;
        this.firstColumn = Math.min(this.firstColumn, column);
        this.lastColumn = Math.max(this.lastColumn, column);
      }
    }
  }

  /** The one warning a sheet read gets, or undefined when nothing was left out. */
  warning(headerEnd: number): ReadWarning | undefined {
    if (this.cellCount === 0) {
      return undefined;
    }
    const first = columnLetters(this.firstColumn);
    const last = columnLetters(this.lastColumn);
    const columns = first === last ? first : `${first}:${last}`;
    const lastHeaderColumn = headerEnd > 0 ? columnLetters(headerEnd - 1) : '';
    const header = headerEnd > 0 ? `its last header, in column ${lastHeaderColumn}` : 'its header row, which names no columns';
    return {
      code: 'DATA_BEYOND_HEADER',
      message: `This sheet has ${this.cellCount} value(s) in ${this.columnCount} column(s) (${columns}) to the right of ${header}. Records stop at the last header, so those values were left out.`,
      ref: this.firstRef,
      detail: { columns, columnCount: this.columnCount, cellCount: this.cellCount, lastHeaderColumn },
    };
  }
}

/**
 * Turns rows into `{ header: value }` records: finds and names the header row's columns, skips everything before
 * the first data row, and fills absent cells with `defval`. Blank rows never reach it - the row reader drops them,
 * which is SheetJS's `blankrows: false` - so without an explicit `headerRow` the first row it sees is the header row
 * (EC-HEADER-ROW-FIRST-NONEMPTY).
 *
 * Columns run from the dimension's first column (EC-HEADER-DIMENSION-START-COLUMN) to the last header cell that
 * holds a value, or are exactly the caller's `headers`. That is where SheetJS and this part ways
 * (EC-DATA-BEYOND-HEADER): SheetJS names every column out to the widest row, so one stray value at XFD gives every
 * record 16,384 keys. A value right of the header is left out and counted for one `DATA_BEYOND_HEADER` warning. A
 * value left of the first named column still gets a blank-header name, as in SheetJS.
 */
class ObjectRowBuilder {
  /** The caller's `headerRow`, `DETECT_HEADER_ROW`, or `NO_HEADER_ROW` when the caller named the columns. */
  private readonly headerRow: number;
  private readonly requestedStartRow: number;
  private firstDataRow: number;
  private readonly declaredFirstColumn: number;
  private readonly defval: CellValue;
  private readonly naming: NonNullable<ObjectRowsOptions['headerNaming']>;
  private readonly dropEmptyHeaders: boolean;
  private readonly utcDates: boolean;
  private readonly namer = new HeaderNamer();
  private readonly columns: HeaderColumn[] = [];
  /** Columns `leftmostNamed` to `headerEnd` are named, the ones `dropEmptyHeaders` left out of `columns` too. */
  private leftmostNamed = 0;
  /** One past the last column a record has a key for; a value from here on is beyond the header. */
  private headerEnd = 0;
  private headersReady = false;
  private recordsBuilt = 0;
  private beyondHeader: DataBeyondHeader | undefined;

  constructor(options: RowsOptions & ObjectRowsOptions, declaredFirstColumn: number, dates: WorksheetReadContext['dates']) {
    this.requestedStartRow = Math.max(1, options.startRow ?? 1);
    this.declaredFirstColumn = declaredFirstColumn;
    // `defval: null` is a caller asking for nulls, not for the default.
    this.defval = options.defval === undefined ? '' : options.defval;
    this.naming = options.headerNaming ?? 'sheetjs';
    this.dropEmptyHeaders = options.dropEmptyHeaders === true;
    this.utcDates = dates === 'utc';
    if (options.headers !== undefined) {
      this.headerRow = NO_HEADER_ROW;
      // A named `headerRow` is then only a row to skip: the data starts after it.
      this.firstDataRow = Math.max(this.requestedStartRow, options.headerRow === undefined ? 1 : Math.max(1, options.headerRow) + 1);
      this.useHeaders(options.headers);
      return;
    }
    this.headerRow = options.headerRow === undefined ? DETECT_HEADER_ROW : Math.max(1, options.headerRow);
    this.firstDataRow =
      this.headerRow === DETECT_HEADER_ROW ? Number.POSITIVE_INFINITY : Math.max(options.startRow ?? 0, this.headerRow + 1);
  }

  /** 1-based row the underlying reader must start at: the header row when the caller named one, else `startRow`. */
  get startRow(): number {
    if (this.headerRow === NO_HEADER_ROW) {
      return this.firstDataRow;
    }
    return this.headerRow === DETECT_HEADER_ROW ? this.requestedStartRow : this.headerRow;
  }

  /** Which row the reader should hand over with error cells as text: the header row, found or named, if any. */
  get errorTextRow(): number | 'first' | undefined {
    if (this.headerRow === NO_HEADER_ROW) {
      return undefined;
    }
    return this.headerRow === DETECT_HEADER_ROW ? 'first' : this.headerRow;
  }

  get headers(): string[] {
    return this.columns.map(column => column.name);
  }

  /** Keys each record built from here on has. */
  get columnCount(): number {
    return this.columns.length;
  }

  /** The record for one row, or null when the row is the header row or sits before the first data row. */
  accept(rowIndex: number, cells: readonly ReadValue[]): Record<string, ReadValue> | null {
    if (!this.headersReady) {
      if (this.headerRow === DETECT_HEADER_ROW) {
        this.firstDataRow = rowIndex + 1;
        this.nameHeaderRow(cells);
        return null;
      }
      if (rowIndex >= this.headerRow) {
        this.nameHeaderRow(rowIndex === this.headerRow ? cells : NO_HEADER_CELLS);
      }
    }
    if (rowIndex < this.firstDataRow) {
      return null;
    }
    if (this.leftmostNamed > 0) {
      this.nameColumnsBefore(firstValueIndex(cells, this.leftmostNamed));
    }
    if (cells.length > this.headerEnd) {
      this.beyondHeader ??= new DataBeyondHeader();
      this.beyondHeader.note(rowIndex, cells, this.headerEnd);
    }
    const record: Record<string, ReadValue> = {};
    for (const column of this.columns) {
      const value = cells[column.column];
      const resolved = value === undefined || value === null ? this.defval : value;
      if (column.definesOwnProperty) {
        setField(record, column, resolved);
      } else {
        record[column.name] = resolved;
      }
    }
    this.recordsBuilt++;
    return record;
  }

  /** The `DATA_BEYOND_HEADER` warning for everything this read left out, or undefined when it left nothing out. */
  dataBeyondHeaderWarning(): ReadWarning | undefined {
    return this.beyondHeader?.warning(this.headerEnd);
  }

  /**
   * Give the records built before a column was named that column too, with `defval`, so every record has every
   * header in the same key order as `headers`. Only `toObjects` can: a streamed record is already in the caller's
   * hands. Only the records built before the last late column are rebuilt, which is none on the common path.
   */
  backfill(records: Record<string, ReadValue>[]): void {
    let affected = 0;
    for (const column of this.columns) {
      affected = Math.max(affected, Math.min(column.firstRecord, records.length));
    }
    for (let index = 0; index < affected; index++) {
      const previous = records[index];
      if (previous === undefined) {
        continue;
      }
      const rebuilt: Record<string, ReadValue> = {};
      for (const column of this.columns) {
        setField(rebuilt, column, Object.hasOwn(previous, column.name) ? (previous[column.name] as ReadValue) : this.defval);
      }
      records[index] = rebuilt;
    }
  }

  /** The caller's names, one per column from A: exactly those keys, whatever the sheet holds. */
  private useHeaders(headers: readonly string[]): void {
    this.headersReady = true;
    for (const [column, header] of headers.entries()) {
      const name = String(header);
      this.columns.push({ name, column, definesOwnProperty: name === PROTOTYPE_KEY, firstRecord: 0 });
    }
    this.headerEnd = headers.length;
  }

  /**
   * Name the header row's columns: from the dimension's first column, or from an earlier one when the header row
   * itself has a value there (a dimension that is wrong about where the data starts), to the last header cell that
   * holds a value.
   */
  private nameHeaderRow(headerCells: readonly ReadValue[]): void {
    this.headersReady = true;
    const end = headerEndOf(headerCells);
    const first = Math.min(firstValueIndex(headerCells, this.declaredFirstColumn), end);
    this.leftmostNamed = first;
    for (let column = first; column < end; column++) {
      const named = this.nameColumn(column, headerCells[column]);
      if (named !== undefined) {
        this.columns.push(named);
      }
    }
    this.headerEnd = end;
  }

  /** A data row with a value before the named columns: name the blank columns from that value's up to them. */
  private nameColumnsBefore(first: number): void {
    const end = this.leftmostNamed;
    if (first >= end) {
      return;
    }
    this.leftmostNamed = first;
    // These columns sit left of every column named so far, so they go in front: `headers` and each record's keys
    // stay in sheet order.
    const added: HeaderColumn[] = [];
    for (let column = first; column < end; column++) {
      const named = this.nameColumn(column, undefined);
      if (named !== undefined) {
        added.push(named);
      }
    }
    this.columns.unshift(...added);
  }

  /** The named column, or undefined when `dropEmptyHeaders` leaves it out. The caller places it in `columns`. */
  private nameColumn(column: number, value: ReadValue | undefined): HeaderColumn | undefined {
    const missing = value === undefined || value === null;
    const text = missing ? '' : headerTextOf(value, this.utcDates);
    // A header cell holding an empty string is an empty name in SheetJS; only an absent cell becomes `__EMPTY`.
    const name = this.naming === 'index' ? String(column) : this.namer.unique(missing ? EMPTY_HEADER_BASE : text);
    if (this.dropEmptyHeaders && text === '') {
      return undefined;
    }
    return { name, column, definesOwnProperty: name === PROTOTYPE_KEY, firstRecord: this.recordsBuilt };
  }
}

// ---------------------------------------------------------------------------------------------------------------------
// Sheets
// ---------------------------------------------------------------------------------------------------------------------

interface SheetParts {
  readonly zip: ZipReader;
  readonly context: WorksheetReadContext;
  /** `limits.maxObjectCells`: records x keys `toObjects()` may hold. */
  readonly maxObjectCells: number;
}

function tooManyObjectCells(maxObjectCells: number, rows: number, columns: number): XlsxError {
  return new XlsxError(
    'LIMIT_EXCEEDED',
    `This sheet has more than ${maxObjectCells} cells to hold as records (${rows} rows of ${columns} columns so far), more than toObjects() holds at once. Read it with rows({ mode: 'object' }) to stream the records instead, or raise limits.maxObjectCells.`,
    { maxObjectCells, rows, columns },
  );
}

class SheetReader implements Sheet<ReadValue> {
  readonly info: SheetInfo;

  private readonly parts: SheetParts;
  private readonly partName: string;
  /** The `<dimension>`'s first column (0 when it declares none), resolved once and only when object mode asks. */
  private declaredFirst: number | undefined;

  constructor(info: SheetInfo, partName: string, parts: SheetParts) {
    this.info = info;
    this.partName = partName;
    this.parts = parts;
  }

  rows(options?: ArrayRowsOptions): AsyncIterable<ReadValue[]>;
  rows(options: ObjectModeRowsOptions): AsyncIterable<Record<string, ReadValue>>;
  rows(options: AnyModeRowsOptions): AsyncIterable<ReadValue[] | Record<string, ReadValue>>;
  rows(options: AnyModeRowsOptions = {}): AsyncIterable<ReadValue[] | Record<string, ReadValue>> {
    if (options.mode === 'object') {
      return this.objectRows(options, { headers: [], truncated: false, builder: undefined });
    }
    return this.arrayRows(options);
  }

  /**
   * Every record at once, under `limits.maxObjectCells` (records x keys): past it this throws `LIMIT_EXCEEDED`
   * pointing at the streaming `rows({ mode: 'object' })`, which holds one record at a time.
   */
  async toObjects(options: RowsOptions & ObjectRowsOptions = {}): Promise<ObjectsResult<ReadValue>> {
    const maxObjectCells = this.parts.maxObjectCells;
    const state: ObjectReadState = { headers: [], truncated: false, builder: undefined };
    const rows: Record<string, ReadValue>[] = [];
    let cells = 0;
    for await (const row of this.objectRows(options, state)) {
      const columns = state.builder?.columnCount ?? 0;
      cells += columns;
      if (cells > maxObjectCells) {
        throw tooManyObjectCells(maxObjectCells, rows.length + 1, columns);
      }
      rows.push(row);
    }
    // Everything is in hand here, unlike a stream: a column a later row named reaches the records before it too.
    const columns = state.headers.length;
    if (rows.length * columns > maxObjectCells) {
      throw tooManyObjectCells(maxObjectCells, rows.length, columns);
    }
    state.builder?.backfill(rows);
    return { rows, headers: state.headers, truncated: state.truncated };
  }

  async head(rowCount: number): Promise<Map<string, RawCell>> {
    return readWorksheetHead(this.partStream(), this.parts.context, rowCount);
  }

  private async *arrayRows(options: RowsOptions): AsyncGenerator<ReadValue[]> {
    for await (const row of readWorksheetRows(this.partStream(), this.parts.context, options)) {
      yield row.cells as ReadValue[];
    }
  }

  private async *objectRows(options: RowsOptions & ObjectRowsOptions, state: ObjectReadState): AsyncGenerator<Record<string, ReadValue>> {
    // Explicit headers name columns from A, so the dimension has nothing to add.
    const declaredFirst = options.headers === undefined ? await this.declaredFirstColumn() : 0;
    const builder = new ObjectRowBuilder(options, declaredFirst, this.parts.context.dates);
    state.builder = builder;
    const maxRows = options.maxRows ?? Number.POSITIVE_INFINITY;
    const rowOptions: WorksheetRowsOptions = {
      startRow: builder.startRow,
      maxColumns: options.maxColumns,
      formulas: options.formulas,
      errorTextRow: builder.errorTextRow,
    };
    let emitted = 0;
    let failed = false;
    try {
      for await (const row of readWorksheetRows(this.partStream(), this.parts.context, rowOptions)) {
        const record = builder.accept(row.index, row.cells as ReadValue[]);
        if (record === null) {
          continue;
        }
        // One record past the limit is what proves the sheet had more to give.
        if (emitted >= maxRows) {
          state.truncated = true;
          return;
        }
        emitted++;
        yield record;
      }
    } catch (error) {
      failed = true;
      throw error;
    } finally {
      state.headers = builder.headers;
      // Once per read, when it ends (or the caller stops it): only then are the columns it left out all known.
      const warning = failed ? undefined : builder.dataBeyondHeaderWarning();
      if (warning !== undefined) {
        this.parts.context.onWarning?.(warning);
      }
    }
  }

  private partStream(): AsyncIterable<Uint8Array> {
    if (!this.parts.zip.has(this.partName)) {
      throw new XlsxError('NOT_XLSX', `The sheet "${this.info.name}" is missing from this workbook; this file may be damaged.`, {
        part: this.partName,
        sheet: this.info.name,
      });
    }
    return this.parts.zip.stream(this.partName);
  }

  /**
   * Where the sheet says its columns start. SheetJS builds its header list from `<dimension>` when the sheet declares
   * one, so a file whose column A is empty (Excel writes `B2:D9`) starts at B; reading the prologue is what keeps
   * object mode's keys identical to `sheet_to_json`'s on the left. On the right the keys stop at the header
   * (EC-DATA-BEYOND-HEADER), so the dimension's last column plays no part.
   */
  private async declaredFirstColumn(): Promise<number> {
    if (this.declaredFirst !== undefined) {
      return this.declaredFirst;
    }
    let declaredFirst = 0;
    if (this.parts.zip.has(this.partName)) {
      const decoder = new TextDecoder('utf-8');
      let prologue = '';
      for await (const chunk of this.parts.zip.stream(this.partName)) {
        prologue += decoder.decode(chunk, { stream: true });
        const ref = DIMENSION_PATTERN.exec(prologue)?.[1];
        if (ref !== undefined) {
          const range = parseRange(ref);
          if (range !== null) {
            declaredFirst = range.start.col;
          }
          break;
        }
        // `<dimension>` is optional, but it always precedes the rows: once they start there is nothing left to find.
        if (prologue.includes('<sheetData') || prologue.length > DIMENSION_SCAN_CHARS) {
          break;
        }
      }
    }
    this.declaredFirst = declaredFirst;
    return declaredFirst;
  }
}

class WorkbookReader implements Workbook<ReadValue> {
  readonly sheets: readonly SheetInfo[];
  readonly date1904: boolean;

  private readonly zip: ZipReader;
  private readonly readers: readonly SheetReader[];

  constructor(zip: ZipReader, date1904: boolean, readers: readonly SheetReader[]) {
    this.zip = zip;
    this.date1904 = date1904;
    this.readers = readers;
    this.sheets = readers.map(reader => reader.info);
  }

  sheet(nameOrIndex: string | number): Sheet<ReadValue> {
    const reader = typeof nameOrIndex === 'number' ? this.readers[nameOrIndex] : this.findByName(nameOrIndex);
    if (!reader) {
      const available = this.readers.map(candidate => `"${candidate.info.name}"`).join(', ');
      throw new XlsxError('SHEET_NOT_FOUND', `This workbook has no sheet ${JSON.stringify(nameOrIndex)}. It has ${available}.`, {
        requested: nameOrIndex,
        available: this.readers.map(candidate => candidate.info.name),
      });
    }
    return reader;
  }

  close(): Promise<void> {
    return this.zip.close();
  }

  /** Exact name first, then case-insensitively: Excel itself compares sheet names without case. */
  private findByName(name: string): SheetReader | undefined {
    const exact = this.readers.find(reader => reader.info.name === name);
    if (exact) {
      return exact;
    }
    const lowered = name.toLowerCase();
    return this.readers.find(reader => reader.info.name.toLowerCase() === lowered);
  }
}

// ---------------------------------------------------------------------------------------------------------------------
// Opening
// ---------------------------------------------------------------------------------------------------------------------

/**
 * Open an xlsx from bytes (`ArrayBuffer`, `SharedArrayBuffer`, any typed array or `DataView`, `Blob`/`File`, or a
 * `RandomAccessSource`; anything else is `NOT_XLSX`). Sniffs the format first and throws a classified `XlsxError`
 * (`ENCRYPTED` with a `password-protected` message, `LEGACY_XLS`, `XLSB`, `ODS`, `NOT_XLSX`) for anything that is
 * not a modern workbook. The workbook lists sheets without touching sheet XML; sheet contents stream on demand, until
 * `close()`, after which every read fails with `ABORTED`. With `errors: 'object'` the row values include `CellError`
 * objects; when `errors` is not a literal the values are typed `ReadValue` (`CellValue | CellError`).
 *
 * `openWorkbook` takes ownership of a `RandomAccessSource` (`fromFile` in Node): if opening fails for any reason the
 * source is closed before the promise rejects, and once it succeeds `workbook.close()` closes it. A source that
 * fails to read is reported as `IO_ERROR` with the source's own error as the `cause`.
 */
export function openWorkbook(input: SourceInput, options: OpenOptions & { readonly errors: 'object' }): Promise<Workbook<ReadValue>>;
export function openWorkbook(input: SourceInput, options?: OpenOptions & { readonly errors?: 'string' | 'null' }): Promise<Workbook>;
export function openWorkbook(input: SourceInput, options?: OpenOptions): Promise<Workbook<ReadValue>>;
export function openWorkbook(input: SourceInput, options: OpenOptions = {}): Promise<Workbook<ReadValue>> {
  return openPackage(input, options);
}

/**
 * Takes ownership of the source: once `sourceFrom` accepted the input, any failure closes it (a close that fails as
 * well is swallowed, so the caller sees the failure that mattered), and success hands it to the workbook.
 */
async function openPackage(input: SourceInput, options: OpenOptions): Promise<Workbook<ReadValue>> {
  const source = guardSource(sourceFrom(input));
  let zip: ZipReader | undefined;
  try {
    throwIfAborted(options.signal);
    const headLength = Math.min(source.size, SNIFF_BYTES);
    const head = headLength > 0 ? await source.read(0, headLength) : new Uint8Array(0);
    const format = sniff(head);
    if (format !== 'zip') {
      rejectSniffedFormat(format);
    }

    const limits = resolveLimits(options.limits);
    zip = await ZipReader.open(source, { maxEntries: limits.maxEntries, maxInflatedBytes: limits.maxInflatedBytes });
    return await readPackage(zip, options, limits);
  } catch (error) {
    try {
      await (zip === undefined ? source.close?.() : zip.close());
    } catch {
      // The failure that got us here is the one worth reporting.
    }
    throw error;
  }
}

async function readPackage(zip: ZipReader, options: OpenOptions, limits: ResolvedLimits): Promise<Workbook<ReadValue>> {
  throwIfAborted(options.signal);
  await rejectOpenDocument(zip, limits);

  let contentTypes: ContentTypes | undefined;
  if (zip.has(CONTENT_TYPES_PART)) {
    const xml = await zip.readText(CONTENT_TYPES_PART, limits.maxPartBytes);
    contentTypes = parsePart(CONTENT_TYPES_PART, () => parseContentTypes(xml, limits.maxEntries));
  }

  const rootRels = await readRels(zip, PACKAGE_ROOT_PART, limits);
  const workbookPart = locateMainPart(zip, rootRels, contentTypes);
  if (workbookPart === undefined) {
    throw notXlsx('zip', 'The archive has no workbook part, so it is not an .xlsx workbook. Open it in Excel and save it as .xlsx.');
  }
  rejectForeignMainPart(workbookPart, contentTypes);

  const workbookXml = await zip.readText(workbookPart, limits.maxPartBytes);
  const parsed = parsePart(workbookPart, () => parseWorkbook(workbookXml, limits.maxEntries));
  const workbookRels = await readRels(zip, workbookPart, limits);
  const workbookDirectory = directoryOf(workbookPart);

  const sharedStringsPart = relationshipTarget(workbookRels, 'sharedStrings') ?? `${workbookDirectory}sharedStrings.xml`;
  const styles = await readStyles(zip, workbookRels, workbookDirectory, limits, options.signal);
  const context: WorksheetReadContext = {
    sharedStrings: sharedStringLoader(zip, sharedStringsPart, limits),
    hasSharedStrings: zip.has(sharedStringsPart),
    isDateByXf: styles.isDateByXf,
    cellFormatsIgnored: styles.formatsIgnored,
    date1904: parsed.date1904,
    dates: options.dates ?? 'local',
    errors: options.errors ?? 'string',
    maxXmlDepth: limits.maxXmlDepth,
    maxTextLength: limits.maxTextLength,
    maxRowChars: limits.maxRowChars,
    signal: options.signal,
  };

  const relById = new Map(workbookRels.map(rel => [rel.id, rel]));
  const readers = parsed.sheets.map((entry, index) => {
    const rel = relById.get(entry.relId);
    const target = rel !== undefined && !rel.external ? rel.target : '';
    const info: SheetInfo = {
      name: entry.name,
      index,
      kind: rel !== undefined && relTypeIs(rel.type, 'chartsheet') ? 'chartsheet' : 'worksheet',
      hidden: entry.state !== 'visible',
    };
    // Without a usable relationship the conventional part name is the only lead left (EC-PART-NONSTANDARD-NAMES).
    const partName = target === '' ? `${workbookDirectory}worksheets/sheet${entry.sheetId}.xml` : target;
    const onWarning = options.onWarning;
    const sheetContext = onWarning === undefined ? context : { ...context, onWarning: sheetWarnings(onWarning, info.name) };
    return new SheetReader(info, partName, { zip, context: sheetContext, maxObjectCells: limits.maxObjectCells });
  });
  return new WorkbookReader(zip, parsed.date1904, readers);
}

/**
 * The caller's warning callback for one sheet: names the sheet, and passes each code on once however many reads of
 * the sheet (rows, head, toObjects) run into it. The row reader already reports each code once per read.
 */
function sheetWarnings(report: (warning: ReadWarning) => void, sheet: string): (warning: ReadWarning) => void {
  const reported = new Set<ReadWarningCode>();
  return (warning: ReadWarning): void => {
    if (reported.has(warning.code)) {
      return;
    }
    reported.add(warning.code);
    try {
      report({ ...warning, sheet });
    } catch (error) {
      // Like every caller callback: a throw stops the read as IO_ERROR, with the callback's error as the cause.
      throw classifyError(error, 'IO_ERROR', 'The onWarning callback threw, so reading the spreadsheet stopped.');
    }
  };
}

/** Without a styles part nothing is a date: every `isDateByXf` lookup on an empty array is `undefined`. */
const NO_STYLES: ParsedStyles = { isDateByXf: new Uint8Array(0), formatsIgnored: false };

/**
 * Streamed, not read whole: a heavily styled workbook's `styles.xml` can be tens of MB, past `maxPartBytes`, and the
 * parser keeps only bounded tables from it, so `maxInflatedBytes` is the cap that applies, as for a worksheet.
 */
async function readStyles(
  zip: ZipReader,
  workbookRels: readonly Relationship[],
  workbookDirectory: string,
  limits: ResolvedLimits,
  signal: AbortSignal | undefined,
): Promise<ParsedStyles> {
  const stylesPart = relationshipTarget(workbookRels, 'styles') ?? `${workbookDirectory}styles.xml`;
  if (!zip.has(stylesPart)) {
    return NO_STYLES;
  }
  return parsePartAsync(stylesPart, () =>
    parseStyles(zip.stream(stylesPart), { maxDepth: limits.maxXmlDepth, maxTextLength: limits.maxTextLength, signal }),
  );
}

/** The shared-string table is parsed on the first `t="s"` cell of the first sheet that has one, then reused. */
function sharedStringLoader(zip: ZipReader, partName: string, limits: ResolvedLimits): () => Promise<readonly string[]> {
  let pending: Promise<readonly string[]> | undefined;
  return () => {
    pending ??= zip.has(partName)
      ? parsePartAsync(partName, () =>
          parseSharedStrings(zip.stream(partName), {
            maxChars: limits.maxSharedStringChars,
            maxDepth: limits.maxXmlDepth,
            maxTextLength: limits.maxTextLength,
          }),
        )
      : Promise.resolve([]);
    return pending;
  };
}
