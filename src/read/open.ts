/**
 * The reader facade: sniff the bytes, open the container, resolve the package parts, and hand out sheet readers
 * that stream rows on demand. Nothing here parses sheet XML - `sml/worksheet-reader` does - and nothing is read
 * eagerly beyond the workbook, its relationships and the styles part.
 */
import { isXlsxError, XlsxError } from '../errors';
import { MAX_COLUMNS, parseRange } from '../sml/cell-ref';
import { parseContentTypes, parseRels, parseWorkbook, relTypeIs, type ContentTypes, type Relationship } from '../sml/package-parts';
import { parseSharedStrings } from '../sml/shared-strings-reader';
import { parseStyles } from '../sml/styles';
import { readWorksheetHead, readWorksheetRows, type WorksheetReadContext, type WorksheetRowsOptions } from '../sml/worksheet-reader';
import type {
  ArrayRowsOptions,
  CellError,
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
import { sourceFrom, type SourceInput } from '../zip/source';
import { ZipReader } from '../zip/zip-reader';
import { SNIFF_BYTES, sniff } from './sniff';

const DEFAULT_MAX_ENTRIES = 10_000;
const DEFAULT_MAX_INFLATED_BYTES = 1024 * 1024 * 1024;
const DEFAULT_MAX_SHARED_STRING_CHARS = 256 * 1024 * 1024;
const DEFAULT_MAX_XML_DEPTH = 256;
const DEFAULT_MAX_TEXT_LENGTH = 64 * 1024 * 1024;

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

/** `<dimension ref="A1:T31"/>`, prefixed or not. It sits in the worksheet prologue, ahead of `<sheetData>`. */
const DIMENSION_PATTERN = /<(?:[\w.-]+:)?dimension[^>]*\sref="([^"]+)"/;
/** How far into a worksheet part to look for `<dimension>` before giving up on it. */
const DIMENSION_SCAN_CHARS = 64 * 1024;

interface ResolvedLimits {
  readonly maxEntries: number;
  readonly maxInflatedBytes: number;
  readonly maxSharedStringChars: number;
  readonly maxXmlDepth: number;
  readonly maxTextLength: number;
}

function resolveLimits(limits: ReadLimits | undefined): ResolvedLimits {
  return {
    maxEntries: limits?.maxEntries ?? DEFAULT_MAX_ENTRIES,
    maxInflatedBytes: limits?.maxInflatedBytes ?? DEFAULT_MAX_INFLATED_BYTES,
    maxSharedStringChars: limits?.maxSharedStringChars ?? DEFAULT_MAX_SHARED_STRING_CHARS,
    maxXmlDepth: limits?.maxXmlDepth ?? DEFAULT_MAX_XML_DEPTH,
    maxTextLength: limits?.maxTextLength ?? DEFAULT_MAX_TEXT_LENGTH,
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

/** Parsers throw `XlsxError` for everything they anticipate; anything else means the part is damaged. */
function damagedPart(partName: string, cause: unknown): XlsxError {
  return new XlsxError('XML_MALFORMED', `The part ${partName} could not be read; this file may be damaged.`, {
    part: partName,
    cause,
  });
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

async function readRels(zip: ZipReader, partName: string): Promise<Relationship[]> {
  const relsPart = relsPartFor(partName);
  if (!zip.has(relsPart)) {
    return [];
  }
  const xml = await zip.readText(relsPart);
  return parsePart(relsPart, () => parseRels(xml, directoryOf(partName)));
}

function relationshipTarget(rels: readonly Relationship[], suffix: string): string | undefined {
  return rels.find(rel => !rel.external && rel.target !== '' && relTypeIs(rel.type, suffix))?.target;
}

/**
 * An OpenDocument package announces itself with a `mimetype` entry, which is the only thing about it a
 * SpreadsheetML reader can recognize: it has no `[Content_Types].xml` at all.
 */
async function rejectOpenDocument(zip: ZipReader): Promise<void> {
  if (!zip.has(ODS_MIMETYPE_ENTRY)) {
    return;
  }
  const mimetype = (await zip.readText(ODS_MIMETYPE_ENTRY)).trim();
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
 * spelling stands in: identical for strings (every header in practice) and Excel's spelling for the other types. An
 * error cell names its column with the error text whatever `errors` says (EC-HEADER-ERROR-CELL); the row reader
 * hands the header row's errors over as text for that reason, and a `CellError` is read the same way.
 */
function headerTextOf(value: ReadValue): string {
  if (typeof value === 'string') {
    return value;
  }
  if (typeof value === 'boolean') {
    return value ? 'TRUE' : 'FALSE';
  }
  if (value instanceof Date) {
    const date = `${value.getFullYear()}-${padTwo(value.getMonth() + 1)}-${padTwo(value.getDate())}`;
    return `${date}T${padTwo(value.getHours())}:${padTwo(value.getMinutes())}:${padTwo(value.getSeconds())}`;
  }
  if (typeof value === 'object' && value !== null) {
    return value.error;
  }
  return String(value);
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
  /** Set once the stream ends, so `toObjects` can give earlier records the columns named after them. */
  builder: ObjectRowBuilder | undefined;
}

/** Where the sheet's `<dimension>` puts its columns: both 0 when it declares none. */
interface DeclaredColumns {
  /** 0-based first column. */
  readonly first: number;
  /** One past the 0-based last column. */
  readonly end: number;
}

const NO_DECLARED_COLUMNS: DeclaredColumns = { first: 0, end: 0 };

/** 1-based header row "the first row with a value", as opposed to a row the caller named. */
const DETECT_HEADER_ROW = 0;

/**
 * Turns rows into `{ header: value }` records: finds and names the header row's columns, skips everything before
 * the first data row, and fills absent cells with `defval`. Blank rows never reach it - the row reader drops them,
 * which is SheetJS's `blankrows: false` - so without an explicit `headerRow` the first row it sees is the header row
 * (EC-HEADER-ROW-FIRST-NONEMPTY).
 *
 * Columns run from the dimension's first column (EC-HEADER-DIMENSION-START-COLUMN) to the wider of the dimension and
 * the header row. Values outside that range are never dropped: the columns they sit in are named as the blanks they
 * are, to the right as a row widens and to the left when a value lies before the dimension's first column.
 */
class ObjectRowBuilder {
  /** The caller's `headerRow`, or `DETECT_HEADER_ROW`. */
  private readonly headerRow: number;
  private readonly requestedStartRow: number;
  private firstDataRow: number;
  private readonly declared: DeclaredColumns;
  private readonly defval: CellValue;
  private readonly naming: NonNullable<ObjectRowsOptions['headerNaming']>;
  private readonly dropEmptyHeaders: boolean;
  private readonly namer = new HeaderNamer();
  private readonly columns: HeaderColumn[] = [];
  /** Columns from `leftmostNamed` up to `namedColumns` have names, including any `dropEmptyHeaders` kept out of `columns`. */
  private leftmostNamed = 0;
  private namedColumns = 0;
  private headersReady = false;
  private recordsBuilt = 0;

  constructor(options: RowsOptions & ObjectRowsOptions, declared: DeclaredColumns) {
    this.requestedStartRow = Math.max(1, options.startRow ?? 1);
    this.headerRow = options.headerRow === undefined ? DETECT_HEADER_ROW : Math.max(1, options.headerRow);
    this.firstDataRow =
      this.headerRow === DETECT_HEADER_ROW ? Number.POSITIVE_INFINITY : Math.max(options.startRow ?? 0, this.headerRow + 1);
    this.declared = declared;
    // `defval: null` is a caller asking for nulls, not for the default.
    this.defval = options.defval === undefined ? '' : options.defval;
    this.naming = options.headerNaming ?? 'sheetjs';
    this.dropEmptyHeaders = options.dropEmptyHeaders === true;
  }

  /** 1-based row the underlying reader must start at: the header row when the caller named one, else `startRow`. */
  get startRow(): number {
    return this.headerRow === DETECT_HEADER_ROW ? this.requestedStartRow : this.headerRow;
  }

  /** Which row the reader should hand over with error cells as text: the header row, found or named. */
  get errorTextRow(): number | 'first' {
    return this.headerRow === DETECT_HEADER_ROW ? 'first' : this.headerRow;
  }

  get headers(): string[] {
    return this.columns.map(column => column.name);
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
    this.nameColumns(cells.length, NO_HEADER_CELLS);
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

  /**
   * Name the header row's columns: from the dimension's first column, or from an earlier one when the header row
   * itself has a value there (a dimension that is wrong about where the data starts), to the wider of the dimension
   * and the header row.
   */
  private nameHeaderRow(headerCells: readonly ReadValue[]): void {
    this.headersReady = true;
    const first = firstValueIndex(headerCells, this.declared.first);
    this.leftmostNamed = first;
    this.namedColumns = first;
    this.nameColumns(Math.max(this.declared.end, headerCells.length), headerCells);
  }

  /** Name every column that is not named yet, up to `width`, reading each name from `headerCells`. */
  private nameColumns(width: number, headerCells: readonly ReadValue[]): void {
    for (let column = this.namedColumns; column < width; column++) {
      const named = this.nameColumn(column, headerCells[column]);
      if (named !== undefined) {
        this.columns.push(named);
      }
    }
    if (width > this.namedColumns) {
      this.namedColumns = width;
    }
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
    const text = missing ? '' : headerTextOf(value);
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
}

class SheetReader implements Sheet<ReadValue> {
  readonly info: SheetInfo;

  private readonly parts: SheetParts;
  private readonly partName: string;
  /** `<dimension>` columns, resolved once per sheet and only when object mode asks for them. */
  private declared: DeclaredColumns | undefined;

  constructor(info: SheetInfo, partName: string, parts: SheetParts) {
    this.info = info;
    this.partName = partName;
    this.parts = parts;
  }

  rows(options?: ArrayRowsOptions): AsyncIterable<ReadValue[]>;
  rows(options: ObjectModeRowsOptions): AsyncIterable<Record<string, ReadValue>>;
  rows(options: ArrayRowsOptions | ObjectModeRowsOptions = {}): AsyncIterable<ReadValue[] | Record<string, ReadValue>> {
    if (options.mode === 'object') {
      return this.objectRows(options, { headers: [], truncated: false, builder: undefined });
    }
    return this.arrayRows(options);
  }

  async toObjects(options: RowsOptions & ObjectRowsOptions = {}): Promise<ObjectsResult<ReadValue>> {
    const state: ObjectReadState = { headers: [], truncated: false, builder: undefined };
    const rows: Record<string, ReadValue>[] = [];
    for await (const row of this.objectRows(options, state)) {
      rows.push(row);
    }
    // Everything is in hand here, unlike a stream: a column a later row named reaches the records before it too.
    state.builder?.backfill(rows);
    return { rows, headers: state.headers, truncated: state.truncated };
  }

  async head(rowCount: number): Promise<Map<string, RawCell>> {
    return readWorksheetHead(this.partStream(), this.parts.context, rowCount);
  }

  private async *arrayRows(options: ArrayRowsOptions): AsyncGenerator<ReadValue[]> {
    for await (const row of readWorksheetRows(this.partStream(), this.parts.context, options)) {
      yield row.cells as ReadValue[];
    }
  }

  private async *objectRows(options: RowsOptions & ObjectRowsOptions, state: ObjectReadState): AsyncGenerator<Record<string, ReadValue>> {
    const builder = new ObjectRowBuilder(options, await this.declaredColumns());
    const maxRows = options.maxRows ?? Number.POSITIVE_INFINITY;
    const rowOptions: WorksheetRowsOptions = {
      startRow: builder.startRow,
      maxColumns: options.maxColumns,
      formulas: options.formulas,
      errorTextRow: builder.errorTextRow,
    };
    let emitted = 0;
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
    } finally {
      state.headers = builder.headers;
      state.builder = builder;
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
   * Which columns the sheet says it has. SheetJS builds its header list from `<dimension>` when the sheet declares
   * one, so a file whose dimension is wider than its data (every Excel, Numbers and Sheets re-save) grows trailing
   * `__EMPTY` columns, and one whose column A is empty (Excel writes `B2:D9`) starts at B; reading the prologue is
   * what keeps object mode identical to `sheet_to_json`.
   */
  private async declaredColumns(): Promise<DeclaredColumns> {
    if (this.declared !== undefined) {
      return this.declared;
    }
    let declared: DeclaredColumns = NO_DECLARED_COLUMNS;
    if (this.parts.zip.has(this.partName)) {
      const decoder = new TextDecoder('utf-8');
      let prologue = '';
      for await (const chunk of this.parts.zip.stream(this.partName)) {
        prologue += decoder.decode(chunk, { stream: true });
        const ref = DIMENSION_PATTERN.exec(prologue)?.[1];
        if (ref !== undefined) {
          const range = parseRange(ref);
          if (range !== null) {
            declared = { first: range.start.col, end: Math.min(range.end.col + 1, MAX_COLUMNS) };
          }
          break;
        }
        // `<dimension>` is optional, but it always precedes the rows: once they start there is nothing left to find.
        if (prologue.includes('<sheetData') || prologue.length > DIMENSION_SCAN_CHARS) {
          break;
        }
      }
    }
    this.declared = declared;
    return declared;
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
 * objects.
 */
export function openWorkbook(
  input: SourceInput,
  options: OpenOptions & { readonly errors: 'object' },
): Promise<Workbook<CellValue | CellError>>;
export function openWorkbook(input: SourceInput, options?: OpenOptions): Promise<Workbook>;
export function openWorkbook(input: SourceInput, options: OpenOptions = {}): Promise<Workbook<CellValue | CellError>> {
  return openPackage(input, options);
}

async function openPackage(input: SourceInput, options: OpenOptions): Promise<Workbook<ReadValue>> {
  throwIfAborted(options.signal);
  const source = sourceFrom(input);
  const headLength = Math.min(source.size, SNIFF_BYTES);
  const head = headLength > 0 ? await source.read(0, headLength) : new Uint8Array(0);
  const format = sniff(head);
  if (format !== 'zip') {
    rejectSniffedFormat(format);
  }

  const limits = resolveLimits(options.limits);
  const zip = await ZipReader.open(source, { maxEntries: limits.maxEntries, maxInflatedBytes: limits.maxInflatedBytes });
  try {
    return await readPackage(zip, options, limits);
  } catch (error) {
    await zip.close();
    throw error;
  }
}

async function readPackage(zip: ZipReader, options: OpenOptions, limits: ResolvedLimits): Promise<Workbook<ReadValue>> {
  throwIfAborted(options.signal);
  await rejectOpenDocument(zip);

  let contentTypes: ContentTypes | undefined;
  if (zip.has(CONTENT_TYPES_PART)) {
    const xml = await zip.readText(CONTENT_TYPES_PART);
    contentTypes = parsePart(CONTENT_TYPES_PART, () => parseContentTypes(xml));
  }

  const rootRels = await readRels(zip, PACKAGE_ROOT_PART);
  const workbookPart = locateMainPart(zip, rootRels, contentTypes);
  if (workbookPart === undefined) {
    throw notXlsx('zip', 'The archive has no workbook part, so it is not an .xlsx workbook. Open it in Excel and save it as .xlsx.');
  }
  rejectForeignMainPart(workbookPart, contentTypes);

  const workbookXml = await zip.readText(workbookPart);
  const parsed = parsePart(workbookPart, () => parseWorkbook(workbookXml));
  const workbookRels = await readRels(zip, workbookPart);
  const workbookDirectory = directoryOf(workbookPart);

  const sharedStringsPart = relationshipTarget(workbookRels, 'sharedStrings') ?? `${workbookDirectory}sharedStrings.xml`;
  const context: WorksheetReadContext = {
    sharedStrings: sharedStringLoader(zip, sharedStringsPart, limits),
    hasSharedStrings: zip.has(sharedStringsPart),
    isDateByXf: await readDateStyles(zip, workbookRels, workbookDirectory),
    date1904: parsed.date1904,
    dates: options.dates ?? 'local',
    errors: options.errors ?? 'string',
    maxXmlDepth: limits.maxXmlDepth,
    maxTextLength: limits.maxTextLength,
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
    return new SheetReader(info, partName, { zip, context: sheetContext });
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
    report({ ...warning, sheet });
  };
}

/** Without a styles part nothing is a date: every `isDateByXf` lookup on an empty array is `undefined`. */
async function readDateStyles(zip: ZipReader, workbookRels: readonly Relationship[], workbookDirectory: string): Promise<Uint8Array> {
  const stylesPart = relationshipTarget(workbookRels, 'styles') ?? `${workbookDirectory}styles.xml`;
  if (!zip.has(stylesPart)) {
    return new Uint8Array(0);
  }
  const xml = await zip.readText(stylesPart);
  return parsePart(stylesPart, () => parseStyles(xml)).isDateByXf;
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
