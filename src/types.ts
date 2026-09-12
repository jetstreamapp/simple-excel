/**
 * Public value model, options and interfaces. Everything here is browser-safe: no Node or DOM types beyond
 * `Uint8Array`, `Blob`, `WritableStream` and `AbortSignal`.
 */

// ---------------------------------------------------------------------------------------------------------------------
// Values
// ---------------------------------------------------------------------------------------------------------------------

/** What a cell reads back as. Dates are only produced when the cell's number format is a date format. */
export type CellValue = string | number | boolean | Date | null;

export type CellErrorCode = '#NULL!' | '#DIV/0!' | '#VALUE!' | '#REF!' | '#NAME?' | '#NUM!' | '#N/A' | '#GETTING_DATA';

/** An Excel error value (`t="e"`). Written as-is; read back as a string, an object or null per `OpenOptions.errors`. */
export interface CellError {
  readonly error: CellErrorCode;
}

/**
 * What the writer accepts. `undefined` and `null` write no `<c>` element; `bigint` becomes a number when safe and a
 * string otherwise; NaN and infinities become `#NUM!`; invalid Dates write nothing.
 */
export type CellInput = CellValue | CellError | bigint | undefined;

/** Index into the workbook's `cellXfs`; 0 is the default style. Obtained from `WorkbookWriter.registerStyle`. */
export type StyleId = number;

/** A cell as returned by `Sheet.head()`: the value plus what the file said about it. */
export interface RawCell {
  readonly value: CellValue;
  /** Present when the cell held an error value; `value` is then the error text. */
  readonly error?: CellErrorCode;
  /** Formula text when the cell has one (cached value is in `value`). */
  readonly formula?: string;
}

// ---------------------------------------------------------------------------------------------------------------------
// Bytes in and out
// ---------------------------------------------------------------------------------------------------------------------

/** Where the writer sends bytes. `write` may apply back-pressure by resolving late. */
export interface ByteSink {
  write(chunk: Uint8Array): Promise<void>;
  close(): Promise<void>;
  abort(reason?: unknown): Promise<void>;
}

/** Where the reader gets bytes: anything that can serve byte ranges on demand. */
export interface RandomAccessSource {
  readonly size: number;
  /** Returns exactly `length` bytes (fewer only at end of input). The returned view must not be reused by the source. */
  read(offset: number, length: number): Promise<Uint8Array>;
  /** Release underlying handles; optional. */
  close?(): Promise<void>;
}

export type ChunkHandler = (chunk: Uint8Array) => Promise<void>;

export interface DeflaterOptions {
  readonly method: 'deflate' | 'store';
  readonly level?: number;
}

/**
 * Streaming compressor: bytes pushed in, compressed bytes handed to `onChunk` in order. Implementations wrap
 * `CompressionStream('deflate-raw')` in the core and `zlib.createDeflateRaw` in the node entry.
 */
export interface Deflater {
  push(chunk: Uint8Array): Promise<void>;
  /** Flush and finish; resolves after the last `onChunk` call settled. */
  finish(): Promise<void>;
  abort(reason?: unknown): Promise<void>;
  readonly bytesIn: number;
  readonly bytesOut: number;
}

export type DeflaterFactory = (onChunk: ChunkHandler, options: DeflaterOptions) => Deflater;

// ---------------------------------------------------------------------------------------------------------------------
// Styles (writer, v1)
// ---------------------------------------------------------------------------------------------------------------------

export type BorderLineStyle = 'thin' | 'medium' | 'thick' | 'dashed' | 'dotted' | 'double' | 'hair';

export interface CellStyle {
  readonly font?: {
    readonly bold?: boolean;
    readonly italic?: boolean;
    readonly underline?: boolean;
    readonly strike?: boolean;
    readonly size?: number;
    /** `#RRGGBB` */
    readonly color?: string;
    readonly name?: string;
  };
  /** Solid fill only. `#RRGGBB` */
  readonly fill?: { readonly color: string };
  readonly border?:
    | BorderLineStyle
    | {
        readonly top?: BorderLineStyle;
        readonly bottom?: BorderLineStyle;
        readonly left?: BorderLineStyle;
        readonly right?: BorderLineStyle;
        /** `#RRGGBB`, applied to every side given */
        readonly color?: string;
      };
  readonly alignment?: {
    readonly horizontal?: 'left' | 'center' | 'right';
    readonly vertical?: 'top' | 'center' | 'bottom';
    readonly wrapText?: boolean;
  };
  /** A number format code (`'0.00'`, `'yyyy-mm-dd'`) or a built-in numFmt id. */
  readonly numFmt?: string | number;
}

// ---------------------------------------------------------------------------------------------------------------------
// Writer
// ---------------------------------------------------------------------------------------------------------------------

export interface SharedStringBudget {
  /** Stop interning new strings once this many unique strings exist. Default 65,536. */
  readonly maxUnique?: number;
  /** Stop interning once the interned text totals this many UTF-16 units. Default 16 Mi. */
  readonly maxChars?: number;
  /** Strings longer than this are always written inline. Default 256. */
  readonly maxLength?: number;
}

export interface WorkbookProperties {
  readonly creator?: string;
  readonly title?: string;
  readonly created?: Date;
}

export interface WriteProgress {
  readonly sheet: string;
  readonly rows: number;
  readonly bytesOut: number;
}

export interface WorkbookWriterOptions {
  /**
   * `'inline'` (default): every string is written in its cell and no shared-string table is built - a constant memory
   * footprint, output within a couple of percent of the table's after compression, and the shape every application
   * reads faithfully (Numbers mangles control characters and escapes in shared strings, ADR-001).
   * `'auto'`: bounded hybrid, short repeated strings go to the table until `sstBudget` is hit, then everything is
   * inline. `'shared'`: intern everything (unbounded; for tests).
   */
  readonly strings?: 'auto' | 'inline' | 'shared';
  readonly sstBudget?: SharedStringBudget;
  /**
   * `'auto'` (default): zip64 headers are emitted only for a sheet whose `rowCount` hint makes a >4 GiB entry
   * plausible (about 40M cells); everything else stays a standard zip, which every reader (including SheetJS, which
   * cannot open zip64) accepts, and a part that still passes 4 GiB fails with `ENTRY_TOO_LARGE`. `true` forces zip64
   * on every streamed part; `false` never emits it.
   */
  readonly zip64?: 'auto' | boolean;
  readonly compression?: 'deflate' | 'store';
  /** Alternate deflate implementation (the node entry provides a zlib-backed one). */
  readonly deflater?: DeflaterFactory;
  /** Reproducible bytes: fixed zip timestamps and docProps dates. Default false. */
  readonly deterministic?: boolean;
  /**
   * Which fields of a `Date` carry the wall-clock time written to the cell. `'local'` (default) writes
   * `getFullYear()/getHours()`; `'utc'` writes `getUTCFullYear()/getUTCHours()`.
   */
  readonly dates?: 'local' | 'utc';
  /** Workbook date system. Default 1900. */
  readonly date1904?: boolean;
  /** Cells over 32,767 characters: truncate (default, with `truncationSuffix`) or throw `CELL_TOO_LONG`. */
  readonly cellOverflow?: 'truncate' | 'throw';
  /** Default `'...(truncated)'`. */
  readonly truncationSuffix?: string;
  readonly onCellTruncated?: (count: number) => void;
  /** Called every 5,000 rows and at each sheet close. */
  readonly onProgress?: (progress: WriteProgress) => void;
  readonly signal?: AbortSignal;
  readonly properties?: WorkbookProperties;
}

export interface ColumnOptions {
  /** Width in character units (Excel's column width). */
  readonly width?: number;
  readonly hidden?: boolean;
  readonly style?: StyleId;
}

export interface SheetOptions {
  /** Written as row 1 with `headerStyle` (bold by default). */
  readonly header?: readonly CellInput[];
  /** `false` writes the header unstyled. */
  readonly headerStyle?: StyleId | false;
  readonly columns?: readonly ColumnOptions[];
  readonly freeze?: { readonly rows?: number; readonly cols?: number };
  /** AutoFilter over the header row and all data columns. */
  readonly autoFilter?: boolean;
  readonly hidden?: boolean;
  /** Enables `<dimension>` and lets zip64 sizing be decided up-front. Data rows only (excludes the header). */
  readonly rowCount?: number;
}

export interface SheetWriteSummary {
  readonly name: string;
  readonly rows: number;
  readonly columns: number;
}

export interface WorkbookWriteResult {
  readonly bytes: number;
  readonly sheets: readonly SheetWriteSummary[];
  readonly truncatedCells: number;
  readonly sharedStrings: { readonly count: number; readonly uniqueCount: number; readonly frozen: boolean };
}

export interface SheetWriter {
  /** The name actually used (after sanitizing and de-duplication). */
  readonly name: string;
  /** 1-based index of the next row that `writeRow` will produce. */
  readonly nextRow: number;
  /** Write one row. `styles` is one id for every cell or one id per cell (undefined = default). */
  writeRow(values: readonly CellInput[], styles?: StyleId | readonly (StyleId | undefined)[]): Promise<void>;
  writeRows(rows: Iterable<readonly CellInput[]> | AsyncIterable<readonly CellInput[]>): Promise<void>;
  /** Register a merged range in A1 notation (`'A1:C1'`). Emitted at close. */
  merge(range: string): void;
  close(): Promise<SheetWriteSummary>;
}

export interface WorkbookWriter {
  /** One sheet may be open at a time; close it before adding the next. */
  addSheet(name: string, options?: SheetOptions): SheetWriter;
  registerStyle(style: CellStyle): StyleId;
  /** Finalize: shared strings, styles, workbook parts, central directory; closes the sink. */
  close(): Promise<WorkbookWriteResult>;
  abort(reason?: unknown): Promise<void>;
}

// ---------------------------------------------------------------------------------------------------------------------
// Reader
// ---------------------------------------------------------------------------------------------------------------------

export type SniffResult = 'xlsx' | 'xlsb' | 'ods' | 'zip' | 'cfb-encrypted' | 'cfb-legacy' | 'xml' | 'html' | 'empty' | 'text' | 'unknown';

export interface ReadLimits {
  /** Per part. Default 1 GiB. */
  readonly maxInflatedBytes?: number;
  /** Default 256 Mi UTF-16 units. */
  readonly maxSharedStringChars?: number;
  /** Default 10,000. */
  readonly maxEntries?: number;
  /** Default 256. */
  readonly maxXmlDepth?: number;
  /** Default 64 Mi UTF-16 units for one text node. */
  readonly maxTextLength?: number;
}

export interface OpenOptions {
  /**
   * `'local'` (default): `new Date(y, m, d, h, mi, s, ms)` from the serial's wall-clock components (what SheetJS +
   * `sheet_to_json` gives Jetstream). `'utc'`: `Date.UTC(...)`. `'serial'`: the raw number, never a Date.
   */
  readonly dates?: 'local' | 'utc' | 'serial';
  /** How error cells come back: their text (default), a `CellError` object, or null. */
  readonly errors?: 'string' | 'object' | 'null';
  readonly limits?: ReadLimits;
  readonly signal?: AbortSignal;
}

export interface SheetInfo {
  readonly name: string;
  /** Position in the workbook's sheet order, from 0. */
  readonly index: number;
  readonly kind: 'worksheet' | 'chartsheet';
  readonly hidden: boolean;
}

export interface RowsOptions {
  /** 1-based first row to yield. Default 1. */
  readonly startRow?: number;
  /** Stop after this many rows. */
  readonly maxRows?: number;
  /** Default false: rows with no cells are skipped (array mode still yields them when true, as `[]`). */
  readonly blankRows?: boolean;
  /** Cap on columns materialized per row (default 16,384). */
  readonly maxColumns?: number;
  /** `'value'` (default) reads the cached value; `'text'` yields formula text for formula cells. */
  readonly formulas?: 'value' | 'text';
}

export interface ObjectRowsOptions {
  /** 1-based row holding the headers. Default `startRow`. */
  readonly headerRow?: number;
  /** Value for cells that are absent from a row. Default `''`. */
  readonly defval?: CellValue;
  /**
   * `'sheetjs'` (default): empty headers become `__EMPTY`, `__EMPTY_1`, ...; duplicates get `_1`, `_2` suffixes.
   * `'index'`: every header is its 0-based column index as a string.
   */
  readonly headerNaming?: 'sheetjs' | 'index';
  /** Drop columns whose header is empty instead of naming them. Default false. */
  readonly dropEmptyHeaders?: boolean;
}

export type ArrayRowsOptions = RowsOptions & { readonly mode?: 'array' };
export type ObjectModeRowsOptions = RowsOptions & ObjectRowsOptions & { readonly mode: 'object' };

/** Element type of a row: `CellValue`, widened to include `CellError` when the workbook was opened with `errors: 'object'`. */
export type ReadValue = CellValue | CellError;

export interface ObjectsResult<V extends ReadValue = CellValue> {
  readonly rows: Record<string, V>[];
  readonly headers: string[];
  /** True when `maxRows` stopped the read before the sheet ended. */
  readonly truncated: boolean;
}

export interface Sheet<V extends ReadValue = CellValue> {
  readonly info: SheetInfo;
  rows(options?: ArrayRowsOptions): AsyncIterable<V[]>;
  rows(options: ObjectModeRowsOptions): AsyncIterable<Record<string, V>>;
  toObjects(options?: RowsOptions & ObjectRowsOptions): Promise<ObjectsResult<V>>;
  /** First `rowCount` rows as A1-keyed cells, then stops inflating. */
  head(rowCount: number): Promise<Map<string, RawCell>>;
}

export interface Workbook<V extends ReadValue = CellValue> {
  readonly sheets: readonly SheetInfo[];
  readonly date1904: boolean;
  sheet(nameOrIndex: string | number): Sheet<V>;
  close(): Promise<void>;
}
