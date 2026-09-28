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

/**
 * An Excel error value (`t="e"`). The writer writes the `CellErrorCode` literals as error cells, another `#` code (a
 * newer Excel error such as `#SPILL!`) as its text, and refuses anything else (EC-ERROR-CODE-UNKNOWN). The reader
 * returns it as a string, an object or null per `OpenOptions.errors`. `error` is typed as the standard codes, but the
 * reader passes through whatever literal the file holds, so under `errors: 'object'` a newer code (`#SPILL!`,
 * `#CALC!`, `#FIELD!`, ...) can appear: do not treat `CellErrorCode` as exhaustive when reading.
 */
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
    /** Points, 1 to 409. */
    readonly size?: number;
    /** `#RRGGBB` */
    readonly color?: string;
    /** 1 to 31 characters. */
    readonly name?: string;
  };
  /** Solid fill only; `color` (`#RRGGBB`) is required. */
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
  /**
   * A number format code of 1 to 255 characters (`'0.00'`, `'yyyy-mm-dd'`) or a built-in numFmt id (an integer
   * 0-163). `registerStyle` validates every field and throws `WRITER_STATE` naming the one it refused.
   */
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

/** Document properties. Characters XML 1.0 forbids (control characters, unpaired surrogates) are dropped. */
export interface WorkbookProperties {
  readonly creator?: string;
  readonly title?: string;
  /** A valid Date in the years 1-9999 (UTC); anything else makes `createWorkbookWriter` throw `WRITER_STATE`. */
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
  /** Default `'...(truncated)'`. The truncated text plus the suffix is at most 32,767 UTF-16 units. */
  readonly truncationSuffix?: string;
  /**
   * Called once, when `close()` has finished the file, with the number of cells truncated across the whole workbook.
   * Not called when nothing was truncated (or with `cellOverflow: 'throw'`).
   * It runs after the file is complete and the sink has closed; an exception it throws still rejects `close()`.
   */
  readonly onCellTruncated?: (count: number) => void;
  /** Called every 5,000 rows and at each sheet close. */
  readonly onProgress?: (progress: WriteProgress) => void;
  readonly signal?: AbortSignal;
  readonly properties?: WorkbookProperties;
}

export interface ColumnOptions {
  /**
   * Width in character units (Excel's column width). Widths past Excel's maximum of 255 are clamped to 255; a NaN,
   * negative or infinite width makes `addSheet` throw `WRITER_STATE`.
   */
  readonly width?: number;
  readonly hidden?: boolean;
  /** A registered style id (from `registerStyle`); anything else makes `addSheet` throw `WRITER_STATE`. */
  readonly style?: StyleId;
}

export interface SheetOptions {
  /** Written as row 1 with `headerStyle` (bold by default). */
  readonly header?: readonly CellInput[];
  /** A registered style id, or `false` to write the header unstyled. */
  readonly headerStyle?: StyleId | false;
  readonly columns?: readonly ColumnOptions[];
  readonly freeze?: { readonly rows?: number; readonly cols?: number };
  /** AutoFilter over the header row and all data columns. */
  readonly autoFilter?: boolean;
  /** Hidden sheets are allowed as long as at least one sheet in the workbook stays visible. */
  readonly hidden?: boolean;
  /**
   * Expected number of data rows (excluding the header): a hint that only decides up front whether the worksheet
   * part needs zip64 (see `zip64: 'auto'`). Writing more or fewer rows is fine, and it never becomes a `<dimension>`
   * element. Must be a non-negative safe integer, or `addSheet` throws `WRITER_STATE`.
   */
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
  /**
   * Write one row. `values` must be an array; `styles` is one id for every cell or one id per cell (undefined =
   * default), each one a registered style id. A row that is not an array, an unsupported value (plain object,
   * function, symbol, ...), an unknown error code or an unregistered style id rejects with `WRITER_STATE` naming the
   * sheet and cell, and - like any failure while writing - aborts the sink and fails every later call with the same
   * error.
   */
  writeRow(values: readonly CellInput[], styles?: StyleId | readonly (StyleId | undefined)[]): Promise<void>;
  writeRows(rows: Iterable<readonly CellInput[]> | AsyncIterable<readonly CellInput[]>): Promise<void>;
  /**
   * Register a merged range in A1 notation (`'A1:C1'`, either case). Validated immediately: a malformed range, a
   * single cell, a range outside the grid, or one that overlaps or repeats a range already merged on this sheet
   * throws `WRITER_STATE` and is not registered. Emitted at close.
   */
  merge(range: string): void;
  close(): Promise<SheetWriteSummary>;
}

export interface WorkbookWriter {
  /** One sheet may be open at a time; close it before adding the next. */
  addSheet(name: string, options?: SheetOptions): SheetWriter;
  registerStyle(style: CellStyle): StyleId;
  /**
   * Finalize: shared strings, styles, workbook parts, central directory; closes the sink. Rejects with
   * `WRITER_STATE` and aborts the sink when no sheet was added or every sheet is hidden. After any earlier failure it
   * rejects with that original error.
   */
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

/**
 * What a `ReadWarning` is about:
 * - `'SST_INDEX_OUT_OF_RANGE'`: a `t="s"` cell points past the end of the shared-string table. It reads as `''`,
 *   which is what Excel shows (EC-SST-INDEX-OUT-OF-RANGE).
 * - `'SHARED_STRINGS_MISSING'`: a `t="s"` cell in a workbook that has no shared-string part at all. It reads as `''`
 *   (EC-SST-ABSENT-INLINE-ONLY).
 */
export type ReadWarningCode = 'SST_INDEX_OUT_OF_RANGE' | 'SHARED_STRINGS_MISSING';

/** Something a sheet says that Excel tolerates and the reader repaired, reported through `OpenOptions.onWarning`. */
export interface ReadWarning {
  readonly code: ReadWarningCode;
  readonly message: string;
  /** Name of the sheet being read. */
  readonly sheet?: string;
  /** A1 reference of the first cell that triggered the warning. */
  readonly ref?: string;
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
  /**
   * Called when a sheet holds something the reader repaired instead of rejecting (see `ReadWarningCode`). Each code
   * is reported at most once per sheet, for the first cell that hit it, however many cells do; the read goes on
   * either way. Without a callback nothing is collected.
   */
  readonly onWarning?: (warning: ReadWarning) => void;
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

/**
 * Object mode, the `sheet_to_json` contract. Columns are named from the header row, starting at the first column of
 * the sheet's `<dimension>` when it declares one (column A otherwise) and running to the wider of the dimension and
 * the header row. A data row with a value outside those columns still keeps it, under a blank-header name
 * (`__EMPTY`, ...). `toObjects()` gives every record every column, filling the ones named after a record was built
 * with `defval`; streaming `rows({ mode: 'object' })` cannot revisit records it already yielded, so there a column
 * first named by a later row is missing from the records before it.
 */
export interface ObjectRowsOptions {
  /**
   * 1-based row holding the headers. By default the header row is the first row at or after `startRow` that holds a
   * value (an empty-string or error cell counts; a row of blank cells does not), so a sheet whose headers sit below
   * blank rows still gets them (EC-HEADER-ROW-FIRST-NONEMPTY). That matches SheetJS whenever the sheet's
   * `<dimension>` is accurate; when it starts at a formatted but empty row, SheetJS names every column `__EMPTY` and
   * this finds the first row with values instead. Data starts on the next row. Given explicitly, that exact row is
   * the header row even when it is blank.
   */
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
