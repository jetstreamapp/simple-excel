import { XlsxError } from '../errors';
import type { CellErrorCode, CellValue, OpenOptions, RawCell, ReadWarning, ReadWarningCode, RowsOptions } from '../types';
import { decodeCellText } from '../xml/escape';
import { XmlTokenizer, type XmlTokenizerHandler } from '../xml/tokenizer';
import { columnIndexOfRef, formatRef, MAX_COLUMNS, MAX_ROWS } from './cell-ref';
import { componentsFromIso, componentsFromSerial, dateFromComponents, serialFromComponents } from './date';

/** What the facade's object mode adds to a row read. */
export interface WorksheetRowsOptions extends RowsOptions {
  /**
   * The header row of an object-mode read: its error cells read as their text whatever `errors` says, since SheetJS
   * names such a column `#N/A` (EC-HEADER-ERROR-CELL). A row number, or `'first'` for the first row that is not
   * blank, which is how object mode finds its header when the caller names none. Every other row follows `errors`.
   */
  readonly errorTextRow?: number | 'first';
}

export interface WorksheetReadContext {
  /** Lazily loaded shared strings; called at most once per sheet read, only when a `t="s"` cell appears. */
  readonly sharedStrings: () => Promise<readonly string[]>;
  /**
   * Whether the package actually has a shared-strings part. When it does not, nothing can resolve through the
   * table and the per-chunk `t="s"` pre-scan is pure cost, so it is skipped: a `t="s"` cell in such a file takes
   * the deferred path instead and resolves against the same empty table.
   */
  readonly hasSharedStrings: boolean;
  readonly isDateByXf: Uint8Array;
  readonly date1904: boolean;
  readonly dates: NonNullable<OpenOptions['dates']>;
  readonly errors: NonNullable<OpenOptions['errors']>;
  readonly maxXmlDepth: number;
  readonly maxTextLength: number;
  /** Checked between chunks; an aborted signal ends the read with `XlsxError('ABORTED')`. Supplied by the facade. */
  readonly signal?: AbortSignal;
  /**
   * Optional diagnostics sink. Each warning code fires at most once per sheet read, however many cells hit it. The
   * facade adds the sheet name; the parser fills in the code, the message and the cell reference.
   */
  readonly onWarning?: (warning: ReadWarning) => void;
}

export interface SheetRow {
  /** 1-based row number from the file (or inferred when `r` is missing). */
  readonly index: number;
  /** Dense, 0-based; trailing empties trimmed; holes are null. */
  readonly cells: CellValue[];
}

// Where the machine is in the document. Integers, so the hot switch compares numbers rather than strings.
const STATE_OUTSIDE = 0;
const STATE_SHEET_DATA = 1;
const STATE_ROW = 2;
const STATE_CELL = 3;

// Which buffer the next `text()` callback belongs to.
const CAPTURE_NONE = 0;
const CAPTURE_VALUE = 1;
const CAPTURE_INLINE = 2;
const CAPTURE_FORMULA = 3;

// The `t` attribute, resolved once per cell.
const TYPE_NUMBER = 0;
const TYPE_SHARED = 1;
const TYPE_STRING = 2;
const TYPE_INLINE = 3;
const TYPE_BOOLEAN = 4;
const TYPE_ERROR = 5;
const TYPE_ISO_DATE = 6;

/** `t="s"` is five characters, so four kept from the previous chunk catch one split across the boundary. */
const SHARED_TYPE_SCAN_TAIL = 4;

const CODE_SPACE = 0x20;
const CODE_DOLLAR = 0x24;
const CODE_ZERO = 0x30;
const CODE_NINE = 0x39;
const CODE_UPPER_A = 0x41;
const CODE_UPPER_Z = 0x5a;
const CODE_LOWER_B = 0x62;
const CODE_LOWER_O = 0x6f;
const CODE_LOWER_X = 0x78;
/** OR-ing an ASCII letter with this lower-cases it. */
const ASCII_LOWER_BIT = 0x20;

// What a numeric `<v>` that starts with a zero or with whitespace turns out to hold; see `leadingNumericKind`.
const NUMERIC_DECIMAL = 0;
const NUMERIC_BLANK = 1;
const NUMERIC_RADIX = 2;

function cellTypeOf(attribute: string): number {
  switch (attribute) {
    case 'n':
      return TYPE_NUMBER;
    case 's':
      return TYPE_SHARED;
    case 'str':
      return TYPE_STRING;
    case 'inlineStr':
      return TYPE_INLINE;
    case 'b':
      return TYPE_BOOLEAN;
    case 'e':
      return TYPE_ERROR;
    case 'd':
      return TYPE_ISO_DATE;
    default:
      // The schema default is `n`, and a `t` we do not know is far more likely a typo than a new type.
      return TYPE_NUMBER;
  }
}

/** Excel only decodes `_xHHHH_` escapes, so text without `_x` never needs the scan (the overwhelming majority). */
function decodeIfEscaped(text: string): string {
  return text.includes('_x') ? decodeCellText(text) : text;
}

function hasSharedType(text: string): boolean {
  return text.includes('t="s"') || text.includes("t='s'");
}

function isXmlWhitespace(code: number): boolean {
  return code === CODE_SPACE || code === 0x09 || code === 0x0a || code === 0x0d;
}

/**
 * Unary `+` accepts more than a spreadsheet number: whitespace alone is 0, and `0x1A`, `0b11`, `0o7` are hex, binary
 * and octal. Excel only writes decimal, so the first is a blank (EC-EMPTY-V-ELEMENT) and the rest are not numbers at
 * all (EC-NUMBER-RADIX-PREFIX). Only text that starts with `0` or whitespace can be either, so only that text is
 * looked at here: after any leading whitespace, nothing at all, or a zero followed by a radix letter.
 */
function leadingNumericKind(text: string): number {
  let position = 0;
  while (position < text.length && isXmlWhitespace(text.charCodeAt(position))) {
    position++;
  }
  if (position === text.length) {
    return NUMERIC_BLANK;
  }
  if (text.charCodeAt(position) !== CODE_ZERO) {
    return NUMERIC_DECIMAL;
  }
  const radixLetter = text.charCodeAt(position + 1) | ASCII_LOWER_BIT;
  return radixLetter === CODE_LOWER_X || radixLetter === CODE_LOWER_B || radixLetter === CODE_LOWER_O ? NUMERIC_RADIX : NUMERIC_DECIMAL;
}

/**
 * The 0-based column of a reference that is well formed but lies past XFD (`XFE2`, `AAAA1`), or -1 for anything
 * else. `columnIndexOfRef` reports such a reference as malformed, and a malformed `r` means "the next column"; told
 * apart here, on that rare path only, so a cell past the grid cannot be moved back into it (EC-REF-BEYOND-XFD).
 */
function columnIndexPastGrid(ref: string): number {
  let position = ref.charCodeAt(0) === CODE_DOLLAR ? 1 : 0;
  const lettersStart = position;
  let oneBasedColumn = 0;
  while (position < ref.length) {
    const code = ref.charCodeAt(position);
    if (code < CODE_UPPER_A || code > CODE_UPPER_Z) {
      break;
    }
    // Saturate rather than overflow: every column past the grid is as unusable as the next.
    oneBasedColumn = Math.min(oneBasedColumn * 26 + (code - CODE_UPPER_A + 1), MAX_COLUMNS + 1);
    position++;
  }
  if (position === lettersStart || oneBasedColumn <= MAX_COLUMNS) {
    return -1;
  }
  if (ref.charCodeAt(position) === CODE_DOLLAR) {
    position++;
  }
  const rowCode = ref.charCodeAt(position);
  return rowCode >= CODE_ZERO && rowCode <= CODE_NINE ? oneBasedColumn - 1 : -1;
}

function limitExceeded(message: string, detail: Readonly<Record<string, unknown>>): XlsxError {
  return new XlsxError('LIMIT_EXCEEDED', message, detail);
}

function throwIfAborted(signal: AbortSignal | undefined): void {
  if (signal?.aborted === true) {
    throw new XlsxError('ABORTED', 'Reading this spreadsheet was cancelled.');
  }
}

interface ParserOptions {
  /** `<f>` text is only captured when someone wants it: `formulas: 'text'`, or `head()`, which always reports it. */
  readonly captureFormulas: boolean;
  readonly maxColumns: number;
  /** How `t="e"` cells materialize. `head()` pins this to `'string'`, since a RawCell carries the text and the code. */
  readonly errorMode: NonNullable<OpenOptions['errors']>;
}

function rawCell(value: CellValue, error: CellErrorCode | undefined, formula: string | undefined): RawCell {
  if (error !== undefined) {
    return formula === undefined ? { value, error } : { value, error, formula };
  }
  return formula === undefined ? { value } : { value, formula };
}

/**
 * The cell state machine both read paths share: `sheetData` -> `row` -> `c` -> (`v` | `is` | `f`), materializing one
 * value per `</c>`. Subclasses decide where those values go (dense rows, or an A1-keyed map for `head()`).
 *
 * Everything a cell needs lives in fields rather than a per-cell object, so a million-cell sheet allocates only the
 * row arrays and the strings it actually returns.
 */
abstract class WorksheetParser implements XmlTokenizerHandler {
  /** Set when no more input is wanted: every callback returns at once and the caller stops pulling chunks. */
  stopped = false;

  protected readonly context: WorksheetReadContext;
  protected readonly tokenizer: XmlTokenizer;
  /** Which fields of a produced `Date` carry the sheet's wall clock. Unused while `serialDates`. */
  protected readonly dateFields: 'local' | 'utc';
  protected readonly serialDates: boolean;
  /** How `t="e"` cells materialize for the row being read; object mode switches it for its header row. */
  protected errorMode: NonNullable<OpenOptions['errors']>;
  protected readonly captureFormulas: boolean;
  protected readonly maxColumns: number;

  /** 1-based; the row currently open, or the last one seen. */
  protected rowNumber = 0;
  /** 0-based column of the cell currently open, or the last one in this row. */
  protected column = -1;
  protected type = TYPE_NUMBER;
  protected style = 0;
  protected valueSeen = false;
  protected valueText = '';
  protected inlineSeen = false;
  protected inlineText = '';
  protected formulaText = '';
  /** False while rows before `startRow` stream past: they are parsed for their numbering and nothing else. */
  protected collecting = false;

  private state = STATE_OUTSIDE;
  private capture = CAPTURE_NONE;
  private inInlineString = false;
  private inPhonetic = false;
  private sharedStrings: readonly string[] | undefined;
  private sharedStringsScanTail = '';
  /** Cells that met `t="s"` before the table was loaded; applied after the chunk, before any row is yielded. */
  private readonly deferredCells: ((strings: readonly string[]) => void)[] = [];
  private readonly warned = new Set<ReadWarningCode>();

  constructor(context: WorksheetReadContext, options: ParserOptions) {
    this.context = context;
    this.tokenizer = new XmlTokenizer(this, { maxDepth: context.maxXmlDepth, maxTextLength: context.maxTextLength });
    this.serialDates = context.dates === 'serial';
    this.dateFields = context.dates === 'utc' ? 'utc' : 'local';
    this.errorMode = options.errorMode;
    this.captureFormulas = options.captureFormulas;
    this.maxColumns = options.maxColumns;
  }

  // -------------------------------------------------------------------------------------------------------------
  // Driving the parse
  // -------------------------------------------------------------------------------------------------------------

  /**
   * `</c>` is handled synchronously inside a tokenizer callback, so the shared-string table has to be in hand before
   * a chunk that needs it is pushed. One `indexOf` pass over the chunk decides, which keeps sheets that never
   * reference the table (inline strings only) from loading it at all. A package with no shared-strings part at all
   * skips even that pass: scanning every chunk of a large inline-string sheet for a table that does not exist cost
   * about 16% of the read (`research/06-performance-baseline.md`).
   */
  async prepareSharedStrings(text: string): Promise<void> {
    if (this.sharedStrings !== undefined || !this.context.hasSharedStrings) {
      return;
    }
    const boundary = this.sharedStringsScanTail;
    this.sharedStringsScanTail = text.length > SHARED_TYPE_SCAN_TAIL ? text.slice(-SHARED_TYPE_SCAN_TAIL) : text;
    if ((boundary.length > 0 && hasSharedType(boundary + text.slice(0, SHARED_TYPE_SCAN_TAIL))) || hasSharedType(text)) {
      await this.loadSharedStrings();
    }
  }

  push(text: string): void {
    this.tokenizer.push(text);
  }

  /** Fill in cells whose `t="s"` the pre-scan missed (an exotic attribute spelling); normally a no-op. */
  async resolveDeferredCells(): Promise<void> {
    if (this.deferredCells.length === 0) {
      return;
    }
    const strings = await this.loadSharedStrings();
    for (const apply of this.deferredCells) {
      apply(strings);
    }
    this.deferredCells.length = 0;
  }

  /** End of input: flush trailing text and let the tokenizer report unclosed elements. */
  finish(): void {
    this.tokenizer.end();
  }

  /** Want no more input. Dropping the capture keeps the rest of the chunk in flight from piling into a buffer. */
  protected stop(): void {
    this.stopped = true;
    this.capture = CAPTURE_NONE;
  }

  // -------------------------------------------------------------------------------------------------------------
  // Tokenizer handler
  // -------------------------------------------------------------------------------------------------------------

  start(name: string): void {
    if (this.stopped) {
      return;
    }
    switch (name) {
      case 'c':
        if (this.collecting && this.state === STATE_ROW) {
          this.beginCell();
        }
        return;
      case 'v':
        if (this.state === STATE_CELL) {
          this.valueSeen = true;
          this.capture = CAPTURE_VALUE;
        }
        return;
      case 't':
        if (this.inInlineString && !this.inPhonetic) {
          this.capture = CAPTURE_INLINE;
        }
        return;
      case 'f':
        if (this.state === STATE_CELL && this.captureFormulas) {
          this.capture = CAPTURE_FORMULA;
        }
        return;
      case 'is':
        if (this.state === STATE_CELL) {
          this.inInlineString = true;
          this.inlineSeen = true;
        }
        return;
      case 'row':
        if (this.state === STATE_SHEET_DATA) {
          this.beginRow();
        }
        return;
      case 'rPh':
        // Phonetic runs are furigana, not the cell's text (EC-RICH-TEXT-RUNS).
        this.inPhonetic = true;
        return;
      case 'sheetData':
        this.state = STATE_SHEET_DATA;
        return;
      default:
        return;
    }
  }

  text(text: string): void {
    switch (this.capture) {
      case CAPTURE_VALUE:
        this.valueText = this.valueText.length === 0 ? text : this.valueText + text;
        return;
      case CAPTURE_INLINE:
        this.inlineText = this.inlineText.length === 0 ? text : this.inlineText + text;
        return;
      case CAPTURE_FORMULA:
        this.formulaText = this.formulaText.length === 0 ? text : this.formulaText + text;
        return;
      default:
        return;
    }
  }

  end(name: string): void {
    if (this.stopped) {
      return;
    }
    switch (name) {
      case 'c':
        if (this.state === STATE_CELL) {
          this.state = STATE_ROW;
          this.finishCell();
        }
        return;
      case 'v':
      case 't':
      case 'f':
        this.capture = CAPTURE_NONE;
        return;
      case 'is':
        this.inInlineString = false;
        return;
      case 'row':
        if (this.state === STATE_ROW) {
          this.state = STATE_SHEET_DATA;
          if (this.collecting) {
            this.onRowEnd(this.rowNumber);
          }
        }
        return;
      case 'rPh':
        this.inPhonetic = false;
        return;
      case 'sheetData':
        // A sheet may hold several `sheetData` elements; row numbering simply continues through them.
        this.state = STATE_OUTSIDE;
        return;
      default:
        return;
    }
  }

  // -------------------------------------------------------------------------------------------------------------
  // Rows and cells
  // -------------------------------------------------------------------------------------------------------------

  /** A missing `<row r>` means the previous row plus one (EC-MISSING-R-ATTRIBUTES). */
  private beginRow(): void {
    const ref = this.tokenizer.attr('r');
    let rowNumber = this.rowNumber + 1;
    if (ref !== undefined) {
      const parsed = +ref;
      if (Number.isInteger(parsed) && parsed >= 1) {
        rowNumber = parsed;
      }
    }
    if (rowNumber > MAX_ROWS) {
      throw limitExceeded(`This sheet has a row numbered ${rowNumber}, past the ${MAX_ROWS} rows a sheet can hold.`, {
        rowNumber,
        maxRows: MAX_ROWS,
      });
    }
    this.rowNumber = rowNumber;
    this.column = -1;
    this.state = STATE_ROW;
    this.onRowStart(rowNumber);
  }

  /**
   * A missing `<c r>` means the next column in this row. A present one supplies the column only: the row came from
   * `<row r>`, so a reference whose digits disagree with it does not move the cell.
   */
  private beginCell(): void {
    const ref = this.tokenizer.attr('r');
    let column = this.column + 1;
    if (ref !== undefined) {
      const parsed = columnIndexOfRef(ref);
      if (parsed >= 0) {
        column = parsed;
      } else {
        // Past XFD is a real position, just not one a sheet can hold: it fails the column check below.
        const pastGrid = columnIndexPastGrid(ref);
        if (pastGrid >= 0) {
          column = pastGrid;
        }
      }
    }
    if (column >= this.maxColumns) {
      throw limitExceeded(`This sheet has cells past column ${this.maxColumns}, more columns than this read allows.`, {
        column: column + 1,
        maxColumns: this.maxColumns,
      });
    }
    this.column = column;
    const type = this.tokenizer.attr('t');
    this.type = type === undefined ? TYPE_NUMBER : cellTypeOf(type);
    const style = this.tokenizer.attr('s');
    if (style === undefined) {
      this.style = 0;
    } else {
      const parsed = +style;
      this.style = Number.isInteger(parsed) && parsed >= 0 ? parsed : 0;
    }
    this.valueSeen = false;
    this.valueText = '';
    this.inlineSeen = false;
    this.inlineText = '';
    this.formulaText = '';
    this.state = STATE_CELL;
  }

  private finishCell(): void {
    if (this.type === TYPE_SHARED) {
      this.finishSharedStringCell();
      return;
    }
    this.onCell(this.cellValue());
  }

  /** `t="s"` is the one type whose value may not be resolvable yet; see `prepareSharedStrings`. */
  private finishSharedStringCell(): void {
    if (this.valueText.length === 0) {
      this.onCell(null);
      return;
    }
    const index = +this.valueText;
    const strings = this.sharedStrings;
    if (strings === undefined) {
      this.onDeferredSharedString(index);
      return;
    }
    this.onCell(this.sharedStringAt(strings, index, this.rowNumber, this.column));
  }

  private cellValue(): CellValue {
    switch (this.type) {
      case TYPE_NUMBER:
        return this.numberValue();
      case TYPE_STRING:
        // An explicitly string-typed cell with an empty `<v>` holds the empty string, not a blank
        // (EC-BLANK-VS-EMPTY-STRING); no `<v>` at all is blank.
        return this.valueSeen ? decodeIfEscaped(this.valueText) : null;
      case TYPE_INLINE:
        return this.inlineValue();
      case TYPE_BOOLEAN:
        return this.valueText.length === 0 ? null : this.valueText === '1' || this.valueText === 'true';
      case TYPE_ERROR:
        return this.errorValue();
      default:
        return this.isoDateValue();
    }
  }

  /**
   * `<v/>`, a `<v>` of only whitespace and a missing `<v>` are all blank, never 0 (EC-EMPTY-V-ELEMENT). A value that
   * is not a finite decimal number comes back as its raw text: Numbers writes `<v>inf</v>` for the max double
   * (EC-NUMBERS-INF-VALUE), and both dropping the cell and returning Infinity would lose what the file actually said.
   * `0x1A` and the other radix spellings unary `+` accepts are text for the same reason (EC-NUMBER-RADIX-PREFIX).
   */
  private numberValue(): CellValue {
    const text = this.valueText;
    if (text.length === 0) {
      return null;
    }
    const parsed = +text;
    if (!Number.isFinite(parsed)) {
      return text;
    }
    const first = text.charCodeAt(0);
    if (first === CODE_ZERO || first <= CODE_SPACE) {
      const kind = leadingNumericKind(text);
      if (kind === NUMERIC_BLANK) {
        return null;
      }
      if (kind === NUMERIC_RADIX) {
        return text;
      }
    }
    // -0 has no spreadsheet representation (EC-NUM-NEGATIVE-ZERO).
    const value = parsed === 0 ? 0 : parsed;
    if (this.serialDates || this.context.isDateByXf[this.style] !== 1) {
      return value;
    }
    const components = componentsFromSerial(value, this.context.date1904);
    // Negative serials are not dates at all (Excel shows ########), so the number is the honest answer.
    return components === null ? value : dateFromComponents(components, this.dateFields);
  }

  /** `<is>` holds either one `<t>` or rich-text runs, which flatten into one string (EC-RICH-TEXT-RUNS). */
  private inlineValue(): CellValue {
    if (this.inlineSeen) {
      return decodeIfEscaped(this.inlineText);
    }
    return this.valueSeen ? decodeIfEscaped(this.valueText) : null;
  }

  private errorValue(): CellValue {
    const text = this.valueText;
    if (text.length === 0 || this.errorMode === 'null') {
      return null;
    }
    if (this.errorMode === 'string') {
      return text;
    }
    // `errors: 'object'` puts a CellError where a CellValue sits; the facade widens the row element type to
    // ReadValue for callers, which is the only place the two views meet.
    return { error: text as CellErrorCode } as unknown as CellValue;
  }

  /** Strict `t="d"`: ISO 8601 with any number of fraction digits, or a bare time (EC-STRICT-ISO-DATE-PRECISION). */
  private isoDateValue(): CellValue {
    const text = this.valueText;
    if (text.length === 0) {
      return null;
    }
    const components = componentsFromIso(text);
    if (components === null) {
      return text;
    }
    if (this.serialDates) {
      const serial = serialFromComponents(components, this.context.date1904);
      return serial === null ? text : serial;
    }
    return dateFromComponents(components, this.dateFields);
  }

  /**
   * Out of range is what Excel itself tolerates: the cell reads as empty text (EC-SST-INDEX-OUT-OF-RANGE), and so
   * does every `t="s"` cell of a package with no table at all (EC-SST-ABSENT-INLINE-ONLY). `rowNumber` and `column`
   * are the cell's own, which a deferred cell no longer shares with the parser.
   */
  protected sharedStringAt(strings: readonly string[], index: number, rowNumber: number, column: number): string {
    const value = strings[index];
    if (value === undefined) {
      this.warnSharedStringMiss(index, strings.length, rowNumber, column);
      return '';
    }
    return value;
  }

  protected defer(apply: (strings: readonly string[]) => void): void {
    this.deferredCells.push(apply);
  }

  protected currentRef(): string {
    return formatRef(this.rowNumber - 1, this.column);
  }

  /** Only ever reached for a cell that missed the table, so the message is built at most once per code. */
  private warnSharedStringMiss(index: number, tableSize: number, rowNumber: number, column: number): void {
    const report = this.context.onWarning;
    if (report === undefined) {
      return;
    }
    const code: ReadWarningCode = this.context.hasSharedStrings ? 'SST_INDEX_OUT_OF_RANGE' : 'SHARED_STRINGS_MISSING';
    if (this.warned.has(code)) {
      return;
    }
    this.warned.add(code);
    const ref = formatRef(rowNumber - 1, column);
    const message =
      code === 'SST_INDEX_OUT_OF_RANGE'
        ? `Cell ${ref} refers to shared string ${index}, but the table holds ${tableSize}; it reads as empty text, as in Excel.`
        : `Cell ${ref} refers to the shared-string table, but this workbook has none; it reads as empty text.`;
    report({ code, message, ref });
  }

  private async loadSharedStrings(): Promise<readonly string[]> {
    this.sharedStrings ??= await this.context.sharedStrings();
    return this.sharedStrings;
  }

  protected abstract onRowStart(rowNumber: number): void;
  protected abstract onRowEnd(rowNumber: number): void;
  protected abstract onCell(value: CellValue): void;
  protected abstract onDeferredSharedString(index: number): void;
}

/** Dense rows for `rows()`: one array per row, holes and trailing empties resolved at `</row>`. */
class RowsParser extends WorksheetParser {
  private readonly startRow: number;
  private readonly maxRows: number;
  private readonly blankRows: boolean;
  private readonly formulaAsValue: boolean;
  private readonly errorTextRow: number | 'first' | undefined;

  private cells: CellValue[] = [];
  private batch: SheetRow[] = [];
  private emitted = 0;
  /** No row with a value has been emitted yet: with `errorTextRow: 'first'`, the next one is the header row. */
  private awaitingFirstRow = true;
  /** With `blankRows`, the next row number that must appear so `rows[i].index === i + startRow` keeps holding. */
  private nextIndex: number;

  constructor(context: WorksheetReadContext, options: WorksheetRowsOptions) {
    super(context, {
      captureFormulas: options.formulas === 'text',
      // No sheet has a column past XFD, whatever cap the caller allows.
      maxColumns: Math.min(options.maxColumns ?? MAX_COLUMNS, MAX_COLUMNS),
      errorMode: context.errors,
    });
    this.startRow = Math.max(1, options.startRow ?? 1);
    this.maxRows = options.maxRows ?? Number.POSITIVE_INFINITY;
    this.blankRows = options.blankRows === true;
    this.formulaAsValue = options.formulas === 'text';
    this.errorTextRow = options.errorTextRow;
    this.nextIndex = this.startRow;
    if (this.maxRows < 1) {
      this.stop();
    }
  }

  takeBatch(): SheetRow[] {
    const batch = this.batch;
    this.batch = [];
    return batch;
  }

  protected override onRowStart(rowNumber: number): void {
    this.collecting = rowNumber >= this.startRow;
    if (this.collecting) {
      this.cells = [];
      if (this.errorTextRow !== undefined) {
        const isHeaderRow = this.errorTextRow === 'first' ? this.awaitingFirstRow : rowNumber === this.errorTextRow;
        this.errorMode = isHeaderRow ? 'string' : this.context.errors;
      }
    }
  }

  protected override onRowEnd(rowNumber: number): void {
    const cells = this.cells;
    let end = cells.length;
    while (end > 0 && cells[end - 1] === null) {
      end--;
    }
    cells.length = end;
    if (end > 0) {
      this.awaitingFirstRow = false;
    }
    if (end === 0 && !this.blankRows) {
      return;
    }
    if (this.blankRows) {
      // Row numbers can jump; the blank rows in between keep the yielded sequence aligned with the sheet.
      for (let index = this.nextIndex; index < rowNumber; index++) {
        this.emit({ index, cells: [] });
        if (this.stopped) {
          return;
        }
      }
    }
    this.emit({ index: rowNumber, cells });
  }

  protected override onCell(value: CellValue): void {
    const formula = this.formulaOverride();
    if (formula !== null) {
      this.place(formula);
      return;
    }
    if (value !== null) {
      this.place(value);
    }
  }

  protected override onDeferredSharedString(index: number): void {
    const formula = this.formulaOverride();
    if (formula !== null) {
      this.place(formula);
      return;
    }
    const cells = this.cells;
    const rowNumber = this.rowNumber;
    const column = this.column;
    // The placeholder must not be null, or the trailing-empty trim could drop the slot before the fixup runs.
    this.place('');
    this.defer(strings => {
      cells[column] = this.sharedStringAt(strings, index, rowNumber, column);
    });
  }

  /**
   * `formulas: 'text'` reports what the sheet computes, whatever type the cached value has. Shared-formula followers
   * (`<f t="shared" si="0"/>`) carry no text of their own, so they keep their cached value.
   */
  private formulaOverride(): string | null {
    return this.formulaAsValue && this.formulaText.length > 0 ? `=${this.formulaText}` : null;
  }

  private place(value: CellValue): void {
    const cells = this.cells;
    const column = this.column;
    if (column < cells.length) {
      cells[column] = value;
      return;
    }
    while (cells.length < column) {
      cells.push(null);
    }
    cells.push(value);
  }

  private emit(row: SheetRow): void {
    this.batch.push(row);
    this.emitted++;
    this.nextIndex = row.index + 1;
    if (this.emitted >= this.maxRows) {
      this.stop();
    }
  }
}

/** A1-keyed raw cells for `head()`: the cached value, plus the formula text whenever the cell has one. */
class HeadParser extends WorksheetParser {
  readonly cells: Map<string, RawCell> = new Map();

  private readonly rowCount: number;

  constructor(context: WorksheetReadContext, rowCount: number) {
    super(context, { captureFormulas: true, maxColumns: MAX_COLUMNS, errorMode: 'string' });
    this.rowCount = rowCount;
    if (rowCount < 1) {
      this.stop();
    }
  }

  protected override onRowStart(rowNumber: number): void {
    this.collecting = rowNumber <= this.rowCount;
    if (!this.collecting) {
      // Rows ascend, so the first row past the head is the end of the read.
      this.stop();
    }
  }

  protected override onRowEnd(): void {
    // Nothing to close: `head()` collects cells, not rows.
  }

  protected override onCell(value: CellValue): void {
    const formula = this.formulaText.length === 0 ? undefined : this.formulaText;
    if (value === null && formula === undefined) {
      return;
    }
    // A RawCell reports an error cell as its text plus the code, whatever `errors` the workbook was opened with.
    const error = this.type === TYPE_ERROR && typeof value === 'string' ? (value as CellErrorCode) : undefined;
    this.cells.set(this.currentRef(), rawCell(value, error, formula));
  }

  protected override onDeferredSharedString(index: number): void {
    const ref = this.currentRef();
    const rowNumber = this.rowNumber;
    const column = this.column;
    const formula = this.formulaText.length === 0 ? undefined : this.formulaText;
    this.defer(strings => {
      this.cells.set(ref, rawCell(this.sharedStringAt(strings, index, rowNumber, column), undefined, formula));
    });
  }
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
  options: WorksheetRowsOptions,
): AsyncIterable<SheetRow> {
  return streamRows(chunks, context, options);
}

async function* streamRows(
  chunks: AsyncIterable<Uint8Array>,
  context: WorksheetReadContext,
  options: WorksheetRowsOptions,
): AsyncGenerator<SheetRow, void, undefined> {
  const parser = new RowsParser(context, options);
  if (parser.stopped) {
    return;
  }
  const decoder = new TextDecoder('utf-8');
  for await (const chunk of chunks) {
    throwIfAborted(context.signal);
    await feed(parser, decoder.decode(chunk, { stream: true }));
    for (const row of parser.takeBatch()) {
      yield row;
    }
    if (parser.stopped) {
      return;
    }
  }
  await feed(parser, decoder.decode());
  parser.finish();
  for (const row of parser.takeBatch()) {
    yield row;
  }
}

/**
 * First `rowCount` rows as A1-keyed raw cells (the `head()` fast path); stops reading after the last wanted row, so
 * a preview of a million-row sheet inflates only its first chunks.
 *
 * Values follow the context's `dates` (the facade's default is `'local'`); formula text is always reported, and an
 * error cell always reads as its text plus its code, whatever `errors` the workbook was opened with. Cells that are
 * blank and carry no formula are left out of the map entirely.
 */
export async function readWorksheetHead(
  chunks: AsyncIterable<Uint8Array>,
  context: WorksheetReadContext,
  rowCount: number,
): Promise<Map<string, RawCell>> {
  const parser = new HeadParser(context, rowCount);
  if (parser.stopped) {
    return parser.cells;
  }
  const decoder = new TextDecoder('utf-8');
  for await (const chunk of chunks) {
    throwIfAborted(context.signal);
    await feed(parser, decoder.decode(chunk, { stream: true }));
    if (parser.stopped) {
      return parser.cells;
    }
  }
  await feed(parser, decoder.decode());
  parser.finish();
  return parser.cells;
}

async function feed(parser: WorksheetParser, text: string): Promise<void> {
  if (text.length === 0) {
    return;
  }
  await parser.prepareSharedStrings(text);
  parser.push(text);
  await parser.resolveDeferredCells();
}
