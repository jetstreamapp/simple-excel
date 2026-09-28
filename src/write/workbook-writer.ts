import { createDeflater, hasNativeDeflate } from '../compress/deflater';
import { XlsxError } from '../errors';
import { fromWritableStream, guardSink } from '../sinks';
import { appXml, contentTypesXml, coreXml, rootRelsXml, type SheetPartInfo, workbookRelsXml, workbookXml } from '../sml/package-parts';
import { SharedStringWriter } from '../sml/shared-strings';
import { sanitizeSheetName } from '../sml/sheet-name';
import { StyleRegistry } from '../sml/styles';
import { MergedRanges, NOTHING_TO_FLUSH, sheetRange, validateSheetOptions, WorksheetWriter } from '../sml/worksheet-writer';
import type {
  ByteSink,
  CellInput,
  CellStyle,
  SharedStringBudget,
  SheetOptions,
  SheetWriter,
  SheetWriteSummary,
  StyleId,
  WorkbookProperties,
  WorkbookWriter,
  WorkbookWriteResult,
  WorkbookWriterOptions,
  WriteProgress,
} from '../types';
import { ZipWriter } from '../zip/zip-writer';

const ENCODER = new TextEncoder();

/** ADR-001: the hybrid table stays small enough to be worth keeping and never unbounded. */
const DEFAULT_SHARED_STRING_BUDGET: Required<SharedStringBudget> = {
  maxUnique: 65_536,
  maxChars: 16 * 1024 * 1024,
  maxLength: 256,
};
/** `strings: 'shared'` interns everything; the budget is still the only knob, so it is simply switched off. */
const UNBOUNDED_SHARED_STRING_BUDGET: Required<SharedStringBudget> = {
  maxUnique: Number.POSITIVE_INFINITY,
  maxChars: Number.POSITIVE_INFINITY,
  maxLength: Number.POSITIVE_INFINITY,
};

/** The timestamp deterministic mode stamps into docProps (the zip entries use the writer's own fixed DOS date). */
const DETERMINISTIC_NOW = new Date('2026-01-01T00:00:00.000Z');
/** `dcterms:created` is an xsd:dateTime with a four-digit year; openpyxl and the validator refuse anything else. */
const FIRST_PROPERTY_YEAR = 1;
const LAST_PROPERTY_YEAR = 9999;
const PROGRESS_ROW_INTERVAL = 5000;
/** Cells above which a worksheet part is assumed to risk the 4 GiB zip32 ceiling and gets a zip64 local header. */
const ZIP64_CELL_THRESHOLD = 40_000_000;
/** Width assumed for a sheet that announced neither a header nor column definitions. */
const ASSUMED_COLUMN_COUNT = 20;

interface ResolvedOptions {
  readonly strings: 'auto' | 'inline' | 'shared';
  readonly budget: Required<SharedStringBudget>;
  readonly zip64: 'auto' | boolean;
  readonly deterministic: boolean;
  readonly dates: 'local' | 'utc';
  readonly date1904: boolean;
  readonly cellOverflow: 'truncate' | 'throw';
  readonly truncationSuffix: string;
  readonly onCellTruncated: ((count: number) => void) | undefined;
  readonly onProgress: ((progress: WriteProgress) => void) | undefined;
  readonly signal: AbortSignal | undefined;
  readonly properties: { readonly creator: string; readonly title?: string; readonly created?: Date };
}

interface SheetRecord {
  readonly name: string;
  readonly sheetId: number;
  readonly relId: string;
  /** Part name relative to `xl/`. */
  readonly path: string;
  readonly hidden: boolean;
  autoFilterRange: string | undefined;
  summary: SheetWriteSummary | undefined;
}

function encode(xml: string): Uint8Array {
  return ENCODER.encode(xml);
}

function isWritableStream(sink: ByteSink | WritableStream<Uint8Array>): sink is WritableStream<Uint8Array> {
  return typeof (sink as WritableStream<Uint8Array>).getWriter === 'function';
}

function toSheetPartInfo(record: SheetRecord): SheetPartInfo {
  return {
    name: record.name,
    sheetId: record.sheetId,
    relId: record.relId,
    path: record.path,
    hidden: record.hidden,
    ...(record.autoFilterRange === undefined ? {} : { autoFilterRange: record.autoFilterRange }),
  };
}

/**
 * `properties.created` becomes `dcterms:created`, an xsd:dateTime with a four-digit year (EC-DOCPROPS-CREATED-RANGE).
 * An Invalid Date or a year outside 1-9999 is refused before any byte is written rather than producing a
 * `docProps/core.xml` that openpyxl rejects.
 */
function validateCreated(created: WorkbookProperties['created']): void {
  if (created === undefined) {
    return;
  }
  // The tag, not `instanceof`: a Date from another realm (an iframe, a vm context, an Electron bridge) is still a Date.
  const year = Object.prototype.toString.call(created) === '[object Date]' ? created.getUTCFullYear() : Number.NaN;
  if (!(year >= FIRST_PROPERTY_YEAR && year <= LAST_PROPERTY_YEAR)) {
    throw new XlsxError(
      'WRITER_STATE',
      `properties.created must be a valid Date between the years ${FIRST_PROPERTY_YEAR} and ${LAST_PROPERTY_YEAR}; got ${String(created)}.`,
      { created },
    );
  }
}

function resolveOptions(options: WorkbookWriterOptions): ResolvedOptions {
  validateCreated(options.properties?.created);
  const strings = options.strings ?? 'inline';
  const budget =
    strings === 'shared'
      ? UNBOUNDED_SHARED_STRING_BUDGET
      : {
          maxUnique: options.sstBudget?.maxUnique ?? DEFAULT_SHARED_STRING_BUDGET.maxUnique,
          maxChars: options.sstBudget?.maxChars ?? DEFAULT_SHARED_STRING_BUDGET.maxChars,
          maxLength: options.sstBudget?.maxLength ?? DEFAULT_SHARED_STRING_BUDGET.maxLength,
        };
  return {
    strings,
    budget,
    zip64: options.zip64 ?? 'auto',
    deterministic: options.deterministic === true,
    dates: options.dates ?? 'local',
    date1904: options.date1904 === true,
    cellOverflow: options.cellOverflow ?? 'truncate',
    truncationSuffix: options.truncationSuffix ?? '...(truncated)',
    onCellTruncated: options.onCellTruncated,
    onProgress: options.onProgress,
    signal: options.signal,
    properties: {
      creator: options.properties?.creator ?? 'simple-excel',
      ...(options.properties?.title === undefined ? {} : { title: options.properties.title }),
      ...(options.properties?.created === undefined ? {} : { created: options.properties.created }),
    },
  };
}

/**
 * A copy of the options a sheet keeps using after `addSheet` returns (EC-ROW-SNAPSHOT): the header row and the
 * column definitions are only written once the zip entry opens, so a caller mutating them in between must not change
 * what lands in the file.
 */
function snapshotSheetOptions(options: SheetOptions): SheetOptions {
  return {
    ...options,
    ...(options.header === undefined ? {} : { header: options.header.slice() }),
    ...(options.columns === undefined ? {} : { columns: options.columns.map(column => ({ ...column })) }),
    ...(options.freeze === undefined ? {} : { freeze: { ...options.freeze } }),
  };
}

/**
 * Whether a worksheet part needs zip64 declared in its local header. Excel repairs an archive whose zip64 is only in
 * the central directory (ADR-002), so the decision has to be made before the first byte of the entry. `'auto'` is
 * conservative: only a sheet whose announced size is plainly huge gets zip64, because SheetJS cannot read zip64
 * archives at all and Excel is picky about their shape. A sheet with an unknown row count stays 32-bit and the
 * writer fails with `ENTRY_TOO_LARGE` (naming `zip64: true` as the fix) in the unlikely case it passes 4 GiB.
 */
function sheetNeedsZip64(mode: 'auto' | boolean, options: SheetOptions): boolean {
  if (mode !== 'auto') {
    return mode;
  }
  if (options.rowCount === undefined) {
    return false;
  }
  const columns = options.header?.length ?? options.columns?.length ?? ASSUMED_COLUMN_COUNT;
  return (options.rowCount + 1) * columns > ZIP64_CELL_THRESHOLD;
}

class WorkbookWriterImpl implements WorkbookWriter {
  readonly zip: ZipWriter;
  readonly styles: StyleRegistry = new StyleRegistry();
  readonly sharedStrings: SharedStringWriter | null;
  readonly options: ResolvedOptions;

  private readonly sheets: SheetRecord[] = [];
  private readonly takenNames = new Set<string>();
  /** Serializes every part: one zip entry is open at a time and the parts must land in the order they were queued. */
  private tail: Promise<void>;
  private openSheet: SheetWriterImpl | undefined;
  private hasVisibleSheet = false;
  private anySheetNeedsZip64 = false;
  private truncatedCells = 0;
  private closed = false;
  /**
   * Latched by the first failure - a sink or deflater error, a refused row, a part that could not be written - or by
   * an abort (EC-WRITER-FAILURE-STICKY). From then on every call rejects with `failure` itself, never with a derived
   * error from the half-written archive.
   */
  private failed = false;
  private failure: unknown;
  /** The one abort the sink ever receives; failure paths wait on it so the sink is torn down before they reject. */
  private sinkAbort: Promise<void> | undefined;

  constructor(sink: ByteSink | WritableStream<Uint8Array>, options: WorkbookWriterOptions) {
    // Validated before the sink is touched, so a refused option leaves a caller's stream unlocked.
    this.options = resolveOptions(options);
    this.sharedStrings = this.options.strings === 'inline' ? null : new SharedStringWriter(this.options.budget);

    const byteSink = guardSink(isWritableStream(sink) ? fromWritableStream(sink) : sink);
    // Without a custom deflater and without CompressionStream the platform deflater silently stores its input, which
    // would leave every local header claiming method 8 over stored bytes.
    const usesNativeDeflate = options.deflater === undefined;
    this.zip = new ZipWriter(byteSink, {
      zip64: this.options.zip64 === true,
      deterministic: this.options.deterministic,
      deflater: options.deflater ?? createDeflater,
      compression: usesNativeDeflate && !hasNativeDeflate() ? 'store' : (options.compression ?? 'deflate'),
    });

    const now = this.options.deterministic ? DETERMINISTIC_NOW : new Date();
    this.tail = Promise.resolve();
    // The package-level relationship and the core properties are fixed, so they lead the archive; the sheet list is
    // only known at close, which is why `[Content_Types].xml` trails it (EC-ZIP-CONTENT-TYPES-LAST). Nobody awaits
    // this step, so a failure here is latched by `enqueue` and reported by the next call, `close()` at the latest.
    void this.enqueue(async () => {
      await this.zip.writeEntry('_rels/.rels', encode(rootRelsXml()));
      await this.zip.writeEntry('docProps/core.xml', encode(coreXml(this.options.properties, now)));
    });
  }

  addSheet(name: string, options: SheetOptions = {}): SheetWriter {
    this.assertOpen();
    if (this.openSheet !== undefined) {
      throw new XlsxError(
        'WRITER_STATE',
        `The sheet "${this.openSheet.name}" is still open. Close it before adding "${name}"; one sheet is written at a time.`,
        { openSheet: this.openSheet.name },
      );
    }
    // Everything that can refuse the sheet runs before any state changes, so a refused addSheet can be retried.
    validateSheetOptions(options, this.styles.count);
    const index = this.sheets.length + 1;
    const record: SheetRecord = {
      name: sanitizeSheetName(name, this.takenNames),
      sheetId: index,
      relId: `rId${index}`,
      path: `worksheets/sheet${index}.xml`,
      hidden: options.hidden === true,
      autoFilterRange: undefined,
      summary: undefined,
    };
    // Sheet XML streams out before later sheets exist, so the selected tab is decided now: the first visible sheet
    // gets `tabSelected`, the same one `workbookXml` makes the active tab (EC-ALL-SHEETS-HIDDEN).
    const tabSelected = !record.hidden && !this.hasVisibleSheet;
    this.hasVisibleSheet ||= !record.hidden;
    this.sheets.push(record);

    const sheet = new SheetWriterImpl(this, record, snapshotSheetOptions(options), tabSelected);
    this.openSheet = sheet;
    return sheet;
  }

  registerStyle(style: CellStyle): StyleId {
    this.assertOpen();
    return this.styles.register(style);
  }

  async close(): Promise<WorkbookWriteResult> {
    this.assertOpen();
    if (this.openSheet !== undefined) {
      throw new XlsxError('WRITER_STATE', `The sheet "${this.openSheet.name}" is still open. Close it before closing the workbook.`, {
        openSheet: this.openSheet.name,
      });
    }
    this.closed = true;
    // A package with no sheet fails the validator and SheetJS (EC-WORKBOOK-NO-SHEETS); one whose sheets are all
    // hidden makes Excel repair it (primer section 12.10, EC-ALL-SHEETS-HIDDEN). Either way the file is not written.
    if (this.sheets.length === 0) {
      throw await this.fail(new XlsxError('WRITER_STATE', 'A workbook needs at least one sheet. Add one with addSheet() before close().'));
    }
    if (!this.hasVisibleSheet) {
      throw await this.fail(
        new XlsxError('WRITER_STATE', 'Every sheet in this workbook is hidden, which Excel refuses. Leave at least one sheet visible.', {
          sheets: this.sheets.map(record => record.name),
        }),
      );
    }

    const result = await this.enqueue(async () => {
      const stats = this.sharedStrings?.stats ?? { count: 0, uniqueCount: 0, frozen: false };
      const hasSharedStrings = this.sharedStrings !== null && stats.uniqueCount > 0;
      if (hasSharedStrings && this.sharedStrings !== null) {
        const entry = await this.zip.beginEntry('xl/sharedStrings.xml', { zip64: this.anySheetNeedsZip64 });
        await this.sharedStrings.writeTo(entry);
        await entry.close();
      }

      const parts = this.sheets.map(record => toSheetPartInfo(record));
      await this.zip.writeEntry('xl/styles.xml', encode(this.styles.toXml()));
      await this.zip.writeEntry('xl/workbook.xml', encode(workbookXml(parts, this.options.date1904)));
      await this.zip.writeEntry('xl/_rels/workbook.xml.rels', encode(workbookRelsXml(parts, hasSharedStrings)));
      await this.zip.writeEntry('docProps/app.xml', encode(appXml(parts)));
      await this.zip.writeEntry('[Content_Types].xml', encode(contentTypesXml(parts, hasSharedStrings)));
      const archive = await this.zip.close();
      return {
        bytes: archive.bytes,
        sheets: this.sheets.map(record => record.summary ?? { name: record.name, rows: 0, columns: 0 }),
        truncatedCells: this.truncatedCells,
        sharedStrings: stats,
      };
    });
    // Once, with the workbook total, and only after the file is complete (EC-CELL-32767-LIMIT).
    if (result.truncatedCells > 0) {
      this.options.onCellTruncated?.(result.truncatedCells);
    }
    return result;
  }

  async abort(reason?: unknown): Promise<void> {
    if (!this.failed) {
      this.latch(new XlsxError('ABORTED', 'Writing the spreadsheet was cancelled.', reason === undefined ? undefined : { reason }));
    }
    await this.abortSink(reason);
  }

  /**
   * Queue one part-producing step; steps run in the order they were queued and never overlap. A step queued behind a
   * failure rejects with that failure without running, and a step that fails latches its error for everyone after.
   */
  enqueue<T>(operation: () => Promise<T>): Promise<T> {
    const running = this.tail.then(async () => {
      this.assertUsable();
      try {
        return await operation();
      } catch (reason) {
        throw await this.fail(reason);
      }
    });
    // The chain itself must stay resolvable so a failed step does not strand the ones behind it; the failure is
    // reported to whoever holds `running` and, through the latch, to every later call.
    this.tail = running.then(
      () => undefined,
      () => undefined,
    );
    return running;
  }

  /**
   * Latch `reason` as the workbook's failure (unless an earlier one already is), abort the sink once, and resolve to
   * the latched failure once the sink has been told. Callers `throw await this.fail(reason)`.
   */
  async fail(reason: unknown): Promise<unknown> {
    if (!this.failed) {
      this.latch(reason);
    }
    await this.abortSink(this.failure);
    return this.failure;
  }

  /**
   * `fail` for the row paths: latch and start the sink abort, but reject at once. The abort is queued behind any
   * write the sink has not accepted yet, and a sink stalled on back-pressure must not keep a refused row pending; the
   * abort reaches the sink when it drains.
   */
  failWith(reason: unknown): Promise<never> {
    if (!this.failed) {
      this.latch(reason);
    }
    void this.abortSink(this.failure);
    return Promise.reject(this.failure);
  }

  /** A sheet stops being the open one the moment `close()` is called, not when its bytes have landed. */
  sheetClosing(): void {
    this.openSheet = undefined;
  }

  sheetClosed(record: SheetRecord, summary: SheetWriteSummary, truncatedCells: number): void {
    record.summary = summary;
    this.truncatedCells += truncatedCells;
  }

  markZip64Sheet(): void {
    this.anySheetNeedsZip64 = true;
  }

  reportProgress(progress: WriteProgress): void {
    this.options.onProgress?.(progress);
  }

  assertOpen(): void {
    this.assertUsable();
    if (this.closed) {
      throw new XlsxError('WRITER_STATE', 'This workbook is closed; nothing more can be written to it.');
    }
  }

  /** Throws the latched failure, latching an abort first when the caller's signal has fired. */
  assertUsable(): void {
    if (this.failed) {
      throw this.failure;
    }
    const { signal } = this.options;
    if (signal?.aborted === true) {
      void this.abort(signal.reason);
      throw this.failure;
    }
  }

  private latch(reason: unknown): void {
    this.failed = true;
    this.failure = reason;
    this.closed = true;
    this.openSheet = undefined;
  }

  private abortSink(reason: unknown): Promise<void> {
    this.sinkAbort ??= this.zip.abort(reason).catch(() => undefined);
    return this.sinkAbort;
  }
}

class SheetWriterImpl implements SheetWriter {
  readonly name: string;

  private readonly workbook: WorkbookWriterImpl;
  private readonly record: SheetRecord;
  private readonly autoFilter: boolean;
  private readonly headerRows: number;
  /** Shared with the worksheet writer, so `merge()` validates now even while the zip entry is still opening. */
  private readonly merges: MergedRanges;
  /** Resolves once the zip entry is open, the prologue is buffered and the header row (if any) is written. */
  private readonly ready: Promise<WorksheetWriter>;
  private worksheet: WorksheetWriter | undefined;
  private rowsIssued = 0;
  private rowsReported = 0;
  private closeRequested = false;

  constructor(workbook: WorkbookWriterImpl, record: SheetRecord, options: SheetOptions, tabSelected: boolean) {
    this.workbook = workbook;
    this.record = record;
    this.name = record.name;
    this.autoFilter = options.autoFilter === true;
    this.headerRows = options.header === undefined ? 0 : 1;
    this.merges = new MergedRanges(record.name);

    const needsZip64 = sheetNeedsZip64(workbook.options.zip64, options);
    if (needsZip64) {
      workbook.markZip64Sheet();
    }
    this.ready = workbook.enqueue(async () => {
      const entry = await workbook.zip.beginEntry(`xl/${record.path}`, { zip64: needsZip64 });
      const worksheet = new WorksheetWriter(
        {
          entry,
          styles: workbook.styles,
          sharedStrings: workbook.sharedStrings,
          date1904: workbook.options.date1904,
          dates: workbook.options.dates,
          cellOverflow: workbook.options.cellOverflow,
          truncationSuffix: workbook.options.truncationSuffix,
          sheetName: record.name,
          tabSelected,
          merges: this.merges,
        },
        options,
      );
      if (options.header !== undefined) {
        const headerStyle = options.headerStyle === false ? undefined : (options.headerStyle ?? workbook.styles.headerStyle);
        await worksheet.writeRow(options.header, headerStyle);
      }
      this.worksheet = worksheet;
      return worksheet;
    });
    // Nobody looks at `ready` until the first row, a close or an abort; marking it handled keeps an abort in between
    // from surfacing as an unhandled rejection. A failure is latched by `enqueue` either way.
    void this.ready.catch(() => undefined);
  }

  /** 1-based index of the next row `writeRow` will produce. */
  get nextRow(): number {
    return this.worksheet?.nextRow ?? 1 + this.headerRows + this.rowsIssued;
  }

  writeRow(values: readonly CellInput[], styles?: StyleId | readonly (StyleId | undefined)[]): Promise<void> {
    try {
      this.assertWritable();
    } catch (reason) {
      return Promise.reject(reason);
    }
    if (!Array.isArray(values)) {
      // An object row (`{ Id, Name }`) would otherwise write an empty row without a word (EC-ROW-NOT-ARRAY).
      return this.workbook.failWith(
        new XlsxError(
          'WRITER_STATE',
          `Sheet "${this.name}" row ${this.nextRow}: a row must be an array of cell values, not a ${values === null ? 'null' : typeof values}. ` +
            'Map each record to an array in column order first.',
          { row: this.nextRow },
        ),
      );
    }
    this.rowsIssued++;
    const worksheet = this.worksheet;
    if (worksheet === undefined) {
      // Queued until the entry opens (bounded: only the rows issued before that). Copy them, so a caller reusing one
      // array without awaiting does not write its last values N times (EC-ROW-SNAPSHOT).
      const snapshot = values.slice();
      const stylesSnapshot = Array.isArray(styles) ? styles.slice() : styles;
      return this.ready.then(opened => {
        this.workbook.assertUsable();
        return this.writeRowTo(opened, snapshot, stylesSnapshot);
      });
    }
    return this.writeRowTo(worksheet, values, styles);
  }

  async writeRows(rows: Iterable<readonly CellInput[]> | AsyncIterable<readonly CellInput[]>): Promise<void> {
    if (Symbol.asyncIterator in rows) {
      for await (const row of rows) {
        await this.writeRow(row);
      }
      return;
    }
    for (const row of rows) {
      await this.writeRow(row);
    }
  }

  merge(range: string): void {
    this.assertWritable();
    this.merges.add(range);
  }

  close(): Promise<SheetWriteSummary> {
    try {
      this.workbook.assertUsable();
    } catch (reason) {
      return Promise.reject(reason);
    }
    if (this.closeRequested) {
      return Promise.reject(new XlsxError('WRITER_STATE', `The sheet "${this.name}" is already closed.`));
    }
    this.closeRequested = true;
    this.workbook.sheetClosing();
    return this.workbook.enqueue(async () => {
      const worksheet = await this.ready;
      const written = await worksheet.close();
      const summary: SheetWriteSummary = { name: this.name, rows: written.rows, columns: written.columns };
      if (this.autoFilter && written.rows > 0 && written.columns > 0) {
        // Excel records the same range as a hidden `_xlnm._FilterDatabase` defined name in workbook.xml.
        this.record.autoFilterRange = sheetRange(written.rows, written.columns);
      }
      this.workbook.sheetClosed(this.record, summary, written.truncatedCells);
      this.reportProgress(written.rows);
      return summary;
    });
  }

  private writeRowTo(
    worksheet: WorksheetWriter,
    values: readonly CellInput[],
    styles: StyleId | readonly (StyleId | undefined)[] | undefined,
  ): Promise<void> {
    const pending = worksheet.writeRow(values, styles);
    // A flush or a refused row: a failure from either has to reach the workbook, which aborts the sink and makes
    // every later call report the same error. Only these paths pay for the extra promise.
    const observed = pending === NOTHING_TO_FLUSH ? pending : pending.then(undefined, (reason: unknown) => this.workbook.failWith(reason));
    const rows = worksheet.nextRow - 1;
    if (rows % PROGRESS_ROW_INTERVAL === 0 && rows > this.rowsReported) {
      this.rowsReported = rows;
      try {
        this.reportProgress(rows);
      } catch (reason) {
        void observed.catch(() => undefined);
        return this.workbook.failWith(reason);
      }
    }
    return observed;
  }

  private reportProgress(rows: number): void {
    this.workbook.reportProgress({ sheet: this.name, rows, bytesOut: this.workbook.zip.bytesWritten });
  }

  private assertWritable(): void {
    this.workbook.assertUsable();
    if (this.closeRequested) {
      throw new XlsxError('WRITER_STATE', `The sheet "${this.name}" is closed; nothing more can be written to it.`);
    }
  }
}

/**
 * Create a streaming workbook writer over a sink (or a `WritableStream<Uint8Array>`). Parts are written in this
 * order: `_rels/.rels`, `docProps/core.xml`, the worksheets (one open at a time), `xl/sharedStrings.xml`,
 * `xl/styles.xml`, `xl/workbook.xml`, `xl/_rels/workbook.xml.rels`, `docProps/app.xml`, `[Content_Types].xml`,
 * central directory. Content types come last because they name every worksheet part, and the sheet list is only
 * final once the caller stops adding sheets; Excel resolves parts through the central directory, and Google Sheets
 * writes its own content types last too (EC-ZIP-CONTENT-TYPES-LAST).
 *
 * Throws `WRITER_STATE` for an invalid `properties.created` before the sink is touched. After any failure the sink
 * is aborted once and every later call rejects with that first error (EC-WRITER-FAILURE-STICKY).
 */
export function createWorkbookWriter(sink: ByteSink | WritableStream<Uint8Array>, options?: WorkbookWriterOptions): WorkbookWriter {
  return new WorkbookWriterImpl(sink, options ?? {});
}
