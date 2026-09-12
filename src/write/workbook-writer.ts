import { createDeflater, hasNativeDeflate } from '../compress/deflater';
import { XlsxError } from '../errors';
import { fromWritableStream, guardSink } from '../sinks';
import { appXml, contentTypesXml, coreXml, rootRelsXml, type SheetPartInfo, workbookRelsXml, workbookXml } from '../sml/package-parts';
import { SharedStringWriter } from '../sml/shared-strings';
import { sanitizeSheetName } from '../sml/sheet-name';
import { StyleRegistry } from '../sml/styles';
import { sheetRange, WorksheetWriter } from '../sml/worksheet-writer';
import type {
  ByteSink,
  CellInput,
  CellStyle,
  SharedStringBudget,
  SheetOptions,
  SheetWriter,
  SheetWriteSummary,
  StyleId,
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

function resolveOptions(options: WorkbookWriterOptions): ResolvedOptions {
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
  private reportedTruncatedCells = 0;
  private closed = false;
  private aborted = false;

  constructor(sink: ByteSink | WritableStream<Uint8Array>, options: WorkbookWriterOptions) {
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
    // only known at close, which is why `[Content_Types].xml` trails it (EC-ZIP-CONTENT-TYPES-LAST).
    this.enqueue(async () => {
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
    const tabSelected = !record.hidden && !this.hasVisibleSheet;
    this.hasVisibleSheet ||= !record.hidden;
    this.sheets.push(record);

    const sheet = new SheetWriterImpl(this, record, options, tabSelected);
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

    return this.enqueue(async () => {
      this.assertNotAborted();
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

      this.reportTruncatedCells();
      return {
        bytes: archive.bytes,
        sheets: this.sheets.map(record => record.summary ?? { name: record.name, rows: 0, columns: 0 }),
        truncatedCells: this.truncatedCells,
        sharedStrings: stats,
      };
    });
  }

  async abort(reason?: unknown): Promise<void> {
    if (this.aborted) {
      return;
    }
    this.aborted = true;
    this.closed = true;
    this.openSheet = undefined;
    await this.zip.abort(reason);
  }

  /** Queue one part-producing step; steps run in the order they were queued and never overlap. */
  enqueue<T>(operation: () => Promise<T>): Promise<T> {
    const running = this.tail.then(operation);
    // The chain itself must stay resolvable so a failed step does not strand the ones behind it; the failure is
    // reported to whoever holds `running`, and the zip writer refuses everything after it anyway.
    this.tail = running.then(
      () => undefined,
      () => undefined,
    );
    return running;
  }

  /** A sheet stops being the open one the moment `close()` is called, not when its bytes have landed. */
  sheetClosing(): void {
    this.openSheet = undefined;
  }

  sheetClosed(record: SheetRecord, summary: SheetWriteSummary): void {
    record.summary = summary;
  }

  /** Every sheet reports its own truncations as it closes; the caller hears the running workbook total. */
  addTruncatedCells(count: number): void {
    this.truncatedCells += count;
    this.reportTruncatedCells();
  }

  markZip64Sheet(): void {
    this.anySheetNeedsZip64 = true;
  }

  reportProgress(progress: WriteProgress): void {
    this.options.onProgress?.(progress);
  }

  assertOpen(): void {
    this.assertNotAborted();
    if (this.closed) {
      throw new XlsxError('WRITER_STATE', 'This workbook is closed; nothing more can be written to it.');
    }
  }

  assertNotAborted(): void {
    if (this.aborted) {
      throw new XlsxError('ABORTED', 'Writing the spreadsheet was cancelled.');
    }
    const { signal } = this.options;
    if (signal?.aborted === true) {
      void this.abort(signal.reason);
      throw new XlsxError('ABORTED', 'Writing the spreadsheet was cancelled.', { reason: signal.reason });
    }
  }

  /**
   * `onCellTruncated` hears the running workbook total as each sheet closes and once more at workbook close, but only
   * when the total actually moved, so the same number is never reported twice.
   */
  private reportTruncatedCells(): void {
    if (this.truncatedCells === this.reportedTruncatedCells) {
      return;
    }
    this.reportedTruncatedCells = this.truncatedCells;
    this.options.onCellTruncated?.(this.truncatedCells);
  }
}

class SheetWriterImpl implements SheetWriter {
  readonly name: string;

  private readonly workbook: WorkbookWriterImpl;
  private readonly record: SheetRecord;
  private readonly autoFilter: boolean;
  private readonly headerRows: number;
  /** Resolves once the zip entry is open, the prologue is buffered and the header row (if any) is written. */
  private readonly ready: Promise<WorksheetWriter>;
  /** Ranges merged before the zip entry finished opening, applied in order once it has. */
  private readonly pendingMerges: string[] = [];
  private worksheet: WorksheetWriter | undefined;
  private rowsIssued = 0;
  private closeRequested = false;

  constructor(workbook: WorkbookWriterImpl, record: SheetRecord, options: SheetOptions, tabSelected: boolean) {
    this.workbook = workbook;
    this.record = record;
    this.name = record.name;
    this.autoFilter = options.autoFilter === true;
    this.headerRows = options.header === undefined ? 0 : 1;

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
          onCellTruncated: (count: number) => workbook.addTruncatedCells(count),
          tabSelected,
        },
        options,
      );
      if (options.header !== undefined) {
        const headerStyle = options.headerStyle === false ? undefined : (options.headerStyle ?? workbook.styles.headerStyle);
        await worksheet.writeRow(options.header, headerStyle);
      }
      this.worksheet = worksheet;
      // Ranges registered while the entry was still opening are applied here, so an invalid one surfaces on the
      // next call rather than disappearing into a floating promise.
      for (const range of this.pendingMerges) {
        worksheet.merge(range);
      }
      this.pendingMerges.length = 0;
      return worksheet;
    });
    // Nobody looks at `ready` until the first row, a close or an abort; marking it handled keeps an abort in between
    // from surfacing as an unhandled rejection.
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
    this.rowsIssued++;
    const worksheet = this.worksheet;
    if (worksheet === undefined) {
      return this.ready.then(opened => this.writeRowTo(opened, values, styles));
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
    if (this.worksheet === undefined) {
      this.pendingMerges.push(range);
      return;
    }
    this.worksheet.merge(range);
  }

  close(): Promise<SheetWriteSummary> {
    if (this.closeRequested) {
      return Promise.reject(new XlsxError('WRITER_STATE', `The sheet "${this.name}" is already closed.`));
    }
    this.closeRequested = true;
    this.workbook.sheetClosing();
    return this.workbook.enqueue(async () => {
      this.workbook.assertNotAborted();
      const worksheet = await this.ready;
      const written = await worksheet.close();
      const summary: SheetWriteSummary = { name: this.name, rows: written.rows, columns: written.columns };
      if (this.autoFilter && written.rows > 0 && written.columns > 0) {
        // Excel records the same range as a hidden `_xlnm._FilterDatabase` defined name in workbook.xml.
        this.record.autoFilterRange = sheetRange(written.rows, written.columns);
      }
      this.workbook.sheetClosed(this.record, summary);
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
    const rows = worksheet.nextRow - 1;
    if (rows % PROGRESS_ROW_INTERVAL === 0) {
      this.reportProgress(rows);
    }
    return pending;
  }

  private reportProgress(rows: number): void {
    this.workbook.reportProgress({ sheet: this.name, rows, bytesOut: this.workbook.zip.bytesWritten });
  }

  private assertWritable(): void {
    this.workbook.assertNotAborted();
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
 */
export function createWorkbookWriter(sink: ByteSink | WritableStream<Uint8Array>, options?: WorkbookWriterOptions): WorkbookWriter {
  return new WorkbookWriterImpl(sink, options ?? {});
}
