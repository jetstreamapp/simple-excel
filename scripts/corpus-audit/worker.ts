/**
 * One file of the real-world corpus audit, run in a child process of `run.ts` so a hang, a crash or a runaway
 * allocation costs one file, not the run. Reports each phase over IPC before it starts (the parent records the phase a
 * timeout or crash happened in) and the result as the last message.
 *
 * Phases, in order:
 *   1. sniff      - what the first bytes say the file is.
 *   2. stream     - our reader over a file source, every worksheet, nothing retained: time, cells, peak RSS.
 *   3. dump       - our reader again, typed dump for comparison (skipped past DUMP_CELL_LIMIT).
 *   4. sheetjs    - SheetJS with Jetstream's options, typed dump the same way `oracle/sheetjs/read-dump.mjs` builds it
 *                   (from the cells SheetJS parsed, not its `!ref`).
 *   -  retime     - two more stream passes when the first was over RETIME_RATIO times SheetJS's read; best of three.
 *   5. compare    - compareDumps(sheetjs, ours).
 *   6. objects    - toObjects() of the first worksheet against sheet_to_json (Jetstream's actual call), by source row.
 *   7. roundtrip  - our dump through our writer and back through our reader.
 *
 *   tsx scripts/corpus-audit/worker.ts <file.xlsx>
 */
import { closeSync, openSync, readFileSync, readSync, statSync } from 'node:fs';
import * as XLSX from 'xlsx';
import { isTyped, toLocalDate, truncateForExcel } from '../../fixtures/canonical/canonical.mjs';
import { collectToBytes, createWorkbookWriter, fromFile, isXlsxError, openWorkbook, sniff, SNIFF_BYTES } from '../../src/node/index';
import type { CellInput, Workbook } from '../../src/types';
import { compareDumps, type DiffReport } from '../../test/helpers/diff';
import { a1, fromJsDate, normalizeRows, toTyped, type TypedDump, type TypedValue } from '../../test/helpers/typed';
import {
  BLANK_HEADER_KEY,
  EMPTY_TEXT_HEADER_KEY,
  ERROR_LITERALS,
  knownSheetjsDivergence,
  RawSheets,
  sameHeader,
  SHEET_TO_JSON_OPTIONS,
  sheetjsKeyColumns,
  writerConvention,
} from './divergence';

/** Past this many cells the comparison dumps (ours and SheetJS's, fully materialized) are skipped. */
const DUMP_CELL_LIMIT = 3_000_000;
const SAMPLE_MISMATCHES = 8;
/**
 * A stream pass slower than this many times SheetJS's read of the same file is timed twice more and the best of the
 * three kept as `bestMs`: one pass of a small file is at the mercy of the other workers, the GC and JIT warm-up, and a
 * single slow run should not make the slowest-reads list.
 */
const RETIME_RATIO = 5;
const RETIME_PASSES = 2;

type Phase = 'sniff' | 'stream' | 'dump' | 'sheetjs' | 'retime' | 'compare' | 'objects' | 'roundtrip';

interface ErrorInfo {
  classified: boolean;
  code: string;
  name: string;
  message: string;
  stack?: string;
}

interface CompareSummary {
  pass: boolean;
  cells: number;
  mismatches: number;
  byCategory: Record<string, number>;
  known: Record<string, number>;
  samples: { sheet: string; cell: string; category: string; expected: unknown; actual: unknown }[];
  sheetRenamed?: { expected: string; actual: string };
}

export interface WorkerResult {
  file: string;
  bytes: number;
  sniff?: string;
  ours?: {
    ok: boolean;
    error?: ErrorInfo;
    ms?: number;
    sheets?: number;
    worksheets?: number;
    rows?: number;
    cells?: number;
    date1904?: boolean;
    baselineRssMb?: number;
    peakRssMb?: number;
    warnings?: string[];
    /** Best of three stream passes, set only when the first was over `RETIME_RATIO` times SheetJS's read. */
    bestMs?: number;
  };
  dumpSkipped?: string;
  sheetjs?: {
    ok: boolean;
    error?: string;
    ms?: number;
    peakRssMb?: number;
    /** Sheets SheetJS lists but left out of `Sheets` because parsing them threw (it swallows the error). */
    dropped?: { sheets: string[]; reason: string | null; error: string };
  };
  parity?: CompareSummary;
  objects?: {
    pass: boolean;
    rows: number;
    headersMatch: boolean;
    mismatches: number;
    known?: Record<string, number>;
    samples: unknown[];
    error?: ErrorInfo;
  };
  roundtrip?: (CompareSummary & { bytes: number }) | { error: ErrorInfo };
  phaseMs: Partial<Record<Phase, number>>;
}

function send(message: unknown): void {
  if (process.send) {
    process.send(message);
  } else {
    console.log(JSON.stringify(message));
  }
}

function errorInfo(error: unknown): ErrorInfo {
  if (isXlsxError(error)) {
    return { classified: true, code: error.code, name: error.name, message: error.message };
  }
  const value = error as { name?: string; message?: string; stack?: string; code?: string };
  return {
    classified: false,
    code: typeof value?.code === 'string' ? value.code : 'UNCLASSIFIED',
    name: value?.name ?? typeof error,
    message: String(value?.message ?? error).slice(0, 500),
    stack: value?.stack?.split('\n').slice(0, 8).join('\n'),
  };
}

/** Peak resident set of this process so far; Node reports maxRSS in kilobytes on every platform. */
function peakRssMb(): number {
  return Math.round(process.resourceUsage().maxRSS / 1024);
}

/** The full value of `cell` (A1) on a sheet of a dump, or undefined when the dump has no such sheet. */
function cellAt(dump: { sheets: { name: string; rows: TypedValue[][] }[] }, sheetName: string, cell: string): TypedValue | undefined {
  const sheet = dump.sheets.find(candidate => candidate.name === sheetName);
  if (!sheet || !cell) {
    return undefined;
  }
  const { r, c } = XLSX.utils.decode_cell(cell);
  return sheet.rows[r]?.[c] ?? null;
}

/**
 * `knownSheets` maps a sheet to the known reason every mismatch on it has: a sheet SheetJS dropped for a documented
 * cause (`droppedSheets`) has nothing to compare with. `dumps` are the two dumps compared: `compareDumps` keeps only a
 * 60-character preview of a long string in the mismatch, so `explain` reads the full values from them instead. `explain`
 * also gets the expected sheet's name and the cell (A1), to prove facts about the expected side's cell.
 */
function summarize(
  report: DiffReport,
  explain?: (expected: unknown, actual: unknown, where: { sheet: string; cell: string }) => string | null,
  knownSheets: ReadonlyMap<string, string> = new Map(),
  dumps?: { expected: { sheets: { name: string; rows: TypedValue[][] }[] }; actual: TypedDump },
): CompareSummary {
  const samples: CompareSummary['samples'] = [];
  const byCategory: Record<string, number> = {};
  const known: Record<string, number> = {};
  let mismatches = 0;
  for (const sheet of report.sheets) {
    const sheetReason = knownSheets.get(sheet.name);
    if (sheetReason !== undefined && sheet.mismatches.length > 0) {
      known[sheetReason] = (known[sheetReason] ?? 0) + sheet.mismatches.length;
      continue;
    }
    // A single sheet renamed between the dumps is compared under the expected name (`compareDumps`'s sheetRenamed).
    const actualName = report.sheetRenamed?.expected === sheet.name ? report.sheetRenamed.actual : sheet.name;
    for (const mismatch of sheet.mismatches) {
      const expected = dumps ? (cellAt(dumps.expected, sheet.name, mismatch.cell) ?? mismatch.expected) : mismatch.expected;
      const actual = dumps ? (cellAt(dumps.actual, actualName, mismatch.cell) ?? mismatch.actual) : mismatch.actual;
      const reason = explain?.(expected, actual, { sheet: sheet.name, cell: mismatch.cell }) ?? null;
      if (reason) {
        known[reason] = (known[reason] ?? 0) + 1;
        continue;
      }
      mismatches++;
      byCategory[mismatch.category] = (byCategory[mismatch.category] ?? 0) + 1;
      if (samples.length < SAMPLE_MISMATCHES) {
        samples.push({ sheet: sheet.name, ...mismatch });
      }
    }
  }
  return {
    pass: mismatches === 0,
    cells: report.totals.cells,
    mismatches,
    byCategory,
    known,
    samples,
    sheetRenamed: report.sheetRenamed,
  };
}

async function timed<T>(result: WorkerResult, phase: Phase, work: () => Promise<T>): Promise<T> {
  send({ phase });
  const start = performance.now();
  try {
    return await work();
  } finally {
    result.phaseMs[phase] = Math.round(performance.now() - start);
  }
}

/** Phase 2: stream every worksheet without keeping anything, the shape a memory-bounded caller reads in. */
async function streamPass(file: string, result: WorkerResult): Promise<boolean> {
  const baselineRssMb = peakRssMb();
  const warnings: string[] = [];
  const start = performance.now();
  const source = await fromFile(file);
  let workbook: Workbook | undefined;
  let rows = 0;
  let cells = 0;
  try {
    workbook = await openWorkbook(source, {
      onWarning: warning => warnings.push(`${warning.code}@${warning.sheet ?? ''}!${warning.ref ?? ''}`),
    });
    let worksheets = 0;
    for (const info of workbook.sheets) {
      if (info.kind !== 'worksheet') {
        continue;
      }
      worksheets++;
      for await (const row of workbook.sheet(info.index).rows()) {
        rows++;
        cells += row.length;
      }
    }
    result.ours = {
      ok: true,
      ms: Math.round(performance.now() - start),
      sheets: workbook.sheets.length,
      worksheets,
      rows,
      cells,
      date1904: workbook.date1904,
      baselineRssMb,
      peakRssMb: peakRssMb(),
      warnings: warnings.slice(0, 10),
    };
    return true;
  } catch (error) {
    result.ours = {
      ok: false,
      error: errorInfo(error),
      ms: Math.round(performance.now() - start),
      rows,
      cells,
      baselineRssMb,
      peakRssMb: peakRssMb(),
    };
    return false;
  } finally {
    await workbook?.close().catch(() => undefined);
    await source.close().catch(() => undefined);
  }
}

/**
 * Sheets SheetJS lists in `SheetNames` but left out of `Sheets`: parsing them threw, and SheetJS swallows sheet errors
 * unless `WTF` is set, so Jetstream gets no rows and no error. A second read with `WTF` surfaces the error for the
 * record; the cause is known when it is EC-NUMFMT-ACCOUNTING-BUILTIN (numFmtId 41 to 44 without a <numFmt>
 * declaration make `fmt_is_date` throw under cellDates without cellText). WTF stops at the first error SheetJS
 * swallows anywhere, so an unrelated one leaves the cause unexplained rather than misattributed.
 */
function droppedSheets(
  workbook: XLSX.WorkBook,
  bytes: Uint8Array,
): NonNullable<NonNullable<WorkerResult['sheetjs']>['dropped']> | undefined {
  const sheets = workbook.SheetNames.filter(name => workbook.Sheets[name] === undefined);
  if (sheets.length === 0) {
    return undefined;
  }
  let error = 'no error with WTF';
  let reason: string | null = null;
  try {
    XLSX.read(bytes, { cellText: false, cellDates: true, type: 'array', WTF: true });
  } catch (thrown) {
    const stack = String((thrown as Error)?.stack ?? thrown);
    error = stack.split('\n').slice(0, 2).join(' | ').slice(0, 300);
    if (/\bfmt_is_date\b/.test(stack)) {
      reason = 'sheetjs-numfmt-accounting';
    }
  }
  return { sheets, reason, error };
}

/** Phase 3: `test/helpers/dump.ts`, over bytes already in memory. */
async function ourDump(bytes: Uint8Array): Promise<TypedDump> {
  const workbook = await openWorkbook(bytes, { dates: 'utc', errors: 'object' });
  const sheets: TypedDump['sheets'] = [];
  try {
    for (const info of workbook.sheets) {
      const rows: TypedValue[][] = [];
      if (info.kind === 'worksheet') {
        for await (const row of workbook.sheet(info.index).rows({ blankRows: true })) {
          rows.push(row.map(value => toTyped(value, 'utc')));
        }
      }
      sheets.push({ name: info.name, hidden: info.hidden, rows: normalizeRows(rows) });
    }
  } finally {
    await workbook.close();
  }
  return { reader: 'simple-excel', sheets };
}

// cellText:false (Jetstream's option) suppresses `w`, so error cells only carry the BIFF error code in `v`.
const SHEETJS_ERRORS: Record<number, string> = {
  0: '#NULL!',
  7: '#DIV/0!',
  15: '#VALUE!',
  23: '#REF!',
  29: '#NAME?',
  36: '#NUM!',
  42: '#N/A',
  43: '#GETTING_DATA',
};

/** `oracle/sheetjs/read-dump.mjs`'s typed cell, minus the formula wrapper (we compare cached values). */
function sheetjsTyped(cell: XLSX.CellObject | undefined): TypedValue {
  if (!cell || cell.t === 'z') {
    return null;
  }
  if (cell.t === 'e') {
    return { $error: cell.w ?? SHEETJS_ERRORS[cell.v as number] ?? String(cell.v) };
  }
  if (cell.v instanceof Date) {
    return fromJsDate(cell.v, 'utc');
  }
  if (cell.t === 'd') {
    return fromJsDate(new Date(cell.v as string), 'utc');
  }
  return (cell.v ?? null) as TypedValue;
}

/**
 * First and last row and column (0-based) of the cells SheetJS actually parsed, or null when it parsed none. Not `!ref`:
 * SheetJS keeps `!ref` from the sheet's `<dimension>` but parses every cell, so a dimension written from a low row-count
 * hint (EC-DIMENSION-FROM-HINT) leaves cells outside it. The workbook is read in SheetJS's default sparse mode (A1 keys).
 */
function parsedExtent(worksheet: XLSX.WorkSheet): { firstRow: number; firstColumn: number; rows: number; columns: number } | null {
  let firstRow = Number.POSITIVE_INFINITY;
  let firstColumn = Number.POSITIVE_INFINITY;
  let lastRow = -1;
  let lastColumn = -1;
  for (const key of Object.keys(worksheet)) {
    if (key.startsWith('!')) {
      continue;
    }
    const { r, c } = XLSX.utils.decode_cell(key);
    firstRow = Math.min(firstRow, r);
    firstColumn = Math.min(firstColumn, c);
    lastRow = Math.max(lastRow, r);
    lastColumn = Math.max(lastColumn, c);
  }
  return lastRow < 0 ? null : { firstRow, firstColumn, rows: lastRow, columns: lastColumn };
}

/**
 * The range sheet_to_json should read: the sheet's `!ref`, widened to every cell SheetJS parsed when some lie outside it
 * (EC-DIMENSION-FROM-HINT: sheet_to_json reads only `!ref`, so a short `<dimension>` cuts records and columns that
 * Jetstream would lose to SheetJS, not to us). `widened` says whether `!ref` fell short. Undefined for a sheet with
 * neither a `!ref` nor a cell.
 */
function objectsRange(worksheet: XLSX.WorkSheet): { range: XLSX.Range; widened: boolean } | undefined {
  const extent = parsedExtent(worksheet);
  const ref = worksheet['!ref'];
  const declared = ref ? XLSX.utils.decode_range(ref) : undefined;
  if (extent === null) {
    return declared ? { range: declared, widened: false } : undefined;
  }
  if (
    declared &&
    extent.firstRow >= declared.s.r &&
    extent.firstColumn >= declared.s.c &&
    extent.rows <= declared.e.r &&
    extent.columns <= declared.e.c
  ) {
    return { range: declared, widened: false };
  }
  const start = declared?.s ?? { r: extent.firstRow, c: extent.firstColumn };
  const end = declared?.e ?? { r: extent.rows, c: extent.columns };
  return {
    range: {
      s: { r: Math.min(start.r, extent.firstRow), c: Math.min(start.c, extent.firstColumn) },
      e: { r: Math.max(end.r, extent.rows), c: Math.max(end.c, extent.columns) },
    },
    widened: true,
  };
}

function sheetjsDump(workbook: XLSX.WorkBook): TypedDump {
  const states = workbook.Workbook?.Sheets ?? [];
  const sheets = workbook.SheetNames.map((name, index) => {
    const worksheet = workbook.Sheets[name];
    const rows: TypedValue[][] = [];
    // EC-CHARTSHEET-CACHED-SERIES: a chartsheet has no cells; SheetJS invents them from the chart's <c:numCache>.
    const extent = worksheet && worksheet['!type'] !== 'chart' ? parsedExtent(worksheet) : null;
    if (worksheet && extent) {
      for (let r = 0; r <= extent.rows; r++) {
        const row: TypedValue[] = [];
        for (let c = 0; c <= extent.columns; c++) {
          row.push(sheetjsTyped(worksheet[a1(r, c)] as XLSX.CellObject | undefined));
        }
        rows.push(row);
      }
    }
    return { name, hidden: Boolean(states[index]?.Hidden), rows: normalizeRows(rows) };
  });
  return { reader: `sheetjs ${XLSX.version}`, sheets };
}

/** sheet_to_json over `range`; it returns nothing for a sheet without `!ref`, so one is lent for the call. */
function sheetToJson(worksheet: XLSX.WorkSheet, range: XLSX.Range): Record<string, unknown>[] {
  const ref = worksheet['!ref'];
  if (ref === undefined) {
    worksheet['!ref'] = XLSX.utils.encode_range(range);
  }
  try {
    return XLSX.utils.sheet_to_json<Record<string, unknown>>(worksheet, { ...SHEET_TO_JSON_OPTIONS, range });
  } finally {
    if (ref === undefined) {
      delete worksheet['!ref'];
    }
  }
}

/**
 * Where `toObjects()` found the header (0-based row) and the source row of each record it built, from a parallel
 * `rows({ blankRows: true })` pass: toObjects() does not say which row a record came from, and with blankRows the i-th
 * row is sheet row i. The header is the first row with a value (EC-HEADER-ROW-FIRST-NONEMPTY) and every later row
 * with one is a record, the rule the object-row builder follows.
 */
async function ourRecordRows(workbook: Workbook): Promise<{ headerRow: number | undefined; recordRows: number[] }> {
  let headerRow: number | undefined;
  const recordRows: number[] = [];
  let index = 0;
  for await (const row of workbook.sheet(0).rows({ blankRows: true })) {
    if (row.length > 0) {
      if (headerRow === undefined) {
        headerRow = index;
      } else {
        recordRows.push(index);
      }
    }
    index++;
  }
  return { headerRow, recordRows };
}

/** Every value of the record is blank (`''` or null), the shape of a record built only from defval. */
function allBlank(record: Record<string, unknown>, keys: readonly string[]): boolean {
  return keys.every(key => record[key] === '' || record[key] === null || record[key] === undefined);
}

/**
 * Phase 6: Jetstream's parseWorkbook against toObjects(), first sheet only, records lined up by source row. `raw` is
 * the raw XML of the workbook's sheets, which proves the shapes the SheetJS-side rules name.
 */
async function objectsParity(
  bytes: Uint8Array,
  sheetjs: XLSX.WorkBook,
  dropped: NonNullable<WorkerResult['sheetjs']>['dropped'],
  raw: RawSheets,
): Promise<NonNullable<WorkerResult['objects']>> {
  const firstName = sheetjs.SheetNames[0];
  const firstSheet = firstName ? sheetjs.Sheets[firstName] : undefined;
  const known: Record<string, number> = {};
  const note = (reason: string, count = 1): void => {
    known[reason] = (known[reason] ?? 0) + count;
  };
  // A first sheet SheetJS dropped for a known cause gives sheet_to_json nothing to compare with.
  if (firstName !== undefined && dropped?.reason && dropped.sheets.includes(firstName)) {
    return { pass: true, rows: 0, headersMatch: true, mismatches: 0, known: { [dropped.reason]: 1 }, samples: [] };
  }
  // EC-CHARTSHEET-CACHED-SERIES: the records SheetJS builds from a chartsheet's invented cells are not data.
  const isChart = firstSheet?.['!type'] === 'chart';
  const sheet = firstSheet && !isChart ? firstSheet : undefined;
  const planned = sheet ? objectsRange(sheet) : undefined;
  let range = planned?.range;
  if (planned?.widened) {
    note('sheetjs-dimension-from-hint');
  }
  let theirs = sheet && range ? sheetToJson(sheet, range) : [];

  // What the SheetJS-side rules need proven about SheetJS's cell at a 0-based row and column of the first sheet.
  const cellOf = (row: number, column: number): XLSX.CellObject | undefined => sheet?.[a1(row, column)] as XLSX.CellObject | undefined;
  // EC-TEXT-CELL-NO-TEXT: '' to SheetJS, and a text-typed cell with no value child in the sheet XML.
  const textCellWithoutValue = (row: number, column: number): boolean =>
    cellOf(row, column)?.v === '' && raw.textCellWithoutValue(0, a1(row, column));
  // No value for us either: no cell, a stub, or the text-typed cell with no value child.
  const sheetjsBlank = (row: number, column: number): boolean => {
    const cell = cellOf(row, column);
    return !cell || cell.t === 'z' || cell.v === undefined || textCellWithoutValue(row, column);
  };

  let beyondHeaderWarned = false;
  const workbook = await openWorkbook(bytes, {
    onWarning: warning => {
      if (warning.code === 'DATA_BEYOND_HEADER' && warning.sheet === firstName) {
        beyondHeaderWarned = true;
      }
    },
  });
  try {
    const firstWorksheet = workbook.sheets.find(info => info.kind === 'worksheet');
    if (!firstWorksheet || firstWorksheet.index !== 0) {
      return {
        pass: theirs.length === 0,
        rows: 0,
        headersMatch: theirs.length === 0,
        mismatches: theirs.length,
        known: isChart ? { 'sheetjs-chartsheet-cells': 1 } : {},
        samples: ['first sheet is not a worksheet'],
      };
    }
    const ours = await workbook.sheet(0).toObjects();
    const { headerRow, recordRows } = await ourRecordRows(workbook);
    // EC-HEADER-ROW-FIRST-NONEMPTY: sheet_to_json takes the first row of its range for the header even when it holds no
    // value (an honest dimension starting at a row of styled blank cells); ask it for ours instead. A '' is no value
    // only for the text-typed cell with no value child (EC-TEXT-CELL-NO-TEXT): a real empty string is a value to us
    // too, so that row is our header as well and nothing is re-run.
    if (sheet && range && headerRow !== undefined && headerRow > range.s.r) {
      const start = range.s;
      let blank = true;
      for (let column = start.c; column <= range.e.c && blank; column++) {
        blank = sheetjsBlank(start.r, column);
      }
      if (blank) {
        range = { s: { r: headerRow, c: start.c }, e: range.e };
        theirs = sheetToJson(sheet, range);
        note('sheetjs-header-row-blank');
      }
    }
    const headerSheetRow = range?.s.r ?? 0;
    const normalize = (value: unknown): unknown => (value instanceof Date ? fromJsDate(value, 'local') : value);
    const theirKeys = theirs.length > 0 ? Object.keys(theirs[0] as object) : [];
    const theirColumnOf = sheet && range ? sheetjsKeyColumns(sheet, range, theirKeys) : new Map<string, number>();
    // The header cell of SheetJS's key at `position` in `theirKeys` names its column '' (then `_N`) for SheetJS and is
    // blank for us: a text-typed cell with no value child (EC-TEXT-CELL-NO-TEXT), or a stub SheetJS made for a comment
    // on a cell with no value. An unknown column proves nothing.
    const headerCellWithoutValue = (position: number): boolean => {
      const column = theirColumnOf.get(theirKeys[position] as string);
      return column !== undefined && (cellOf(headerSheetRow, column)?.t === 'z' || textCellWithoutValue(headerSheetRow, column));
    };
    // A SheetJS record whose values are all blank, each '' proven to be no value for us either.
    const theirRecordBlank = (record: Record<string, unknown>, keys: readonly string[]): boolean => {
      const row = (record as { __rowNum__?: number }).__rowNum__;
      return keys.every(key => {
        const value = record[key];
        if (value === null || value === undefined) {
          return true;
        }
        const column = theirColumnOf.get(key);
        return value === '' && typeof row === 'number' && column !== undefined && sheetjsBlank(row, column);
      });
    };
    let ourHeaders = ours.rows.length > 0 ? Object.keys(ours.rows[0] as object) : [];
    // Keys compared column by column, so a column both sides name differently for a known reason still lines up.
    const headerReasons = new Set<string>();
    const headersLineUp = (keys: readonly string[]): boolean => {
      headerReasons.clear();
      if (theirKeys.length < keys.length) {
        return false;
      }
      for (let column = 0; column < keys.length; column++) {
        const reason = sameHeader(theirKeys[column] as string, keys[column] as string, () => headerCellWithoutValue(column));
        if (reason === null) {
          return false;
        }
        if (reason !== 'same') {
          headerReasons.add(reason);
        }
      }
      return true;
    };
    let headersMatch = headersLineUp(ourHeaders);
    // EC-HEADER-DIMENSION-START-COLUMN for a sheet part with no <dimension>: our columns start at A, SheetJS's at the
    // first column of the cells it parsed. Leading blank-header columns that are '' in every record are that
    // difference. With a <dimension>, both start at its first column, so the rule does not apply.
    let leading = 0;
    while (leading < ourHeaders.length && BLANK_HEADER_KEY.test(ourHeaders[leading] as string)) {
      leading++;
    }
    const leadingBlank = ourHeaders.slice(0, leading);
    if (
      !headersMatch &&
      leading > 0 &&
      raw.lacksDimension(0) &&
      ours.rows.every(record => allBlank(record, leadingBlank)) &&
      headersLineUp(ourHeaders.slice(leading))
    ) {
      headersMatch = true;
      ourHeaders = ourHeaders.slice(leading);
      headerReasons.add('no-dimension-leading-blank-column');
    }
    // EC-DATA-BEYOND-HEADER: our records stop at the last header cell with a value, where SheetJS adds a blank-header
    // key for every column out to the widest row. Those trailing keys are the intended divergence when we reported the
    // values we left out, or when there were none to leave out (a dimension wider than the data). A trailing '' or
    // `_N` key is a text-typed header cell with no value child (EC-TEXT-CELL-NO-TEXT), blank to us.
    const extra = theirKeys.slice(ourHeaders.length);
    if (headersMatch && extra.length > 0) {
      headersMatch =
        extra.every(
          (key, index) =>
            BLANK_HEADER_KEY.test(key) || (EMPTY_TEXT_HEADER_KEY.test(key) && headerCellWithoutValue(ourHeaders.length + index)),
        ) &&
        (beyondHeaderWarned || theirs.every(record => theirRecordBlank(record, extra)));
      if (headersMatch) {
        note('data-beyond-header');
      }
    }
    for (const reason of headerReasons) {
      note(reason);
    }
    // With the headers lined up, each of our keys reads SheetJS's key for the same column; otherwise both sides' keys
    // are compared by name.
    const theirHeaders = headersMatch ? theirKeys.slice(0, ourHeaders.length) : theirKeys;
    const theirKeyFor = new Map(headersMatch ? ourHeaders.map((key, column) => [key, theirKeys[column] as string]) : []);
    let mismatches = 0;
    const samples: unknown[] = [];
    const sample = (item: unknown): void => {
      if (samples.length < SAMPLE_MISMATCHES) {
        samples.push(item);
      }
    };
    // EC-OBJECTS-ERROR-ONLY-ROW: SheetJS's row holds at least one error and nothing else with a value.
    const sheetjsErrorOnlyRow = (row: number): boolean => {
      if (!sheet || !range) {
        return false;
      }
      let errors = 0;
      for (let column = range.s.c; column <= range.e.c; column++) {
        if (cellOf(row, column)?.t === 'e') {
          errors++;
        } else if (!sheetjsBlank(row, column)) {
          return false;
        }
      }
      return errors > 0;
    };
    // Line the records up by source row (`__rowNum__` is sheet_to_json's 0-based row) rather than by position, so a
    // record only one side builds shows up as that record instead of shifting every record after it.
    const aligned = recordRows.length === ours.rows.length;
    const theirByRow = new Map<number, Record<string, unknown>>();
    theirs.forEach((record, index) => {
      const row = (record as { __rowNum__?: number }).__rowNum__;
      theirByRow.set(aligned && typeof row === 'number' ? row : index, record);
    });
    if (!aligned) {
      sample({ unaligned: { recordRows: recordRows.length, records: ours.rows.length } });
    }
    for (const [index, ourRow] of ours.rows.entries()) {
      const row = aligned ? (recordRows[index] as number) : index;
      const theirRow = theirByRow.get(row);
      if (theirRow === undefined) {
        // EC-OBJECTS-ERROR-ONLY-ROW: with errors: 'string' a row whose only values are errors is a record; SheetJS
        // drops it (its errors become defval, so the row has no value), as we do with errors: 'null'.
        // The errors may all sit right of the header (EC-DATA-BEYOND-HEADER), leaving a record of defval.
        if (
          aligned &&
          sheetjsErrorOnlyRow(row) &&
          Object.values(ourRow).every(value => value === '' || (typeof value === 'string' && ERROR_LITERALS.has(value)))
        ) {
          note('objects-error-only-row');
        } else {
          mismatches++;
          sample({ row: row + 1, oursOnly: ourRow });
        }
        continue;
      }
      theirByRow.delete(row);
      // SheetJS's own row for the record, whatever the alignment, for the facts about its cells.
      const sheetRow = (theirRow as { __rowNum__?: number }).__rowNum__;
      const keys = headersMatch ? ourHeaders : [...new Set([...theirHeaders, ...Object.keys(ourRow)])];
      for (const key of keys) {
        const theirKey = theirKeyFor.get(key) ?? key;
        const theirValue = normalize(theirRow[theirKey]);
        const ourValue = normalize(ourRow[key]);
        if (JSON.stringify(theirValue) === JSON.stringify(ourValue)) {
          continue;
        }
        if (
          typeof theirValue === 'number' &&
          typeof ourValue === 'number' &&
          Math.abs(theirValue - ourValue) <= Math.abs(theirValue) * 1e-15
        ) {
          continue;
        }
        const column = theirColumnOf.get(theirKey);
        const facts =
          typeof sheetRow === 'number' && column !== undefined
            ? {
                textCellWithoutValue: () => textCellWithoutValue(sheetRow, column),
                sheetjsError: () => cellOf(sheetRow, column)?.t === 'e',
                strCellText: () => raw.strCellText(0, a1(sheetRow, column)),
              }
            : {};
        const reason = knownSheetjsDivergence(theirValue, ourValue, facts);
        if (reason) {
          note(reason);
          continue;
        }
        mismatches++;
        sample({ row: row + 1, key, sheetjs: theirValue, ours: ourValue });
      }
    }
    for (const [row, theirRow] of theirByRow) {
      // EC-TEXT-CELL-NO-TEXT: a row whose only cells are text-typed cells with no value child is blank for us and a
      // record of '' for SheetJS.
      if (theirRecordBlank(theirRow, Object.keys(theirRow))) {
        note('sheetjs-empty-text-cell');
        continue;
      }
      mismatches++;
      sample({ row: row + 1, sheetjsOnly: theirRow });
    }
    if (theirs.length !== ours.rows.length) {
      samples.unshift({ rowCount: { sheetjs: theirs.length, ours: ours.rows.length } });
    }
    if (!headersMatch) {
      samples.unshift({ headers: { sheetjs: theirHeaders.slice(0, 20), ours: ourHeaders.slice(0, 20) } });
    }
    return { pass: mismatches === 0 && headersMatch, rows: ours.rows.length, headersMatch, mismatches, known, samples };
  } finally {
    await workbook.close();
  }
}

function toCellInput(value: TypedValue): CellInput {
  if (value === null || typeof value !== 'object') {
    return value;
  }
  if (isTyped(value, '$error')) {
    return { error: (value as { $error: string }).$error as never };
  }
  return toLocalDate(value as { $date?: string; $datetime?: string; $time?: string });
}

/** Phase 7: our dump through our writer (local dates, matching `toLocalDate`) and back through our reader. */
async function roundTrip(dump: TypedDump, date1904: boolean): Promise<NonNullable<WorkerResult['roundtrip']>> {
  const sink = collectToBytes();
  const writer = createWorkbookWriter(sink, { deterministic: true, dates: 'local', date1904, strings: 'auto' });
  const everyHidden = dump.sheets.every(sheet => sheet.hidden);
  const written: TypedDump['sheets'] = [];
  for (const [index, sheetSpec] of dump.sheets.entries()) {
    const hidden = sheetSpec.hidden && !(everyHidden && index === 0);
    const sheet = writer.addSheet(sheetSpec.name, { hidden });
    for (const row of sheetSpec.rows) {
      await sheet.writeRow(row.map(value => truncateForExcel(toCellInput(value)) as CellInput));
    }
    await sheet.close();
    written.push({ name: sheet.name, hidden, rows: sheetSpec.rows });
  }
  await writer.close();
  const bytes = sink.result();
  const reread = await ourDump(bytes);
  return {
    ...summarize(compareDumps({ sheets: written }, reread, new Set(['truncate-32767'])), writerConvention, new Map(), {
      expected: { sheets: written },
      actual: reread,
    }),
    bytes: bytes.byteLength,
  };
}

async function main(): Promise<void> {
  const file = process.argv[2];
  if (!file) {
    console.error('usage: worker.ts <file.xlsx>');
    process.exit(2);
  }
  const result: WorkerResult = { file, bytes: statSync(file).size, phaseMs: {} };
  result.sniff = await timed(result, 'sniff', async () => {
    const handle = openSync(file, 'r');
    try {
      const head = new Uint8Array(SNIFF_BYTES);
      return sniff(head.subarray(0, readSync(handle, head, 0, SNIFF_BYTES, 0)));
    } finally {
      closeSync(handle);
    }
  });
  // Before the whole file is loaded for the comparison phases, so the peak RSS it records is the reader's alone.
  const streamed = await timed(result, 'stream', () => streamPass(file, result));

  const buffer = readFileSync(file);
  const bytes = new Uint8Array(buffer.buffer, buffer.byteOffset, buffer.byteLength);

  let ours: TypedDump | undefined;
  if (streamed && (result.ours?.cells ?? 0) <= DUMP_CELL_LIMIT) {
    try {
      ours = await timed(result, 'dump', () => ourDump(bytes));
    } catch (error) {
      result.ours = { ...result.ours, ok: false, error: { ...errorInfo(error), message: `dump pass: ${errorInfo(error).message}` } };
    }
  } else if (streamed) {
    result.dumpSkipped = `${result.ours?.cells} cells > ${DUMP_CELL_LIMIT}`;
  }

  let sheetjs: XLSX.WorkBook | undefined;
  if (result.dumpSkipped === undefined) {
    await timed(result, 'sheetjs', async () => {
      const start = performance.now();
      try {
        sheetjs = XLSX.read(bytes, { cellText: false, cellDates: true, type: 'array' });
        result.sheetjs = { ok: true, ms: Math.round(performance.now() - start), peakRssMb: peakRssMb() };
        const dropped = droppedSheets(sheetjs, bytes);
        if (dropped) {
          result.sheetjs.dropped = dropped;
        }
      } catch (error) {
        result.sheetjs = {
          ok: false,
          error: String((error as Error)?.message ?? error).slice(0, 300),
          ms: Math.round(performance.now() - start),
        };
      }
    });
  }

  const firstPassMs = result.ours?.ms;
  const sheetjsMs = result.sheetjs?.ok ? result.sheetjs.ms : undefined;
  if (result.ours?.ok && firstPassMs !== undefined && sheetjsMs !== undefined && firstPassMs > RETIME_RATIO * sheetjsMs) {
    await timed(result, 'retime', async () => {
      let best = firstPassMs;
      for (let pass = 0; pass < RETIME_PASSES; pass++) {
        const scratch: WorkerResult = { file, bytes: result.bytes, phaseMs: {} };
        if (await streamPass(file, scratch)) {
          best = Math.min(best, scratch.ours?.ms ?? best);
        }
      }
      if (result.ours) {
        result.ours.bestMs = best;
      }
    });
  }

  if (ours && sheetjs) {
    const theirs = sheetjsDump(sheetjs);
    const dropped = result.sheetjs?.dropped;
    const knownSheets = new Map(dropped?.reason ? dropped.sheets.map(name => [name, dropped.reason as string]) : []);
    const dumps = { expected: theirs, actual: ours };
    const workbook = sheetjs;
    const raw = new RawSheets(bytes);
    // The facts a rule needs about SheetJS's cell, proven from that cell and the raw sheet XML only when asked.
    const explain = (expected: unknown, actual: unknown, where: { sheet: string; cell: string }): string | null => {
      const index = workbook.SheetNames.indexOf(where.sheet);
      const cell = where.cell ? (workbook.Sheets[where.sheet]?.[where.cell] as XLSX.CellObject | undefined) : undefined;
      return knownSheetjsDivergence(expected, actual, {
        textCellWithoutValue: () => index >= 0 && cell?.v === '' && raw.textCellWithoutValue(index, where.cell),
        sheetjsError: () => cell?.t === 'e',
        strCellText: () => (index >= 0 && where.cell ? raw.strCellText(index, where.cell) : undefined),
      });
    };
    result.parity = await timed(result, 'compare', async () => summarize(compareDumps(theirs, ours), explain, knownSheets, dumps));
    try {
      result.objects = await timed(result, 'objects', () => objectsParity(bytes, workbook, result.sheetjs?.dropped, raw));
    } catch (error) {
      result.objects = { pass: false, rows: 0, headersMatch: false, mismatches: 0, samples: [], error: errorInfo(error) };
    }
  }

  if (ours) {
    const dump = ours;
    try {
      result.roundtrip = await timed(result, 'roundtrip', () => roundTrip(dump, result.ours?.date1904 ?? false));
    } catch (error) {
      result.roundtrip = { error: errorInfo(error) };
    }
  }

  send({ result });
}

main().catch(error => {
  send({ fatal: errorInfo(error) });
  process.exitCode = 1;
});
