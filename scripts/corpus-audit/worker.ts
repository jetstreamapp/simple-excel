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

/** The writer's documented conventions (ADR-003 pre-1900 dates as ISO text, EC-DATE-DST-GAP), not round-trip losses. */
function writerConvention(expected: unknown, actual: unknown): string | null {
  const expectedTime = temporalMs(expected);
  if (expectedTime !== undefined && typeof actual === 'string' && expectedTime < Date.parse('1900-03-01T00:00:00.000Z')) {
    return 'writer-pre-1900-as-text';
  }
  const actualTime = temporalMs(actual);
  if (expectedTime !== undefined && actualTime !== undefined && Math.abs(actualTime - expectedTime) === 3_600_000) {
    return 'writer-dst-gap';
  }
  return null;
}

function temporalMs(value: unknown): number | undefined {
  if (typeof value !== 'object' || value === null) {
    return undefined;
  }
  const record = value as Record<string, unknown>;
  if (typeof record.$datetime === 'string') {
    return Date.parse(`${record.$datetime}Z`);
  }
  if (typeof record.$date === 'string') {
    return Date.parse(`${record.$date}T00:00:00.000Z`);
  }
  if (typeof record.$time === 'string') {
    return Date.parse(`1899-12-30T${record.$time}Z`);
  }
  return undefined;
}

/** A character past U+FFFF. */
const ASTRAL = /[\u{10000}-\u{10FFFF}]/gu;

/** Year of a `$date` or `$datetime` value, read off its text: `Date.parse` cannot read a five-digit year. */
function temporalYear(value: unknown): number | undefined {
  if (typeof value !== 'object' || value === null) {
    return undefined;
  }
  const record = value as Record<string, unknown>;
  const text = typeof record.$datetime === 'string' ? record.$datetime : typeof record.$date === 'string' ? record.$date : undefined;
  const year = text === undefined ? undefined : /^(-?\d+)-/.exec(text)?.[1];
  return year === undefined ? undefined : Number(year);
}

/** 9999-12-31 in the 1904 date system, the lower of the two systems' last serials (2,958,465 in the 1900 system). */
const LAST_SERIAL_1904 = 2_957_003;
const MARKER_DAY_START = Date.parse('1899-12-31T00:00:00.000Z');
const MARKER_DAY_LENGTH = 86_400_000;
const ERROR_LITERALS = new Set([
  '#NULL!',
  '#DIV/0!',
  '#VALUE!',
  '#REF!',
  '#NAME?',
  '#NUM!',
  '#N/A',
  '#GETTING_DATA',
  '#SPILL!',
  '#CALC!',
  '#FIELD!',
  '#BLOCKED!',
  '#UNKNOWN!',
  '#CONNECT!',
  '#BUSY!',
]);

/**
 * SheetJS behaviour already documented in `test/corpus-policies.json` and the catalog, so a difference it explains is
 * not a finding: `theirs` is SheetJS's value, `ours` ours. Returns the reason, or null when the difference is unexplained.
 */
function knownSheetjsDivergence(theirs: unknown, ours: unknown): string | null {
  const theirTime = temporalMs(theirs);
  const ourTime = temporalMs(ours);
  if (theirTime !== undefined && ourTime !== undefined) {
    // EC-TIME-ONLY-HAS-DATE-PART: SheetJS hangs a time of day on 1899-12-31, we on 1899-12-30.
    if (
      theirTime >= MARKER_DAY_START &&
      theirTime < MARKER_DAY_START + MARKER_DAY_LENGTH &&
      Math.abs(theirTime - MARKER_DAY_LENGTH - ourTime) <= 1
    ) {
      return 'sheetjs-time-only-marker';
    }
    // SheetJS truncates the serial-to-milliseconds conversion where we round.
    if (Math.abs(theirTime - ourTime) <= 1) {
      return 'sheetjs-ms-truncation';
    }
    return null;
  }
  // EC-DATE-SERIAL-OVER-9999: past 9999-12-31 (serial 2,958,465, or 2,957,003 in the 1904 system) SheetJS still makes
  // a Date, of year 10000 or later or invalid, where we return the number. The lower 1904 bound covers both systems;
  // the year SheetJS gives is what proves the serial is past the end.
  if (typeof ours === 'number' && ours > LAST_SERIAL_1904 && (theirs === '$invalidDate' || (temporalYear(theirs) ?? 0) > 9999)) {
    return 'sheetjs-date-over-9999';
  }
  // Errors become defval in sheet_to_json with cellText:false (objects mode reads them as their text), except #NULL!
  // (BIFF code 0), which sheet_to_json turns into null whatever defval says.
  const ourError = typeof ours === 'object' && ours !== null && '$error' in ours ? (ours as { $error: string }).$error : ours;
  if ((theirs === '' || (theirs === null && ourError === '#NULL!')) && typeof ourError === 'string' && ERROR_LITERALS.has(ourError)) {
    return 'sheetjs-error-defval';
  }
  // EC-TEXT-CELL-NO-TEXT: a text-typed cell with no text (`<c t="inlineStr"/>`, `<c t="str"/>`) is blank for us and
  // '' for SheetJS.
  if (theirs === '' && ours === null) {
    return 'sheetjs-empty-text-cell';
  }
  if (typeof theirs === 'string' && typeof ours === 'string') {
    // EC-CRLF-NORMALIZED: SheetJS drops the CR of a CRLF (or a decoded _x000D_).
    if (ours.replace(/\r\n/g, '\n').replace(/\r/g, '\n') === theirs.replace(/\r/g, '\n') || ours.replace(/\r/g, '') === theirs) {
      return 'sheetjs-cr-collapse';
    }
    // EC-XML-NCR-ABOVE-BMP: SheetJS keeps the low 16 bits of a character reference past U+FFFF (U+1F603 -> U+F603),
    // or reads it as U+0000 in a t="str" value.
    if (
      ours.replace(ASTRAL, character => String.fromCharCode((character.codePointAt(0) as number) & 0xffff)) === theirs ||
      ours.replace(ASTRAL, '\0') === theirs
    ) {
      return 'sheetjs-ncr-above-bmp';
    }
    // EC-XML-ESCAPE-CASE: SheetJS also decodes `_X0041_`, which Excel leaves literal.
    const decoded = ours.replace(/_[xX]([0-9A-Fa-f]{4})_/g, (_match, hex: string) => String.fromCharCode(Number.parseInt(hex, 16)));
    if (decoded === theirs || decoded.replace(/\r/g, '') === theirs) {
      return 'sheetjs-escape-case';
    }
    return null;
  }
  // EC-DATE-PRE-1900 / EC-DATE-TIME-ONLY-NEGATIVE-SERIAL: serial 60 and negative serials stay numbers for us.
  if (theirTime !== undefined && typeof ours === 'number' && (ours === 60 || ours < 0)) {
    return 'sheetjs-fake-serial';
  }
  // SheetJS refuses to date serial 60 (Excel's fake 1900-02-29) and leaves it a number; we read 1900-03-01.
  if (
    typeof theirs === 'number' &&
    theirs >= 60 &&
    theirs < 61 &&
    ourTime !== undefined &&
    ourTime >= Date.parse('1900-02-28T00:00:00.000Z') &&
    ourTime < Date.parse('1900-03-02T00:00:00.000Z')
  ) {
    return 'sheetjs-serial-60';
  }
  // A number cell holding text past the double range ('inf', '1.797693134862316e+308'): SheetJS drops it, we keep the text.
  if (theirs === null && typeof ours === 'string' && /^(-?inf|nan|-?\d(\.\d+)?e\+30[89])$/i.test(ours)) {
    return 'sheetjs-unparseable-number';
  }
  return null;
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
 * 60-character preview of a long string in the mismatch, so `explain` reads the full values from them instead.
 */
function summarize(
  report: DiffReport,
  explain?: (expected: unknown, actual: unknown) => string | null,
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
      const reason = explain?.(expected, actual) ?? null;
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
 * Last row and column (0-based) of the cells SheetJS actually parsed, or null when it parsed none. Not `!ref`: SheetJS
 * keeps `!ref` from the sheet's `<dimension>` but parses every cell, so a dimension written from a low row-count hint
 * (EC-DIMENSION-FROM-HINT) leaves cells outside it. The workbook is read in SheetJS's default sparse mode (A1 keys).
 */
function parsedExtent(worksheet: XLSX.WorkSheet): { rows: number; columns: number } | null {
  let lastRow = -1;
  let lastColumn = -1;
  for (const key of Object.keys(worksheet)) {
    if (key.startsWith('!')) {
      continue;
    }
    const { r, c } = XLSX.utils.decode_cell(key);
    lastRow = Math.max(lastRow, r);
    lastColumn = Math.max(lastColumn, c);
  }
  return lastRow < 0 ? null : { rows: lastRow, columns: lastColumn };
}

/** True when SheetJS parsed cells below or right of the sheet's `!ref` (EC-DIMENSION-FROM-HINT). */
function dimensionUndersized(worksheet: XLSX.WorkSheet): boolean {
  const extent = parsedExtent(worksheet);
  const ref = worksheet['!ref'];
  if (extent === null) {
    return false;
  }
  if (!ref) {
    return true;
  }
  const range = XLSX.utils.decode_range(ref);
  return extent.rows > range.e.r || extent.columns > range.e.c;
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

const SHEET_TO_JSON_OPTIONS: XLSX.Sheet2JSONOpts = {
  dateNF: 'yyyy"-"mm"-"dd"T"hh:mm:ss',
  defval: '',
  blankrows: false,
  rawNumbers: true,
};

/** A key SheetJS gives a column whose header cell is blank: `__EMPTY`, `__EMPTY_1`, ... */
const BLANK_HEADER_KEY = /^__EMPTY(?:_\d+)?$/;

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

/**
 * True when the first row of the sheet's `!ref`, which sheet_to_json takes for the header, holds no value. A '' counts
 * as none: SheetJS reads a text-typed cell with no text as '' where it is blank to us (EC-TEXT-CELL-NO-TEXT), so a row
 * of only those is the header to SheetJS and skipped by our header search. A real empty string is a value to us too,
 * so that row is our header as well and nothing is re-run.
 */
function sheetjsHeaderRowBlank(worksheet: XLSX.WorkSheet): boolean {
  const ref = worksheet['!ref'];
  if (!ref) {
    return false;
  }
  const range = XLSX.utils.decode_range(ref);
  for (let c = range.s.c; c <= range.e.c; c++) {
    const cell = worksheet[a1(range.s.r, c)] as XLSX.CellObject | undefined;
    if (cell && cell.t !== 'z' && cell.v !== undefined && cell.v !== '') {
      return false;
    }
  }
  return true;
}

/** SheetJS's key for a header cell whose text is '' (a text-typed cell with no text): '', then `_1`, `_2`, ... */
const EMPTY_TEXT_HEADER_KEY = /^(?:_\d+)?$/;
/** SheetJS's dateNF text of a time-only header cell: Excel's "day 0" of January 1900, as Excel shows it (`_N` dedupes). */
const TIME_ONLY_HEADER_KEY = /^1900-01-00T(\d\d:\d\d:\d\d(?:_\d+)?)$/;

/**
 * 'same' when SheetJS's key and ours name the same column, the known reason when they differ for one, or null.
 * - Two blank-header names (`__EMPTY_N`) are the same column even when an earlier difference shifted the numbering.
 * - EC-TEXT-CELL-NO-TEXT: a text-typed header cell with no text is '' (then `_1`, ...) to SheetJS, blank to us.
 * - A time-only header cell (a serial below 1) is named by its wall clock on Excel's day 0 of 1900 by SheetJS and on
 *   1899-12-30 by us, the marker day our values use (EC-TIME-ONLY-HAS-DATE-PART).
 */
function sameHeader(theirKey: string, ourKey: string): string | null {
  if (theirKey === ourKey || (BLANK_HEADER_KEY.test(theirKey) && BLANK_HEADER_KEY.test(ourKey))) {
    return 'same';
  }
  if (BLANK_HEADER_KEY.test(ourKey) && EMPTY_TEXT_HEADER_KEY.test(theirKey)) {
    return 'sheetjs-empty-text-cell';
  }
  const time = TIME_ONLY_HEADER_KEY.exec(theirKey)?.[1];
  if (time !== undefined && ourKey === `1899-12-30T${time}`) {
    return 'header-time-only-marker';
  }
  return null;
}

/** Every value of the record is blank (`''` or null), the shape of a record built only from defval. */
function allBlank(record: Record<string, unknown>, keys: readonly string[]): boolean {
  return keys.every(key => record[key] === '' || record[key] === null || record[key] === undefined);
}

/** Phase 6: Jetstream's parseWorkbook against toObjects(), first sheet only, records lined up by source row. */
async function objectsParity(
  bytes: Uint8Array,
  sheetjs: XLSX.WorkBook,
  dropped: NonNullable<WorkerResult['sheetjs']>['dropped'],
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
  let theirs = firstSheet && !isChart ? XLSX.utils.sheet_to_json<Record<string, unknown>>(firstSheet, SHEET_TO_JSON_OPTIONS) : [];
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
    // EC-HEADER-ROW-FIRST-NONEMPTY: sheet_to_json takes the first row of `!ref` for the header even when it holds no
    // value (an honest dimension starting at a row of styled blank cells); ask it for ours instead.
    if (firstSheet && headerRow !== undefined && sheetjsHeaderRowBlank(firstSheet)) {
      const start = XLSX.utils.decode_range(firstSheet['!ref'] as string).s.r;
      if (headerRow > start) {
        theirs = XLSX.utils.sheet_to_json<Record<string, unknown>>(firstSheet, { ...SHEET_TO_JSON_OPTIONS, range: headerRow });
        note('sheetjs-header-row-blank');
      }
    }
    const normalize = (value: unknown): unknown => (value instanceof Date ? fromJsDate(value, 'local') : value);
    const theirKeys = theirs.length > 0 ? Object.keys(theirs[0] as object) : [];
    let ourHeaders = ours.rows.length > 0 ? Object.keys(ours.rows[0] as object) : [];
    // Keys compared column by column, so a column both sides name differently for a known reason still lines up.
    const headerReasons = new Set<string>();
    const headersLineUp = (keys: readonly string[]): boolean => {
      headerReasons.clear();
      if (theirKeys.length < keys.length) {
        return false;
      }
      for (let column = 0; column < keys.length; column++) {
        const reason = sameHeader(theirKeys[column] as string, keys[column] as string);
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
    // EC-HEADER-DIMENSION-START-COLUMN without a <dimension>: our columns start at A, SheetJS's at the first column
    // of the cells it parsed. Leading blank-header columns that are '' in every record are that difference.
    let leading = 0;
    while (leading < ourHeaders.length && BLANK_HEADER_KEY.test(ourHeaders[leading] as string)) {
      leading++;
    }
    const leadingBlank = ourHeaders.slice(0, leading);
    if (
      !headersMatch &&
      leading > 0 &&
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
    // `_N` key is a text-typed header cell with no text (EC-TEXT-CELL-NO-TEXT), blank to us.
    const extra = theirKeys.slice(ourHeaders.length);
    if (headersMatch && extra.length > 0) {
      headersMatch =
        extra.every(key => BLANK_HEADER_KEY.test(key) || EMPTY_TEXT_HEADER_KEY.test(key)) &&
        (beyondHeaderWarned || theirs.every(record => allBlank(record, extra)));
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
        if (Object.values(ourRow).every(value => value === '' || (typeof value === 'string' && ERROR_LITERALS.has(value)))) {
          note('objects-error-only-row');
        } else {
          mismatches++;
          sample({ row: row + 1, oursOnly: ourRow });
        }
        continue;
      }
      theirByRow.delete(row);
      const keys = headersMatch ? ourHeaders : [...new Set([...theirHeaders, ...Object.keys(ourRow)])];
      for (const key of keys) {
        const theirValue = normalize(theirRow[theirKeyFor.get(key) ?? key]);
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
        const reason = knownSheetjsDivergence(theirValue, ourValue);
        if (reason) {
          note(reason);
          continue;
        }
        mismatches++;
        sample({ row: row + 1, key, sheetjs: theirValue, ours: ourValue });
      }
    }
    for (const [row, theirRow] of theirByRow) {
      // EC-TEXT-CELL-NO-TEXT: a row whose only cells are text-typed cells with no text is blank for us and a record
      // of '' for SheetJS.
      if (allBlank(theirRow, Object.keys(theirRow))) {
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
    // EC-DIMENSION-FROM-HINT: sheet_to_json reads only the rows and columns of `!ref`, so with cells beyond it SheetJS
    // gives Jetstream fewer records or columns than the sheet has. A known SheetJS divergence, not ours to match.
    if (firstSheet && (mismatches > 0 || !headersMatch) && dimensionUndersized(firstSheet)) {
      note('sheetjs-dimension-from-hint', Math.max(mismatches, 1));
      return { pass: true, rows: ours.rows.length, headersMatch, mismatches: 0, known, samples };
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
    result.parity = await timed(result, 'compare', async () =>
      summarize(compareDumps(theirs, ours), knownSheetjsDivergence, knownSheets, dumps),
    );
    const workbook = sheetjs;
    try {
      result.objects = await timed(result, 'objects', () => objectsParity(bytes, workbook, result.sheetjs?.dropped));
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
