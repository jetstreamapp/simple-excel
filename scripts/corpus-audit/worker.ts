/**
 * One file of the real-world corpus audit, run in a child process of `run.ts` so a hang, a crash or a runaway
 * allocation costs one file, not the run. Reports each phase over IPC before it starts (the parent records the phase a
 * timeout or crash happened in) and the result as the last message.
 *
 * Phases, in order:
 *   1. sniff      - what the first bytes say the file is.
 *   2. stream     - our reader over a file source, every worksheet, nothing retained: time, cells, peak RSS.
 *   3. dump       - our reader again, typed dump for comparison (skipped past DUMP_CELL_LIMIT).
 *   4. sheetjs    - SheetJS with Jetstream's options, typed dump the same way `oracle/sheetjs/read-dump.mjs` builds it.
 *   5. compare    - compareDumps(sheetjs, ours).
 *   6. objects    - toObjects() of the first worksheet against sheet_to_json (Jetstream's actual call).
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

type Phase = 'sniff' | 'stream' | 'dump' | 'sheetjs' | 'compare' | 'objects' | 'roundtrip';

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
  };
  dumpSkipped?: string;
  sheetjs?: { ok: boolean; error?: string; ms?: number; peakRssMb?: number };
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
  // Errors become defval in sheet_to_json with cellText:false (objects mode reads them as their text).
  const ourError = typeof ours === 'object' && ours !== null && '$error' in ours ? (ours as { $error: string }).$error : ours;
  if (theirs === '' && typeof ourError === 'string' && ERROR_LITERALS.has(ourError)) {
    return 'sheetjs-error-defval';
  }
  if (typeof theirs === 'string' && typeof ours === 'string') {
    // EC-CRLF-NORMALIZED: SheetJS drops the CR of a CRLF (or a decoded _x000D_).
    if (ours.replace(/\r\n/g, '\n').replace(/\r/g, '\n') === theirs.replace(/\r/g, '\n') || ours.replace(/\r/g, '') === theirs) {
      return 'sheetjs-cr-collapse';
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

function summarize(report: DiffReport, explain?: (expected: unknown, actual: unknown) => string | null): CompareSummary {
  const samples: CompareSummary['samples'] = [];
  const byCategory: Record<string, number> = {};
  const known: Record<string, number> = {};
  let mismatches = 0;
  for (const sheet of report.sheets) {
    for (const mismatch of sheet.mismatches) {
      const reason = explain?.(mismatch.expected, mismatch.actual) ?? null;
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

function sheetjsDump(workbook: XLSX.WorkBook): TypedDump {
  const states = workbook.Workbook?.Sheets ?? [];
  const sheets = workbook.SheetNames.map((name, index) => {
    const worksheet = workbook.Sheets[name];
    const rows: TypedValue[][] = [];
    const ref = worksheet?.['!ref'];
    if (worksheet && ref) {
      const range = XLSX.utils.decode_range(ref);
      for (let r = 0; r <= range.e.r; r++) {
        const row: TypedValue[] = [];
        for (let c = 0; c <= range.e.c; c++) {
          row.push(sheetjsTyped(worksheet[a1(r, c)] as XLSX.CellObject | undefined));
        }
        rows.push(row);
      }
    }
    return { name, hidden: Boolean(states[index]?.Hidden), rows: normalizeRows(rows) };
  });
  return { reader: `sheetjs ${XLSX.version}`, sheets };
}

/** Phase 6: Jetstream's parseWorkbook against toObjects(), first sheet only. */
async function objectsParity(bytes: Uint8Array, sheetjs: XLSX.WorkBook): Promise<NonNullable<WorkerResult['objects']>> {
  const firstName = sheetjs.SheetNames[0];
  const theirs = firstName
    ? XLSX.utils.sheet_to_json<Record<string, unknown>>(sheetjs.Sheets[firstName] as XLSX.WorkSheet, {
        dateNF: 'yyyy"-"mm"-"dd"T"hh:mm:ss',
        defval: '',
        blankrows: false,
        rawNumbers: true,
      })
    : [];
  const workbook = await openWorkbook(bytes);
  try {
    const firstWorksheet = workbook.sheets.find(info => info.kind === 'worksheet');
    if (!firstWorksheet || firstWorksheet.index !== 0) {
      return {
        pass: theirs.length === 0,
        rows: 0,
        headersMatch: theirs.length === 0,
        mismatches: theirs.length,
        samples: ['first sheet is not a worksheet'],
      };
    }
    const ours = await workbook.sheet(0).toObjects();
    const normalize = (value: unknown): unknown => (value instanceof Date ? fromJsDate(value, 'local') : value);
    const theirHeaders = theirs.length > 0 ? Object.keys(theirs[0] as object) : [];
    const ourHeaders = ours.rows.length > 0 ? Object.keys(ours.rows[0] as object) : [];
    const headersMatch = JSON.stringify(theirHeaders) === JSON.stringify(ourHeaders);
    let mismatches = Math.abs(theirs.length - ours.rows.length);
    const known: Record<string, number> = {};
    const samples: unknown[] = [];
    const count = Math.min(theirs.length, ours.rows.length);
    for (let index = 0; index < count; index++) {
      const theirRow = theirs[index] as Record<string, unknown>;
      const ourRow = ours.rows[index] as Record<string, unknown>;
      for (const key of new Set([...Object.keys(theirRow), ...Object.keys(ourRow)])) {
        const theirValue = normalize(theirRow[key]);
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
          known[reason] = (known[reason] ?? 0) + 1;
          continue;
        }
        mismatches++;
        if (samples.length < SAMPLE_MISMATCHES) {
          samples.push({ row: index, key, sheetjs: theirValue, ours: ourValue });
        }
      }
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
    ...summarize(compareDumps({ sheets: written }, reread, new Set(['truncate-32767'])), writerConvention),
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
      } catch (error) {
        result.sheetjs = {
          ok: false,
          error: String((error as Error)?.message ?? error).slice(0, 300),
          ms: Math.round(performance.now() - start),
        };
      }
    });
  }

  if (ours && sheetjs) {
    const theirs = sheetjsDump(sheetjs);
    result.parity = await timed(result, 'compare', async () => summarize(compareDumps(theirs, ours), knownSheetjsDivergence));
    const workbook = sheetjs;
    try {
      result.objects = await timed(result, 'objects', () => objectsParity(bytes, workbook));
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
