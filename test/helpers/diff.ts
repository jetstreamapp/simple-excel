/**
 * Typed-dump comparison, ported from `oracle/diff.mjs` so the corpus suite classifies mismatches with the same
 * vocabulary the compatibility matrix uses. Keep the two in sync.
 */
import { EXCEL_MAX_CELL_CHARS, truncateForExcel } from '../../fixtures/canonical/canonical.mjs';
import { a1, isTyped, normalizeRows, type TypedDump, type TypedValue } from './typed';

export type MismatchCategory = string;

export interface CellMismatch {
  cell: string;
  category: MismatchCategory;
  expected: unknown;
  actual: unknown;
}

export interface SheetReport {
  name: string;
  present: boolean;
  cells: number;
  matches: number;
  mismatches: CellMismatch[];
}

export interface DiffReport {
  pass: boolean;
  sheets: SheetReport[];
  totals: { cells: number; matches: number; mismatches: number };
  byCategory: Record<string, number>;
  sheetRenamed?: { expected: string; actual: string };
}

export interface ExpectedWorkbook {
  sheets: { name: string; hidden?: boolean; rows: TypedValue[][] }[];
}

/** Value of a typed-encoding key (`$error`, `$date`, ...) or undefined when the value is not that shape. */
function field(value: TypedValue, key: string): string | undefined {
  return isTyped(value, key) ? String((value as Record<string, unknown>)[key]) : undefined;
}

function unwrapFormula(value: TypedValue): TypedValue {
  return isTyped(value, '$formula') ? (value as { cached: TypedValue }).cached : value;
}

function sameNumber(a: number, b: number): boolean {
  if (Object.is(a, b) || (a === 0 && b === 0)) {
    return true;
  }
  const scale = Math.max(Math.abs(a), Math.abs(b), 1e-300);
  return Math.abs(a - b) / scale < 1e-15;
}

const TEMPORAL_KEYS = ['$date', '$datetime', '$time'] as const;

function temporalKey(value: TypedValue): (typeof TEMPORAL_KEYS)[number] | undefined {
  return TEMPORAL_KEYS.find(key => isTyped(value, key));
}

function temporalText(value: TypedValue, key: string): string {
  return field(value, key) ?? '';
}

/** Returns null when the two cells are equal, else a classification string. */
export function classify(expected: TypedValue, actual: TypedValue, policies: ReadonlySet<string>): MismatchCategory | null {
  const e = unwrapFormula(expected);
  const a = unwrapFormula(actual);
  if (isTyped(expected, '$formula') && !isTyped(actual, '$formula') && a === null) {
    return 'formula-dropped';
  }
  if (e === a) {
    return null;
  }
  if (isTyped(e, '$error') && isTyped(a, '$error')) {
    return field(e, '$error') === field(a, '$error') ? null : 'error-mismatch';
  }
  if (isTyped(e, '$datetime') && isTyped(a, '$date') && field(e, '$datetime') === `${field(a, '$date')}T00:00:00.000`) {
    return null;
  }
  if (isTyped(e, '$date') && isTyped(a, '$datetime') && field(a, '$datetime') === `${field(e, '$date')}T00:00:00.000`) {
    return null;
  }
  if (e === null && a === '') {
    return 'blank-vs-empty-string';
  }
  if (e === '' && a === null) {
    return 'empty-string-vs-blank';
  }
  if (e === null || a === null) {
    return a === null ? 'cell-missing' : 'unexpected-cell';
  }
  if (typeof e === 'number' && typeof a === 'number') {
    return sameNumber(e, a) ? null : 'number-precision';
  }
  if (typeof e === 'number' && typeof a === 'string') {
    return Number(a) === e ? 'number-as-text' : 'type-mismatch';
  }
  if (typeof e === 'string' && typeof a === 'number') {
    return 'text-as-number';
  }
  if (typeof e === 'boolean' && typeof a !== 'boolean') {
    return String(a).toUpperCase() === String(e).toUpperCase() || a === (e ? 1 : 0) ? 'boolean-as-text-or-number' : 'type-mismatch';
  }
  if (typeof e === 'string' && typeof a === 'string') {
    if (policies.has('truncate-32767') && e.length > EXCEL_MAX_CELL_CHARS && a === truncateForExcel(e)) {
      return null;
    }
    if (a === e.trim()) {
      return 'whitespace-trimmed';
    }
    // oxlint-disable-next-line no-control-regex -- matching control characters is the point here
    if (a === e.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '')) {
      return 'control-chars-stripped';
    }
    if (a.replace(/\r\n/g, '\n') === e.replace(/\r\n/g, '\n')) {
      return 'crlf-normalized';
    }
    if (e.length > EXCEL_MAX_CELL_CHARS && a.length <= EXCEL_MAX_CELL_CHARS) {
      return 'truncated-differently';
    }
    if (/_x[0-9A-Fa-f]{4}_/.test(e) || /_x[0-9A-Fa-f]{4}_/.test(a)) {
      return 'escape-sequence-mangled';
    }
    return 'string-mismatch';
  }
  if (isTyped(e, '$error') || isTyped(a, '$error')) {
    if (isTyped(e, '$error') && typeof a === 'string' && a === field(e, '$error')) {
      return 'error-as-text';
    }
    return 'error-mismatch';
  }
  const eTemporal = temporalKey(e);
  const aTemporal = temporalKey(a);
  if (eTemporal && aTemporal) {
    const eText = temporalText(e, eTemporal);
    const aText = temporalText(a, aTemporal);
    if (eText === aText) {
      return null;
    }
    if (eTemporal === '$time' && aTemporal === '$datetime' && /^1899-12-3[01]T/.test(aText) && aText.slice(11) === eText) {
      return 'time-only-has-date-part';
    }
    if (eTemporal === '$datetime' && aTemporal === '$datetime') {
      const deltaMs = Date.parse(`${aText}Z`) - Date.parse(`${eText}Z`);
      if (Math.abs(deltaMs) === 3600000) {
        return 'dst-gap-shift-1h';
      }
      if (Math.abs(deltaMs) < 3600000 * 24 && deltaMs % 60000 !== 0) {
        return 'historical-tz-offset-shift';
      }
    }
    if (eTemporal === '$date' && aTemporal === '$datetime') {
      const deltaMs = Date.parse(`${aText}Z`) - Date.parse(`${eText}T00:00:00.000Z`);
      if (Math.abs(deltaMs) < 3600000 * 24 && deltaMs !== 0) {
        return 'date-tz-offset-shift';
      }
      if (aText.startsWith(eText)) {
        return aText.endsWith('T00:00:00.000') ? null : 'date-gained-time';
      }
    }
    if (eTemporal === '$datetime' && aTemporal === '$datetime' && eText.slice(0, 19) === aText.slice(0, 19)) {
      return 'datetime-ms-lost';
    }
    return 'temporal-mismatch';
  }
  if (eTemporal && typeof a === 'number') {
    return 'date-as-serial';
  }
  if (eTemporal && typeof a === 'string') {
    return 'date-as-text';
  }
  if (aTemporal && typeof e === 'string') {
    return 'text-as-date';
  }
  return 'mismatch';
}

function summarize(value: unknown): unknown {
  return typeof value === 'string' && value.length > 60 ? `${value.slice(0, 57)}... (${value.length} chars)` : value;
}

export function compareDumps(expected: ExpectedWorkbook, actual: TypedDump, policies: ReadonlySet<string> = new Set()): DiffReport {
  const report: DiffReport = { pass: false, sheets: [], totals: { cells: 0, matches: 0, mismatches: 0 }, byCategory: {} };
  const count = (category: string): void => {
    report.byCategory[category] = (report.byCategory[category] ?? 0) + 1;
    report.totals.mismatches++;
  };
  const actualByName = new Map(actual.sheets.map(sheet => [sheet.name, sheet]));
  const firstExpected = expected.sheets[0];
  const firstActual = actual.sheets[0];
  if (expected.sheets.length === 1 && actual.sheets.length === 1 && firstExpected && firstActual && !actualByName.has(firstExpected.name)) {
    report.sheetRenamed = { expected: firstExpected.name, actual: firstActual.name };
    actualByName.set(firstExpected.name, firstActual);
  }
  for (const expSheet of expected.sheets) {
    const actSheet = actualByName.get(expSheet.name);
    const sheetReport: SheetReport = { name: expSheet.name, present: Boolean(actSheet), cells: 0, matches: 0, mismatches: [] };
    report.sheets.push(sheetReport);
    if (!actSheet) {
      sheetReport.mismatches.push({ cell: '', category: 'sheet-missing', expected: expSheet.name, actual: null });
      count('sheet-missing');
      continue;
    }
    if (expSheet.hidden !== undefined && expSheet.hidden !== actSheet.hidden) {
      sheetReport.mismatches.push({ cell: '', category: 'hidden-state', expected: expSheet.hidden, actual: actSheet.hidden });
      count('hidden-state');
    }
    const expRows = normalizeRows(expSheet.rows);
    const actRows = normalizeRows(actSheet.rows);
    const rowCount = Math.max(expRows.length, actRows.length);
    for (let r = 0; r < rowCount; r++) {
      const expRow = expRows[r] ?? [];
      const actRow = actRows[r] ?? [];
      const colCount = Math.max(expRow.length, actRow.length);
      for (let c = 0; c < colCount; c++) {
        const exp = expRow[c] ?? null;
        const act = actRow[c] ?? null;
        sheetReport.cells++;
        report.totals.cells++;
        const category = classify(exp, act, policies);
        if (category === null) {
          sheetReport.matches++;
          report.totals.matches++;
        } else {
          sheetReport.mismatches.push({ cell: a1(r, c), category, expected: summarize(exp), actual: summarize(act) });
          count(category);
        }
      }
    }
  }
  for (const actSheet of actual.sheets) {
    if (report.sheetRenamed && actSheet.name === report.sheetRenamed.actual) {
      continue;
    }
    if (!expected.sheets.some(sheet => sheet.name === actSheet.name)) {
      report.sheets.push({
        name: actSheet.name,
        present: true,
        cells: 0,
        matches: 0,
        mismatches: [{ cell: '', category: 'sheet-unexpected', expected: null, actual: actSheet.name }],
      });
      count('sheet-unexpected');
    }
  }
  report.pass = report.totals.mismatches === 0;
  return report;
}

/** Human-readable summary of the first mismatches, for assertion messages. */
export function describeReport(report: DiffReport, limit = 15): string {
  const lines: string[] = [`${report.totals.matches}/${report.totals.cells} cells match; categories ${JSON.stringify(report.byCategory)}`];
  for (const sheet of report.sheets) {
    for (const mismatch of sheet.mismatches.slice(0, limit)) {
      lines.push(
        `  ${sheet.name}!${mismatch.cell} ${mismatch.category}: expected ${JSON.stringify(mismatch.expected)} got ${JSON.stringify(mismatch.actual)}`,
      );
    }
    if (sheet.mismatches.length > limit) {
      lines.push(`  ${sheet.name}: ... ${sheet.mismatches.length - limit} more`);
    }
  }
  return lines.join('\n');
}
