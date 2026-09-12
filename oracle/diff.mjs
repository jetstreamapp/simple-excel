#!/usr/bin/env node
// Compare a reader's typed dump against the expected workbook (canonical.json or another dump).
//
//   node diff.mjs <expected.json> <actual-dump.json> [--policy truncate-32767] [--sheets Data,Features] [--out report.json]
//
// Exit code 0 when every compared cell matches (after policies), 1 otherwise. Mismatches are classified so the
// edge-case catalog can tell "reader bug" from "generator limitation" from "documented policy".
import { readFileSync, writeFileSync } from 'node:fs';
import { parseArgs } from 'node:util';
import { EXCEL_MAX_CELL_CHARS, isTyped, truncateForExcel } from '../fixtures/canonical/canonical.mjs';
import { a1, normalizeRows } from './lib/typed.mjs';

const { values: options, positionals } = parseArgs({
  allowPositionals: true,
  options: {
    policy: { type: 'string', multiple: true, default: [] },
    sheets: { type: 'string' },
    out: { type: 'string' },
    quiet: { type: 'boolean', default: false },
  },
});
const [expectedPath, actualPath] = positionals;
if (!expectedPath || !actualPath) {
  console.error('usage: diff.mjs <expected.json> <actual-dump.json> [--policy truncate-32767] [--sheets a,b] [--out report.json]');
  process.exit(2);
}
const expected = JSON.parse(readFileSync(expectedPath, 'utf8'));
const actual = JSON.parse(readFileSync(actualPath, 'utf8'));
const policies = new Set(options.policy);
const onlySheets = options.sheets ? new Set(options.sheets.split(',')) : null;

function unwrapFormula(value) {
  return isTyped(value, '$formula') ? value.cached : value;
}

function sameNumber(a, b) {
  if (Object.is(a, b)) {
    return true;
  }
  if (a === 0 && b === 0) {
    return true; // -0 vs 0: Excel has no negative zero
  }
  const scale = Math.max(Math.abs(a), Math.abs(b), 1e-300);
  return Math.abs(a - b) / scale < 1e-15;
}

/** Returns null when equal, else a classification string. */
function classify(exp, act) {
  const e = unwrapFormula(exp);
  const a = unwrapFormula(act);
  if (isTyped(exp, '$formula') && !isTyped(act, '$formula') && a === null) {
    return 'formula-dropped';
  }
  if (e === a) {
    return null;
  }
  if (isTyped(e, '$error') && isTyped(a, '$error')) {
    return e.$error === a.$error ? null : 'error-mismatch';
  }
  // a datetime at exactly midnight and a date-only value are the same cell to Excel
  if (isTyped(e, '$datetime') && isTyped(a, '$date') && e.$datetime === `${a.$date}T00:00:00.000`) {
    return null;
  }
  if (isTyped(e, '$date') && isTyped(a, '$datetime') && a.$datetime === `${e.$date}T00:00:00.000`) {
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
    if (isTyped(e, '$error') && typeof a === 'string' && a === e.$error) {
      return 'error-as-text';
    }
    return 'error-mismatch';
  }
  const temporalKeys = ['$date', '$datetime', '$time'];
  const eTemporal = temporalKeys.find(key => isTyped(e, key));
  const aTemporal = temporalKeys.find(key => isTyped(a, key));
  if (eTemporal && aTemporal) {
    if (e[eTemporal] === a[aTemporal]) {
      return null;
    }
    // time-only values come back with a date part (1899-12-30 or Excel's "Jan 0 1900" = 1899-12-31) in some readers
    if (eTemporal === '$time' && aTemporal === '$datetime' && /^1899-12-3[01]T/.test(a.$datetime) && a.$datetime.slice(11) === e.$time) {
      return 'time-only-has-date-part';
    }
    if (eTemporal === '$datetime' && aTemporal === '$datetime') {
      const deltaMs = Date.parse(`${a.$datetime}Z`) - Date.parse(`${e.$datetime}Z`);
      if (Math.abs(deltaMs) === 3600000) {
        return 'dst-gap-shift-1h';
      }
      if (Math.abs(deltaMs) < 3600000 * 24 && deltaMs % 60000 !== 0) {
        return 'historical-tz-offset-shift';
      }
    }
    if (eTemporal === '$date' && aTemporal === '$datetime') {
      const deltaMs = Date.parse(`${a.$datetime}Z`) - Date.parse(`${e.$date}T00:00:00.000Z`);
      if (Math.abs(deltaMs) < 3600000 * 24 && deltaMs !== 0) {
        return 'date-tz-offset-shift';
      }
    }
    if (eTemporal === '$date' && aTemporal === '$datetime' && a.$datetime.startsWith(e.$date)) {
      return a.$datetime.endsWith('T00:00:00.000') ? null : 'date-gained-time';
    }
    if (eTemporal === '$datetime' && aTemporal === '$datetime' && e.$datetime.slice(0, 19) === a.$datetime.slice(0, 19)) {
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

const report = {
  expected: expectedPath,
  actual: actualPath,
  reader: actual.reader,
  policies: [...policies],
  sheets: [],
  totals: { cells: 0, matches: 0, mismatches: 0 },
  byCategory: {},
};
const actualByName = new Map(actual.sheets.map(sheet => [sheet.name, sheet]));
// A single-sheet workbook derived from a CSV is named after the file by the importing application
// ('canonical', 'canonical.csv', 'Sheet1'...): compare by position and record the rename instead of failing.
if (expected.sheets.length === 1 && actual.sheets.length === 1 && !actualByName.has(expected.sheets[0].name)) {
  report.sheetRenamed = { expected: expected.sheets[0].name, actual: actual.sheets[0].name };
  actualByName.set(expected.sheets[0].name, actual.sheets[0]);
}

for (const expSheet of expected.sheets) {
  if (onlySheets && !onlySheets.has(expSheet.name)) {
    continue;
  }
  const actSheet = actualByName.get(expSheet.name);
  const sheetReport = { name: expSheet.name, present: Boolean(actSheet), cells: 0, matches: 0, mismatches: [] };
  report.sheets.push(sheetReport);
  if (!actSheet) {
    sheetReport.mismatches.push({ category: 'sheet-missing' });
    report.byCategory['sheet-missing'] = (report.byCategory['sheet-missing'] ?? 0) + 1;
    report.totals.mismatches++;
    continue;
  }
  if (expSheet.hidden !== undefined && actSheet.hidden !== null && actSheet.hidden !== undefined && expSheet.hidden !== actSheet.hidden) {
    sheetReport.mismatches.push({ category: 'hidden-state', expected: expSheet.hidden, actual: actSheet.hidden });
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
      const category = classify(exp, act);
      if (category === null) {
        sheetReport.matches++;
        report.totals.matches++;
      } else {
        const summarize = value =>
          typeof value === 'string' && value.length > 60 ? `${value.slice(0, 57)}... (${value.length} chars)` : value;
        sheetReport.mismatches.push({ cell: a1(r, c), category, expected: summarize(exp), actual: summarize(act) });
        report.byCategory[category] = (report.byCategory[category] ?? 0) + 1;
        report.totals.mismatches++;
      }
    }
  }
  if (expSheet.features) {
    const structural = {};
    const feature = expSheet.features;
    if (actSheet.merges) {
      structural.merges =
        actSheet.merges.length === feature.merges.length ? 'ok' : `expected ${feature.merges.length}, got ${actSheet.merges.length}`;
    }
    if (actSheet.autoFilter !== undefined) {
      structural.autoFilter =
        actSheet.autoFilter === feature.autoFilter ? 'ok' : `expected ${feature.autoFilter}, got ${actSheet.autoFilter}`;
    }
    if (actSheet.hyperlinks) {
      structural.hyperlink = actSheet.hyperlinks.some(link => link.url === feature.hyperlink.url) ? 'ok' : 'missing';
    }
    if (actSheet.comments) {
      structural.note = actSheet.comments.some(comment => comment.text?.includes(feature.note.text)) ? 'ok' : 'missing';
    }
    if (actSheet.hiddenRows) {
      structural.hiddenRows = feature.hiddenRows.every(row => actSheet.hiddenRows.includes(row)) ? 'ok' : 'missing';
    }
    if (actSheet.hiddenColumns) {
      structural.hiddenColumns = feature.hiddenColumns.every(col => actSheet.hiddenColumns.includes(col)) ? 'ok' : 'missing';
    }
    if (actSheet.freeze !== undefined) {
      structural.freeze = actSheet.freeze ? 'ok' : 'missing';
    }
    sheetReport.structural = structural;
  }
}
for (const actSheet of actual.sheets) {
  if (report.sheetRenamed && actSheet.name === report.sheetRenamed.actual) {
    continue;
  }
  if (!expected.sheets.some(sheet => sheet.name === actSheet.name) && !(onlySheets && !onlySheets.has(actSheet.name))) {
    report.sheets.push({ name: actSheet.name, present: true, extra: true, mismatches: [{ category: 'sheet-unexpected' }] });
    report.byCategory['sheet-unexpected'] = (report.byCategory['sheet-unexpected'] ?? 0) + 1;
  }
}
report.pass = report.totals.mismatches === 0 && !report.byCategory['sheet-missing'];

if (options.out) {
  writeFileSync(options.out, JSON.stringify(report, null, 2) + '\n');
}
if (!options.quiet) {
  const { cells, matches, mismatches } = report.totals;
  console.log(
    `${report.pass ? 'PASS' : 'FAIL'} ${actual.reader}: ${matches}/${cells} cells match, ${mismatches} mismatches ${JSON.stringify(report.byCategory)}`,
  );
  for (const sheet of report.sheets) {
    for (const mismatch of sheet.mismatches.slice(0, 12)) {
      console.log(
        `  ${sheet.name}!${mismatch.cell ?? ''} ${mismatch.category}: expected ${JSON.stringify(mismatch.expected)} got ${JSON.stringify(mismatch.actual)}`,
      );
    }
    if (sheet.mismatches.length > 12) {
      console.log(`  ${sheet.name}: ... ${sheet.mismatches.length - 12} more`);
    }
    if (sheet.structural) {
      console.log(`  ${sheet.name} structural: ${JSON.stringify(sheet.structural)}`);
    }
  }
}
process.exit(report.pass ? 0 : 1);
