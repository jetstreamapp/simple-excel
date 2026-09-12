/**
 * Jetstream reads spreadsheets with SheetJS `sheet_to_json` (dateNF, defval '', blankrows false, rawNumbers). Our
 * object mode must produce the same rows for every golden, except cells listed as `parityExemptions` in
 * `test/corpus-policies.json` (documented SheetJS defects such as serial 60 or undecoded escapes) and the two
 * whole-corpus date conventions in `sameValue` below.
 */
import { describe, expect, it } from 'vitest';
import * as XLSX from 'xlsx';
import { openWorkbook } from '../src/index';
import type { CellValue } from '../src/types';
import { corpusPolicies, fixturesWithTag, readFixture } from './helpers/fixtures';
import { READER_READY } from './helpers/ready';
import { a1, columnLetter, fromJsDate, isTyped, type TypedValue } from './helpers/typed';

const policies = corpusPolicies();

/** How far apart two datetimes may be and still be the same instant: SheetJS truncates where we round to the ms. */
const MILLISECOND_TOLERANCE = 1;

function typed(value: unknown): TypedValue {
  if (value instanceof Date) {
    return fromJsDate(value, 'local');
  }
  return value as TypedValue;
}

function temporalText(value: TypedValue, key: string): string | undefined {
  return isTyped(value, key) ? String((value as Record<string, unknown>)[key]) : undefined;
}

/**
 * The time of day of a value that carries no real date: a serial below 1. We hang those on 1899-12-30 and SheetJS on
 * 1899-12-31 (Excel's "Jan 0 1900"), and the catalog blesses both markers and tells consumers to accept either
 * (EC-TIME-ONLY-HAS-DATE-PART), so the two spellings of the same time of day compare equal here.
 */
function timeOnly(value: TypedValue): string | undefined {
  const time = temporalText(value, '$time');
  if (time !== undefined) {
    return time;
  }
  const dateTime = temporalText(value, '$datetime');
  if (dateTime !== undefined && dateTime.startsWith('1899-12-31T')) {
    return dateTime.slice(11);
  }
  return temporalText(value, '$date') === '1899-12-31' ? '00:00:00.000' : undefined;
}

/** Milliseconds between two `$datetime` values, or null when they are not both datetimes. */
function millisecondsApart(ours: TypedValue, theirs: TypedValue): number | null {
  const ourText = temporalText(ours, '$datetime');
  const theirText = temporalText(theirs, '$datetime');
  if (ourText === undefined || theirText === undefined) {
    return null;
  }
  const difference = Date.parse(`${ourText}Z`) - Date.parse(`${theirText}Z`);
  return Number.isNaN(difference) ? null : Math.abs(difference);
}

/**
 * Equal for parity purposes. Beyond deep equality, two documented date conventions run through the whole corpus and
 * would otherwise need an exemption on hundreds of cells: the time-only marker day, and SheetJS truncating the
 * serial-to-milliseconds conversion where we round to the nearest millisecond (`datetime-ms-lost` in the corpus
 * comparator). Everything else has to be listed per fixture.
 */
function sameValue(ours: TypedValue, theirs: TypedValue): boolean {
  if (JSON.stringify(ours) === JSON.stringify(theirs)) {
    return true;
  }
  const ourTime = timeOnly(ours);
  if (ourTime !== undefined && ourTime === timeOnly(theirs)) {
    return true;
  }
  const apart = millisecondsApart(ours, theirs);
  return apart !== null && apart <= MILLISECOND_TOLERANCE;
}

describe.skipIf(!READER_READY)('parity: object mode equals SheetJS sheet_to_json with Jetstream options', () => {
  for (const fixture of fixturesWithTag('kind:golden')) {
    const policy = policies[fixture.id];
    const skipped = policy?.skip ?? policy?.paritySkip;
    if (skipped) {
      it.skip(`${fixture.id} (${skipped})`, () => {});
      continue;
    }
    it(fixture.id, async () => {
      const bytes = readFixture(fixture);
      const reference = XLSX.read(bytes, { cellText: false, cellDates: true, type: 'array' });
      const exemptions = new Set(policy?.parityExemptions ?? []);
      const workbook = await openWorkbook(bytes);
      try {
        for (const info of workbook.sheets) {
          if (info.kind !== 'worksheet') {
            continue;
          }
          const referenceSheet = reference.Sheets[info.name];
          expect(referenceSheet, `SheetJS did not see sheet ${info.name}`).toBeDefined();
          const expectedRows = XLSX.utils.sheet_to_json<Record<string, unknown>>(referenceSheet!, {
            dateNF: 'yyyy"-"mm"-"dd"T"hh:mm:ss',
            defval: '',
            blankrows: false,
            rawNumbers: true,
          });
          const { rows, headers } = await workbook.sheet(info.index).toObjects();
          const expectedHeaders = expectedRows.length > 0 ? Object.keys(expectedRows[0] ?? {}) : [];
          expect(headers, `${info.name} headers`).toEqual(expectedHeaders);
          expect(rows.length, `${info.name} row count`).toBe(expectedRows.length);
          rows.forEach((row: Record<string, CellValue>, rowIndex: number) => {
            const expectedRow = expectedRows[rowIndex] ?? {};
            headers.forEach((header: string, columnIndex: number) => {
              const cell = `${info.name}!${a1(rowIndex + 1, columnIndex)}`;
              // An exemption is a cell (`Data!F7`) or, where a whole column of the canonical dataset probes the same
              // quirk, the column it lives in (`Data!F`).
              if (exemptions.has(cell) || exemptions.has(`${info.name}!${columnLetter(columnIndex)}`)) {
                return;
              }
              const ours = typed(row[header]);
              const theirs = typed(expectedRow[header]);
              if (!sameValue(ours, theirs)) {
                expect(ours, cell).toEqual(theirs);
              }
            });
          });
        }
      } finally {
        await workbook.close();
      }
    });
  }
});
