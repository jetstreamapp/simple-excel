/**
 * Jetstream reads spreadsheets with SheetJS `sheet_to_json` (dateNF, defval '', blankrows false, rawNumbers). Our
 * object mode must produce the same rows for every golden, except cells listed as `parityExemptions` in
 * `test/corpus-policies.json` (documented SheetJS defects such as serial 60 or undecoded escapes).
 */
import { describe, expect, it } from 'vitest';
import * as XLSX from 'xlsx';
import { openWorkbook } from '../src/index';
import type { CellValue } from '../src/types';
import { corpusPolicies, fixturesWithTag, readFixture } from './helpers/fixtures';
import { READER_READY } from './helpers/ready';
import { a1, fromJsDate, type TypedValue } from './helpers/typed';

const policies = corpusPolicies();

function typed(value: unknown): TypedValue {
  if (value instanceof Date) {
    return fromJsDate(value, 'local');
  }
  return value as TypedValue;
}

describe.skipIf(!READER_READY)('parity: object mode equals SheetJS sheet_to_json with Jetstream options', () => {
  for (const fixture of fixturesWithTag('kind:golden')) {
    const policy = policies[fixture.id];
    if (policy?.skip) {
      it.skip(`${fixture.id} (${policy.skip})`, () => {});
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
              if (exemptions.has(cell)) {
                return;
              }
              expect(typed(row[header]), cell).toEqual(typed(expectedRow[header]));
            });
          });
        }
      } finally {
        await workbook.close();
      }
    });
  }
});
