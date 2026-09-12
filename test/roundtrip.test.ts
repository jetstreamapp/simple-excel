/**
 * Write with our writer, read with our reader: every canonical value survives (after the documented 32,767
 * truncation policy), across both string strategies and both date modes.
 */
import { describe, expect, it } from 'vitest';
import { canonicalWorkbook } from '../fixtures/canonical/canonical.mjs';
import { writeCanonicalWorkbook } from './helpers/canonical-writer';
import { compareDumps, describeReport, type CellMismatch } from './helpers/diff';
import { dumpWorkbook } from './helpers/dump';
import { READER_READY, WRITER_READY } from './helpers/ready';

const TRUNCATE = new Set(['truncate-32767']);

/**
 * What a round trip cannot restore, all of it the writer's documented policy: formulas are written as their cached
 * value, hyperlink and note cells as plain text, a date before 1900 as ISO text (ADR-003) and a `Date` that falls in
 * a DST gap an hour late (EC-DATE-DST-GAP). `test/corpus-policies.json` allows the same two date categories for the
 * `golden-canonical-simple-excel` fixtures, where sheetjs, openpyxl and calamine report them on the same bytes.
 */
const WRITER_CONVENTIONS = new Set(['formula-dropped', 'date-as-text', 'dst-gap-shift-1h']);

/**
 * A time of day is a bare fraction of a day in both date systems, so a 1904 workbook reads one back as that time on
 * 1904-01-01: the wall clock survives, the epoch day it hangs on cannot (Excel does the same with the same bytes).
 * Matched cell by cell rather than by category, so a real 1904 date bug still fails the round trip.
 */
function isTimeOnlyOnEpochDay(mismatch: CellMismatch): boolean {
  const time = (mismatch.expected as { $time?: unknown })?.$time;
  if (typeof time !== 'string') {
    return false;
  }
  const actual = mismatch.actual as { $date?: unknown; $datetime?: unknown };
  return actual?.$datetime === `1904-01-01T${time}` || (actual?.$date === '1904-01-01' && time === '00:00:00.000');
}

describe.skipIf(!WRITER_READY || !READER_READY)('round trip: canonical workbook through our writer and reader', () => {
  for (const strings of ['auto', 'inline', 'shared'] as const) {
    it(`strings: ${strings}`, async () => {
      const bytes = await writeCanonicalWorkbook({ strings });
      const report = compareDumps(canonicalWorkbook(), await dumpWorkbook(bytes), TRUNCATE);
      const unexpected = Object.keys(report.byCategory).filter(category => !WRITER_CONVENTIONS.has(category));
      expect(unexpected, describeReport(report)).toEqual([]);
    });
  }

  it('crossing the shared-string budget mid-sheet keeps every value', async () => {
    const bytes = await writeCanonicalWorkbook({ strings: 'auto', sstBudget: { maxUnique: 8, maxChars: 200, maxLength: 12 } });
    const report = compareDumps(canonicalWorkbook(), await dumpWorkbook(bytes), TRUNCATE);
    const unexpected = Object.keys(report.byCategory).filter(category => !WRITER_CONVENTIONS.has(category));
    expect(unexpected, describeReport(report)).toEqual([]);
  });

  it('date1904 workbooks read back to the same wall clock', async () => {
    const bytes = await writeCanonicalWorkbook({ date1904: true });
    const report = compareDumps(canonicalWorkbook(), await dumpWorkbook(bytes), TRUNCATE);
    const unexpected = report.sheets
      .flatMap(sheet => sheet.mismatches)
      .filter(mismatch => !WRITER_CONVENTIONS.has(mismatch.category) && !isTimeOnlyOnEpochDay(mismatch));
    expect(unexpected, describeReport(report)).toEqual([]);
  });
});
