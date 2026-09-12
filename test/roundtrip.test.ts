/**
 * Write with our writer, read with our reader: every canonical value survives (after the documented 32,767
 * truncation policy), across both string strategies and both date modes.
 */
import { describe, expect, it } from 'vitest';
import { canonicalWorkbook } from '../fixtures/canonical/canonical.mjs';
import { writeCanonicalWorkbook } from './helpers/canonical-writer';
import { compareDumps, describeReport } from './helpers/diff';
import { dumpWorkbook } from './helpers/dump';
import { READER_READY, WRITER_READY } from './helpers/ready';

const TRUNCATE = new Set(['truncate-32767']);

describe.skipIf(!WRITER_READY || !READER_READY)('round trip: canonical workbook through our writer and reader', () => {
  for (const strings of ['auto', 'inline', 'shared'] as const) {
    it(`strings: ${strings}`, async () => {
      const bytes = await writeCanonicalWorkbook({ strings });
      const report = compareDumps(canonicalWorkbook(), await dumpWorkbook(bytes), TRUNCATE);
      // formulas are written as their cached values; hyperlink/note cells are plain text
      const unexpected = Object.keys(report.byCategory).filter(category => category !== 'formula-dropped');
      expect(unexpected, describeReport(report)).toEqual([]);
    });
  }

  it('crossing the shared-string budget mid-sheet keeps every value', async () => {
    const bytes = await writeCanonicalWorkbook({ strings: 'auto', sstBudget: { maxUnique: 8, maxChars: 200, maxLength: 12 } });
    const report = compareDumps(canonicalWorkbook(), await dumpWorkbook(bytes), TRUNCATE);
    const unexpected = Object.keys(report.byCategory).filter(category => category !== 'formula-dropped');
    expect(unexpected, describeReport(report)).toEqual([]);
  });

  it('date1904 workbooks read back to the same wall clock', async () => {
    const bytes = await writeCanonicalWorkbook({ date1904: true });
    const report = compareDumps(canonicalWorkbook(), await dumpWorkbook(bytes), TRUNCATE);
    const unexpected = Object.keys(report.byCategory).filter(category => category !== 'formula-dropped');
    expect(unexpected, describeReport(report)).toEqual([]);
  });
});
