/**
 * Every `kind:hostile` fixture must be rejected with a classified `XlsxError` (never a bare Error, never a hang,
 * never unbounded memory) either at open or while streaming a sheet.
 */
import { describe, expect, it } from 'vitest';
import { isXlsxError, openWorkbook, type XlsxErrorCode } from '../src/index';
import { fixturesWithTag, readFixture } from './helpers/fixtures';
import { READER_READY } from './helpers/ready';

/** The manifest records reader-agnostic error names; map them to our codes (several are acceptable for some). */
const EXPECTED_CODES: Record<string, XlsxErrorCode[]> = {
  XML_DOCTYPE: ['XML_DOCTYPE'],
  TRUNCATED: ['ZIP_TRUNCATED'],
  CRC_MISMATCH: ['ZIP_CRC_MISMATCH'],
  DUPLICATE_ENTRY: ['ZIP_DUPLICATE_ENTRY'],
  LIMIT_EXCEEDED: ['LIMIT_EXCEEDED'],
  NOT_XLSX: ['NOT_XLSX', 'ODS', 'LEGACY_XLS', 'XLSB'],
  ENCRYPTED: ['ENCRYPTED'],
  ZIP_BOMB: ['ZIP_BOMB'],
};

async function openAndStreamEverything(bytes: Uint8Array): Promise<void> {
  const workbook = await openWorkbook(bytes, { limits: { maxInflatedBytes: 8 * 1024 * 1024, maxSharedStringChars: 4 * 1024 * 1024 } });
  try {
    for (const info of workbook.sheets) {
      for await (const row of workbook.sheet(info.index).rows()) {
        void row;
      }
    }
  } finally {
    await workbook.close();
  }
}

describe.skipIf(!READER_READY)('hostile: every hostile fixture is rejected with a classified error', () => {
  for (const fixture of fixturesWithTag('kind:hostile')) {
    it(`${fixture.id} -> ${fixture.expectedError}`, async () => {
      const acceptable = EXPECTED_CODES[fixture.expectedError ?? ''];
      expect(acceptable, `manifest error ${fixture.expectedError} has no mapping`).toBeDefined();
      let caught: unknown;
      try {
        await openAndStreamEverything(readFixture(fixture));
      } catch (thrown) {
        caught = thrown;
      }
      expect(caught, 'expected a rejection').toBeDefined();
      expect(isXlsxError(caught), `expected XlsxError, got ${String(caught)}`).toBe(true);
      if (isXlsxError(caught)) {
        expect(acceptable).toContain(caught.code);
        if (caught.code === 'ENCRYPTED') {
          expect(caught.message).toContain('password-protected');
        }
      }
    });
  }
});
