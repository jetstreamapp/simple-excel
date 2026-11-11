/**
 * Every `kind:hostile` fixture must fail with a classified `XlsxError` (never a bare Error, never a hang, never
 * unbounded memory) either at open or while streaming a sheet, when the manifest names an `expectedError`. The few
 * hostile fixtures without one are legal workbooks shaped to exhaust memory; the right outcome for those is a read to
 * the end in bounded memory.
 */
import { setFlagsFromString } from 'node:v8';
import { runInNewContext } from 'node:vm';
import { describe, expect, it } from 'vitest';
import { isXlsxError, openWorkbook, type ArrayRowsOptions, type OpenOptions, type XlsxErrorCode } from '../src/index';
import { fixturesWithTag, readFixture, type Fixture } from './helpers/fixtures';
import { READER_READY } from './helpers/ready';

/** The manifest records reader-agnostic error names; map them to our codes (several are acceptable for some). */
const EXPECTED_CODES: Record<string, XlsxErrorCode[]> = {
  XML_DOCTYPE: ['XML_DOCTYPE'],
  XML_MALFORMED: ['XML_MALFORMED'],
  TRUNCATED: ['ZIP_TRUNCATED'],
  CRC_MISMATCH: ['ZIP_CRC_MISMATCH'],
  DUPLICATE_ENTRY: ['ZIP_DUPLICATE_ENTRY'],
  LIMIT_EXCEEDED: ['LIMIT_EXCEEDED'],
  NOT_XLSX: ['NOT_XLSX', 'ODS', 'LEGACY_XLS', 'XLSB'],
  ENCRYPTED: ['ENCRYPTED'],
  ZIP_BOMB: ['ZIP_BOMB'],
};

/** The budget a host application sets for untrusted uploads. */
const TIGHT_LIMITS: OpenOptions['limits'] = {
  maxInflatedBytes: 8 * 1024 * 1024,
  maxPartBytes: 4 * 1024 * 1024,
  maxSharedStringChars: 4 * 1024 * 1024,
};

/** The reader DoS family: these must hold with the limits a caller gets by default, not only with tight ones. */
const DEFAULT_LIMIT_CATALOG_IDS: readonly string[] = [
  'EC-PACKAGE-ELEMENT-FLOOD',
  'EC-PART-SIZE-PADDING',
  'EC-SST-EMPTY-ITEM-FLOOD',
  'EC-CELL-TEXT-RUN-SPLIT',
  'EC-XML-LONG-TAG-RESCAN',
  'EC-WIDE-SPARSE-ROWS',
  'EC-BLANK-ROWS-GAP-FILL',
];

/**
 * Fixtures whose part inflates past the tight `maxInflatedBytes` before the limit they probe is reached: that cap
 * rejects them first under the tight limits, and the default-limit block below checks the limit itself.
 */
const TIGHT_LIMIT_CODES_BY_CATALOG_ID: Readonly<Record<string, XlsxErrorCode[]>> = {
  'EC-SST-EMPTY-ITEM-FLOOD': ['ZIP_BOMB'],
};

/** Read options that make a read-flat fixture hostile, by catalog id. */
const ROW_OPTIONS_BY_CATALOG_ID: Readonly<Record<string, ArrayRowsOptions>> = {
  'EC-BLANK-ROWS-GAP-FILL': { blankRows: true },
};

const MB = 1024 * 1024;
/**
 * Live heap a read-flat fixture may hold at any point of its read, measured after a full collection. A batch is
 * bounded at about 8 MB of row slots; before that budget, one inflate chunk of `wide-sparse-rows` held 40 MB of rows
 * and the `blank-rows-gap` gap over 100 MB.
 */
const READ_FLAT_LIVE_HEAP_BUDGET = 32 * MB;
/** A fixture that must fail fast does so in well under this, whatever machine runs the suite. */
const FAIL_FAST_MS = 10_000;

async function openAndStreamEverything(
  bytes: Uint8Array,
  options: OpenOptions = {},
  rowOptions: ArrayRowsOptions = {},
  onRow: (rows: number) => void = () => {},
): Promise<number> {
  const workbook = await openWorkbook(bytes, options);
  let rows = 0;
  try {
    for (const info of workbook.sheets) {
      for await (const row of workbook.sheet(info.index).rows(rowOptions)) {
        void row;
        onRow(++rows);
      }
    }
  } finally {
    await workbook.close();
  }
  return rows;
}

/** A full garbage collection, so heap samples measure what is live rather than what the collector has not reached. */
function garbageCollector(): () => void {
  if (typeof globalThis.gc === 'function') {
    return globalThis.gc;
  }
  setFlagsFromString('--expose-gc');
  return runInNewContext('gc') as () => void;
}

/** Sample often while the first batches go by, where an unbounded one shows, then sparsely: a full GC is not free. */
function shouldSample(rows: number): boolean {
  return rows <= 4096 ? rows % 32 === 0 : rows % 16_384 === 0;
}

async function rejection(operation: () => Promise<unknown>): Promise<unknown> {
  try {
    await operation();
  } catch (thrown) {
    return thrown;
  }
  return undefined;
}

/** `caught` must be the classified error the manifest names. The caller has checked that something was thrown. */
function expectClassified(fixture: Fixture, caught: unknown, acceptable = EXPECTED_CODES[fixture.expectedError ?? '']): void {
  expect(acceptable, `manifest error ${fixture.expectedError} has no mapping`).toBeDefined();
  expect(isXlsxError(caught), `expected XlsxError, got ${String(caught)}`).toBe(true);
  if (isXlsxError(caught)) {
    expect(acceptable).toContain(caught.code);
    if (caught.code === 'ENCRYPTED') {
      expect(caught.message).toContain('password-protected');
    }
  }
}

function rowOptionsFor(fixture: Fixture): ArrayRowsOptions {
  const catalogId = fixture.tags.find(tag => ROW_OPTIONS_BY_CATALOG_ID[tag] !== undefined);
  return catalogId === undefined ? {} : (ROW_OPTIONS_BY_CATALOG_ID[catalogId] ?? {});
}

describe.skipIf(!READER_READY)('hostile: every hostile fixture is rejected with a classified error', () => {
  for (const fixture of fixturesWithTag('kind:hostile').filter(candidate => candidate.expectedError !== undefined)) {
    it(`${fixture.id} -> ${fixture.expectedError}`, async () => {
      const caught = await rejection(() => openAndStreamEverything(readFixture(fixture), { limits: TIGHT_LIMITS }, rowOptionsFor(fixture)));
      expect(caught, 'expected a rejection').toBeDefined();
      const tightCodes = fixture.tags.map(tag => TIGHT_LIMIT_CODES_BY_CATALOG_ID[tag]).find(codes => codes !== undefined);
      expectClassified(fixture, caught, tightCodes);
    });
  }
});

describe.skipIf(!READER_READY)('hostile: the reader DoS fixtures hold with the default limits', () => {
  const fixtures = fixturesWithTag('kind:hostile').filter(fixture => fixture.tags.some(tag => DEFAULT_LIMIT_CATALOG_IDS.includes(tag)));

  it('covers every catalog id in the family', () => {
    const covered = new Set(fixtures.flatMap(fixture => fixture.tags));
    expect(DEFAULT_LIMIT_CATALOG_IDS.filter(catalogId => !covered.has(catalogId))).toEqual([]);
  });

  for (const fixture of fixtures) {
    if (fixture.expectedError !== undefined) {
      it(`${fixture.id} -> ${fixture.expectedError}, fast`, async () => {
        const started = performance.now();
        const caught = await rejection(() => openAndStreamEverything(readFixture(fixture), {}, rowOptionsFor(fixture)));
        expect(caught, 'expected a rejection').toBeDefined();
        expectClassified(fixture, caught);
        expect(performance.now() - started).toBeLessThan(FAIL_FAST_MS);
      });
      continue;
    }
    it(`${fixture.id} reads to the end in bounded memory`, async () => {
      const collectGarbage = garbageCollector();
      const bytes = readFixture(fixture);
      collectGarbage();
      const heapBefore = process.memoryUsage().heapUsed;
      let peakLive = 0;
      // Sampled from the consumer, mid-batch: a batch is yielded without an event-loop turn, so a timer would miss it.
      const rows = await openAndStreamEverything(bytes, {}, rowOptionsFor(fixture), count => {
        if (shouldSample(count)) {
          collectGarbage();
          peakLive = Math.max(peakLive, process.memoryUsage().heapUsed - heapBefore);
        }
      });
      expect(rows).toBeGreaterThan(1000);
      expect(peakLive, `live heap reached ${(peakLive / MB).toFixed(1)} MB`).toBeLessThan(READ_FLAT_LIVE_HEAP_BUDGET);
    });
  }
});
