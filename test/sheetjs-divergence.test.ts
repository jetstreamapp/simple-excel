/**
 * Catalog entries where we and SheetJS 0.20.3 (with Jetstream's options) read the same bytes differently on purpose,
 * found by the real-world corpus audit (`scripts/corpus-audit/`). Each test pins our side of the contract on the
 * entry's edge fixture and the SheetJS behaviour its `seenIn` records, so a SheetJS upgrade that changes it shows up
 * here, and the audit harness's known-divergence rules can be revisited.
 */
import { describe, expect, it } from 'vitest';
import * as XLSX from 'xlsx';
import { openWorkbook } from '../src/index';
import type { OpenOptions } from '../src/types';
import { fixtureById, readFixture } from './helpers/fixtures';
import { READER_READY } from './helpers/ready';

const SHEET_TO_JSON_OPTIONS: XLSX.Sheet2JSONOpts = {
  dateNF: 'yyyy"-"mm"-"dd"T"hh:mm:ss',
  defval: '',
  blankrows: false,
  rawNumbers: true,
};

function sheetjs(id: string): XLSX.WorkBook {
  return XLSX.read(readFixture(fixtureById(id)), { cellText: false, cellDates: true, type: 'array' });
}

async function ourRows(id: string, sheet: number | string = 0): Promise<unknown[][]> {
  const workbook = await openWorkbook(readFixture(fixtureById(id)));
  try {
    const rows: unknown[][] = [];
    for await (const row of workbook.sheet(sheet).rows()) {
      rows.push(row);
    }
    return rows;
  } finally {
    await workbook.close();
  }
}

async function ourObjects(id: string, errors?: OpenOptions['errors']): Promise<Record<string, unknown>[]> {
  const workbook = await openWorkbook(readFixture(fixtureById(id)), { errors });
  try {
    return (await workbook.sheet(0).toObjects()).rows;
  } finally {
    await workbook.close();
  }
}

/** Each character's code point, `U+1F603` style. */
function codePoints(text: unknown): string[] {
  return [...String(text)].map(character => `U+${(character.codePointAt(0) ?? 0).toString(16).toUpperCase()}`);
}

describe.skipIf(!READER_READY)('known divergences from SheetJS 0.20.3', () => {
  it('EC-TEXT-CELL-NO-TEXT: a text-typed cell with no value child is blank, so its row is no record', async () => {
    const rows = await ourRows('edge-text-cell-no-text');
    expect(rows.map(row => row.length)).toEqual([3, 1, 3]);
    expect(await ourObjects('edge-text-cell-no-text')).toHaveLength(2);

    const sheet = sheetjs('edge-text-cell-no-text').Sheets.Data as XLSX.WorkSheet;
    expect([sheet.B2?.v, sheet.C2?.v, sheet.A4?.v, sheet.B4?.v, sheet.C4]).toEqual(['', '', '', '', undefined]);
    const records = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, SHEET_TO_JSON_OPTIONS);
    expect(records).toHaveLength(4);
    expect(records.slice(2)).toEqual([
      { Name: '', Amount: '', When: '' },
      { Name: '', Amount: '', When: '' },
    ]);
  });

  it('EC-NUMFMT-ACCOUNTING-BUILTIN: numFmtId 41 to 44 are plain numbers; SheetJS drops the sheet without an error', async () => {
    const rows = await ourRows('edge-accounting-numfmt');
    expect(rows[1]?.[1]).toBe(12.5);
    expect(rows[3]).toEqual([1234.5, -1234.5, 0.25]);

    const workbook = sheetjs('edge-accounting-numfmt');
    expect(workbook.SheetNames).toEqual(['Data']);
    expect(workbook.Sheets.Data).toBeUndefined();
  });

  it('EC-XML-NCR-ABOVE-BMP: a character reference past U+FFFF decodes to the astral character', async () => {
    const rows = await ourRows('edge-ncr-above-bmp');
    expect(rows[1]?.[0]).toBe('Zoë \u{1F603}');
    expect(rows[2]?.[0]).toBe('\u{1F600}x');
    expect(rows[3]?.[0]).toBe('\u{1F600}');

    const sheet = sheetjs('edge-ncr-above-bmp').Sheets.Data as XLSX.WorkSheet;
    expect(codePoints(sheet.A2?.v).at(-1)).toBe('U+F603');
    expect(codePoints(sheet.A3?.v)[0]).toBe('U+F600');
    expect(codePoints(sheet.A4?.v)).toEqual(['U+0']);
  });

  it('EC-CHARTSHEET-CACHED-SERIES: a chartsheet has no rows; SheetJS invents cells from the cached series', async () => {
    const workbook = await openWorkbook(readFixture(fixtureById('edge-chartsheet-cached-series')));
    try {
      expect(workbook.sheets.map(info => [info.name, info.kind])).toEqual([
        ['Data', 'worksheet'],
        ['Chart1', 'chartsheet'],
      ]);
    } finally {
      await workbook.close();
    }
    expect(await ourRows('edge-chartsheet-cached-series', 'Chart1')).toEqual([]);

    const chart = sheetjs('edge-chartsheet-cached-series').Sheets.Chart1 as XLSX.WorkSheet;
    expect(chart['!type']).toBe('chart');
    expect([chart.A1?.v, chart.A2?.v]).toEqual([5, 7]);
  });

  it("EC-OBJECTS-ERROR-ONLY-ROW: a row of errors is a record with errors: 'string' and none with errors: 'null'", async () => {
    const stringRecords = await ourObjects('edge-error-only-rows', 'string');
    expect(stringRecords).toHaveLength(4);
    expect(stringRecords.slice(2)).toEqual([
      { Name: '#N/A', Amount: '#DIV/0!', When: '' },
      { Name: '#NULL!', Amount: '', When: '' },
    ]);

    const nullRecords = await ourObjects('edge-error-only-rows', 'null');
    expect(nullRecords.map(record => record.Name)).toEqual(['Zoë', 'x']);

    const sheet = sheetjs('edge-error-only-rows').Sheets.Data as XLSX.WorkSheet;
    const records = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, SHEET_TO_JSON_OPTIONS);
    expect(records.map(record => record.Name)).toEqual(['Zoë', 'x']);
  });
});
