/**
 * The corpus audit's exemption rules (`scripts/corpus-audit/divergence.ts`). Each rule must match only the SheetJS-side
 * shape it names: the shape still exempts, and the same values produced by a regression in our reader do not.
 */
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import * as XLSX from 'xlsx';
import {
  hasDimensionElement,
  knownSheetjsDivergence,
  RawSheets,
  sameHeader,
  SHEET_TO_JSON_OPTIONS,
  sheetjsKeyColumns,
  strCellTexts,
  textCellsWithoutValue,
  wallClockInLocalGap,
  writerConvention,
} from '../scripts/corpus-audit/divergence';
import { fixtureById, readFixture } from './helpers/fixtures';

const proven = (): boolean => true;
const disproven = (): boolean => false;

describe('knownSheetjsDivergence', () => {
  it("EC-TEXT-CELL-NO-TEXT: exempts SheetJS '' against our blank only for a proven text-typed cell with no value child", () => {
    expect(knownSheetjsDivergence('', null, { textCellWithoutValue: proven })).toBe('sheetjs-empty-text-cell');
    // An empty shared string, a formula returning "" or <is><t/></is> read as blank by us: '' to both readers.
    expect(knownSheetjsDivergence('', null)).toBeNull();
    expect(knownSheetjsDivergence('', null, { textCellWithoutValue: disproven })).toBeNull();
  });

  it('EC-XML-ESCAPE-CASE: exempts only the uppercase _X escape SheetJS decodes', () => {
    expect(knownSheetjsDivergence('aAb', 'a_X0041_b')).toBe('sheetjs-escape-case');
    // We stopped decoding _xHHHH_ (EC-XML-ESCAPE-UNDECODED): ours to fix, not SheetJS's.
    expect(knownSheetjsDivergence('aAb', 'a_x0041_b')).toBeNull();
    expect(knownSheetjsDivergence('p\nq', 'p_x000A_q')).toBeNull();
  });

  it('EC-XML-ESCAPE-LITERAL: exempts the double decoding of a t="str" value only when the raw text proves both sides', () => {
    expect(knownSheetjsDivergence('A', '_x0041_', { strCellText: () => '_x005F_x0041_' })).toBe('sheetjs-str-escape-twice');
    expect(knownSheetjsDivergence('a\rb', 'a_x000D_b', { strCellText: () => 'a_x005F_x000D_b' })).toBe('sheetjs-str-escape-twice');
    // The raw text is a real escape that we left undecoded.
    expect(knownSheetjsDivergence('A', '_x0041_', { strCellText: () => '_x0041_' })).toBeNull();
    expect(knownSheetjsDivergence('A', '_x0041_', { strCellText: () => undefined })).toBeNull();
    expect(knownSheetjsDivergence('A', '_x0041_')).toBeNull();
  });

  it("errors as defval: exempts SheetJS '' against our error text only for a cell SheetJS parsed as an error", () => {
    expect(knownSheetjsDivergence('', '#N/A', { sheetjsError: proven })).toBe('sheetjs-error-defval');
    expect(knownSheetjsDivergence(null, '#NULL!', { sheetjsError: proven })).toBe('sheetjs-error-defval');
    expect(knownSheetjsDivergence('', '#N/A')).toBeNull();
    expect(knownSheetjsDivergence('', '#N/A', { sheetjsError: disproven })).toBeNull();
  });

  it('keeps the pinned rules narrow', () => {
    expect(knownSheetjsDivergence({ $datetime: '1899-12-31T06:00:00.000' }, { $datetime: '1899-12-30T06:00:00.000' })).toBe(
      'sheetjs-time-only-marker',
    );
    expect(knownSheetjsDivergence({ $datetime: '2024-01-01T00:00:00.000' }, { $datetime: '2024-01-01T00:00:00.001' })).toBe(
      'sheetjs-ms-truncation',
    );
    expect(knownSheetjsDivergence({ $datetime: '2024-01-01T00:00:00.000' }, { $datetime: '2024-01-01T01:00:00.000' })).toBeNull();
    expect(knownSheetjsDivergence('a\nb', 'a\r\nb')).toBe('sheetjs-cr-collapse');
    expect(knownSheetjsDivergence('x', 'y')).toBeNull();
  });
});

describe('sameHeader', () => {
  it("EC-TEXT-CELL-NO-TEXT: a SheetJS '' header against our blank header needs the proven shape", () => {
    expect(sameHeader('', '__EMPTY', proven)).toBe('sheetjs-empty-text-cell');
    expect(sameHeader('_1', '__EMPTY_2', proven)).toBe('sheetjs-empty-text-cell');
    // A header cell that is an empty shared string names its column '' for us too.
    expect(sameHeader('', '__EMPTY')).toBeNull();
    expect(sameHeader('_1', '__EMPTY_2', disproven)).toBeNull();
    expect(sameHeader('__EMPTY_1', '__EMPTY_3')).toBe('same');
  });
});

describe('sheetjsKeyColumns', () => {
  it('maps each sheet_to_json key to its column, even when array-index keys are listed first', () => {
    const sheet = XLSX.utils.aoa_to_sheet([
      [null, 'Name', 8, null, 'Name', 2024],
      [1, 2, 3, 4, 5, 6],
    ]);
    const range = XLSX.utils.decode_range(sheet['!ref'] as string);
    const keys = Object.keys(XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { ...SHEET_TO_JSON_OPTIONS, range })[0] as object);
    expect(keys).toEqual(['8', '2024', '__EMPTY', 'Name', '__EMPTY_1', 'Name_1']);
    expect(Object.fromEntries(sheetjsKeyColumns(sheet, range, keys))).toEqual({
      __EMPTY: 0,
      Name: 1,
      8: 2,
      __EMPTY_1: 3,
      Name_1: 4,
      2024: 5,
    });
    // Without array-index keys the position is the column.
    expect([...sheetjsKeyColumns(sheet, range, ['a', 'b'])]).toEqual([
      ['a', 0],
      ['b', 1],
    ]);
    // Keys that are not what sheet_to_json would name prove no column.
    expect(sheetjsKeyColumns(sheet, range, ['8', 'other']).size).toBe(0);
  });
});

describe('writerConvention', () => {
  let savedZone: string | undefined;
  beforeAll(() => {
    savedZone = process.env.TZ;
    process.env.TZ = 'America/New_York';
  });
  afterAll(() => {
    if (savedZone === undefined) {
      delete process.env.TZ;
    } else {
      process.env.TZ = savedZone;
    }
  });

  it('runs in a zone with a spring-forward gap', () => {
    expect(new Date(2024, 2, 10, 2, 30).getHours()).toBe(3);
  });

  it('EC-DATE-DST-GAP: excuses one hour forward only for a wall clock that does not exist locally', () => {
    expect(wallClockInLocalGap({ $datetime: '2024-03-10T02:30:00.000' })).toBe(true);
    expect(wallClockInLocalGap({ $datetime: '2024-03-10T03:30:00.000' })).toBe(false);
    expect(wallClockInLocalGap({ $datetime: '2024-11-03T01:30:00.000' })).toBe(false);
    expect(writerConvention({ $datetime: '2024-03-10T02:30:00.000' }, { $datetime: '2024-03-10T03:30:00.000' })).toBe('writer-dst-gap');
    // The wrong direction for a gap.
    expect(writerConvention({ $datetime: '2024-03-10T02:30:00.000' }, { $datetime: '2024-03-10T01:30:00.000' })).toBeNull();
    // A one-hour shift of a wall clock that exists: the local/UTC offset bug, not the gap.
    expect(writerConvention({ $datetime: '2024-07-01T12:00:00.000' }, { $datetime: '2024-07-01T13:00:00.000' })).toBeNull();
    expect(writerConvention({ $datetime: '2024-11-03T01:30:00.000' }, { $datetime: '2024-11-03T02:30:00.000' })).toBeNull();
  });

  it('ADR-003: a date before 1900-03-01 comes back as text', () => {
    expect(writerConvention({ $date: '1899-01-01' }, '1899-01-01')).toBe('writer-pre-1900-as-text');
  });
});

describe('sheet XML facts', () => {
  it('EC-TEXT-CELL-NO-TEXT: finds only text-typed cells with no value child and an r attribute', () => {
    const xml =
      '<worksheet><sheetData><row r="1">' +
      '<c r="A1" t="inlineStr"/><c r="B1" s="2" t="str" /><c r="C1" t="str"></c><x:c r="H1" t=\'inlineStr\'/>' +
      '<c r="D1" t="inlineStr"><is><t/></is></c><c r="E1" t="str"><v/></c><c r="F1" t="str"><f>""</f><v></v></c>' +
      '<c r="G1" t="s"/><c t="str"/><c r="I1" t="inlineStr"><is/></c><c r="J1"/>' +
      '</row></sheetData></worksheet>';
    expect([...textCellsWithoutValue(xml)].toSorted()).toEqual(['A1', 'B1', 'C1', 'H1']);
  });

  it('reads the raw <v> text of t="str" cells, entities decoded and escapes kept', () => {
    const xml =
      '<sheetData><row r="1"><c r="A1" t="str"><f>X</f><v>a &amp; _x005F_x0041_&#x41;</v></c><c r="B1" t="s"><v>0</v></c>' +
      '<x:c r="C1" t="str"><x:v>q</x:v></x:c><c r="D1" t="str"/><c r="E1" t="str"><f>Y</f></c></row></sheetData>';
    expect([...strCellTexts(xml)]).toEqual([
      ['A1', 'a & _x005F_x0041_A'],
      ['C1', 'q'],
    ]);
  });

  it('EC-HEADER-DIMENSION-START-COLUMN: sees a <dimension> before <sheetData>, prefixed or not', () => {
    expect(hasDimensionElement('<worksheet><dimension ref="B3:E6"/><sheetData/></worksheet>')).toBe(true);
    expect(hasDimensionElement('<x:worksheet><x:dimension ref="A1"/><x:sheetData/></x:worksheet>')).toBe(true);
    expect(hasDimensionElement('<worksheet><sheetViews/><sheetData><row r="1"/></sheetData></worksheet>')).toBe(false);
  });

  it('reads the raw sheet parts of a workbook in workbook order', () => {
    const textCells = new RawSheets(readFixture(fixtureById('edge-text-cell-no-text')));
    for (const ref of ['B2', 'C2', 'A4', 'B4', 'A5', 'B5']) {
      expect(textCells.textCellWithoutValue(0, ref), ref).toBe(true);
    }
    for (const ref of ['A1', 'C3', 'C4', 'Z9']) {
      expect(textCells.textCellWithoutValue(0, ref), ref).toBe(false);
    }
    expect(textCells.lacksDimension(0)).toBe(false);
    expect(textCells.textCellWithoutValue(5, 'A1')).toBe(false);

    const hidden = new RawSheets(readFixture(fixtureById('edge-hidden-and-veryhidden-sheets')));
    expect([0, 1, 2].map(index => hidden.lacksDimension(index))).toEqual([false, true, true]);
    expect(hidden.sheetXml(1)).toContain('>hidden<');
    expect(hidden.sheetXml(2)).toContain('>veryHidden<');
    expect(new RawSheets(readFixture(fixtureById('edge-no-dimension'))).lacksDimension(0)).toBe(true);
    expect(new RawSheets(new Uint8Array([1, 2, 3])).lacksDimension(0)).toBe(false);
  });
});
