import { describe, expect, it } from 'vitest';
import { compareDumps, describeReport, type ExpectedWorkbook } from '../../../test/helpers/diff';
import { fixturesWithTag, fixturePolicies, readExpected, readFixture, type Fixture } from '../../../test/helpers/fixtures';
import { fromJsDate, toTyped, type TypedValue } from '../../../test/helpers/typed';
import { isXlsxError, XlsxError } from '../../errors';
import type { CellValue, RawCell, RowsOptions } from '../../types';
import { decodeCellText } from '../../xml/escape';
import { XmlTokenizer } from '../../xml/tokenizer';
import { sourceFrom } from '../../zip/source';
import { ZipReader } from '../../zip/zip-reader';
import { isBuiltinDateId, isDateFormatCode } from '../numfmt';
import { readWorksheetHead, readWorksheetRows, type SheetRow, type WorksheetReadContext, type WorksheetWarning } from '../worksheet-reader';

const encoder = new TextEncoder();

// ---------------------------------------------------------------------------------------------------------------
// Harness
// ---------------------------------------------------------------------------------------------------------------

function context(overrides: Partial<WorksheetReadContext> = {}): WorksheetReadContext {
  return {
    sharedStrings: () => Promise.resolve([]),
    hasSharedStrings: true,
    isDateByXf: new Uint8Array(0),
    date1904: false,
    dates: 'utc',
    errors: 'string',
    maxXmlDepth: 256,
    maxTextLength: 64 * 1024 * 1024,
    ...overrides,
  };
}

/** Wrap row XML in the smallest worksheet that carries it. */
function sheet(rows: string): string {
  return `<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${rows}</sheetData></worksheet>`;
}

async function* bytesOf(text: string, chunkSize = Number.POSITIVE_INFINITY): AsyncGenerator<Uint8Array> {
  const bytes = encoder.encode(text);
  const size = Number.isFinite(chunkSize) ? chunkSize : bytes.byteLength;
  for (let offset = 0; offset < bytes.byteLength; offset += size) {
    yield bytes.subarray(offset, Math.min(offset + size, bytes.byteLength));
  }
}

async function collect(rows: AsyncIterable<SheetRow>): Promise<SheetRow[]> {
  const out: SheetRow[] = [];
  for await (const row of rows) {
    out.push(row);
  }
  return out;
}

/** Rows of a hand-written sheet, as plain cell arrays. */
async function readCells(xml: string, options: RowsOptions = {}, overrides: Partial<WorksheetReadContext> = {}): Promise<CellValue[][]> {
  const rows = await collect(readWorksheetRows(bytesOf(xml), context(overrides), options));
  return rows.map(row => row.cells);
}

async function readRows(xml: string, options: RowsOptions = {}, overrides: Partial<WorksheetReadContext> = {}): Promise<SheetRow[]> {
  return collect(readWorksheetRows(bytesOf(xml), context(overrides), options));
}

/** UTC fields of a Date as an ISO-ish string, so date assertions do not depend on the runner's timezone. */
function utcText(value: CellValue): string {
  return value instanceof Date ? value.toISOString() : `not a date: ${String(value)}`;
}

// ---------------------------------------------------------------------------------------------------------------
// Structure: rows, columns, references
// ---------------------------------------------------------------------------------------------------------------

describe('structure', () => {
  it('EC-MISSING-R-ATTRIBUTES infers row and column positions', async () => {
    const cells = await readCells(
      sheet(
        '<row><c><v>1</v></c><c><v>2</v></c></row>' +
          '<row><c><v>3</v></c><c r="D2"><v>4</v></c><c><v>5</v></c></row>' +
          '<row r="7"><c><v>6</v></c></row>',
      ),
    );
    expect(cells).toEqual([[1, 2], [3, null, null, 4, 5], [6]]);
  });

  it('trusts <row r> over the row digits of <c r>', async () => {
    const rows = await readRows(sheet('<row r="5"><c r="B99"><v>1</v></c></row>'));
    expect(rows).toEqual([{ index: 5, cells: [null, 1] }]);
  });

  it('falls back to the next column when <c r> is malformed', async () => {
    const cells = await readCells(sheet('<row r="1"><c r="A1"><v>1</v></c><c r="nonsense"><v>2</v></c></row>'));
    expect(cells).toEqual([[1, 2]]);
  });

  it('trims trailing empty cells and keeps interior holes', async () => {
    const cells = await readCells(sheet('<row r="1"><c r="A1"><v>1</v></c><c r="C1"><v>3</v></c><c r="E1"/></row>'));
    expect(cells).toEqual([[1, null, 3]]);
  });

  it('EC-PREFIXED-ELEMENTS reads namespace-prefixed worksheets', async () => {
    const xml =
      '<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:sheetData>' +
      '<x:row r="1"><x:c r="A1" t="inlineStr"><x:is><x:t>a</x:t></x:is></x:c><x:c r="B1"><x:v>2</x:v></x:c></x:row>' +
      '</x:sheetData></x:worksheet>';
    expect(await readCells(xml)).toEqual([['a', 2]]);
  });

  it('ignores everything outside sheetData and reads every sheetData element', async () => {
    const xml =
      '<worksheet><dimension ref="A1:B2"/><cols><col min="1" max="1" width="8"/></cols>' +
      '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>' +
      '<sheetData><row r="2"><c r="A2"><v>2</v></c></row></sheetData>' +
      '<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>' +
      '<extLst><ext><row r="99"><c r="A99"><v>404</v></c></row></ext></extLst></worksheet>';
    expect(await readRows(xml)).toEqual([
      { index: 1, cells: [1] },
      { index: 2, cells: [2] },
    ]);
  });

  it('yields nothing for a chartsheet-shaped document', async () => {
    expect(await readRows('<chartsheet><sheetViews><sheetView workbookViewId="0"/></sheetViews></chartsheet>')).toEqual([]);
  });

  it('reads rows the same however the bytes are chunked', async () => {
    const xml = sheet(
      '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="inlineStr"><is><t>Zoë</t></is></c></row>' +
        '<row r="2"><c r="A2"><v>12.5</v></c><c r="B2" t="b"><v>1</v></c></row>',
    );
    const expected = [
      ['shared', 'Zoë'],
      [12.5, true],
    ];
    for (const chunkSize of [1, 2, 3, 7, 16, 64, 4096]) {
      const rows = await collect(
        readWorksheetRows(bytesOf(xml, chunkSize), context({ sharedStrings: () => Promise.resolve(['shared']) }), {}),
      );
      expect(
        rows.map(row => row.cells),
        `chunk size ${chunkSize}`,
      ).toEqual(expected);
    }
  });
});

// ---------------------------------------------------------------------------------------------------------------
// Cell types
// ---------------------------------------------------------------------------------------------------------------

describe('cell values', () => {
  it('EC-EMPTY-V-ELEMENT reads <v/> and a childless <c> as blank', async () => {
    const rows = await readRows(
      sheet('<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v/></c><c r="C1" s="1"/><c r="D1"><v>4</v></c></row>'),
      {},
      {
        isDateByXf: Uint8Array.from([0, 1]),
      },
    );
    expect(rows).toEqual([{ index: 1, cells: [1, null, null, 4] }]);
  });

  it('EC-NUM-* parses doubles, normalizes -0 and keeps non-finite text verbatim', async () => {
    const cells = await readCells(
      sheet(
        '<row r="1">' +
          '<c r="A1"><v>0.30000000000000004</v></c>' +
          '<c r="B1"><v>-0</v></c>' +
          '<c r="C1"><v>1E+21</v></c>' +
          '<c r="D1"><v>1.7976931348623157E+308</v></c>' +
          '<c r="E1"><v>inf</v></c>' +
          '<c r="F1"><v>1e400</v></c>' +
          '<c r="G1"><v>twelve</v></c>' +
          '</row>',
      ),
    );
    const row = cells[0] ?? [];
    expect(row[0]).toBe(0.30000000000000004);
    expect(Object.is(row[1], 0)).toBe(true);
    expect(row[2]).toBe(1e21);
    expect(row[3]).toBe(1.7976931348623157e308);
    // EC-NUMBERS-INF-VALUE: Numbers writes the max double as `inf`; the raw text is the only honest answer.
    expect(row[4]).toBe('inf');
    expect(row[5]).toBe('1e400');
    expect(row[6]).toBe('twelve');
  });

  it('reads t="str" as literal text, with an empty <v> as the empty string', async () => {
    const cells = await readCells(
      sheet(
        '<row r="1"><c r="A1" t="str"><v>Id</v></c><c r="B1" t="str"><v></v></c><c r="C1" t="str"/><c r="D1" t="str"><v>4</v></c></row>',
      ),
    );
    expect(cells).toEqual([['Id', '', null, '4']]);
  });

  it('EC-INLINE-STRINGS-CDATA flattens CDATA, rich-text runs and skips rPh phonetics', async () => {
    const cells = await readCells(
      sheet(
        '<row r="1">' +
          '<c r="A1" t="inlineStr"><is><t><![CDATA[Amount & <more>]]></t></is></c>' +
          '<c r="B1" t="inlineStr"><is><r><t>Wh</t></r><r><rPr><b/></rPr><t>en</t></r></is></c>' +
          '<c r="C1" t="inlineStr"><is><t>漢字</t><rPh sb="0" eb="2"><t>かんじ</t></rPh><phoneticPr fontId="1"/></is></c>' +
          '<c r="D1" t="inlineStr"><is><t xml:space="preserve"> padded </t></is></c>' +
          '<c r="E1" t="inlineStr"><is><t/></is></c>' +
          '<c r="F1" t="inlineStr"/>' +
          '<c r="G1" t="inlineStr"><is><t>tail</t></is></c>' +
          '</row>',
      ),
    );
    expect(cells).toEqual([['Amount & <more>', 'When', '漢字', ' padded ', '', null, 'tail']]);
  });

  it('reads booleans and treats a missing value as blank', async () => {
    const cells = await readCells(
      sheet(
        '<row r="1"><c r="A1" t="b"><v>1</v></c><c r="B1" t="b"><v>0</v></c><c r="C1" t="b"><v>true</v></c><c r="D1" t="b"><v/></c><c r="E1" t="b"><v>2</v></c></row>',
      ),
    );
    expect(cells).toEqual([[true, false, true, null, false]]);
  });

  it('EC-ERROR-CELL-AS-TEXT returns error cells per the errors option', async () => {
    const xml = sheet(
      '<row r="1"><c r="A1" t="e"><v>#N/A</v></c><c r="B1" t="e"><v>#DIV/0!</v></c><c r="C1" t="e"><v/></c><c r="D1"><v>1</v></c></row>',
    );
    expect(await readCells(xml, {}, { errors: 'string' })).toEqual([['#N/A', '#DIV/0!', null, 1]]);
    expect(await readCells(xml, {}, { errors: 'object' })).toEqual([[{ error: '#N/A' }, { error: '#DIV/0!' }, null, 1]]);
    // `errors: 'null'` drops them entirely, leaving the row with only its non-error cells.
    expect(await readCells(xml, {}, { errors: 'null' })).toEqual([[null, null, null, 1]]);
  });

  it('EC-XML-ESCAPE-* decodes _xHHHH_ escapes only where Excel does', async () => {
    const cells = await readCells(
      sheet(
        '<row r="1">' +
          '<c r="A1" t="str"><v>a_x0001_b</v></c>' +
          '<c r="B1" t="str"><v>x_x001f_y</v></c>' +
          '<c r="C1" t="str"><v>_X0041_</v></c>' +
          '<c r="D1" t="str"><v>_x005F_x0041_</v></c>' +
          '<c r="E1" t="str"><v>_xZZZZ_</v></c>' +
          '<c r="F1" t="inlineStr"><is><t>crlf_x000D_line2</t></is></c>' +
          '</row>',
      ),
    );
    expect(cells).toEqual([['ab', 'xy', '_X0041_', '_x0041_', '_xZZZZ_', 'crlf\rline2']]);
  });

  it('EC-CRLF-NORMALIZED folds a raw CR to LF and keeps an escaped one', async () => {
    // XML line-end normalization is why Excel writes a carriage return as `_x000D_`: a raw one cannot survive a
    // parser, so a file that means CRLF has to escape it. SheetJS, office-kit and Excel all read these two cells
    // this way (05 matrix, D6/D7 of the Excel re-saves).
    const cells = await readCells(
      sheet(
        '<row r="1">' +
          '<c r="A1" t="inlineStr"><is><t xml:space="preserve">line1\r\nline2</t></is></c>' +
          '<c r="B1" t="inlineStr"><is><t xml:space="preserve">crlf_x000D_\r\nline2</t></is></c>' +
          '<c r="C1" t="inlineStr"><is><t xml:space="preserve">old\rmac</t></is></c>' +
          '</row>',
      ),
    );
    expect(cells).toEqual([['line1\nline2', 'crlf\r\nline2', 'old\nmac']]);
  });

  it('EC-CRLF-NORMALIZED folds a CRLF split across two decoded chunks', async () => {
    const xml = sheet('<row r="1"><c r="A1" t="inlineStr"><is><t xml:space="preserve">line1\r\nline2</t></is></c></row>');
    const splitInsideLineEnd = xml.indexOf('\r\n') + 1;
    const rows = await collect(
      readWorksheetRows(
        (async function* (): AsyncGenerator<Uint8Array> {
          yield encoder.encode(xml.slice(0, splitInsideLineEnd));
          yield encoder.encode(xml.slice(splitInsideLineEnd));
        })(),
        context(),
        {},
      ),
    );
    expect(rows.map(row => row.cells)).toEqual([['line1\nline2']]);
  });
});

// ---------------------------------------------------------------------------------------------------------------
// Shared strings
// ---------------------------------------------------------------------------------------------------------------

describe('shared strings', () => {
  it('resolves t="s" indexes through the table', async () => {
    const cells = await readCells(
      sheet('<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>2</v></c></row>'),
      {},
      { sharedStrings: () => Promise.resolve(['Name', 'Amount', 'When']) },
    );
    expect(cells).toEqual([['Name', 'When']]);
  });

  it('EC-SST-INDEX-OUT-OF-RANGE reads out-of-range indexes as blank text and warns once', async () => {
    const warnings: WorksheetWarning[] = [];
    const cells = await readCells(
      sheet('<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>9</v></c><c r="C1" t="s"><v>-1</v></c></row>'),
      {},
      { sharedStrings: () => Promise.resolve(['Name']), onWarning: warning => warnings.push(warning) },
    );
    expect(cells).toEqual([['Name', '', '']]);
    expect(warnings).toEqual([{ code: 'shared-string-index-out-of-range', ref: 'B1' }]);
  });

  it('never loads the table for a sheet without t="s" cells', async () => {
    let loads = 0;
    const cells = await readCells(
      sheet('<row r="1"><c r="A1" t="inlineStr"><is><t>a</t></is></c><c r="B1"><v>1</v></c></row>'),
      {},
      {
        sharedStrings: () => {
          loads++;
          return Promise.resolve([]);
        },
      },
    );
    expect(cells).toEqual([['a', 1]]);
    expect(loads).toBe(0);
  });

  it('loads the table once, even when t="s" is split across chunks', async () => {
    let loads = 0;
    const xml = sheet('<row r="1"><c r="A1" t="s"><v>0</v></c></row><row r="2"><c r="A2" t="s"><v>1</v></c></row>');
    for (const chunkSize of [1, 5, 9, 13, 32]) {
      loads = 0;
      const rows = await collect(
        readWorksheetRows(
          bytesOf(xml, chunkSize),
          context({
            sharedStrings: () => {
              loads++;
              return Promise.resolve(['one', 'two']);
            },
          }),
          {},
        ),
      );
      expect(
        rows.map(row => row.cells),
        `chunk size ${chunkSize}`,
      ).toEqual([['one'], ['two']]);
      expect(loads, `chunk size ${chunkSize}`).toBe(1);
    }
  });

  it('skips the per-chunk pre-scan when the package has no shared-strings part, and still resolves the cell', async () => {
    let loads = 0;
    const cells = await readCells(
      sheet('<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>1</v></c></row>'),
      {},
      {
        hasSharedStrings: false,
        sharedStrings: () => {
          loads++;
          return Promise.resolve([]);
        },
      },
    );
    // The deferred path still asks for the (empty) table once, and an unresolvable index reads as blank text.
    expect(cells).toEqual([['', 1]]);
    expect(loads).toBe(1);
  });

  it('still resolves t="s" when the pre-scan cannot see it (spaced attribute)', async () => {
    const cells = await readCells(
      sheet('<row r="1"><c r="A1" t = "s"><v>1</v></c><c r="B1" t = "s"><f>X()</f><v>0</v></c></row>'),
      { formulas: 'text' },
      { sharedStrings: () => Promise.resolve(['one', 'two']) },
    );
    expect(cells).toEqual([['two', '=X()']]);
  });
});

// ---------------------------------------------------------------------------------------------------------------
// Dates
// ---------------------------------------------------------------------------------------------------------------

describe('dates', () => {
  const dateStyles = Uint8Array.from([0, 1]);

  it('EC-DATE-DETECTION-VIA-NUMFMT converts serials only for date-formatted styles', async () => {
    const cells = await readCells(
      sheet('<row r="1"><c r="A1" s="1"><v>45351.5</v></c><c r="B1"><v>45351.5</v></c></row>'),
      {},
      {
        isDateByXf: dateStyles,
      },
    );
    expect(utcText(cells[0]?.[0] ?? null)).toBe('2024-02-29T12:00:00.000Z');
    expect(cells[0]?.[1]).toBe(45351.5);
  });

  it('EC-DATE-1900-LEAP-BUG maps serial 60 to 1900-03-01 and shifts the serials below it', async () => {
    const cells = await readCells(
      sheet(
        '<row r="1"><c r="A1" s="1"><v>1</v></c><c r="B1" s="1"><v>59</v></c><c r="C1" s="1"><v>60</v></c><c r="D1" s="1"><v>61</v></c></row>',
      ),
      {},
      { isDateByXf: dateStyles },
    );
    expect((cells[0] ?? []).map(value => utcText(value).slice(0, 10))).toEqual(['1900-01-01', '1900-02-28', '1900-03-01', '1900-03-01']);
  });

  it('EC-TIME-ONLY-HAS-DATE-PART puts time-only serials on 1899-12-30', async () => {
    const cells = await readCells(sheet('<row r="1"><c r="A1" s="1"><v>0.5</v></c></row>'), {}, { isDateByXf: dateStyles });
    expect(utcText(cells[0]?.[0] ?? null)).toBe('1899-12-30T12:00:00.000Z');
  });

  it('EC-DATE-TIME-ONLY-NEGATIVE-SERIAL keeps a negative serial as a number', async () => {
    const cells = await readCells(sheet('<row r="1"><c r="A1" s="1"><v>-0.4757316087962963</v></c></row>'), {}, { isDateByXf: dateStyles });
    expect(cells[0]?.[0]).toBe(-0.4757316087962963);
  });

  it('EC-DATE-1904 honours the workbook date system', async () => {
    const cells = await readCells(
      sheet('<row r="1"><c r="A1" s="1"><v>43889</v></c></row>'),
      {},
      {
        isDateByXf: dateStyles,
        date1904: true,
      },
    );
    expect(utcText(cells[0]?.[0] ?? null)).toBe('2024-02-29T00:00:00.000Z');
  });

  it('dates: "serial" never builds a Date', async () => {
    const cells = await readCells(
      sheet('<row r="1"><c r="A1" s="1"><v>45351.5</v></c></row>'),
      {},
      {
        isDateByXf: dateStyles,
        dates: 'serial',
      },
    );
    expect(cells[0]?.[0]).toBe(45351.5);
  });

  it('dates: "local" builds Dates from local fields', async () => {
    const cells = await readCells(
      sheet('<row r="1"><c r="A1" s="1"><v>45351.5</v></c></row>'),
      {},
      {
        isDateByXf: dateStyles,
        dates: 'local',
      },
    );
    const value = cells[0]?.[0];
    expect(value).toBeInstanceOf(Date);
    expect(fromJsDate(value as Date, 'local')).toEqual({ $datetime: '2024-02-29T12:00:00.000' });
  });

  it('EC-STRICT-ISO-DATE-PRECISION parses t="d" with any fraction digits and bare times', async () => {
    const cells = await readCells(
      sheet(
        '<row r="1">' +
          '<c r="A1" t="d"><v>2024-02-29</v></c>' +
          '<c r="B1" t="d"><v>2024-03-10T09:29:59.9999995809048412</v></c>' +
          '<c r="C1" t="d"><v>19:34:56.789</v></c>' +
          '<c r="D1" t="d"><v>2024-02-29T12:00:00Z</v></c>' +
          '<c r="E1" t="d"><v>not a date</v></c>' +
          '<c r="F1" t="d"><v/></c>' +
          '</row>',
      ),
    );
    const row = cells[0] ?? [];
    expect(utcText(row[0] ?? null)).toBe('2024-02-29T00:00:00.000Z');
    expect(utcText(row[1] ?? null)).toBe('2024-03-10T09:30:00.000Z');
    expect(utcText(row[2] ?? null)).toBe('1899-12-30T19:34:56.789Z');
    expect(utcText(row[3] ?? null)).toBe('2024-02-29T12:00:00.000Z');
    expect(row[4]).toBe('not a date');
    expect(row[5]).toBeUndefined();
  });

  it('t="d" with dates: "serial" returns the serial number', async () => {
    const cells = await readCells(
      sheet('<row r="1"><c r="A1" t="d"><v>2024-02-29T12:00:00</v></c><c r="B1" t="d"><v>1800-01-01</v></c></row>'),
      {},
      { dates: 'serial' },
    );
    expect(cells[0]?.[0]).toBe(45351.5);
    // Pre-epoch dates have no serial, so the text survives instead of becoming a wrong number (EC-DATE-PRE-1900).
    expect(cells[0]?.[1]).toBe('1800-01-01');
  });
});

// ---------------------------------------------------------------------------------------------------------------
// Formulas
// ---------------------------------------------------------------------------------------------------------------

describe('formulas', () => {
  const formulaSheet = sheet(
    '<row r="1">' +
      '<c r="A1"><f>B2*2</f><v>25</v></c>' +
      '<c r="B1" t="str"><f>CONCAT("a","b")</f><v>ab</v></c>' +
      '<c r="C1"><f t="shared" ref="C1:C9" si="0">A1+1</f><v>26</v></c>' +
      '<c r="D1"><f t="shared" si="0"/><v>27</v></c>' +
      '<c r="E1" t="e"><f>1/0</f><v>#DIV/0!</v></c>' +
      '</row>',
  );

  it('returns cached values by default', async () => {
    expect(await readCells(formulaSheet)).toEqual([[25, 'ab', 26, 27, '#DIV/0!']]);
  });

  it('EC-FORMULA-CACHED-ERROR returns formula text with formulas: "text", cached values for shared followers', async () => {
    expect(await readCells(formulaSheet, { formulas: 'text' })).toEqual([['=B2*2', '=CONCAT("a","b")', '=A1+1', 27, '=1/0']]);
  });
});

// ---------------------------------------------------------------------------------------------------------------
// Row windows and limits
// ---------------------------------------------------------------------------------------------------------------

describe('row windows and limits', () => {
  const gapped = sheet(
    '<row r="1"><c r="A1"><v>1</v></c></row>' +
      '<row r="2"/>' +
      '<row r="5"><c r="A5"><v>5</v></c></row>' +
      '<row r="6"><c r="A6"/><c r="B6"/></row>' +
      '<row r="7"><c r="A7"><v>7</v></c></row>',
  );

  it('skips blank rows and row gaps by default', async () => {
    expect(await readRows(gapped)).toEqual([
      { index: 1, cells: [1] },
      { index: 5, cells: [5] },
      { index: 7, cells: [7] },
    ]);
  });

  it('fills blank rows and gaps so indexes stay aligned with blankRows', async () => {
    const rows = await readRows(gapped, { blankRows: true });
    expect(rows.map(row => row.index)).toEqual([1, 2, 3, 4, 5, 6, 7]);
    expect(rows.every((row, offset) => row.index === offset + 1)).toBe(true);
    expect(rows.map(row => row.cells)).toEqual([[1], [], [], [], [5], [], [7]]);
  });

  it('startRow skips earlier rows and keeps blankRows aligned to it', async () => {
    expect(await readRows(gapped, { startRow: 5 })).toEqual([
      { index: 5, cells: [5] },
      { index: 7, cells: [7] },
    ]);
    const aligned = await readRows(gapped, { startRow: 5, blankRows: true });
    expect(aligned.map(row => row.index)).toEqual([5, 6, 7]);
  });

  it('maxRows stops the read after that many rows', async () => {
    expect(await readRows(gapped, { maxRows: 2 })).toEqual([
      { index: 1, cells: [1] },
      { index: 5, cells: [5] },
    ]);
    expect(await readRows(gapped, { maxRows: 3, blankRows: true })).toHaveLength(3);
    expect(await readRows(gapped, { maxRows: 0 })).toEqual([]);
  });

  it('maxColumns rejects a cell past the limit', async () => {
    await expect(readRows(sheet('<row r="1"><c r="C1"><v>3</v></c></row>'), { maxColumns: 2 })).rejects.toThrow(XlsxError);
    await expect(readRows(sheet('<row r="1"><c r="B1"><v>2</v></c></row>'), { maxColumns: 2 })).resolves.toHaveLength(1);
    const error = await readRows(sheet('<row r="1"><c r="C1"><v>3</v></c></row>'), { maxColumns: 2 }).catch((thrown: unknown) => thrown);
    expect(isXlsxError(error) && error.code).toBe('LIMIT_EXCEEDED');
  });

  it('rejects a row number past the sheet limit', async () => {
    const error = await readRows(sheet('<row r="1048577"><c r="A1048577"><v>1</v></c></row>')).catch((thrown: unknown) => thrown);
    expect(isXlsxError(error) && error.code).toBe('LIMIT_EXCEEDED');
  });

  it('an aborted signal ends the read', async () => {
    const controller = new AbortController();
    controller.abort();
    const error = await readRows(sheet('<row r="1"><c r="A1"><v>1</v></c></row>'), {}, { signal: controller.signal }).catch(
      (thrown: unknown) => thrown,
    );
    expect(isXlsxError(error) && error.code).toBe('ABORTED');
  });

  it('reports malformed XML rather than truncating silently', async () => {
    const error = await readRows('<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row>').catch((thrown: unknown) => thrown);
    expect(isXlsxError(error) && error.code).toBe('XML_MALFORMED');
  });
});

// ---------------------------------------------------------------------------------------------------------------
// Streaming
// ---------------------------------------------------------------------------------------------------------------

describe('streaming', () => {
  /** A chunk source that records how much of itself the reader pulled and whether it was closed early. */
  function countingSource(
    text: string,
    chunkSize: number,
  ): { chunks: AsyncIterable<Uint8Array>; pulled: () => number; closed: () => boolean } {
    const bytes = encoder.encode(text);
    let pulled = 0;
    let closed = false;
    async function* generate(): AsyncGenerator<Uint8Array> {
      try {
        for (let offset = 0; offset < bytes.byteLength; offset += chunkSize) {
          pulled++;
          yield bytes.subarray(offset, Math.min(offset + chunkSize, bytes.byteLength));
        }
      } finally {
        closed = true;
      }
    }
    return { chunks: generate(), pulled: () => pulled, closed: () => closed };
  }

  const manyRows = sheet(
    Array.from({ length: 200 }, (_unused, index) => `<row r="${index + 1}"><c r="A${index + 1}"><v>${index}</v></c></row>`).join(''),
  );

  it('stops pulling chunks when the consumer breaks', async () => {
    const source = countingSource(manyRows, 256);
    for await (const row of readWorksheetRows(source.chunks, context(), {})) {
      expect(row.index).toBe(1);
      break;
    }
    expect(source.pulled()).toBe(1);
    expect(source.closed()).toBe(true);
  });

  it('stops pulling chunks once maxRows is reached', async () => {
    const source = countingSource(manyRows, 256);
    const rows = await collect(readWorksheetRows(source.chunks, context(), { maxRows: 2 }));
    expect(rows).toHaveLength(2);
    expect(source.pulled()).toBe(1);
    expect(source.closed()).toBe(true);
  });

  it('yields rows in batches as chunks arrive rather than buffering the part', async () => {
    const source = countingSource(manyRows, 512);
    let seen = 0;
    for await (const _row of readWorksheetRows(source.chunks, context(), {})) {
      seen++;
      if (seen === 1) {
        // The first row is available long before the last chunk has been pulled.
        expect(source.pulled()).toBeLessThan(Math.ceil(encoder.encode(manyRows).byteLength / 512));
      }
    }
    expect(seen).toBe(200);
  });
});

// ---------------------------------------------------------------------------------------------------------------
// head()
// ---------------------------------------------------------------------------------------------------------------

describe('readWorksheetHead', () => {
  const headSheet = sheet(
    '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="C1" t="inlineStr"><is><t>Name</t></is></c></row>' +
      '<row r="2"><c r="A2"><v>12.5</v></c><c r="B2" s="1"><v>45351</v></c><c r="C2"><f>A2*2</f><v>25</v></c><c r="D2" t="e"><v>#N/A</v></c><c r="E2"/></row>' +
      '<row r="3"><c r="A3"><v>3</v></c></row>',
  );

  async function head(rowCount: number, overrides: Partial<WorksheetReadContext> = {}): Promise<Map<string, RawCell>> {
    return readWorksheetHead(
      bytesOf(headSheet, 32),
      context({ sharedStrings: () => Promise.resolve(['Id']), isDateByXf: Uint8Array.from([0, 1]), ...overrides }),
      rowCount,
    );
  }

  it('keys cells by A1 reference and stops after the requested rows', async () => {
    const cells = await head(2);
    expect([...cells.keys()]).toEqual(['A1', 'C1', 'A2', 'B2', 'C2', 'D2']);
    expect(cells.get('A1')).toEqual({ value: 'Id' });
    expect(cells.get('C1')).toEqual({ value: 'Name' });
    expect(cells.get('A2')).toEqual({ value: 12.5 });
    expect(utcText(cells.get('B2')?.value ?? null)).toBe('2024-02-29T00:00:00.000Z');
  });

  it('always carries formula text and the error code alongside the cached value', async () => {
    const cells = await head(2);
    expect(cells.get('C2')).toEqual({ value: 25, formula: 'A2*2' });
    expect(cells.get('D2')).toEqual({ value: '#N/A', error: '#N/A' });
  });

  it('reports error cells as text plus code whatever the errors option is', async () => {
    for (const errors of ['string', 'object', 'null'] as const) {
      const cells = await head(2, { errors });
      expect(cells.get('D2'), `errors: ${errors}`).toEqual({ value: '#N/A', error: '#N/A' });
    }
  });

  it('returns an empty map for a row count below one and reads nothing', async () => {
    let pulled = 0;
    async function* counted(): AsyncGenerator<Uint8Array> {
      pulled++;
      yield encoder.encode(headSheet);
    }
    expect(await readWorksheetHead(counted(), context(), 0)).toEqual(new Map());
    expect(pulled).toBe(0);
  });

  it('stops consuming chunks after the last wanted row', async () => {
    let pulled = 0;
    let closed = false;
    const bytes = encoder.encode(headSheet);
    async function* counted(): AsyncGenerator<Uint8Array> {
      try {
        for (let offset = 0; offset < bytes.byteLength; offset += 64) {
          pulled++;
          yield bytes.subarray(offset, Math.min(offset + 64, bytes.byteLength));
        }
      } finally {
        closed = true;
      }
    }
    const cells = await readWorksheetHead(counted(), context({ sharedStrings: () => Promise.resolve(['Id']) }), 1);
    expect([...cells.keys()]).toEqual(['A1', 'C1']);
    expect(pulled).toBeLessThan(Math.ceil(bytes.byteLength / 64));
    expect(closed).toBe(true);
  });
});

// ---------------------------------------------------------------------------------------------------------------
// Real fixtures
// ---------------------------------------------------------------------------------------------------------------

/** Read a zip part into one string (test-only convenience; the reader itself never buffers a part). */
async function partText(zip: ZipReader, name: string): Promise<string> {
  return zip.readText(name);
}

/**
 * `sharedStrings.xml` flattened the way the real reader will: `<r>` runs concatenated, `<rPh>` skipped, `_xHHHH_`
 * decoded. Hand-rolled here because the shared-string reader is another work package.
 */
function parseSharedStringsXml(xml: string): string[] {
  const strings: string[] = [];
  let current = '';
  let capturing = false;
  let phonetic = false;
  const tokenizer = new XmlTokenizer({
    start(name: string): void {
      if (name === 'si') {
        current = '';
      } else if (name === 'rPh') {
        phonetic = true;
      } else if (name === 't' && !phonetic) {
        capturing = true;
      }
    },
    text(text: string): void {
      if (capturing) {
        current += text;
      }
    },
    end(name: string): void {
      if (name === 'si') {
        strings.push(current.includes('_x') ? decodeCellText(current) : current);
      } else if (name === 'rPh') {
        phonetic = false;
      } else if (name === 't') {
        capturing = false;
      }
    },
  });
  tokenizer.push(xml);
  tokenizer.end();
  return strings;
}

/** `cellXfs[i]` resolved to date-ness through the built-in ids and the custom `numFmt` codes. */
function parseIsDateByXf(xml: string): Uint8Array {
  const customCodes = new Map<number, string>();
  const numberFormatIds: number[] = [];
  let inCellXfs = false;
  const tokenizer: XmlTokenizer = new XmlTokenizer({
    start(name: string): void {
      if (name === 'cellXfs') {
        inCellXfs = true;
      } else if (name === 'numFmt') {
        customCodes.set(Number(tokenizer.attr('numFmtId') ?? '0'), tokenizer.attr('formatCode') ?? '');
      } else if (name === 'xf' && inCellXfs) {
        numberFormatIds.push(Number(tokenizer.attr('numFmtId') ?? '0'));
      }
    },
    text(): void {},
    end(name: string): void {
      if (name === 'cellXfs') {
        inCellXfs = false;
      }
    },
  });
  tokenizer.push(xml);
  tokenizer.end();
  const isDateByXf = new Uint8Array(numberFormatIds.length);
  for (const [index, id] of numberFormatIds.entries()) {
    const code = customCodes.get(id);
    isDateByXf[index] = (code === undefined ? isBuiltinDateId(id) : isDateFormatCode(code)) ? 1 : 0;
  }
  return isDateByXf;
}

interface OpenFixture {
  rows: () => AsyncIterable<SheetRow>;
  close: () => Promise<void>;
}

async function openFirstWorksheet(fixture: Fixture, dates: 'local' | 'utc' | 'serial' = 'utc'): Promise<OpenFixture> {
  const zip = await ZipReader.open(sourceFrom(readFixture(fixture)), { maxEntries: 10_000, maxInflatedBytes: 1 << 30 });
  const names = [...zip.entries.keys()];
  const sheetPart = names.includes('xl/worksheets/sheet1.xml')
    ? 'xl/worksheets/sheet1.xml'
    : names.find(name => name.startsWith('xl/worksheets/') && name.endsWith('.xml'));
  if (sheetPart === undefined) {
    throw new Error(`${fixture.id}: no worksheet part`);
  }
  const date1904 = names.includes('xl/workbook.xml') && (await partText(zip, 'xl/workbook.xml')).includes('date1904="1"');
  const isDateByXf = names.includes('xl/styles.xml') ? parseIsDateByXf(await partText(zip, 'xl/styles.xml')) : new Uint8Array(0);
  const sharedStrings = async (): Promise<readonly string[]> =>
    names.includes('xl/sharedStrings.xml') ? parseSharedStringsXml(await partText(zip, 'xl/sharedStrings.xml')) : [];
  const readContext = context({ sharedStrings, isDateByXf, date1904, dates, errors: 'object' });
  return {
    rows: () => readWorksheetRows(zip.stream(sheetPart), readContext, { blankRows: true }),
    close: () => zip.close(),
  };
}

async function dumpFirstWorksheet(fixture: Fixture): Promise<TypedValue[][]> {
  const opened = await openFirstWorksheet(fixture);
  try {
    const dump: TypedValue[][] = [];
    for await (const row of opened.rows()) {
      dump.push(row.cells.map(value => toTyped(value, 'utc')));
    }
    return dump;
  } finally {
    await opened.close();
  }
}

interface FixtureDeviation {
  /** Mismatch categories the comparison is expected to report. */
  readonly categories: readonly string[];
  /** Why the file, not the reader, produces them. */
  readonly reason: string;
}

const EDGE_DEVIATIONS: Record<string, FixtureDeviation> = {
  'edge-sst-index-out-of-range': {
    categories: ['string-mismatch'],
    reason:
      'EC-SST-INDEX-OUT-OF-RANGE: the fixture pairs a 5-string sheet with a 2-string table, and indexes past the ' +
      'table read as blank text, which is what Excel shows',
  },
  'edge-date1904': {
    categories: ['temporal-mismatch'],
    reason: 'EC-DATE-1904: the fixture reuses the 1900-system serial with date1904="1", so honouring the flag moves the date',
  },
};

describe('fixtures/edge', () => {
  for (const fixture of fixturesWithTag('kind:edge')) {
    it(`${fixture.id} matches its expected dump`, async () => {
      const expected = readExpected(fixture);
      const firstSheet = expected.sheets[0];
      expect(firstSheet, `${fixture.id} has no expected sheet`).toBeDefined();
      const rows = await dumpFirstWorksheet(fixture);
      const expectedWorkbook: ExpectedWorkbook = { sheets: [{ name: 'Data', hidden: false, rows: firstSheet?.rows ?? [] }] };
      const report = compareDumps(
        expectedWorkbook,
        { reader: 'ours', sheets: [{ name: 'Data', hidden: false, rows }] },
        fixturePolicies(fixture),
      );
      const allowed = EDGE_DEVIATIONS[fixture.id]?.categories ?? [];
      expect(Object.keys(report.byCategory).toSorted(), describeReport(report)).toEqual(allowed.toSorted());
    });
  }
});

/**
 * What SheetJS's own writer did to the canonical data: catalog-recorded generator behaviour, not reader behaviour.
 * Every other cell of the 31x20 sheet, formula cells compared against their cached value, has to match exactly.
 */
const SHEETJS_GOLDEN_DEVIATION: FixtureDeviation = {
  categories: ['date-as-serial', 'dst-gap-shift-1h', 'escape-sequence-mangled', 'temporal-mismatch'],
  reason:
    'date-as-serial (30, the Time column): EC-DATE-TIME-ONLY-NEGATIVE-SERIAL - SheetJS writes time-only Dates as ' +
    'negative serials, which are not dates at all, so the number is what comes back. ' +
    'dst-gap-shift-1h (5, DateTime 2024-03-10T02:30): EC-DATE-DST-GAP - the wall clock does not exist in the US ' +
    'zone the fixture was written in, so the serial in the file is 03:30. ' +
    'escape-sequence-mangled (14, EscapeLiteral): EC-XML-ESCAPE-LITERAL and EC-XML-ESCAPE-OVERLAP - SheetJS writes ' +
    '_x0041_ undefused, so every reader (Excel included) decodes it back to "A". ' +
    'temporal-mismatch (6, Date 1899-12-31 and 1900-02-28): EC-DATE-PRE-1900 (written as serial 0) and ' +
    "EC-DATE-LEAP-WRITER-OFF-BY-ONE (1900-02-28 written as serial 60, Excel's fake leap day).",
};

describe('fixtures/golden', () => {
  it('golden-canonical-sheetjs matches the canonical dump apart from the generator deviations', async () => {
    const fixture = fixturesWithTag('kind:golden').find(candidate => candidate.id === 'golden-canonical-sheetjs');
    expect(fixture).toBeDefined();
    if (!fixture) {
      return;
    }
    const expected = readExpected(fixture);
    const dataSheet = expected.sheets.find(candidate => candidate.name === 'Data');
    expect(dataSheet).toBeDefined();
    const rows = await dumpFirstWorksheet(fixture);
    const report = compareDumps(
      { sheets: [{ name: 'Data', rows: dataSheet?.rows ?? [] }] },
      { reader: 'ours', sheets: [{ name: 'Data', hidden: false, rows }] },
      fixturePolicies(fixture),
    );
    expect(Object.keys(report.byCategory).toSorted(), describeReport(report, 40)).toEqual([...SHEETJS_GOLDEN_DEVIATION.categories]);
  });
});

// ---------------------------------------------------------------------------------------------------------------
// Throughput (printed, never asserted)
// ---------------------------------------------------------------------------------------------------------------

describe('throughput', () => {
  it('reads a 200k x 20 sheet', async () => {
    const rowCount = 200_000;
    const chunkBytes = 1024 * 1024;
    const chunks: Uint8Array[] = [];
    let pending = '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>';
    const flush = (force: boolean): void => {
      if (pending.length >= chunkBytes || force) {
        chunks.push(encoder.encode(pending));
        pending = '';
      }
    };
    for (let index = 1; index <= rowCount; index++) {
      pending +=
        `<row r="${index}">` +
        `<c r="A${index}"><v>${index}</v></c>` +
        `<c r="B${index}" t="s"><v>${index % 8}</v></c>` +
        `<c r="C${index}" t="inlineStr"><is><t>row ${index}</t></is></c>` +
        `<c r="D${index}" t="b"><v>${index % 2}</v></c>` +
        `<c r="E${index}" s="1"><v>${45351 + (index % 365)}.5</v></c>` +
        Array.from(
          { length: 15 },
          (_unused, offset) => `<c r="${String.fromCharCode(70 + offset)}${index}"><v>${index + offset}</v></c>`,
        ).join('') +
        '</row>';
      flush(false);
    }
    pending += '</sheetData></worksheet>';
    flush(true);
    const bytes = chunks.reduce((total, chunk) => total + chunk.byteLength, 0);
    const strings = Array.from({ length: 8 }, (_unused, index) => `shared ${index}`);

    async function* feedChunks(): AsyncGenerator<Uint8Array> {
      for (const chunk of chunks) {
        yield chunk;
      }
    }

    const started = performance.now();
    let rows = 0;
    let cells = 0;
    for await (const row of readWorksheetRows(
      feedChunks(),
      context({ sharedStrings: () => Promise.resolve(strings), isDateByXf: Uint8Array.from([0, 1]) }),
      {},
    )) {
      rows++;
      cells += row.cells.length;
    }
    const seconds = (performance.now() - started) / 1000;
    expect(rows).toBe(rowCount);
    expect(cells).toBe(rowCount * 20);
    // Reported, never asserted: machines differ, and a perf gate belongs in the bench harness. Written straight
    // to stderr because vitest swallows console output in this project's reporter setup.
    process.stderr.write(
      `worksheet-reader throughput: ${rows} rows x 20 cols, ${(bytes / (1024 * 1024)).toFixed(1)} MiB in ${seconds.toFixed(2)}s ` +
        `= ${Math.round(rows / seconds).toLocaleString('en-US')} rows/s, ${Math.round(cells / seconds).toLocaleString('en-US')} cells/s\n`,
    );
  }, 240_000);
});
