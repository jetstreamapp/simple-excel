import { describe, expect, it } from 'vitest';
import { isXlsxError, type XlsxErrorCode } from '../../errors';
import type { CellInput, SheetOptions, StyleId } from '../../types';
import type { ZipEntrySummary, ZipEntryWriter } from '../../zip/zip-writer';
import { columnLetters } from '../cell-ref';
import { SharedStringWriter } from '../shared-strings';
import { StyleRegistry } from '../styles';
import { MergedRanges, validateSheetOptions, WorksheetWriter, type WorksheetWriterContext } from '../worksheet-writer';

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const WORKSHEET_OPEN =
  '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' +
  ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';
const PAGE_MARGINS = '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>';
const MAX_CELL_CHARS = 32_767;

class CollectingEntry implements ZipEntryWriter {
  readonly chunks: string[] = [];
  closed = false;

  async write(chunk: Uint8Array): Promise<void> {
    this.chunks.push(new TextDecoder().decode(chunk));
  }

  async close(): Promise<ZipEntrySummary> {
    this.closed = true;
    return { name: 'xl/worksheets/sheet1.xml', crc32: 0, compressedSize: 0, uncompressedSize: 0, offset: 0 };
  }

  get xml(): string {
    return this.chunks.join('');
  }
}

interface Harness {
  readonly entry: CollectingEntry;
  readonly styles: StyleRegistry;
  readonly worksheet: WorksheetWriter;
}

/** A registry holding `count` styles in all (ids 0 to count - 1), so tests can use small ids without caring what they are. */
function registryWithStyles(count: number): StyleRegistry {
  const styles = new StyleRegistry();
  for (let size = 1; styles.count < count; size++) {
    styles.register({ font: { size } });
  }
  return styles;
}

function createWorksheet(options: SheetOptions = {}, context: Partial<WorksheetWriterContext> = {}): Harness {
  const entry = new CollectingEntry();
  const styles = context.styles ?? registryWithStyles(8);
  const worksheet = new WorksheetWriter(
    {
      entry,
      styles,
      sharedStrings: null,
      date1904: false,
      dates: 'local',
      cellOverflow: 'truncate',
      truncationSuffix: '...(truncated)',
      sheetName: 'Data',
      ...context,
    },
    options,
  );
  return { entry, styles, worksheet };
}

/** The `<sheetData>` payload of a sheet written from one batch of rows. */
async function rowsXml(rows: readonly (readonly CellInput[])[], styles?: StyleId | readonly (StyleId | undefined)[]): Promise<string> {
  const { entry, worksheet } = createWorksheet();
  for (const row of rows) {
    await worksheet.writeRow(row, styles);
  }
  await worksheet.close();
  const [, body = ''] = /<sheetData>(.*)<\/sheetData>/s.exec(entry.xml) ?? [];
  return body;
}

async function oneCell(value: CellInput, styles?: StyleId | readonly (StyleId | undefined)[]): Promise<string> {
  return rowsXml([[value]], styles);
}

/** The `XlsxErrorCode` a call rejects (or throws) with, so the assertion stays in the test that cares about it. */
async function errorCode(run: () => unknown): Promise<XlsxErrorCode | string> {
  try {
    await run();
  } catch (error) {
    return isXlsxError(error) ? error.code : `not an XlsxError: ${String(error)}`;
  }
  return 'no error was thrown';
}

/** What a call rejects (or throws) with, for assertions on the message as well as the code. */
async function rejection(run: () => unknown): Promise<unknown> {
  try {
    await run();
  } catch (error) {
    return error;
  }
  return new Error('no error was thrown');
}

describe('worksheet prologue (CT_Worksheet element order)', () => {
  it('writes sheetViews, sheetFormatPr and cols in schema order', async () => {
    const { entry, worksheet } = createWorksheet(
      {
        header: ['Id', 'Name', 'Amount'],
        rowCount: 2,
        columns: [{ width: 16 }, { width: 24, hidden: true }, { style: 3 }],
        freeze: { rows: 1 },
      },
      { tabSelected: true },
    );
    await worksheet.writeRow(['Id', 'Name', 'Amount']);
    await worksheet.writeRow(['a', 'b', 1]);
    await worksheet.writeRow(['c', 'd', 2]);
    await worksheet.close();

    expect(entry.xml.slice(0, entry.xml.indexOf('<sheetData>'))).toBe(
      `${XML_DECLARATION}${WORKSHEET_OPEN}` +
        '<sheetViews><sheetView workbookViewId="0" tabSelected="1">' +
        '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>' +
        '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>' +
        '</sheetView></sheetViews>' +
        '<sheetFormatPr defaultRowHeight="15"/>' +
        '<cols><col min="1" max="1" width="16" customWidth="1"/><col min="2" max="2" width="24" customWidth="1" hidden="1"/>' +
        '<col min="3" max="3" style="3"/></cols>',
    );
  });

  it('EC-DIMENSION-FROM-HINT: never writes <dimension>, whatever the hints say', async () => {
    for (const options of [
      { header: ['Id'] },
      { rowCount: 10 },
      { rowCount: 2, columns: Array.from({ length: 20 }, () => ({})), header: ['Id', 'Name'] },
      { rowCount: 0, header: ['Id'] },
    ] satisfies SheetOptions[]) {
      const { entry, worksheet } = createWorksheet(options);
      // More rows and wider rows than the hints announce: a dimension written up front could only be wrong.
      for (let row = 0; row < 5; row++) {
        await worksheet.writeRow(['a', 'b', 'c', 'd']);
      }
      const summary = await worksheet.close();
      expect(entry.xml).not.toContain('<dimension');
      expect(summary.dimension, 'the summary still reports the real extent').toBe('A1:D5');
    }
  });

  it('EC-COLS-WIDTH-OVER-255: clamps a width past 255 and keeps the rest as given', async () => {
    const { entry, worksheet } = createWorksheet({ columns: [{ width: 300 }, { width: 255 }, { width: 0 }, { width: 8.43 }] });
    await worksheet.close();
    expect(entry.xml).toContain(
      '<cols><col min="1" max="1" width="255" customWidth="1"/><col min="2" max="2" width="255" customWidth="1"/>' +
        '<col min="3" max="3" width="0" customWidth="1"/><col min="4" max="4" width="8.43" customWidth="1"/></cols>',
    );
  });

  it('freezes columns and both axes with the matching active pane', async () => {
    const columnsOnly = createWorksheet({ freeze: { cols: 2 } });
    await columnsOnly.worksheet.close();
    expect(columnsOnly.entry.xml).toContain('<pane xSplit="2" topLeftCell="C1" activePane="topRight" state="frozen"/>');

    const both = createWorksheet({ freeze: { rows: 2, cols: 1 } });
    await both.worksheet.close();
    expect(both.entry.xml).toContain('<pane xSplit="1" ySplit="2" topLeftCell="B3" activePane="bottomRight" state="frozen"/>');
  });

  it('writes a bare sheetView when nothing is frozen and no cols when nothing is set', async () => {
    const { entry, worksheet } = createWorksheet({ columns: [{}, {}] });
    await worksheet.close();
    expect(entry.xml).toContain('<sheetViews><sheetView workbookViewId="0"/></sheetViews>');
    expect(entry.xml).not.toContain('<cols>');
  });

  it('closes the entry with the epilogue in schema order', async () => {
    const { entry, worksheet } = createWorksheet({ autoFilter: true });
    await worksheet.writeRow(['Id', 'Name']);
    await worksheet.writeRow(['a', 'b']);
    worksheet.merge('A1:B1');
    const summary = await worksheet.close();

    expect(entry.closed).toBe(true);
    expect(entry.xml.slice(entry.xml.indexOf('</sheetData>'))).toBe(
      `</sheetData><autoFilter ref="A1:B2"/><mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>${PAGE_MARGINS}</worksheet>`,
    );
    expect(summary).toEqual({ rows: 2, columns: 2, truncatedCells: 0, dimension: 'A1:B2' });
  });
});

describe('cell serialization', () => {
  it('writes numbers as the shortest round-trip xsd:double (EC-NUM-NEGATIVE-ZERO)', async () => {
    expect(await oneCell(123.45)).toBe('<row r="1"><c r="A1"><v>123.45</v></c></row>');
    expect(await oneCell(-0), 'Excel has no negative zero').toBe('<row r="1"><c r="A1"><v>0</v></c></row>');
    expect(await oneCell(0.1 + 0.2)).toContain('<v>0.30000000000000004</v>');
    expect(await oneCell(1e21), 'Excel writes the exponent uppercase').toContain('<v>1E+21</v>');
    expect(await oneCell(1e-7)).toContain('<v>1E-7</v>');
    expect(await oneCell(1.7976931348623157e308)).toContain('<v>1.7976931348623157E+308</v>');
    expect(await oneCell(-123.456)).toContain('<v>-123.456</v>');
  });

  it('turns NaN and the infinities into #NUM! error cells', async () => {
    expect(await oneCell(Number.NaN)).toContain('<c r="A1" t="e"><v>#NUM!</v></c>');
    expect(await oneCell(Number.POSITIVE_INFINITY)).toContain('<c r="A1" t="e"><v>#NUM!</v></c>');
    expect(await oneCell(Number.NEGATIVE_INFINITY)).toContain('<c r="A1" t="e"><v>#NUM!</v></c>');
  });

  it('writes booleans and error values', async () => {
    expect(await oneCell(true)).toContain('<c r="A1" t="b"><v>1</v></c>');
    expect(await oneCell(false)).toContain('<c r="A1" t="b"><v>0</v></c>');
    expect(await oneCell({ error: '#N/A' })).toContain('<c r="A1" t="e"><v>#N/A</v></c>');
    expect(await oneCell({ error: '#DIV/0!' })).toContain('<c r="A1" t="e"><v>#DIV/0!</v></c>');
  });

  it('EC-ERROR-CODE-UNKNOWN: writes every CellErrorCode, other #-codes as text, and refuses anything else', async () => {
    for (const code of ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA'] as const) {
      expect(await oneCell({ error: code })).toContain(`<c r="A1" t="e"><v>${code}</v></c>`);
    }
    // Newer Excel errors the reader can return, and misspellings, never reach t="e": they are written as text.
    for (const code of ['#SPILL!', '#CALC!', '#n/a']) {
      const xml = await oneCell({ error: code } as unknown as CellInput);
      expect(xml).toContain(`<t>${code}</t>`);
      expect(xml).not.toContain('t="e"');
    }
    for (const code of ['any text', '', '<b>']) {
      const failure = await rejection(() => oneCell({ error: code } as unknown as CellInput));
      expect(failure).toMatchObject({ code: 'WRITER_STATE' });
      expect(String((failure as Error).message)).toContain('Sheet "Data" cell A1');
    }
  });

  it('never infers a formula from text (EC-FORMULA-LIKE-TEXT-WRITTEN-AS-FORMULA)', async () => {
    const xml = await rowsXml([['=SUM(A1)', '=cmd|calc', '+1', '@x']]);
    expect(xml).toContain('<c r="A1" t="inlineStr"><is><t>=SUM(A1)</t></is></c>');
    expect(xml).toContain('<c r="B1" t="inlineStr"><is><t>=cmd|calc</t></is></c>');
    expect(xml).not.toContain('<f>');
  });

  it('escapes cell text and preserves whitespace', async () => {
    expect(await oneCell('  padded  ')).toContain('<t xml:space="preserve">  padded  </t>');
    expect(await oneCell('a & b <c>')).toContain('<t>a &amp; b &lt;c&gt;</t>');
    expect(await oneCell('_x0041_')).toContain('<t>_x005F_x0041_</t>');
    expect(await oneCell('nul ?')).toContain('<t>nul_x0000_?</t>');
    expect(await oneCell('')).toBe('<row r="1"><c r="A1" t="inlineStr"><is><t></t></is></c></row>');
  });

  it('writes shared-string indexes when the table accepts the string', async () => {
    const sharedStrings = new SharedStringWriter({ maxUnique: 2, maxChars: 1024, maxLength: 32 });
    const { entry, worksheet } = createWorksheet({}, { sharedStrings });
    await worksheet.writeRow(['Account', 'Contact', 'Account', 'Lead']);
    await worksheet.close();

    expect(entry.xml).toContain('<c r="A1" t="s"><v>0</v></c>');
    expect(entry.xml).toContain('<c r="B1" t="s"><v>1</v></c>');
    expect(entry.xml).toContain('<c r="C1" t="s"><v>0</v></c>');
    expect(entry.xml, 'the budget froze, so the new string is inline in the same row').toContain(
      '<c r="D1" t="inlineStr"><is><t>Lead</t></is></c>',
    );
    expect(sharedStrings.stats).toEqual({ count: 3, uniqueCount: 2, frozen: true });
  });

  it('writes blanks only when they carry a style', async () => {
    expect(await rowsXml([[null, undefined, 1]])).toBe('<row r="1"><c r="C1"><v>1</v></c></row>');
    expect(await rowsXml([[null]], 4)).toBe('<row r="1"><c r="A1" s="4"/></row>');
    expect(await rowsXml([[null]], 0), 'style 0 is the default every cell already has').toBe('<row r="1"></row>');
  });

  it('writes a bigint as a number while it is exact and as text past 2^53', async () => {
    expect(await oneCell(42n)).toContain('<c r="A1"><v>42</v></c>');
    expect(await oneCell(-9007199254740992n)).toContain('<c r="A1"><v>-9007199254740992</v></c>');
    expect(await oneCell(123456789012345678901n)).toContain('<c r="A1" t="inlineStr"><is><t>123456789012345678901</t></is></c>');
  });

  it('EC-STYLE-ID-RANGE: refuses a style id that was never registered, scalar or per cell', async () => {
    const styles = registryWithStyles(3);
    for (const badStyles of [3, -1, 1.5, Number.NaN, [undefined, 3], [0, 2, 99]] as (StyleId | (StyleId | undefined)[])[]) {
      const { worksheet } = createWorksheet({}, { styles });
      const failure = await rejection(() => worksheet.writeRow(['a', 'b', 'c'], badStyles));
      expect(failure, String(badStyles)).toMatchObject({ code: 'WRITER_STATE' });
      expect(String((failure as Error).message)).toContain('is not a registered style id (this workbook has ids 0 to 2)');
    }

    const perCell = createWorksheet({}, { styles });
    await perCell.worksheet.writeRow(['a', 'b', 'c'], [0, 1, 2]);
    const failure = await rejection(() => perCell.worksheet.writeRow(['x', 'y'], [2, 7]));
    expect(String((failure as Error).message), 'the message names the cell').toContain('Sheet "Data" cell B2');

    const { worksheet } = createWorksheet({}, { styles });
    expect(await errorCode(() => worksheet.writeRow(['a'], '1' as unknown as StyleId))).toBe('WRITER_STATE');
  });

  it('applies per-cell styles and omits the default style', async () => {
    expect(await rowsXml([['a', 'b']], [undefined, 7])).toBe(
      '<row r="1"><c r="A1" t="inlineStr"><is><t>a</t></is></c><c r="B1" s="7" t="inlineStr"><is><t>b</t></is></c></row>',
    );
    expect(await rowsXml([[1, 2]], 5)).toBe('<row r="1"><c r="A1" s="5"><v>1</v></c><c r="B1" s="5"><v>2</v></c></row>');
  });

  it('EC-CELL-UNSUPPORTED-TYPE: refuses a value that is not a cell value, naming the sheet, cell and type', async () => {
    const cases: ReadonlyArray<readonly [unknown, string]> = [
      [{ Id: 1 }, 'object'],
      [[1, 2], 'array'],
      [() => 1, 'function'],
      [Symbol('nope'), 'symbol'],
      [new String('boxed'), 'object (String)'],
      [new Map(), 'object (Map)'],
    ];
    for (const [value, type] of cases) {
      const { worksheet } = createWorksheet();
      await worksheet.writeRow(['first row']);
      const failure = await rejection(() => worksheet.writeRow(['a', 'b', value as CellInput]));
      expect(failure).toMatchObject({ code: 'WRITER_STATE' });
      expect((failure as Error).message).toBe(
        `Sheet "Data" cell C2: a value of type ${type} cannot be written. Convert it to a string, number, boolean, Date or null first.`,
      );
    }
  });
});

describe('date cells', () => {
  it('EC-DATE-STYLE-MERGE: merges the default date format into a caller style that has none', async () => {
    const styles = new StyleRegistry();
    const bold = styles.register({ font: { bold: true }, fill: { color: '#FFFF00' }, alignment: { horizontal: 'center' } });
    const { entry, worksheet } = createWorksheet({}, { styles });
    await worksheet.writeRow([new Date(2024, 1, 29, 12, 0, 0), 'label'], bold);
    await worksheet.writeRow([new Date(2024, 2, 1)], bold);
    await worksheet.close();

    expect(styles.count, 'derived once, then reused').toBe(3);
    const derived = styles.dateStyleFor(bold);
    expect(derived).not.toBe(bold);
    expect(derived, 'not the bare default date style either').not.toBe(styles.defaultDateStyle);
    expect(styles.isDateStyle(derived)).toBe(true);
    expect(entry.xml).toContain(`<c r="A1" s="${derived}"><v>45351.5</v></c>`);
    expect(entry.xml, 'the text cell keeps the caller style').toContain(`<c r="B1" s="${bold}" t="inlineStr">`);
    expect(entry.xml).toContain(`<c r="A2" s="${derived}"><v>45352</v></c>`);
    expect(styles.toXml()).toContain(
      `<xf numFmtId="164" fontId="1" fillId="2" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyAlignment="1">` +
        '<alignment horizontal="center"/></xf>',
    );
  });

  it('EC-DATE-STYLE-MERGE: a caller style with a non-date format gets the date format instead of it', async () => {
    const styles = new StyleRegistry();
    const money = styles.register({ numFmt: '0.00', font: { italic: true } });
    const { entry, worksheet } = createWorksheet({}, { styles });
    await worksheet.writeRow([new Date(2024, 1, 29)], money);
    await worksheet.close();
    expect(entry.xml).toContain(`<c r="A1" s="${styles.dateStyleFor(money)}"><v>45351</v></c>`);
    expect(styles.isDateStyle(styles.dateStyleFor(money))).toBe(true);
  });

  it('uses the default date style for a Date under style 0 or no style', async () => {
    const styles = new StyleRegistry();
    const { entry, worksheet } = createWorksheet({}, { styles });
    await worksheet.writeRow([new Date(2024, 1, 29), new Date(2024, 1, 29)], [0, undefined]);
    await worksheet.close();
    expect(entry.xml).toContain(`<c r="A1" s="${styles.defaultDateStyle}"><v>45351</v></c>`);
    expect(entry.xml).toContain(`<c r="B1" s="${styles.defaultDateStyle}"><v>45351</v></c>`);
    expect(styles.count).toBe(2);
  });

  it('keeps a caller style that already carries a date format', async () => {
    const styles = new StyleRegistry();
    const dateStyle = styles.register({ numFmt: 'yyyy-mm-dd' });
    const { entry, worksheet } = createWorksheet({}, { styles });
    await worksheet.writeRow([new Date(2024, 1, 29)], dateStyle);
    await worksheet.close();

    expect(entry.xml).toContain(`<c r="A1" s="${dateStyle}"><v>45351</v></c>`);
  });

  it('reads the UTC fields when dates: utc', async () => {
    const local = createWorksheet();
    await local.worksheet.writeRow([new Date(Date.UTC(2024, 1, 29, 12, 0, 0))]);
    await local.worksheet.close();

    const utc = createWorksheet({}, { dates: 'utc' });
    await utc.worksheet.writeRow([new Date(Date.UTC(2024, 1, 29, 12, 0, 0))]);
    await utc.worksheet.close();

    expect(utc.entry.xml).toContain('<v>45351.5</v>');
    // The same instant read through local fields only matches where that instant's offset is zero (ADR-003). The
    // offset has to be the one on that February date, not today's: Europe/London is UTC then but UTC+1 in summer.
    const instant = new Date(Date.UTC(2024, 1, 29, 12, 0, 0));
    expect(local.entry.xml === utc.entry.xml).toBe(instant.getTimezoneOffset() === 0);
  });

  it('writes the serials the 1900 leap-year bug demands (EC-DATE-1900-LEAP-BUG)', async () => {
    const { entry, worksheet } = createWorksheet();
    await worksheet.writeRow([new Date(1900, 0, 1), new Date(1900, 1, 28), new Date(1900, 2, 1)]);
    await worksheet.close();

    expect(entry.xml).toContain('<v>1</v>');
    expect(entry.xml, '1900-02-28 is serial 59, not 60').toContain('<v>59</v>');
    expect(entry.xml).toContain('<v>61</v>');
  });

  it('writes time-only values as a fraction of a day (EC-DATE-TIME-ONLY-NEGATIVE-SERIAL)', async () => {
    expect(await oneCell(new Date(1899, 11, 30, 12, 0, 0))).toContain('<v>0.5</v>');
  });

  it('writes pre-epoch dates as ISO text (EC-DATE-PRE-1900)', async () => {
    expect(await oneCell(new Date(1899, 11, 31))).toContain('<c r="A1" t="inlineStr"><is><t>1899-12-31T00:00:00</t></is></c>');
    expect(await oneCell(new Date(1850, 5, 4, 7, 8, 9))).toContain('<t>1850-06-04T07:08:09</t>');
  });

  it('EC-DATE-PRE-1900: writes a negative year as expanded ISO text with a sign', async () => {
    const { entry, worksheet } = createWorksheet({}, { dates: 'utc' });
    const yearMinus50 = new Date(Date.UTC(2000, 0, 1));
    yearMinus50.setUTCFullYear(-50);
    const yearZero = new Date(Date.UTC(2000, 5, 15, 1, 2, 3));
    yearZero.setUTCFullYear(0);
    await worksheet.writeRow([yearMinus50, yearZero]);
    await worksheet.close();
    expect(entry.xml).toContain('<c r="A1" t="inlineStr"><is><t>-0050-01-01T00:00:00</t></is></c>');
    expect(entry.xml).toContain('<c r="B1" t="inlineStr"><is><t>0000-06-15T01:02:03</t></is></c>');
  });

  it('EC-DATE-SERIAL-OVER-9999: writes a Date past 9999 as ISO text, not a number', async () => {
    const { entry, worksheet } = createWorksheet({}, { dates: 'utc' });
    await worksheet.writeRow([new Date(Date.UTC(10_000, 0, 1)), new Date(Date.UTC(9999, 11, 31, 23, 59, 59))]);
    await worksheet.close();
    expect(entry.xml).toContain('<c r="A1" t="inlineStr"><is><t>10000-01-01T00:00:00</t></is></c>');
    expect(entry.xml, 'the last Excel day is still a serial').toMatch(/<c r="B1" s="\d+"><v>2958465\.99998/);
  });

  it('shifts the epoch for a 1904 workbook (EC-DATE-1904)', async () => {
    const { entry, worksheet } = createWorksheet({}, { date1904: true });
    await worksheet.writeRow([new Date(1904, 0, 1), new Date(2024, 1, 29)]);
    await worksheet.close();

    expect(entry.xml).toContain('<v>0</v>');
    expect(entry.xml).toContain('<v>43889</v>');
  });

  it('writes nothing for an invalid Date', async () => {
    expect(await rowsXml([[new Date('nonsense'), 1]])).toBe('<row r="1"><c r="B1"><v>1</v></c></row>');
  });
});

describe('cell length policy (EC-CELL-32767-LIMIT)', () => {
  const long = 'x'.repeat(40_000);

  it('truncates with a visible suffix and counts the cell', async () => {
    const { entry, worksheet } = createWorksheet();
    await worksheet.writeRow([long, 'short', long]);
    const summary = await worksheet.close();

    const written = /<t>(x+\.\.\.\(truncated\))<\/t>/.exec(entry.xml)?.[1] ?? '';
    expect(written.length).toBe(MAX_CELL_CHARS);
    expect(written.endsWith('...(truncated)')).toBe(true);
    expect(summary.truncatedCells).toBe(2);
  });

  it('honours a custom suffix', async () => {
    const { entry, worksheet } = createWorksheet({}, { truncationSuffix: '[cut]' });
    await worksheet.writeRow([long]);
    await worksheet.close();
    expect(entry.xml).toContain('[cut]</t>');
  });

  it('throws CELL_TOO_LONG when asked to', async () => {
    const { worksheet } = createWorksheet({}, { cellOverflow: 'throw' });
    expect(await errorCode(() => worksheet.writeRow([long]))).toBe('CELL_TOO_LONG');
  });

  it('EC-TRUNCATION-SURROGATE: never splits a surrogate pair at the cut', async () => {
    const suffix = '...(truncated)';
    // The last kept unit would be the high half of an emoji: the cut moves one unit earlier.
    const straddling = `${'x'.repeat(MAX_CELL_CHARS - suffix.length - 1)}${'😀'.repeat(10)}`;
    const { entry, worksheet } = createWorksheet();
    await worksheet.writeRow([straddling]);
    await worksheet.close();
    const written = /<t>([^<]*)<\/t>/.exec(entry.xml)?.[1] ?? '';
    expect(written).not.toContain('\uFFFD');
    expect(written.endsWith(`x${suffix}`)).toBe(true);
    expect(written.length).toBe(MAX_CELL_CHARS - 1);

    // A pair ending exactly at the cut is kept whole.
    const aligned = `${'x'.repeat(MAX_CELL_CHARS - suffix.length - 2)}${'😀'.repeat(10)}`;
    const second = createWorksheet();
    await second.worksheet.writeRow([aligned]);
    await second.worksheet.close();
    const kept = /<t>([^<]*)<\/t>/.exec(second.entry.xml)?.[1] ?? '';
    expect(kept.endsWith(`😀${suffix}`)).toBe(true);
    expect(kept.length).toBe(MAX_CELL_CHARS);
  });

  it('EC-TRUNCATION-SURROGATE: a suffix longer than the limit is cut on a pair boundary too', async () => {
    const { entry, worksheet } = createWorksheet({}, { truncationSuffix: `${'s'.repeat(MAX_CELL_CHARS - 1)}😀` });
    await worksheet.writeRow(['y'.repeat(40_000)]);
    await worksheet.close();
    const written = /<t>([^<]*)<\/t>/.exec(entry.xml)?.[1] ?? '';
    expect(written).toBe('s'.repeat(MAX_CELL_CHARS - 1));
  });

  it('leaves a cell exactly at the limit alone', async () => {
    const { worksheet } = createWorksheet();
    await worksheet.writeRow(['y'.repeat(MAX_CELL_CHARS)]);
    const summary = await worksheet.close();
    expect(summary.truncatedCells).toBe(0);
  });
});

describe('merges, autofilter and limits', () => {
  it('normalizes a range and emits the merges in order', async () => {
    const { entry, worksheet } = createWorksheet();
    await worksheet.writeRow(['a', 'b', 'c']);
    worksheet.merge('C1:A1');
    worksheet.merge('A2:B3');
    await worksheet.close();

    expect(entry.xml).toContain('<mergeCells count="2"><mergeCell ref="A1:C1"/><mergeCell ref="A2:B3"/></mergeCells>');
  });

  it('rejects a malformed range, a single-cell merge and a merge after close', async () => {
    const { worksheet } = createWorksheet();
    expect(await errorCode(() => worksheet.merge('not a range'))).toBe('WRITER_STATE');
    expect(await errorCode(() => worksheet.merge('A1:A1'))).toBe('WRITER_STATE');
    await worksheet.close();
    expect(await errorCode(() => worksheet.merge('A1:B1'))).toBe('WRITER_STATE');
  });

  it('EC-MERGE-OVERLAP: accepts lowercase, refuses ranges outside the grid, overlaps and duplicates', () => {
    const merges = new MergedRanges('Accounts');
    merges.add('a1:c1');
    merges.add('$D$1:E2');
    merges.add('A3:B1048576');
    expect(merges.ranges).toEqual(['A1:C1', 'D1:E2', 'A3:B1048576']);

    const refused: ReadonlyArray<readonly [unknown, string]> = [
      ['A1:XFE1', 'is not a range like A1:C1 inside A1:XFD1048576'],
      ['A1:A1048577', 'is not a range like A1:C1 inside A1:XFD1048576'],
      ['A1', 'merges a single cell'],
      [42, 'a number is not a range'],
      ['C1:A1', 'is already merged'],
      ['B1:B2', 'overlaps the merged range "A1:C1"'],
      ['E2:F3', 'overlaps the merged range "D1:E2"'],
      ['C500000:B500001', 'overlaps the merged range "A3:B1048576"'],
      ['a1:a2', 'overlaps the merged range "A1:C1"'],
    ];
    for (const [range, message] of refused) {
      let failure: unknown;
      try {
        merges.add(range as string);
      } catch (error) {
        failure = error;
      }
      expect(failure, String(range)).toMatchObject({ code: 'WRITER_STATE' });
      expect((failure as Error).message, String(range)).toContain('Sheet "Accounts"');
      expect((failure as Error).message, String(range)).toContain(message);
    }
    expect(merges.ranges, 'a refused range is not registered').toHaveLength(3);
    merges.add('C2:C3');
    expect(merges.ranges).toHaveLength(4);
  });

  it('EC-MERGE-OVERLAP: checks thousands of merges quickly and exactly', () => {
    const merges = new MergedRanges('Permissions');
    const started = performance.now();
    // A grid of 2x2 merges, 60 across and 100 down: every cell of A1:DH200 is merged exactly once.
    for (let row = 1; row <= 200; row += 2) {
      for (let column = 0; column < 120; column += 2) {
        merges.add(`${columnLetters(column)}${row}:${columnLetters(column + 1)}${row + 1}`);
      }
    }
    expect(merges.ranges).toHaveLength(6000);
    expect(() => merges.add('B2:C3')).toThrow(/overlaps/);
    expect(() => merges.add('A1:A1048576')).toThrow(/overlaps/);
    merges.add('A201:DH201');
    expect(performance.now() - started, 'no quadratic scan').toBeLessThan(2000);
  });

  it('covers the header row alone when that is all there is', async () => {
    const { entry, worksheet } = createWorksheet({ autoFilter: true });
    await worksheet.writeRow(['Id', 'Name']);
    await worksheet.close();
    expect(entry.xml).toContain('<autoFilter ref="A1:B1"/>');
  });

  it('omits the autofilter for an empty sheet', async () => {
    const { entry, worksheet } = createWorksheet({ autoFilter: true });
    const summary = await worksheet.close();
    expect(entry.xml).not.toContain('<autoFilter');
    expect(summary).toEqual({ rows: 0, columns: 0, truncatedCells: 0, dimension: null });
  });

  it('rejects a row wider than the sheet', async () => {
    const { worksheet } = createWorksheet();
    expect(await errorCode(() => worksheet.writeRow(Array.from({ length: 16_385 }, () => 1)))).toBe('ROW_OUT_OF_RANGE');
  });

  it('counts rows monotonically and refuses writes after close', async () => {
    const { worksheet } = createWorksheet();
    expect(worksheet.nextRow).toBe(1);
    await worksheet.writeRow(['a']);
    await worksheet.writeRow(['b']);
    expect(worksheet.nextRow).toBe(3);
    await worksheet.close();
    expect(await errorCode(() => worksheet.writeRow(['c']))).toBe('WRITER_STATE');
    expect(await errorCode(() => worksheet.close())).toBe('WRITER_STATE');
  });

  it('reports the same failure to every later call', async () => {
    const { worksheet } = createWorksheet({}, { cellOverflow: 'throw' });
    expect(await errorCode(() => worksheet.writeRow(['z'.repeat(40_000)]))).toBe('CELL_TOO_LONG');
    expect(await errorCode(() => worksheet.writeRow(['fine']))).toBe('CELL_TOO_LONG');
    expect(await errorCode(() => worksheet.close())).toBe('CELL_TOO_LONG');
  });
});

describe('validateSheetOptions', () => {
  const codeOf = (options: unknown, styleCount = 3): string => {
    try {
      validateSheetOptions(options as SheetOptions, styleCount);
    } catch (error) {
      return isXlsxError(error) ? error.code : `not an XlsxError: ${String(error)}`;
    }
    return 'accepted';
  };

  it('EC-DIMENSION-FROM-HINT: rowCount must be a non-negative safe integer', () => {
    for (const rowCount of [0, 1, 1_048_575, Number.MAX_SAFE_INTEGER]) {
      expect(codeOf({ rowCount }), String(rowCount)).toBe('accepted');
    }
    for (const rowCount of [-1, 1.5, Number.NaN, Number.POSITIVE_INFINITY, 2 ** 53, '10', null]) {
      expect(codeOf({ rowCount }), String(rowCount)).toBe('WRITER_STATE');
    }
  });

  it('EC-COLS-WIDTH-OVER-255: accepts any finite width of 0 or more and refuses NaN, negative and infinite ones', () => {
    for (const width of [0, 8.43, 255, 300, 10_000]) {
      expect(codeOf({ columns: [{ width }] }), String(width)).toBe('accepted');
    }
    for (const width of [-1, Number.NaN, Number.POSITIVE_INFINITY, Number.NEGATIVE_INFINITY, '20']) {
      expect(codeOf({ columns: [{ width }] }), String(width)).toBe('WRITER_STATE');
    }
  });

  it('EC-STYLE-ID-RANGE: headerStyle and column styles must be registered ids', () => {
    expect(codeOf({ headerStyle: 2, columns: [{ style: 0 }, { style: 2 }] })).toBe('accepted');
    expect(codeOf({ headerStyle: false })).toBe('accepted');
    for (const style of [3, -1, 0.5, Number.NaN, '1']) {
      expect(codeOf({ headerStyle: style }), `headerStyle ${String(style)}`).toBe('WRITER_STATE');
      expect(codeOf({ columns: [{}, { style }] }), `column style ${String(style)}`).toBe('WRITER_STATE');
    }
  });

  it('refuses options, headers and columns of the wrong shape', () => {
    expect(codeOf(null)).toBe('WRITER_STATE');
    expect(codeOf({ header: { Id: 1 } })).toBe('WRITER_STATE');
    expect(codeOf({ columns: { width: 3 } })).toBe('WRITER_STATE');
    expect(codeOf({ columns: [null] })).toBe('WRITER_STATE');
    expect(codeOf({ columns: Array.from({ length: 16_385 }, () => ({})) })).toBe('ROW_OUT_OF_RANGE');
  });
});

describe('chunked output', () => {
  it('only touches the entry at flush boundaries', async () => {
    const { entry, worksheet } = createWorksheet();
    for (let i = 0; i < 20; i++) {
      await worksheet.writeRow(['a', 'b', 'c']);
    }
    expect(entry.chunks.length, 'twenty short rows are nowhere near 64 Ki characters').toBe(0);

    for (let i = 0; i < 2000; i++) {
      await worksheet.writeRow([`row ${i}`, i, 'some text that adds up over two thousand rows']);
    }
    expect(entry.chunks.length).toBeGreaterThan(0);
    for (const chunk of entry.chunks) {
      expect(chunk.length).toBeLessThan(64 * 1024 + 4096);
    }

    await worksheet.close();
    expect(entry.xml.startsWith(XML_DECLARATION)).toBe(true);
    expect(entry.xml.endsWith('</worksheet>')).toBe(true);
    expect(entry.xml.match(/<row /g)?.length).toBe(2020);
  });
});
