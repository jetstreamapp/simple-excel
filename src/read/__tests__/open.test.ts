/**
 * The reader facade: what each kind of input is rejected as, how the package parts are located, and the row, object
 * and `head()` shapes callers get. Packages are hand-built from raw XML so a single part can be wrong at a time; the
 * real corpus is covered by `test/corpus.test.ts`.
 */
import { runInNewContext } from 'node:vm';
import { describe, expect, it } from 'vitest';
import { fixtureById, readFixture } from '../../../test/helpers/fixtures';
import { isXlsxError, XlsxError } from '../../errors';
import type {
  CellError,
  CellValue,
  ObjectRowsOptions,
  ObjectsResult,
  OpenOptions,
  RandomAccessSource,
  RawCell,
  ReadValue,
  ReadWarning,
  RowsOptions,
} from '../../types';
import { type SourceInput, sourceFrom } from '../../zip/source';
import { ZipReader } from '../../zip/zip-reader';
import { openWorkbook } from '../open';

const MINIZIP = new URL('../../../fixtures/generators/edge/minizip.mjs', import.meta.url).href;

const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_PACKAGE_RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const NS_CONTENT_TYPES = 'http://schemas.openxmlformats.org/package/2006/content-types';
const SPREADSHEET_MAIN = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml';

interface ZipEntry {
  name: string;
  data: string;
}

interface ZipOptions {
  truncateAt?: number;
  corruptCrc?: boolean;
}

async function zip(parts: Record<string, string>, options?: ZipOptions): Promise<Uint8Array> {
  const { buildZip } = (await import(MINIZIP)) as { buildZip: (entries: ZipEntry[], options?: ZipOptions) => Uint8Array };
  return buildZip(
    Object.entries(parts).map(([name, data]) => ({ name, data })),
    options,
  );
}

// ---------------------------------------------------------------------------------------------------------------
// Package parts
// ---------------------------------------------------------------------------------------------------------------

function contentTypes(overrides: Record<string, string> = {}): string {
  const types = { '/xl/workbook.xml': SPREADSHEET_MAIN, ...overrides };
  const entries = Object.entries(types)
    .map(([partName, contentType]) => `<Override PartName="${partName}" ContentType="${contentType}"/>`)
    .join('');
  return (
    `${XML}<Types xmlns="${NS_CONTENT_TYPES}">` +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    `<Default Extension="xml" ContentType="application/xml"/>${entries}</Types>`
  );
}

function relationships(...rels: readonly { id: string; type: string; target: string }[]): string {
  const entries = rels.map(rel => `<Relationship Id="${rel.id}" Type="${NS_R}/${rel.type}" Target="${rel.target}"/>`).join('');
  return `${XML}<Relationships xmlns="${NS_PACKAGE_RELS}">${entries}</Relationships>`;
}

const ROOT_RELS = relationships({ id: 'rId1', type: 'officeDocument', target: 'xl/workbook.xml' });

interface SheetEntry {
  name: string;
  sheetId?: number;
  relId?: string;
  state?: string;
}

function workbookXml(sheets: readonly SheetEntry[], date1904 = false): string {
  const entries = sheets
    .map((sheet, index) => {
      const relId = sheet.relId === undefined ? ` r:id="rId${index + 1}"` : sheet.relId === '' ? '' : ` r:id="${sheet.relId}"`;
      const state = sheet.state === undefined ? '' : ` state="${sheet.state}"`;
      return `<sheet name="${sheet.name}" sheetId="${sheet.sheetId ?? index + 1}"${state}${relId}/>`;
    })
    .join('');
  const properties = date1904 ? '<workbookPr date1904="1"/>' : '';
  return `${XML}<workbook xmlns="${NS}" xmlns:r="${NS_R}">${properties}<sheets>${entries}</sheets></workbook>`;
}

/** numFmtId 14 on cellXfs 1, so `s="1"` marks a date cell and nothing else does. */
const STYLES =
  `${XML}<styleSheet xmlns="${NS}"><cellXfs count="2">` +
  '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
  '<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
  '</cellXfs></styleSheet>';

function sharedStrings(...items: readonly string[]): string {
  const entries = items.map(text => `<si><t>${text}</t></si>`).join('');
  return `${XML}<sst xmlns="${NS}" count="${items.length}" uniqueCount="${items.length}">${entries}</sst>`;
}

function sheetXml(rows: string, dimension = ''): string {
  const dimensionElement = dimension === '' ? '' : `<dimension ref="${dimension}"/>`;
  return `${XML}<worksheet xmlns="${NS}">${dimensionElement}<sheetData>${rows}</sheetData></worksheet>`;
}

/** One row of inline strings, numbers and blanks: `null` writes no `<c>` at all. */
function row(rowNumber: number, values: readonly (string | number | boolean | null)[]): string {
  const cells = values
    .map((value, index) => {
      const ref = `${String.fromCharCode(65 + index)}${rowNumber}`;
      if (value === null) {
        return '';
      }
      if (typeof value === 'number') {
        return `<c r="${ref}"><v>${value}</v></c>`;
      }
      if (typeof value === 'boolean') {
        return `<c r="${ref}" t="b"><v>${value ? 1 : 0}</v></c>`;
      }
      return `<c r="${ref}" t="inlineStr"><is><t>${value}</t></is></c>`;
    })
    .join('');
  return `<row r="${rowNumber}">${cells}</row>`;
}

/** A complete, valid package; every part can be replaced or removed (with `undefined`) by an override. */
function packageParts(overrides: Record<string, string | undefined> = {}): Record<string, string> {
  const parts: Record<string, string | undefined> = {
    '[Content_Types].xml': contentTypes(),
    '_rels/.rels': ROOT_RELS,
    'xl/workbook.xml': workbookXml([{ name: 'Data' }]),
    'xl/_rels/workbook.xml.rels': relationships(
      { id: 'rId1', type: 'worksheet', target: 'worksheets/sheet1.xml' },
      { id: 'rId2', type: 'styles', target: 'styles.xml' },
      { id: 'rId3', type: 'sharedStrings', target: 'sharedStrings.xml' },
    ),
    'xl/styles.xml': STYLES,
    'xl/sharedStrings.xml': sharedStrings('Name', 'Amount'),
    'xl/worksheets/sheet1.xml': sheetXml(`${row(1, ['Name', 'Amount'])}${row(2, ['Zoë', 12.5])}`),
    ...overrides,
  };
  const present: Record<string, string> = {};
  for (const [name, data] of Object.entries(parts)) {
    if (data !== undefined) {
      present[name] = data;
    }
  }
  return present;
}

function workbookBytes(overrides: Record<string, string | undefined> = {}): Promise<Uint8Array> {
  return zip(packageParts(overrides));
}

/** A workbook whose only sheet is `rows`, with no shared strings and no styles unless asked for. */
function sheetOnlyBytes(rows: string, dimension = ''): Promise<Uint8Array> {
  return zip(
    packageParts({
      'xl/sharedStrings.xml': undefined,
      'xl/worksheets/sheet1.xml': sheetXml(rows, dimension),
    }),
  );
}

async function caught(action: () => Promise<unknown>): Promise<XlsxError> {
  try {
    await action();
  } catch (error) {
    if (isXlsxError(error)) {
      return error;
    }
    throw error;
  }
  throw new Error('expected the call to throw an XlsxError');
}

async function collect<T>(rows: AsyncIterable<T>): Promise<T[]> {
  const out: T[] = [];
  for await (const value of rows) {
    out.push(value);
  }
  return out;
}

async function readRows(bytes: Uint8Array, sheet: string | number = 0): Promise<CellValue[][]> {
  const workbook = await openWorkbook(bytes);
  try {
    return await collect(workbook.sheet(sheet).rows());
  } finally {
    await workbook.close();
  }
}

const encoder = new TextEncoder();

// ---------------------------------------------------------------------------------------------------------------
// What is not an xlsx
// ---------------------------------------------------------------------------------------------------------------

describe('openWorkbook: inputs that are not a modern workbook', () => {
  it('rejects an empty input', async () => {
    const error = await caught(() => openWorkbook(new Uint8Array(0)));
    expect(error.code).toBe('NOT_XLSX');
    expect(error.detail?.format).toBe('empty');
    expect(error.message).toBe('The file is empty.');
  });

  it('rejects an empty zip archive as empty', async () => {
    const empty = await zip({});
    const error = await caught(() => openWorkbook(empty));
    expect(error.code).toBe('NOT_XLSX');
    expect(error.detail?.format).toBe('empty');
  });

  it('rejects csv bytes and names the fix', async () => {
    const error = await caught(() => openWorkbook(encoder.encode('Id,Name\n001,Zoë\n')));
    expect(error.code).toBe('NOT_XLSX');
    expect(error.detail?.format).toBe('text');
    expect(error.message).toContain('csv');
    expect(error.message).toContain('.xlsx');
  });

  it('rejects an XML document (SpreadsheetML 2003) and an HTML export', async () => {
    const xml = await caught(() =>
      openWorkbook(encoder.encode('<?xml version="1.0"?><Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"/>')),
    );
    expect([xml.code, xml.detail?.format]).toEqual(['NOT_XLSX', 'xml']);
    const html = await caught(() => openWorkbook(encoder.encode('<html><body><table><tr><td>1</td></tr></table></body></html>')));
    expect([html.code, html.detail?.format]).toEqual(['NOT_XLSX', 'html']);
  });

  it('rejects bytes that are neither text nor a container we know', async () => {
    const noise = new Uint8Array(512);
    for (let i = 0; i < noise.length; i++) {
      noise[i] = (i * 7) % 256;
    }
    const error = await caught(() => openWorkbook(noise));
    expect(error.code).toBe('NOT_XLSX');
    expect(error.detail?.format).toBe('unknown');
  });

  it('EC-INPUT-ENCRYPTED: an encrypted workbook says password-protected, which the UI matches on', async () => {
    const error = await caught(() => openWorkbook(readFixture(fixtureById('hostile-encrypted-password-test'))));
    expect(error.code).toBe('ENCRYPTED');
    expect(error.message).toContain('password-protected');
  });

  it('EC-INPUT-XLS-RENAMED: a legacy .xls workbook says so instead of claiming to be encrypted', async () => {
    const error = await caught(() => openWorkbook(readFixture(fixtureById('hostile-biff8-xls-renamed'))));
    expect(error.code).toBe('LEGACY_XLS');
    expect(error.message).toContain('.xls');
    expect(error.message).not.toContain('password-protected');
  });

  it('EC-INPUT-ODS-RENAMED: an OpenDocument package is recognized by its mimetype entry', async () => {
    const error = await caught(() => openWorkbook(readFixture(fixtureById('hostile-ods-renamed'))));
    expect(error.code).toBe('ODS');
    expect(error.detail?.mimetype).toBe('application/vnd.oasis.opendocument.spreadsheet');
  });

  it('EC-INPUT-XLSB-RENAMED: a binary workbook is recognized by its main content type', async () => {
    const error = await caught(() => openWorkbook(readFixture(fixtureById('hostile-xlsb-renamed'))));
    expect(error.code).toBe('XLSB');
    expect(error.message).toContain('.xlsb');
  });

  it('rejects the other OOXML families by their main part', async () => {
    const word = await zip({
      '[Content_Types].xml': contentTypes({
        '/word/document.xml': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
      }),
      '_rels/.rels': relationships({ id: 'rId1', type: 'officeDocument', target: 'word/document.xml' }),
      'word/document.xml': `${XML}<document/>`,
    });
    expect((await caught(() => openWorkbook(word))).detail?.format).toBe('docx');

    const slides = await zip({
      '[Content_Types].xml': contentTypes({
        '/ppt/presentation.xml': 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',
      }),
      '_rels/.rels': relationships({ id: 'rId1', type: 'officeDocument', target: 'ppt/presentation.xml' }),
      'ppt/presentation.xml': `${XML}<presentation/>`,
    });
    expect((await caught(() => openWorkbook(slides))).detail?.format).toBe('pptx');
  });

  it('rejects a zip with no workbook part', async () => {
    const bytes = await zip({ 'readme.txt': 'not a workbook' });
    const error = await caught(() => openWorkbook(bytes));
    expect(error.code).toBe('NOT_XLSX');
    expect(error.detail?.format).toBe('zip');
    expect(error.message).toContain('no workbook part');
  });

  it('lets zip errors through unchanged', async () => {
    const truncated = await caught(() => openWorkbook(readFixture(fixtureById('hostile-truncated-central-directory'))));
    expect(truncated.code).toBe('ZIP_TRUNCATED');
    const duplicate = await caught(() => openWorkbook(readFixture(fixtureById('hostile-duplicate-sheet-entries'))));
    expect(duplicate.code).toBe('ZIP_DUPLICATE_ENTRY');
  });

  it('applies the entry limit from options', async () => {
    const bytes = await workbookBytes();
    const error = await caught(() => openWorkbook(bytes, { limits: { maxEntries: 3 } }));
    expect(error.code).toBe('LIMIT_EXCEEDED');
    expect(error.detail?.maxEntries).toBe(3);
  });

  it('reports a damaged part as XML_MALFORMED rather than a bare error', async () => {
    const bytes = await workbookBytes({ 'xl/workbook.xml': `${XML}<workbook><sheets>` });
    const error = await caught(() => openWorkbook(bytes));
    expect(error.code).toBe('XML_MALFORMED');
  });

  it('refuses a DOCTYPE in any part it reads', async () => {
    const bytes = await workbookBytes({ 'xl/workbook.xml': `${XML}<!DOCTYPE workbook [<!ENTITY x "y">]><workbook/>` });
    const error = await caught(() => openWorkbook(bytes));
    expect(error.code).toBe('XML_DOCTYPE');
  });
});

// ---------------------------------------------------------------------------------------------------------------
// Locating the parts
// ---------------------------------------------------------------------------------------------------------------

describe('openWorkbook: package layout', () => {
  it('accepts the macro-enabled and template main content types', async () => {
    for (const contentType of [
      'application/vnd.ms-excel.sheet.macroEnabled.main+xml',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml',
    ]) {
      const bytes = await workbookBytes({ '[Content_Types].xml': contentTypes({ '/xl/workbook.xml': contentType }) });
      expect(await readRows(bytes)).toEqual([
        ['Name', 'Amount'],
        ['Zoë', 12.5],
      ]);
    }
  });

  it('EC-PART-NONSTANDARD-NAMES: follows the officeDocument relationship wherever it points', async () => {
    const bytes = await zip({
      '[Content_Types].xml': contentTypes({ '/spreadsheet/book.xml': SPREADSHEET_MAIN }),
      '_rels/.rels': relationships({ id: 'rId9', type: 'officeDocument', target: '/spreadsheet/book.xml' }),
      'spreadsheet/book.xml': workbookXml([{ name: 'Data' }]),
      'spreadsheet/_rels/book.xml.rels': relationships({ id: 'rId1', type: 'worksheet', target: 'sheets/one.xml' }),
      'spreadsheet/sheets/one.xml': sheetXml(row(1, ['moved'])),
    });
    expect(await readRows(bytes)).toEqual([['moved']]);
  });

  it('falls back to xl/workbook.xml when the package has no root relationships', async () => {
    const bytes = await workbookBytes({ '_rels/.rels': undefined });
    expect(await readRows(bytes)).toEqual([
      ['Name', 'Amount'],
      ['Zoë', 12.5],
    ]);
  });

  it('falls back to the conventional sheet part when the sheet relationship is missing', async () => {
    const bytes = await workbookBytes({
      'xl/workbook.xml': workbookXml([{ name: 'Data', sheetId: 1, relId: '' }]),
      'xl/_rels/workbook.xml.rels': relationships({ id: 'rId2', type: 'styles', target: 'styles.xml' }),
      'xl/sharedStrings.xml': undefined,
    });
    expect(await readRows(bytes)).toEqual([
      ['Name', 'Amount'],
      ['Zoë', 12.5],
    ]);
  });

  it('reads sheet order, kind and hidden state from the workbook', async () => {
    const bytes = await workbookBytes({
      'xl/workbook.xml': workbookXml([
        { name: 'Data' },
        { name: 'Chart', relId: 'rId4' },
        { name: 'Hidden', relId: 'rId5', state: 'hidden' },
        { name: 'Secret', relId: 'rId6', state: 'veryHidden' },
      ]),
      'xl/_rels/workbook.xml.rels': relationships(
        { id: 'rId1', type: 'worksheet', target: 'worksheets/sheet1.xml' },
        { id: 'rId4', type: 'chartsheet', target: 'chartsheets/sheet1.xml' },
        { id: 'rId5', type: 'worksheet', target: 'worksheets/sheet2.xml' },
        { id: 'rId6', type: 'worksheet', target: 'worksheets/sheet3.xml' },
      ),
      'xl/chartsheets/sheet1.xml': `${XML}<chartsheet xmlns="${NS}"/>`,
      'xl/worksheets/sheet2.xml': sheetXml(row(1, ['hidden'])),
      'xl/worksheets/sheet3.xml': sheetXml(row(1, ['very hidden'])),
    });
    const workbook = await openWorkbook(bytes);
    try {
      expect(workbook.sheets).toEqual([
        { name: 'Data', index: 0, kind: 'worksheet', hidden: false },
        { name: 'Chart', index: 1, kind: 'chartsheet', hidden: false },
        { name: 'Hidden', index: 2, kind: 'worksheet', hidden: true },
        { name: 'Secret', index: 3, kind: 'worksheet', hidden: true },
      ]);
      expect(workbook.date1904).toBe(false);
    } finally {
      await workbook.close();
    }
  });

  it('EC-DATE-1904: the workbook date system reaches the row reader', async () => {
    const rows = `${row(1, ['When'])}<row r="2"><c r="A2" s="1"><v>45351</v></c></row>`;
    const parts = { 'xl/sharedStrings.xml': undefined, 'xl/worksheets/sheet1.xml': sheetXml(rows) };
    const workbook1900 = await openWorkbook(await zip(packageParts(parts)), { dates: 'utc' });
    const workbook1904 = await openWorkbook(
      await zip(packageParts({ ...parts, 'xl/workbook.xml': workbookXml([{ name: 'Data' }], true) })),
      { dates: 'utc' },
    );
    try {
      expect(workbook1904.date1904).toBe(true);
      const [, under1900 = []] = await collect(workbook1900.sheet(0).rows());
      const [, under1904 = []] = await collect(workbook1904.sheet(0).rows());
      expect((under1900[0] as Date).toISOString()).toBe('2024-02-29T00:00:00.000Z');
      expect((under1904[0] as Date).toISOString()).toBe('2028-03-01T00:00:00.000Z');
    } finally {
      await workbook1900.close();
      await workbook1904.close();
    }
  });

  it('reads dates only when a styles part says so', async () => {
    const rows = `<row r="1"><c r="A1" s="1"><v>45351</v></c></row>`;
    const styled = await zip(packageParts({ 'xl/sharedStrings.xml': undefined, 'xl/worksheets/sheet1.xml': sheetXml(rows) }));
    expect((await readRows(styled))[0]?.[0]).toBeInstanceOf(Date);

    const unstyled = await zip(
      packageParts({
        'xl/styles.xml': undefined,
        'xl/sharedStrings.xml': undefined,
        'xl/worksheets/sheet1.xml': sheetXml(rows),
      }),
    );
    expect(await readRows(unstyled)).toEqual([[45351]]);
  });

  it('EC-SST-ABSENT-INLINE-ONLY: a missing shared-string part leaves t="s" cells blank', async () => {
    const rows = `<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>1</v></c></row>`;
    const bytes = await zip(packageParts({ 'xl/sharedStrings.xml': undefined, 'xl/worksheets/sheet1.xml': sheetXml(rows) }));
    expect(await readRows(bytes)).toEqual([['', 1]]);
  });

  it('parses the shared-string part only when a cell asks for it', async () => {
    // A part that would throw the moment it is parsed proves the table is never touched for an inline-only sheet.
    const poisoned = `${XML}<!DOCTYPE sst><sst xmlns="${NS}"/>`;
    const inlineOnly = await workbookBytes({ 'xl/sharedStrings.xml': poisoned });
    expect(await readRows(inlineOnly)).toEqual([
      ['Name', 'Amount'],
      ['Zoë', 12.5],
    ]);

    const shared = await workbookBytes({
      'xl/sharedStrings.xml': poisoned,
      'xl/worksheets/sheet1.xml': sheetXml(`<row r="1"><c r="A1" t="s"><v>0</v></c></row>`),
    });
    expect((await caught(() => readRows(shared))).code).toBe('XML_DOCTYPE');
  });

  it('parses the shared-string part once for the whole workbook', async () => {
    const bytes = await workbookBytes({
      'xl/workbook.xml': workbookXml([{ name: 'One' }, { name: 'Two', relId: 'rId4' }]),
      'xl/_rels/workbook.xml.rels': relationships(
        { id: 'rId1', type: 'worksheet', target: 'worksheets/sheet1.xml' },
        { id: 'rId4', type: 'worksheet', target: 'worksheets/sheet2.xml' },
        { id: 'rId2', type: 'styles', target: 'styles.xml' },
        { id: 'rId3', type: 'sharedStrings', target: 'sharedStrings.xml' },
      ),
      'xl/worksheets/sheet1.xml': sheetXml(`<row r="1"><c r="A1" t="s"><v>0</v></c></row>`),
      'xl/worksheets/sheet2.xml': sheetXml(`<row r="1"><c r="A1" t="s"><v>1</v></c></row>`),
    });
    const container = await ZipReader.open(sourceFrom(bytes), { maxEntries: 100, maxInflatedBytes: 1024 * 1024 });
    const sharedStringsHeader = container.entries.get('xl/sharedStrings.xml')?.localHeaderOffset;
    expect(sharedStringsHeader).toBeDefined();

    const reads: number[] = [];
    const source: RandomAccessSource = {
      size: bytes.byteLength,
      read(offset: number, length: number): Promise<Uint8Array> {
        reads.push(offset);
        return Promise.resolve(bytes.subarray(offset, Math.min(offset + length, bytes.byteLength)));
      },
    };
    const workbook = await openWorkbook(source);
    try {
      expect(await collect(workbook.sheet(0).rows())).toEqual([['Name']]);
      expect(await collect(workbook.sheet(1).rows())).toEqual([['Amount']]);
    } finally {
      await workbook.close();
    }
    // Streaming an entry starts by reading its local header, so a second parse would show up as a second read.
    expect(reads.filter(offset => offset === sharedStringsHeader)).toHaveLength(1);
  });

  it('closes the source', async () => {
    const bytes = await workbookBytes();
    let closed = 0;
    const source: RandomAccessSource = {
      size: bytes.byteLength,
      read: (offset: number, length: number) => Promise.resolve(bytes.subarray(offset, offset + length)),
      close: () => {
        closed++;
        return Promise.resolve();
      },
    };
    const workbook = await openWorkbook(source);
    await workbook.close();
    expect(closed).toBe(1);
  });
});

// ---------------------------------------------------------------------------------------------------------------
// Sheet lookup
// ---------------------------------------------------------------------------------------------------------------

describe('workbook.sheet', () => {
  const bytes = (): Promise<Uint8Array> =>
    workbookBytes({
      'xl/workbook.xml': workbookXml([{ name: 'Data' }, { name: 'Summary', relId: 'rId4' }]),
      'xl/_rels/workbook.xml.rels': relationships(
        { id: 'rId1', type: 'worksheet', target: 'worksheets/sheet1.xml' },
        { id: 'rId4', type: 'worksheet', target: 'worksheets/sheet2.xml' },
        { id: 'rId2', type: 'styles', target: 'styles.xml' },
      ),
      'xl/sharedStrings.xml': undefined,
      'xl/worksheets/sheet2.xml': sheetXml(row(1, ['second'])),
    });

  it('finds a sheet by exact name, by name ignoring case, and by index', async () => {
    const workbook = await openWorkbook(await bytes());
    try {
      expect(workbook.sheet('Summary').info.index).toBe(1);
      expect(workbook.sheet('summary').info.index).toBe(1);
      expect(workbook.sheet(1).info.name).toBe('Summary');
    } finally {
      await workbook.close();
    }
  });

  it('lists the sheets it does have when the name or index is unknown', async () => {
    const workbook = await openWorkbook(await bytes());
    try {
      const byName = await caught(() => Promise.resolve(workbook.sheet('Nope')));
      expect(byName.code).toBe('SHEET_NOT_FOUND');
      expect(byName.message).toContain('"Data"');
      expect(byName.message).toContain('"Summary"');
      expect((await caught(() => Promise.resolve(workbook.sheet(7)))).code).toBe('SHEET_NOT_FOUND');
    } finally {
      await workbook.close();
    }
  });

  it('reports a sheet whose part is missing instead of returning nothing', async () => {
    const workbook = await openWorkbook(await workbookBytes({ 'xl/worksheets/sheet1.xml': undefined }));
    try {
      const error = await caught(() => collect(workbook.sheet(0).rows()));
      expect(error.code).toBe('NOT_XLSX');
      expect(error.detail?.sheet).toBe('Data');
    } finally {
      await workbook.close();
    }
  });
});

// ---------------------------------------------------------------------------------------------------------------
// Array mode
// ---------------------------------------------------------------------------------------------------------------

describe('sheet.rows: array mode', () => {
  const SHEET = `${row(1, ['Id', 'Name'])}${row(2, ['001', 'Zoë'])}<row r="3"/>${row(4, ['003', 'Ünïcödé'])}`;

  it('yields dense rows and skips blank ones', async () => {
    expect(await readRows(await sheetOnlyBytes(SHEET))).toEqual([
      ['Id', 'Name'],
      ['001', 'Zoë'],
      ['003', 'Ünïcödé'],
    ]);
  });

  it('passes startRow, maxRows and blankRows through to the row reader', async () => {
    const workbook = await openWorkbook(await sheetOnlyBytes(SHEET));
    try {
      const sheet = workbook.sheet(0);
      expect(await collect(sheet.rows({ startRow: 2 }))).toEqual([
        ['001', 'Zoë'],
        ['003', 'Ünïcödé'],
      ]);
      expect(await collect(sheet.rows({ maxRows: 1 }))).toEqual([['Id', 'Name']]);
      expect(await collect(sheet.rows({ startRow: 2, blankRows: true }))).toEqual([['001', 'Zoë'], [], ['003', 'Ünïcödé']]);
    } finally {
      await workbook.close();
    }
  });

  it('caps the columns it materializes', async () => {
    const wide = await sheetOnlyBytes(row(1, ['a', 'b', 'c']));
    const workbook = await openWorkbook(wide);
    try {
      expect((await caught(() => collect(workbook.sheet(0).rows({ maxColumns: 2 })))).code).toBe('LIMIT_EXCEEDED');
    } finally {
      await workbook.close();
    }
  });

  it('reports formula text when asked and the cached value otherwise', async () => {
    const bytes = await sheetOnlyBytes(`<row r="1"><c r="A1"><f>B2*2</f><v>25</v></c></row>`);
    const workbook = await openWorkbook(bytes);
    try {
      expect(await collect(workbook.sheet(0).rows())).toEqual([[25]]);
      expect(await collect(workbook.sheet(0).rows({ formulas: 'text' }))).toEqual([['=B2*2']]);
    } finally {
      await workbook.close();
    }
  });

  it('materializes error cells the way `errors` asks', async () => {
    const bytes = await sheetOnlyBytes(`<row r="1"><c r="A1" t="e"><v>#DIV/0!</v></c><c r="B1"><v>1</v></c></row>`);
    const asText = await openWorkbook(bytes);
    const asObject = await openWorkbook(bytes, { errors: 'object' });
    const asNull = await openWorkbook(bytes, { errors: 'null' });
    try {
      expect(await collect(asText.sheet(0).rows())).toEqual([['#DIV/0!', 1]]);
      const objectRows: (CellValue | CellError)[][] = await collect(asObject.sheet(0).rows());
      expect(objectRows).toEqual([[{ error: '#DIV/0!' }, 1]]);
      expect(await collect(asNull.sheet(0).rows())).toEqual([[null, 1]]);
    } finally {
      await asText.close();
      await asObject.close();
      await asNull.close();
    }
  });

  it('honours the dates option', async () => {
    const bytes = await sheetOnlyBytes(`<row r="1"><c r="A1" s="1"><v>45351</v></c></row>`);
    const local = await openWorkbook(bytes);
    const utc = await openWorkbook(bytes, { dates: 'utc' });
    const serial = await openWorkbook(bytes, { dates: 'serial' });
    try {
      const [[localValue] = []] = await collect(local.sheet(0).rows());
      const [[utcValue] = []] = await collect(utc.sheet(0).rows());
      expect((localValue as Date).getFullYear()).toBe(2024);
      expect((utcValue as Date).toISOString()).toBe('2024-02-29T00:00:00.000Z');
      expect(await collect(serial.sheet(0).rows())).toEqual([[45351]]);
    } finally {
      await local.close();
      await utc.close();
      await serial.close();
    }
  });

  it('stops on an aborted signal, at open and while streaming', async () => {
    const bytes = await sheetOnlyBytes(SHEET);
    const before = new AbortController();
    before.abort();
    expect((await caught(() => openWorkbook(bytes, { signal: before.signal }))).code).toBe('ABORTED');

    const during = new AbortController();
    const workbook = await openWorkbook(bytes, { signal: during.signal });
    try {
      during.abort();
      expect((await caught(() => collect(workbook.sheet(0).rows()))).code).toBe('ABORTED');
    } finally {
      await workbook.close();
    }
  });
});

// ---------------------------------------------------------------------------------------------------------------
// Object mode
// ---------------------------------------------------------------------------------------------------------------

async function objects(rows: string, options: RowsOptions & ObjectRowsOptions = {}, dimension = ''): Promise<ObjectsResult<CellValue>> {
  const workbook = await openWorkbook(await sheetOnlyBytes(rows, dimension));
  try {
    return await workbook.sheet(0).toObjects(options);
  } finally {
    await workbook.close();
  }
}

describe('sheet.toObjects: the sheet_to_json contract', () => {
  it('keys each row by its header and fills absent cells with defval', async () => {
    const result = await objects(`${row(1, ['Id', 'Name'])}${row(2, ['001', 'Zoë'])}${row(3, ['002', null])}`);
    expect(result.headers).toEqual(['Id', 'Name']);
    expect(result.rows).toEqual([
      { Id: '001', Name: 'Zoë' },
      { Id: '002', Name: '' },
    ]);
    expect(result.truncated).toBe(false);
  });

  it('takes defval from the options', async () => {
    const result = await objects(`${row(1, ['Id', 'Name'])}${row(2, ['001', null])}`, { defval: null });
    expect(result.rows).toEqual([{ Id: '001', Name: null }]);
  });

  it('names a blank header __EMPTY, then __EMPTY_1, and de-duplicates repeats with _1', async () => {
    const result = await objects(`${row(1, ['Id', null, 'Name', null, 'Name'])}${row(2, ['001', 'x', 'a', 'y', 'b'])}`);
    expect(result.headers).toEqual(['Id', '__EMPTY', 'Name', '__EMPTY_1', 'Name_1']);
    expect(result.rows).toEqual([{ Id: '001', __EMPTY: 'x', Name: 'a', __EMPTY_1: 'y', Name_1: 'b' }]);
  });

  it('skips a suffix that is already taken, as SheetJS does', async () => {
    const result = await objects(`${row(1, ['Name', 'Name_1', 'Name'])}${row(2, ['a', 'b', 'c'])}`);
    expect(result.headers).toEqual(['Name', 'Name_1', 'Name_2']);
  });

  it('keeps an empty-string header as an empty name, unlike an absent cell', async () => {
    const result = await objects(`${row(1, ['Id', ''])}${row(2, ['001', 'x'])}`);
    expect(result.headers).toEqual(['Id', '']);
  });

  it('numbers the headers by column with headerNaming: index', async () => {
    const result = await objects(`${row(1, ['Id', 'Name'])}${row(2, ['001', 'Zoë'])}`, { headerNaming: 'index' });
    expect(result.headers).toEqual(['0', '1']);
    expect(result.rows).toEqual([{ '0': '001', '1': 'Zoë' }]);
  });

  it('drops blank-header columns with dropEmptyHeaders', async () => {
    const result = await objects(`${row(1, ['Id', null, 'Name'])}${row(2, ['001', 'x', 'Zoë'])}`, { dropEmptyHeaders: true });
    expect(result.headers).toEqual(['Id', 'Name']);
    expect(result.rows).toEqual([{ Id: '001', Name: 'Zoë' }]);
  });

  it('takes the header from headerRow and the data from startRow', async () => {
    const rows = `${row(1, ['Title'])}${row(2, ['Id', 'Name'])}${row(3, ['001', 'Zoë'])}${row(4, ['002', 'Bob'])}`;
    expect((await objects(rows, { headerRow: 2 })).headers).toEqual(['Id', 'Name']);
    expect((await objects(rows, { headerRow: 2 })).rows).toEqual([
      { Id: '001', Name: 'Zoë' },
      { Id: '002', Name: 'Bob' },
    ]);
    expect((await objects(rows, { headerRow: 2, startRow: 4 })).rows).toEqual([{ Id: '002', Name: 'Bob' }]);
    // Without headerRow the header is the start row itself.
    expect((await objects(rows, { startRow: 2 })).headers).toEqual(['Id', 'Name']);
  });

  it('skips blank rows and reports truncation', async () => {
    const rows = `${row(1, ['Id'])}${row(2, ['001'])}<row r="3"/>${row(4, ['002'])}${row(5, ['003'])}`;
    expect((await objects(rows)).rows).toEqual([{ Id: '001' }, { Id: '002' }, { Id: '003' }]);
    const limited = await objects(rows, { maxRows: 2 });
    expect(limited.rows).toEqual([{ Id: '001' }, { Id: '002' }]);
    expect(limited.truncated).toBe(true);
    expect((await objects(rows, { maxRows: 3 })).truncated).toBe(false);
  });

  it('gives a column the sheet declares but never fills a blank header', async () => {
    const result = await objects(`${row(1, ['Id', 'Name'])}${row(2, ['001', 'Zoë'])}`, {}, 'A1:C2');
    expect(result.headers).toEqual(['Id', 'Name', '__EMPTY']);
    expect(result.rows).toEqual([{ Id: '001', Name: 'Zoë', __EMPTY: '' }]);
  });

  it('keeps values from a row wider than both the header and the dimension', async () => {
    const result = await objects(`${row(1, ['Id'])}${row(2, ['001', 'extra'])}`, {}, 'A1:A2');
    expect(result.headers).toEqual(['Id', '__EMPTY']);
    expect(result.rows).toEqual([{ Id: '001', __EMPTY: 'extra' }]);
  });

  it('streams the same records through rows({ mode: "object" })', async () => {
    const workbook = await openWorkbook(await sheetOnlyBytes(`${row(1, ['Id'])}${row(2, ['001'])}${row(3, ['002'])}`));
    try {
      expect(await collect(workbook.sheet(0).rows({ mode: 'object' }))).toEqual([{ Id: '001' }, { Id: '002' }]);
      expect(await collect(workbook.sheet(0).rows({ mode: 'object', maxRows: 1 }))).toEqual([{ Id: '001' }]);
    } finally {
      await workbook.close();
    }
  });
});

async function objectsWith(
  rows: string,
  openOptions: OpenOptions,
  options: RowsOptions & ObjectRowsOptions = {},
  dimension = '',
): Promise<ObjectsResult<ReadValue>> {
  const workbook = await openWorkbook(await sheetOnlyBytes(rows, dimension), openOptions);
  try {
    return await workbook.sheet(0).toObjects(options);
  } finally {
    await workbook.close();
  }
}

async function streamedObjects(
  rows: string,
  options: RowsOptions & ObjectRowsOptions = {},
  dimension = '',
): Promise<Record<string, CellValue>[]> {
  const workbook = await openWorkbook(await sheetOnlyBytes(rows, dimension));
  try {
    return await collect(workbook.sheet(0).rows({ ...options, mode: 'object' }));
  } finally {
    await workbook.close();
  }
}

describe('object mode: which row holds the headers', () => {
  const HEADERS_ON_ROW_2 = `${row(2, ['Id', 'Name'])}${row(3, ['001', 'Zoë'])}${row(4, ['002', 'Bob'])}`;

  it('EC-HEADER-ROW-FIRST-NONEMPTY: a blank row 1 puts the headers on the first row with a value', async () => {
    // Google Sheets writes no dimension; Excel writes one that starts at the header row. Both read the same.
    for (const dimension of ['', 'A2:B4']) {
      const result = await objects(HEADERS_ON_ROW_2, {}, dimension);
      expect(result.headers, dimension).toEqual(['Id', 'Name']);
      expect(result.rows, dimension).toEqual([
        { Id: '001', Name: 'Zoë' },
        { Id: '002', Name: 'Bob' },
      ]);
    }
  });

  it('EC-HEADER-ROW-FIRST-NONEMPTY: blank rows 1 and 2, and rows of blank cells, are skipped too', async () => {
    const rows = `<row r="1"><c r="A1" s="1"/><c r="B1"><v/></c></row><row r="2"/>${row(3, ['Id', 'Name'])}${row(4, ['001', 'Zoë'])}`;
    const result = await objects(rows, {}, 'A3:B4');
    expect(result.headers).toEqual(['Id', 'Name']);
    expect(result.rows).toEqual([{ Id: '001', Name: 'Zoë' }]);
    expect(await streamedObjects(rows)).toEqual([{ Id: '001', Name: 'Zoë' }]);
  });

  it('EC-HEADER-ROW-FIRST-NONEMPTY: an empty-string cell makes its row the header row, as in SheetJS', async () => {
    const rows = `<row r="1"><c r="A1" t="str"><v></v></c></row>${row(2, ['Id'])}${row(3, ['001'])}`;
    const result = await objects(rows);
    expect(result.headers).toEqual(['']);
    expect(result.rows).toEqual([{ '': 'Id' }, { '': '001' }]);
  });

  it('EC-HEADER-ROW-FIRST-NONEMPTY: startRow without headerRow looks for the header from startRow on', async () => {
    const rows = `${row(1, ['Title'])}<row r="3"/>${row(4, ['Id', 'Name'])}${row(5, ['001', 'Zoë'])}`;
    const result = await objects(rows, { startRow: 2 });
    expect(result.headers).toEqual(['Id', 'Name']);
    expect(result.rows).toEqual([{ Id: '001', Name: 'Zoë' }]);
    // A startRow that lands on a row with values uses that row, as before (the wider row 4 adds a blank header).
    expect((await objects(rows, { startRow: 1 })).headers).toEqual(['Title', '__EMPTY']);
  });

  it('EC-HEADER-ROW-FIRST-NONEMPTY: an explicit headerRow is that exact row, blank or not', async () => {
    const result = await objects(HEADERS_ON_ROW_2, { headerRow: 1 });
    expect(result.headers).toEqual(['__EMPTY', '__EMPTY_1']);
    expect(result.rows).toEqual([
      { __EMPTY: 'Id', __EMPTY_1: 'Name' },
      { __EMPTY: '001', __EMPTY_1: 'Zoë' },
      { __EMPTY: '002', __EMPTY_1: 'Bob' },
    ]);
    expect((await objects(HEADERS_ON_ROW_2, { headerRow: 2 })).headers).toEqual(['Id', 'Name']);
  });

  it("EC-HEADER-ROW-FIRST-NONEMPTY: Jetstream's options (errors: 'null', dropEmptyHeaders) find the headers below a blank row", async () => {
    const rows =
      `${row(3, ['Id', null, 'Name'])}` +
      `<row r="4"><c r="A4" t="inlineStr"><is><t>001</t></is></c><c r="B4"><v>7</v></c><c r="C4" t="e"><v>#N/A</v></c></row>` +
      `${row(5, ['002', null, 'Bob'])}`;
    const result = await objectsWith(rows, { errors: 'null' }, { dropEmptyHeaders: true }, 'A3:C5');
    expect(result.headers).toEqual(['Id', 'Name']);
    expect(result.rows).toEqual([
      { Id: '001', Name: '' },
      { Id: '002', Name: 'Bob' },
    ]);
  });
});

describe('object mode: which columns get names', () => {
  it('EC-HEADER-DIMENSION-START-COLUMN: a dimension that starts at B names no column for the empty column A', async () => {
    const rows =
      `<row r="2">${'<c r="B2" t="inlineStr"><is><t>Id</t></is></c><c r="C2" t="inlineStr"><is><t>Name</t></is></c>'}</row>` +
      `<row r="3"><c r="B3" t="inlineStr"><is><t>001</t></is></c><c r="C3" t="inlineStr"><is><t>Zoë</t></is></c></row>`;
    const result = await objects(rows, {}, 'B2:D3');
    // The dimension runs to D, so D is a trailing blank header, exactly as SheetJS names it.
    expect(result.headers).toEqual(['Id', 'Name', '__EMPTY']);
    expect(result.rows).toEqual([{ Id: '001', Name: 'Zoë', __EMPTY: '' }]);
    expect((await objects(rows, { dropEmptyHeaders: true }, 'B2:D3')).headers).toEqual(['Id', 'Name']);
    expect((await objects(rows, { headerNaming: 'index' }, 'B2:D3')).headers).toEqual(['1', '2', '3']);
    // Without a dimension the columns start at A.
    expect((await objects(rows)).headers).toEqual(['__EMPTY', 'Id', 'Name']);
  });

  it('EC-HEADER-DIMENSION-START-COLUMN: a value left of a wrong dimension is kept under a blank header', async () => {
    const rows =
      `<row r="1"><c r="B1" t="inlineStr"><is><t>Id</t></is></c></row>` +
      `<row r="2"><c r="B2" t="inlineStr"><is><t>001</t></is></c></row>` +
      `<row r="3"><c r="A3" t="inlineStr"><is><t>stray</t></is></c><c r="B3" t="inlineStr"><is><t>002</t></is></c></row>`;
    const result = await objects(rows, {}, 'B1:B3');
    // Sheet order: the stray column A comes before Id in `headers` and in every record's keys.
    expect(result.headers).toEqual(['__EMPTY', 'Id']);
    expect(result.rows).toEqual([
      { __EMPTY: '', Id: '001' },
      { __EMPTY: 'stray', Id: '002' },
    ]);
    expect(result.rows.map(record => Object.keys(record))).toEqual([
      ['__EMPTY', 'Id'],
      ['__EMPTY', 'Id'],
    ]);
    // Streaming cannot reach back to the record it already yielded.
    expect(await streamedObjects(rows, {}, 'B1:B3')).toEqual([{ Id: '001' }, { Id: '002', __EMPTY: 'stray' }]);
  });

  it('EC-HEADER-DIMENSION-START-COLUMN: a header left of a wrong dimension still names its column', async () => {
    const rows = `${row(1, ['Id', 'Name'])}${row(2, ['001', 'Zoë'])}`;
    const result = await objects(rows, {}, 'B1:B2');
    expect(result.headers).toEqual(['Id', 'Name']);
    expect(result.rows).toEqual([{ Id: '001', Name: 'Zoë' }]);
  });

  it('EC-HEADER-DIMENSION-START-COLUMN: toObjects gives every record the columns a later, wider row named', async () => {
    const rows = `${row(1, ['Id'])}${row(2, ['001'])}${row(3, ['002', 'wide'])}${row(4, ['003', null, 'wider'])}`;
    const result = await objects(rows);
    expect(result.headers).toEqual(['Id', '__EMPTY', '__EMPTY_1']);
    expect(result.rows).toEqual([
      { Id: '001', __EMPTY: '', __EMPTY_1: '' },
      { Id: '002', __EMPTY: 'wide', __EMPTY_1: '' },
      { Id: '003', __EMPTY: '', __EMPTY_1: 'wider' },
    ]);
    for (const record of result.rows) {
      expect(Object.keys(record)).toEqual(result.headers);
    }
    expect((await objects(rows, { defval: null })).rows[0]).toEqual({ Id: '001', __EMPTY: null, __EMPTY_1: null });
    // The stream yields each record as it was when built.
    expect(await streamedObjects(rows)).toEqual([
      { Id: '001' },
      { Id: '002', __EMPTY: 'wide' },
      { Id: '003', __EMPTY: '', __EMPTY_1: 'wider' },
    ]);
  });
});

describe('object mode: header text', () => {
  it('EC-HEADER-PROTO-KEY: __proto__, constructor, toString and hasOwnProperty are ordinary header names', async () => {
    const rows =
      `${row(1, ['__proto__', 'constructor', 'toString', 'hasOwnProperty', '__proto__', 'constructor'])}` +
      `${row(2, ['a', 'b', 'c', 'd', 'e', 'f'])}` +
      `<row r="3"><c r="A3" s="1"><v>45351</v></c><c r="B3"><v>2</v></c></row>`;
    const result = await objects(rows);
    const names = ['__proto__', 'constructor', 'toString', 'hasOwnProperty', '__proto___1', 'constructor_1'];
    expect(result.headers).toEqual(names);
    const [first, second] = result.rows;
    for (const record of result.rows) {
      expect(Object.getPrototypeOf(record)).toBe(Object.prototype);
      expect(Object.keys(record)).toEqual(names);
    }
    expect(Object.getOwnPropertyDescriptor(first, '__proto__')).toEqual({
      value: 'a',
      writable: true,
      enumerable: true,
      configurable: true,
    });
    expect(first?.constructor).toBe('b');
    expect(first?.toString).toBe('c');
    expect(first?.hasOwnProperty).toBe('d');
    expect(first?.__proto___1).toBe('e');
    expect(JSON.stringify(first)).toBe(
      '{"__proto__":"a","constructor":"b","toString":"c","hasOwnProperty":"d","__proto___1":"e","constructor_1":"f"}',
    );
    // A Date under __proto__ is a value, not a new prototype for the record.
    expect(Object.getOwnPropertyDescriptor(second, '__proto__')?.value).toBeInstanceOf(Date);
    expect(Object.getPrototypeOf(second)).toBe(Object.prototype);
  });

  it('EC-HEADER-PROTO-KEY: a __proto__ column filled with defval stays an own property', async () => {
    const result = await objects(`${row(1, ['__proto__'])}${row(2, ['001'])}${row(3, ['002', 'x'])}`, { defval: null });
    expect(result.headers).toEqual(['__proto__', '__EMPTY']);
    expect(Object.getOwnPropertyDescriptor(result.rows[0], '__proto__')?.value).toBe('001');
    expect(Object.getPrototypeOf(result.rows[0])).toBe(Object.prototype);
    const blankProto = await objects(`${row(1, [null, '__proto__'])}${row(2, [null, null])}${row(3, ['v', null])}`);
    expect(blankProto.headers).toEqual(['__EMPTY', '__proto__']);
    expect(Object.getOwnPropertyDescriptor(blankProto.rows[0], '__proto__')?.value).toBe('');
  });

  it('EC-HEADER-ERROR-CELL: an error header names its column with the error text, whatever errors says', async () => {
    const rows =
      `<row r="1"><c r="A1" t="e"><v>#N/A</v></c><c r="B1" t="inlineStr"><is><t>Name</t></is></c></row>` +
      `<row r="2"><c r="A2" t="e"><v>#DIV/0!</v></c><c r="B2" t="inlineStr"><is><t>Zoë</t></is></c></row>`;
    const asString = await objectsWith(rows, {});
    expect(asString.headers).toEqual(['#N/A', 'Name']);
    expect(asString.rows).toEqual([{ '#N/A': '#DIV/0!', Name: 'Zoë' }]);
    const asNull = await objectsWith(rows, { errors: 'null' }, { dropEmptyHeaders: true });
    expect(asNull.headers).toEqual(['#N/A', 'Name']);
    expect(asNull.rows).toEqual([{ '#N/A': '', Name: 'Zoë' }]);
    const asObject = await objectsWith(rows, { errors: 'object' });
    expect(asObject.headers).toEqual(['#N/A', 'Name']);
    expect(asObject.rows).toEqual([{ '#N/A': { error: '#DIV/0!' }, Name: 'Zoë' }]);
    // An explicit headerRow gets the same treatment, and a row of nothing but errors can be the header row.
    expect((await objectsWith(`<row r="1"/>${rows.replaceAll('1"', '5"')}`, { errors: 'null' }, { headerRow: 2 })).headers).toEqual([
      '__EMPTY',
      '__EMPTY_1',
    ]);
    const errorsOnly = `<row r="2"><c r="A2" t="e"><v>#REF!</v></c></row><row r="3"><c r="A3"><v>1</v></c></row>`;
    expect(await objectsWith(errorsOnly, { errors: 'null' })).toEqual({ rows: [{ '#REF!': 1 }], headers: ['#REF!'], truncated: false });
  });
});

describe('openWorkbook: part names that differ in case', () => {
  it('EC-PART-NAME-CASE: finds the shared-string part when the relationship and the entry disagree on case', async () => {
    const bytes = await workbookBytes({
      'xl/sharedStrings.xml': undefined,
      'xl/SharedStrings.xml': sharedStrings('Name', 'Amount'),
      'xl/worksheets/sheet1.xml': sheetXml(`<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>`),
    });
    expect(await readRows(bytes)).toEqual([['Name', 'Amount']]);
  });

  it('EC-PART-NAME-CASE: finds a worksheet, the workbook and its relationships whatever their case', async () => {
    const bytes = await zip({
      '[Content_Types].xml': contentTypes({ '/xl/workbook.xml': SPREADSHEET_MAIN }),
      '_rels/.rels': ROOT_RELS,
      'XL/Workbook.xml': workbookXml([{ name: 'Data' }]),
      'XL/_rels/Workbook.xml.rels': relationships(
        { id: 'rId1', type: 'worksheet', target: 'worksheets/sheet1.xml' },
        { id: 'rId3', type: 'sharedStrings', target: 'sharedStrings.xml' },
      ),
      'XL/sharedstrings.XML': sharedStrings('shared'),
      'XL/Worksheets/Sheet1.xml': sheetXml(`<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>2</v></c></row>`, 'A1:B1'),
    });
    expect(await readRows(bytes)).toEqual([['shared', 2]]);
    const workbook = await openWorkbook(bytes);
    try {
      expect(await workbook.sheet(0).toObjects()).toEqual({ rows: [], headers: ['shared', '2'], truncated: false });
    } finally {
      await workbook.close();
    }
  });

  it('EC-PART-NAME-CASE: entries that differ only in case are not duplicates, and the exact name wins', async () => {
    const bytes = await workbookBytes({
      'xl/worksheets/sheet1.xml': sheetXml(row(1, ['exact'])),
      'xl/worksheets/Sheet1.xml': sheetXml(row(1, ['other case'])),
    });
    expect(await readRows(bytes)).toEqual([['exact']]);
  });
});

describe('openWorkbook: warnings', () => {
  const twoSheets = (): Promise<Uint8Array> =>
    workbookBytes({
      'xl/workbook.xml': workbookXml([{ name: 'One' }, { name: 'Two', relId: 'rId4' }]),
      'xl/_rels/workbook.xml.rels': relationships(
        { id: 'rId1', type: 'worksheet', target: 'worksheets/sheet1.xml' },
        { id: 'rId4', type: 'worksheet', target: 'worksheets/sheet2.xml' },
        { id: 'rId3', type: 'sharedStrings', target: 'sharedStrings.xml' },
      ),
      'xl/worksheets/sheet1.xml': sheetXml(
        `<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>5</v></c></row><row r="2"><c r="A2" t="s"><v>9</v></c></row>`,
      ),
      'xl/worksheets/sheet2.xml': sheetXml(`<row r="1"><c r="C1" t="s"><v>7</v></c></row>`),
    });

  it('EC-SST-INDEX-OUT-OF-RANGE: onWarning hears about an out-of-range index once per sheet, naming the sheet and cell', async () => {
    const warnings: ReadWarning[] = [];
    const workbook = await openWorkbook(await twoSheets(), { onWarning: warning => warnings.push(warning) });
    try {
      expect(await collect(workbook.sheet(0).rows())).toEqual([['Name', ''], ['']]);
      // Reading the sheet again repeats nothing; the other sheet reports its own.
      await collect(workbook.sheet(0).rows());
      await workbook.sheet(0).toObjects();
      await workbook.sheet(0).head(2);
      expect(await collect(workbook.sheet('Two').rows())).toEqual([[null, null, '']]);
    } finally {
      await workbook.close();
    }
    expect(warnings.map(({ code, sheet, ref }) => ({ code, sheet, ref }))).toEqual([
      { code: 'SST_INDEX_OUT_OF_RANGE', sheet: 'One', ref: 'B1' },
      { code: 'SST_INDEX_OUT_OF_RANGE', sheet: 'Two', ref: 'C1' },
    ]);
    expect(warnings[0]?.message).toContain('B1');
  });

  it('EC-SST-INDEX-OUT-OF-RANGE: a t="s" cell in a workbook without a shared-string part warns SHARED_STRINGS_MISSING', async () => {
    const warnings: ReadWarning[] = [];
    const bytes = await zip(
      packageParts({
        'xl/sharedStrings.xml': undefined,
        'xl/worksheets/sheet1.xml': sheetXml(`<row r="1"><c r="A1"><v>1</v></c><c r="B1" t="s"><v>0</v></c></row>`),
      }),
    );
    const workbook = await openWorkbook(bytes, { onWarning: warning => warnings.push(warning) });
    try {
      expect(await collect(workbook.sheet(0).rows())).toEqual([[1, '']]);
    } finally {
      await workbook.close();
    }
    expect(warnings.map(({ code, sheet, ref }) => ({ code, sheet, ref }))).toEqual([
      { code: 'SHARED_STRINGS_MISSING', sheet: 'Data', ref: 'B1' },
    ]);
  });

  it('reads the same with no callback', async () => {
    expect(await readRows(await twoSheets())).toEqual([['Name', ''], ['']]);
  });
});

describe('openWorkbook: what it accepts as input', () => {
  it('EC-INPUT-TYPE: anything that is not bytes is NOT_XLSX, naming what is accepted', async () => {
    for (const input of [null, undefined, 42, { data: 'x' }, ['P', 'K'], true]) {
      const error = await caught(() => openWorkbook(input as unknown as SourceInput));
      expect(error.code, String(input)).toBe('NOT_XLSX');
      expect(error.message, String(input)).toContain('ArrayBuffer');
      expect(error.message, String(input)).toContain('Blob');
    }
    const text = await caught(() => openWorkbook('PK\u0003\u0004binary string' as unknown as SourceInput));
    expect(text.code).toBe('NOT_XLSX');
    expect(text.message).toContain('convert a binary string to bytes');
    expect(text.detail?.received).toBe('string');
  });

  it('EC-INPUT-TYPE: any ArrayBufferView or SharedArrayBuffer reads without a copy', async () => {
    const bytes = await workbookBytes();
    const expected = [
      ['Name', 'Amount'],
      ['Zoë', 12.5],
    ];
    const padded = new Uint8Array(bytes.byteLength + 8);
    padded.set(bytes, 3);
    const shared = new SharedArrayBuffer(bytes.byteLength);
    new Uint8Array(shared).set(bytes);
    for (const input of [
      new DataView(padded.buffer, 3, bytes.byteLength),
      new Int8Array(padded.buffer, 3, bytes.byteLength),
      new Uint8ClampedArray(padded.buffer, 3, bytes.byteLength),
      shared,
      new Uint8Array(shared),
    ]) {
      expect(await readRows(input as unknown as Uint8Array), input.constructor.name).toEqual(expected);
    }
  });

  it('EC-INPUT-TYPE: buffers from another realm (an iframe, a worker, an Electron bridge) are accepted', async () => {
    const bytes = await workbookBytes();
    const foreign = runInNewContext(`new ArrayBuffer(${bytes.byteLength})`) as ArrayBuffer;
    expect(foreign instanceof ArrayBuffer).toBe(false);
    const foreignView = runInNewContext('buffer => new Uint8Array(buffer)')(foreign) as Uint8Array;
    foreignView.set(bytes);
    expect(foreignView instanceof Uint8Array).toBe(false);
    expect(await readRows(foreign as unknown as Uint8Array)).toEqual([
      ['Name', 'Amount'],
      ['Zoë', 12.5],
    ]);
    expect(await readRows(foreignView)).toHaveLength(2);
  });
});

describe('workbook.close', () => {
  it('fails every read after close with ABORTED, for in-memory bytes and for any other source', async () => {
    const bytes = await workbookBytes();
    let closed = false;
    const fragile: RandomAccessSource = {
      size: bytes.byteLength,
      read: (offset: number, length: number) =>
        closed ? Promise.reject(new Error('EBADF: bad file descriptor, read')) : Promise.resolve(bytes.subarray(offset, offset + length)),
      close: () => {
        closed = true;
        return Promise.resolve();
      },
    };
    for (const input of [bytes, fragile]) {
      const workbook = await openWorkbook(input);
      await workbook.close();
      const sheet = workbook.sheet(0);
      for (const read of [
        () => collect(sheet.rows()),
        () => sheet.toObjects(),
        () => sheet.head(1),
        () => collect(sheet.rows({ mode: 'object' })),
      ]) {
        const error = await caught(read);
        expect(error.code).toBe('ABORTED');
        expect(error.message).toContain('closed');
      }
      // Closing twice is harmless.
      await workbook.close();
    }
  });

  it('fails a read that is still streaming when the workbook closes under it', async () => {
    // Big enough to inflate in several chunks, so the read has more to pull after the first rows arrive.
    const rows = Array.from({ length: 20_000 }, (_, index) => row(index + 1, [`row ${index}`, index * 7919])).join('');
    const bytes = await sheetOnlyBytes(rows);
    const workbook = await openWorkbook(bytes);
    const error = await caught(async () => {
      for await (const _ of workbook.sheet(0).rows()) {
        await workbook.close();
      }
    });
    expect(error.code).toBe('ABORTED');
  });
});

// ---------------------------------------------------------------------------------------------------------------
// head()
// ---------------------------------------------------------------------------------------------------------------

describe('sheet.head', () => {
  it('returns the first rows as A1-keyed cells, with formulas and error codes', async () => {
    const rows =
      `${row(1, ['Id', 'Name'])}` +
      `<row r="2"><c r="A2"><f>1+1</f><v>2</v></c><c r="B2" t="e"><v>#N/A</v></c></row>` +
      `${row(3, ['skipped'])}`;
    const workbook = await openWorkbook(await sheetOnlyBytes(rows));
    try {
      const head: Map<string, RawCell> = await workbook.sheet(0).head(2);
      expect([...head.keys()]).toEqual(['A1', 'B1', 'A2', 'B2']);
      expect(head.get('A2')).toEqual({ value: 2, formula: '1+1' });
      expect(head.get('B2')).toEqual({ value: '#N/A', error: '#N/A' });
      expect(await workbook.sheet(0).head(0)).toEqual(new Map());
    } finally {
      await workbook.close();
    }
  });

  it('EC-MULTI-OBJECT-TEMPLATE: reads the Jetstream template header block and its data rows', async () => {
    // Replaces `worksheet['B1'|'B2'|'B3'|'A5'].v` plus `sheet_to_json({ range: 4, header: 1 })` (01 B3).
    const workbook = await openWorkbook(readFixture(fixtureById('jetstream-multi-object-template-gsheets')));
    try {
      const sheet = workbook.sheet('Create Accounts');
      const head = await sheet.head(5);
      expect(head.get('B1')?.value).toBe('Account');
      expect(head.get('B2')?.value).toBe('Insert');
      expect(head.get('B3')).toBeUndefined();
      expect(head.get('A5')?.value).toBe('Reference Id');
      expect(await collect(sheet.rows({ startRow: 5 }))).toEqual([
        ['Reference Id', 'Name', '{ParentId}'],
        ['account1', 'Account 1'],
        ['account2', 'Account 2'],
        ['account3', 'Account 3', 'account1'],
      ]);
    } finally {
      await workbook.close();
    }
  });
});
