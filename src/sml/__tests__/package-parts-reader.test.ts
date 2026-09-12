import { describe, expect, it } from 'vitest';
import { allFixtures, readFixture, readExpected, type Fixture } from '../../../test/helpers/fixtures';
import { sourceFrom } from '../../zip/source';
import { ZipReader } from '../../zip/zip-reader';
import { parseContentTypes, parseRels, parseWorkbook, relTypeIs, type Relationship } from '../package-parts';

const RELS_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const STRICT_RELS_NS = 'http://purl.oclc.org/ooxml/officeDocument/relationships';
const PACKAGE_RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

function relsXml(...relationships: string[]): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PACKAGE_RELS_NS}">${relationships.join('')}</Relationships>`;
}

function relationshipXml(id: string, type: string, target: string, targetMode = ''): string {
  return `<Relationship Id="${id}" Type="${type}" Target="${target}"${targetMode === '' ? '' : ` TargetMode="${targetMode}"`}/>`;
}

function targetOf(relationships: readonly Relationship[], id: string): string | undefined {
  return relationships.find(candidate => candidate.id === id)?.target;
}

describe('parseRels', () => {
  it('resolves a relative target against the source part folder', () => {
    const parsed = parseRels(relsXml(relationshipXml('rId1', `${RELS_NS}/worksheet`, 'worksheets/sheet1.xml')), 'xl/');
    expect(parsed).toEqual([{ id: 'rId1', type: `${RELS_NS}/worksheet`, target: 'xl/worksheets/sheet1.xml', external: false }]);
  });

  it('resolves a root-rels target against the package root', () => {
    const parsed = parseRels(relsXml(relationshipXml('rId1', `${RELS_NS}/officeDocument`, 'xl/workbook.xml')), '');
    expect(targetOf(parsed, 'rId1')).toBe('xl/workbook.xml');
  });

  it('EC-ABSOLUTE-REL-TARGETS: a leading slash means a package-root part name', () => {
    const parsed = parseRels(relsXml(relationshipXml('rId1', `${RELS_NS}/worksheet`, '/xl/worksheets/sheet1.xml')), 'xl/');
    expect(targetOf(parsed, 'rId1')).toBe('xl/worksheets/sheet1.xml');
  });

  it('EC-BACKSLASH-REL-TARGETS: Windows separators become slashes', () => {
    const parsed = parseRels(relsXml(relationshipXml('rId1', `${RELS_NS}/worksheet`, 'worksheets\\sheet1.xml')), 'xl/');
    expect(targetOf(parsed, 'rId1')).toBe('xl/worksheets/sheet1.xml');
  });

  it('collapses . and .. segments (the Google Sheets sheet rels point at ../comments1.xml)', () => {
    const parsed = parseRels(
      relsXml(
        relationshipXml('rId1', `${RELS_NS}/comments`, '../comments1.xml'),
        relationshipXml('rId2', `${RELS_NS}/vmlDrawing`, '../drawings/vmlDrawing1.vml'),
        relationshipXml('rId3', `${RELS_NS}/styles`, './styles.xml'),
      ),
      'xl/worksheets/',
    );
    expect(targetOf(parsed, 'rId1')).toBe('xl/comments1.xml');
    expect(targetOf(parsed, 'rId2')).toBe('xl/drawings/vmlDrawing1.vml');
    expect(targetOf(parsed, 'rId3')).toBe('xl/worksheets/styles.xml');
  });

  it('never produces a leading slash, whatever the separators', () => {
    const parsed = parseRels(relsXml(relationshipXml('rId1', `${RELS_NS}/worksheet`, '\\xl\\worksheets\\sheet1.xml')), '');
    expect(targetOf(parsed, 'rId1')).toBe('xl/worksheets/sheet1.xml');
  });

  it('tolerates a base path written without its trailing slash', () => {
    const parsed = parseRels(relsXml(relationshipXml('rId1', `${RELS_NS}/worksheet`, 'worksheets/sheet1.xml')), 'xl');
    expect(targetOf(parsed, 'rId1')).toBe('xl/worksheets/sheet1.xml');
  });

  it('keeps an external target verbatim and flags it', () => {
    const parsed = parseRels(
      relsXml(relationshipXml('rId1', `${RELS_NS}/hyperlink`, 'https://getjetstream.app/app/feedback', 'External')),
      'xl/worksheets/',
    );
    expect(parsed[0]).toEqual({
      id: 'rId1',
      type: `${RELS_NS}/hyperlink`,
      target: 'https://getjetstream.app/app/feedback',
      external: true,
    });
  });

  it('keeps the raw type string so Strict and vendor types survive', () => {
    const parsed = parseRels(
      relsXml(
        relationshipXml('rId1', `${STRICT_RELS_NS}/officeDocument`, 'xl/workbook.xml'),
        relationshipXml('rId7', 'http://customschemas.google.com/relationships/workbookmetadata', 'metadata'),
      ),
      '',
    );
    expect(parsed[0]?.type).toBe(`${STRICT_RELS_NS}/officeDocument`);
    expect(parsed[1]?.type).toBe('http://customschemas.google.com/relationships/workbookmetadata');
    expect(targetOf(parsed, 'rId7')).toBe('metadata');
  });

  it('ignores unknown elements and attributes', () => {
    const xml = `<?xml version="1.0"?><Relationships xmlns="${PACKAGE_RELS_NS}"><!-- noise --><Unknown Target="nope.xml"/>${relationshipXml(
      'rId1',
      `${RELS_NS}/worksheet`,
      'worksheets/sheet1.xml',
    ).replace('/>', ' Unexpected="1" TargetMode="Internal"/>')}</Relationships>`;
    const parsed = parseRels(xml, 'xl/');
    expect(parsed).toHaveLength(1);
    expect(parsed[0]?.external).toBe(false);
  });

  it('defaults a missing Id, Type and Target to empty strings rather than throwing', () => {
    const parsed = parseRels(`<Relationships xmlns="${PACKAGE_RELS_NS}"><Relationship/></Relationships>`, 'xl/');
    expect(parsed).toEqual([{ id: '', type: '', target: '', external: false }]);
  });
});

describe('relTypeIs', () => {
  it('EC-STRICT-NAMESPACES: matches both namespace families by suffix', () => {
    expect(relTypeIs(`${RELS_NS}/officeDocument`, 'officeDocument')).toBe(true);
    expect(relTypeIs(`${STRICT_RELS_NS}/officeDocument`, 'officeDocument')).toBe(true);
    expect(relTypeIs(`${RELS_NS}/worksheet`, 'officeDocument')).toBe(false);
    expect(relTypeIs('http://customschemas.google.com/relationships/workbookmetadata', 'worksheet')).toBe(false);
  });
});

describe('parseContentTypes', () => {
  const xml =
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="RELS" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension=".xml" ContentType="application/xml"/>' +
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
    '<Override ContentType="application/binary" PartName="/xl/metadata"/>' +
    '</Types>';

  it('lower-cases extensions and drops a leading dot', () => {
    const contentTypes = parseContentTypes(xml);
    expect(contentTypes.defaults.get('rels')).toBe('application/vnd.openxmlformats-package.relationships+xml');
    expect(contentTypes.defaults.get('xml')).toBe('application/xml');
  });

  it('normalizes override part names so zip entry names resolve', () => {
    const contentTypes = parseContentTypes(xml);
    expect(contentTypes.overrides.get('xl/workbook.xml')).toContain('spreadsheetml.sheet.main');
    expect(contentTypes.typeOf('xl/workbook.xml')).toContain('spreadsheetml.sheet.main');
    expect(contentTypes.typeOf('/xl/workbook.xml')).toContain('spreadsheetml.sheet.main');
  });

  it('EC-PART-NONSTANDARD-NAMES: an extensionless part resolves through its override only', () => {
    const contentTypes = parseContentTypes(xml);
    expect(contentTypes.typeOf('xl/metadata')).toBe('application/binary');
    expect(contentTypes.typeOf('xl/commentsmeta0')).toBeUndefined();
  });

  it('falls back to the extension default, case-insensitively', () => {
    const contentTypes = parseContentTypes(xml);
    expect(contentTypes.typeOf('xl/theme/theme1.XML')).toBe('application/xml');
    expect(contentTypes.typeOf('_rels/.rels')).toBe('application/vnd.openxmlformats-package.relationships+xml');
    expect(contentTypes.typeOf('xl/media/image1.png')).toBeUndefined();
  });
});

describe('parseWorkbook', () => {
  const workbookXml = (body: string): string =>
    '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' +
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    `${body}</workbook>`;

  it('reads sheets in document order with their rel ids', () => {
    const parsed = parseWorkbook(
      workbookXml(
        '<sheets><sheet name="Instructions" sheetId="1" r:id="rId4"/><sheet name="Create Accounts" sheetId="2" r:id="rId5"/></sheets>',
      ),
    );
    expect(parsed.sheets).toEqual([
      { name: 'Instructions', sheetId: 1, relId: 'rId4', state: 'visible' },
      { name: 'Create Accounts', sheetId: 2, relId: 'rId5', state: 'visible' },
    ]);
  });

  it('decodes XML entities in sheet names but leaves _xHHHH_ alone', () => {
    const parsed = parseWorkbook(workbookXml('<sheets><sheet name="It&apos;s a &amp; _x0041_ name" sheetId="1" r:id="rId1"/></sheets>'));
    expect(parsed.sheets[0]?.name).toBe("It's a & _x0041_ name");
  });

  it('EC-HIDDEN-SHEETS: reports hidden and veryHidden, defaulting to visible', () => {
    const parsed = parseWorkbook(
      workbookXml(
        '<sheets><sheet name="Data" sheetId="1" r:id="rId1"/>' +
          '<sheet name="Hidden" sheetId="2" state="hidden" r:id="rId2"/>' +
          '<sheet state="veryHidden" name="VeryHidden" sheetId="3" r:id="rId3"/></sheets>',
      ),
    );
    expect(parsed.sheets.map(sheet => sheet.state)).toEqual(['visible', 'hidden', 'veryHidden']);
  });

  it('falls back to position + 1 for a missing sheetId and to an empty rel id', () => {
    const parsed = parseWorkbook(workbookXml('<sheets><sheet name="A"/><sheet name="B"/></sheets>'));
    expect(parsed.sheets).toEqual([
      { name: 'A', sheetId: 1, relId: '', state: 'visible' },
      { name: 'B', sheetId: 2, relId: '', state: 'visible' },
    ]);
  });

  it('reads date1904 from 1 and true, and nothing else', () => {
    expect(parseWorkbook(workbookXml('<workbookPr date1904="1"/><sheets/>')).date1904).toBe(true);
    expect(parseWorkbook(workbookXml('<workbookPr date1904="true"/><sheets/>')).date1904).toBe(true);
    // EC-POI-BOOLEAN-ATTRIBUTE-SPELLING: Apache POI and LibreOffice spell the negative case out.
    expect(parseWorkbook(workbookXml('<workbookPr date1904="false"/><sheets/>')).date1904).toBe(false);
    expect(parseWorkbook(workbookXml('<workbookPr date1904="0"/><sheets/>')).date1904).toBe(false);
    expect(parseWorkbook(workbookXml('<workbookPr codeName="ThisWorkbook"/><sheets/>')).date1904).toBe(false);
    expect(parseWorkbook(workbookXml('<sheets/>')).date1904).toBe(false);
  });

  it('ignores sheet elements outside the sheet list', () => {
    const parsed = parseWorkbook(
      workbookXml(
        '<sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>' +
          '<customWorkbookViews><customWorkbookView name="view"><sheet name="Not a sheet"/></customWorkbookView></customWorkbookViews>',
      ),
    );
    expect(parsed.sheets.map(sheet => sheet.name)).toEqual(['Data']);
  });

  it('EC-STRICT-NAMESPACES: reads a Strict workbook the same way', () => {
    const parsed = parseWorkbook(
      '<?xml version="1.0"?><workbook xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"' +
        ' xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships" conformance="strict">' +
        '<workbookPr/><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>',
    );
    expect(parsed.sheets).toEqual([{ name: 'Data', sheetId: 1, relId: 'rId1', state: 'visible' }]);
  });

  it('EC-PREFIXED-ELEMENTS: reads a prefixed workbook the same way', () => {
    const parsed = parseWorkbook(
      '<x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' +
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<x:sheets><x:sheet name="Data" sheetId="1" r:id="rId1"/></x:sheets></x:workbook>',
    );
    expect(parsed.sheets).toEqual([{ name: 'Data', sheetId: 1, relId: 'rId1', state: 'visible' }]);
  });
});

// ---- the corpus ----------------------------------------------------------------------------------------------------

interface ExpectedSheet {
  readonly name: string;
  readonly hidden: boolean;
}

/**
 * Sheet lists the producer changed on the way in, so the fixture's expected dump (which describes the workbook that
 * was fed to it) no longer matches what the file says.
 */
const SHEET_LIST_OVERRIDES: Readonly<Record<string, readonly ExpectedSheet[]>> = {
  // write-excel-file has no hidden-sheet option, so the Hidden sheet never made it into the file.
  'golden-canonical-write-excel-file': [
    { name: 'Data', hidden: false },
    { name: 'Features', hidden: false },
    { name: "It's a very long sheet name 001", hidden: false },
  ],
  // Google Sheets strips the apostrophe from the long sheet name on import.
  'golden-canonical-gsheets-resave': [
    { name: 'Data', hidden: false },
    { name: 'Features', hidden: false },
    { name: 'Its a very long sheet name 001', hidden: false },
    { name: 'Hidden', hidden: true },
  ],
  // Numbers drops the hidden state on export (EC-HIDDEN-SHEETS, "import banner" note in the manifest).
  'golden-canonical-numbers': [
    { name: 'Data', hidden: false },
    { name: 'Features', hidden: false },
    { name: "It's a very long sheet name 001", hidden: false },
    { name: 'Hidden', hidden: false },
  ],
  // The same two application quirks on the re-exports of our own goldens.
  'golden-canonical-gsheets-resave-simple-excel': [
    { name: 'Data', hidden: false },
    { name: 'Features', hidden: false },
    { name: 'Its a very long sheet name 001', hidden: false },
    { name: 'Hidden', hidden: true },
  ],
  'golden-canonical-gsheets-resave-simple-excel-inline': [
    { name: 'Data', hidden: false },
    { name: 'Features', hidden: false },
    { name: 'Its a very long sheet name 001', hidden: false },
    { name: 'Hidden', hidden: true },
  ],
  'golden-canonical-numbers-resave-simple-excel': [
    { name: 'Data', hidden: false },
    { name: 'Features', hidden: false },
    { name: "It's a very long sheet name 001", hidden: false },
    { name: 'Hidden', hidden: false },
  ],
  'golden-canonical-numbers-resave-simple-excel-inline': [
    { name: 'Data', hidden: false },
    { name: 'Features', hidden: false },
    { name: "It's a very long sheet name 001", hidden: false },
    { name: 'Hidden', hidden: false },
  ],
  // Google Sheets names the imported sheet after the whole file name, extension included.
  'golden-canonical-gsheets-from-csv': [{ name: 'canonical.csv', hidden: false }],
};

function expectedSheets(fixture: Fixture): readonly ExpectedSheet[] | undefined {
  const override = SHEET_LIST_OVERRIDES[fixture.id];
  if (override) {
    return override;
  }
  if (!fixture.expected) {
    return undefined;
  }
  return readExpected(fixture).sheets.map(sheet => ({ name: sheet.name, hidden: sheet.hidden === true }));
}

/** The folder a part's own `.rels` targets resolve against (`xl/workbook.xml` -> `xl/`). */
function folderOf(partName: string): string {
  const lastSlash = partName.lastIndexOf('/');
  return lastSlash < 0 ? '' : partName.slice(0, lastSlash + 1);
}

function relsPathFor(partName: string): string {
  return `${folderOf(partName)}_rels/${partName.slice(folderOf(partName).length)}.rels`;
}

const corpusFixtures = allFixtures().filter(fixture => !fixture.tags.includes('kind:hostile') && fixture.path.endsWith('.xlsx'));

async function openFixture(fixture: Fixture): Promise<ZipReader> {
  return ZipReader.open(sourceFrom(readFixture(fixture)), { maxEntries: 10_000, maxInflatedBytes: 1 << 30 });
}

describe('every corpus package parses', () => {
  it.each(corpusFixtures.map(fixture => [fixture.id, fixture] as const))('%s', async (_id, fixture) => {
    const zip = await openFixture(fixture);

    const rootRels = parseRels(await zip.readText('_rels/.rels'), '');
    const officeDocument = rootRels.find(relationship => relTypeIs(relationship.type, 'officeDocument'));
    expect(officeDocument, 'root rels must name the workbook part').toBeDefined();
    const workbookPart = officeDocument?.target ?? '';
    expect(zip.has(workbookPart), `${workbookPart} must exist`).toBe(true);

    // Google Sheets writes [Content_Types].xml last in the archive; entry order never matters to a reader.
    expect(zip.has('[Content_Types].xml')).toBe(true);
    const contentTypes = parseContentTypes(await zip.readText('[Content_Types].xml'));
    expect(contentTypes.typeOf(workbookPart)).toContain('spreadsheetml.sheet.main');

    const workbook = parseWorkbook(await zip.readText(workbookPart));
    expect(workbook.sheets.length).toBeGreaterThan(0);
    expect(workbook.date1904).toBe(fixture.tags.includes('EC-DATE-1904'));

    const expected = expectedSheets(fixture);
    if (expected) {
      expect(workbook.sheets.map(sheet => ({ name: sheet.name, hidden: sheet.state !== 'visible' }))).toEqual(expected);
    }

    const workbookRels = parseRels(await zip.readText(relsPathFor(workbookPart)), folderOf(workbookPart));
    for (const relationship of workbookRels) {
      if (relTypeIs(relationship.type, 'worksheet') || relTypeIs(relationship.type, 'chartsheet')) {
        expect(zip.has(relationship.target), `${relationship.id} -> ${relationship.target}`).toBe(true);
      }
    }
    for (const sheet of workbook.sheets) {
      const target = workbookRels.find(relationship => relationship.id === sheet.relId)?.target;
      expect(target, `${sheet.name} (${sheet.relId}) must resolve`).toBeDefined();
      expect(zip.has(target ?? ''), `${sheet.name} -> ${target}`).toBe(true);
    }
    await zip.close();
  });
});
