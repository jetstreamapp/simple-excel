import { describe, expect, it } from 'vitest';
import { XlsxError } from '../../errors';
import {
  appXml,
  contentTypesXml,
  coreXml,
  relTypeIs,
  rootRelsXml,
  workbookRelsXml,
  workbookXml,
  type SheetPartInfo,
} from '../package-parts';

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PACKAGE_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
/** Strict OOXML swaps the namespace family but keeps the suffix (EC-STRICT-NAMESPACES). */
const STRICT_REL_NS = 'http://purl.oclc.org/ooxml/officeDocument/relationships';

const sheet = (overrides: Partial<SheetPartInfo> = {}): SheetPartInfo => ({
  name: 'Sheet1',
  sheetId: 1,
  relId: 'rId1',
  path: 'worksheets/sheet1.xml',
  hidden: false,
  ...overrides,
});

const TWO_SHEETS: readonly SheetPartInfo[] = [sheet(), sheet({ name: 'Sheet2', sheetId: 2, relId: 'rId2', path: 'worksheets/sheet2.xml' })];

describe('contentTypesXml', () => {
  it('overrides every SpreadsheetML part', () => {
    expect(contentTypesXml(TWO_SHEETS, true)).toBe(
      XML_DECLARATION +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
        '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
        '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' +
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
        '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' +
        '</Types>',
    );
  });

  it('leaves out the shared strings override when there is no table', () => {
    expect(contentTypesXml([sheet()], false)).not.toContain('sharedStrings');
  });
});

describe('rootRelsXml', () => {
  it('wires the workbook and both docProps parts', () => {
    expect(rootRelsXml()).toBe(
      XML_DECLARATION +
        `<Relationships xmlns="${PACKAGE_REL_NS}">` +
        `<Relationship Id="rId1" Type="${REL_NS}/officeDocument" Target="xl/workbook.xml"/>` +
        `<Relationship Id="rId2" Type="${PACKAGE_REL_NS}/metadata/core-properties" Target="docProps/core.xml"/>` +
        `<Relationship Id="rId3" Type="${REL_NS}/extended-properties" Target="docProps/app.xml"/>` +
        '</Relationships>',
    );
  });
});

describe('workbookXml', () => {
  it('writes a workbookView and one sheet element per sheet', () => {
    expect(workbookXml([sheet()], false)).toBe(
      XML_DECLARATION +
        `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${REL_NS}">` +
        '<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="20000" windowHeight="10000"/></bookViews>' +
        '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>' +
        '</workbook>',
    );
  });

  it('switches the date system with workbookPr', () => {
    expect(workbookXml([sheet()], true)).toContain('<workbookPr date1904="1"/><bookViews>');
  });

  it('marks hidden sheets and moves the active tab off them (EC-HIDDEN-SHEETS)', () => {
    const sheets = [
      sheet({ name: 'Gone', hidden: true }),
      sheet({ name: 'Shown', sheetId: 2, relId: 'rId2', path: 'worksheets/sheet2.xml' }),
    ];
    const xml = workbookXml(sheets, false);
    expect(xml).toContain('windowHeight="10000" activeTab="1"/>');
    expect(xml).toContain('<sheet name="Gone" sheetId="1" state="hidden" r:id="rId1"/>');
    expect(xml).toContain('<sheet name="Shown" sheetId="2" r:id="rId2"/>');
  });

  it('escapes sheet names in attributes', () => {
    const xml = workbookXml([sheet({ name: 'A & B <"C"> ’D’' })], false);
    expect(xml).toContain('<sheet name="A &amp; B &lt;&quot;C&quot;&gt; ’D’" sheetId="1" r:id="rId1"/>');
  });

  it('records an autofilter as a hidden _xlnm._FilterDatabase name', () => {
    const sheets = [
      sheet({ name: 'No filter' }),
      sheet({ name: "Bob's & co", sheetId: 2, relId: 'rId2', path: 'worksheets/sheet2.xml', autoFilterRange: 'A1:C12' }),
    ];
    expect(workbookXml(sheets, false)).toContain(
      '<definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="1" hidden="1">' +
        "'Bob''s &amp; co'!$A$1:$C$12</definedName></definedNames>",
    );
  });

  it('accepts a range that is already absolute and rejects one that is not a range', () => {
    expect(workbookXml([sheet({ autoFilterRange: '$a$1:$c$12' })], false)).toContain(">'Sheet1'!$A$1:$C$12<");
    expect(() => workbookXml([sheet({ autoFilterRange: 'A1' })], false)).toThrow(XlsxError);
    expect(() => workbookXml([sheet({ autoFilterRange: 'Sheet1!A1:C12' })], false)).toThrow(XlsxError);
  });

  it('omits definedNames when no sheet has an autofilter', () => {
    expect(workbookXml(TWO_SHEETS, false)).not.toContain('definedNames');
  });
});

describe('workbookRelsXml', () => {
  it('keeps the sheet rel ids and takes the next two for styles and shared strings', () => {
    expect(workbookRelsXml(TWO_SHEETS, true)).toBe(
      XML_DECLARATION +
        `<Relationships xmlns="${PACKAGE_REL_NS}">` +
        `<Relationship Id="rId1" Type="${REL_NS}/worksheet" Target="worksheets/sheet1.xml"/>` +
        `<Relationship Id="rId2" Type="${REL_NS}/worksheet" Target="worksheets/sheet2.xml"/>` +
        `<Relationship Id="rId3" Type="${REL_NS}/styles" Target="styles.xml"/>` +
        `<Relationship Id="rId4" Type="${REL_NS}/sharedStrings" Target="sharedStrings.xml"/>` +
        '</Relationships>',
    );
  });

  it('numbers styles after the highest sheet id, however the caller numbered them', () => {
    const sheets = [sheet({ relId: 'rId7' })];
    const xml = workbookRelsXml(sheets, false);
    expect(xml).toContain(`<Relationship Id="rId8" Type="${REL_NS}/styles" Target="styles.xml"/>`);
    expect(xml).not.toContain('sharedStrings');
  });
});

describe('coreXml', () => {
  const now = new Date('2026-09-12T15:04:05.678Z');

  it('defaults the creator and stamps both dates to the second', () => {
    expect(coreXml({}, now)).toBe(
      XML_DECLARATION +
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"' +
        ' xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"' +
        ' xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
        '<dc:creator>simple-excel</dc:creator>' +
        '<dcterms:created xsi:type="dcterms:W3CDTF">2026-09-12T15:04:05Z</dcterms:created>' +
        '<dcterms:modified xsi:type="dcterms:W3CDTF">2026-09-12T15:04:05Z</dcterms:modified>' +
        '</cp:coreProperties>',
    );
  });

  it('keeps the caller’s creator, title and creation date', () => {
    const xml = coreXml({ creator: 'Ada & Co', title: 'Q3 <report>', created: new Date('2020-01-02T03:04:05Z') }, now);
    expect(xml).toContain('<dc:title>Q3 &lt;report&gt;</dc:title><dc:creator>Ada &amp; Co</dc:creator>');
    expect(xml).toContain('<dcterms:created xsi:type="dcterms:W3CDTF">2020-01-02T03:04:05Z</dcterms:created>');
    expect(xml).toContain('<dcterms:modified xsi:type="dcterms:W3CDTF">2026-09-12T15:04:05Z</dcterms:modified>');
  });

  it('EC-XML-CONTROL-CHARS-METADATA: drops characters XML forbids from the title and creator', () => {
    const xml = coreXml({ creator: 'Ada\u0000 Lovelace\uD800', title: 'Q3\u0001 report\uFFFF' }, now);
    expect(xml).toContain('<dc:title>Q3 report</dc:title><dc:creator>Ada Lovelace</dc:creator>');
  });

  it('EC-DOCPROPS-CREATED-RANGE: refuses a timestamp without a four-digit year instead of writing a bad one', () => {
    const tenThousand = new Date(Date.UTC(10_000, 0, 1));
    const negative = new Date(Date.UTC(2000, 0, 1));
    negative.setUTCFullYear(-1);
    const yearZero = new Date(Date.UTC(2000, 0, 1));
    yearZero.setUTCFullYear(0);
    for (const created of [new Date(Number.NaN), tenThousand, negative, yearZero]) {
      expect(() => coreXml({ created }, now), String(created)).toThrowError(expect.objectContaining({ code: 'WRITER_STATE' }));
    }
    const yearOne = new Date(Date.UTC(2000, 0, 1));
    yearOne.setUTCFullYear(1);
    expect(coreXml({ created: yearOne }, now)).toContain(
      '<dcterms:created xsi:type="dcterms:W3CDTF">0001-01-01T00:00:00Z</dcterms:created>',
    );
  });
});

describe('workbookXml and appXml metadata (EC-XML-CONTROL-CHARS-METADATA)', () => {
  it('never writes a raw control character, even for a name that skipped sanitizing', () => {
    const sheets = [sheet({ name: 'Bad\u0001Name\uDC00' })];
    expect(workbookXml(sheets, false)).toContain('<sheet name="BadName" sheetId="1" r:id="rId1"/>');
    expect(appXml(sheets)).toContain('<vt:lpstr>BadName</vt:lpstr>');
  });
});

describe('appXml', () => {
  it('lists the worksheets as a heading pair and a title vector', () => {
    expect(appXml(TWO_SHEETS)).toBe(
      XML_DECLARATION +
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"' +
        ' xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
        '<Application>simple-excel</Application>' +
        '<HeadingPairs><vt:vector size="2" baseType="variant">' +
        '<vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>2</vt:i4></vt:variant>' +
        '</vt:vector></HeadingPairs>' +
        '<TitlesOfParts><vt:vector size="2" baseType="lpstr">' +
        '<vt:lpstr>Sheet1</vt:lpstr><vt:lpstr>Sheet2</vt:lpstr>' +
        '</vt:vector></TitlesOfParts>' +
        '</Properties>',
    );
  });

  it('escapes sheet names in the title vector', () => {
    expect(appXml([sheet({ name: 'A & <B>' })])).toContain('<vt:lpstr>A &amp; &lt;B&gt;</vt:lpstr>');
  });
});

describe('relTypeIs (EC-STRICT-NAMESPACES)', () => {
  it.each(['worksheet', 'styles', 'sharedStrings', 'officeDocument'])('matches the %s type in both namespace families', suffix => {
    expect(relTypeIs(`${REL_NS}/${suffix}`, suffix)).toBe(true);
    expect(relTypeIs(`${STRICT_REL_NS}/${suffix}`, suffix)).toBe(true);
  });

  it('does not match a different or partial suffix', () => {
    expect(relTypeIs(`${REL_NS}/chartsheet`, 'worksheet')).toBe(false);
    expect(relTypeIs(`${REL_NS}/worksheet`, 'sheet')).toBe(false);
    expect(relTypeIs('http://schemas.openxmlformats.org/officeDocument/2006/worksheet', 'worksheet')).toBe(false);
    expect(relTypeIs(`${PACKAGE_REL_NS}/metadata/core-properties`, 'core-properties')).toBe(false);
  });

  it('matches the package core-properties type by its full suffix', () => {
    expect(relTypeIs(`${PACKAGE_REL_NS}/metadata/core-properties`, 'metadata/core-properties')).toBe(true);
  });
});
