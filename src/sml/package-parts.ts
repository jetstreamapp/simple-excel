import { XlsxError } from '../errors';
import { notImplemented } from '../internal/not-implemented';
import type { WorkbookProperties } from '../types';
import { escapeAttr, escapeText } from '../xml/escape';

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const SPREADSHEET_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const CONTENT_TYPES_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
const PACKAGE_RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OFFICE_REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CORE_PROPERTIES_REL_TYPE = `${PACKAGE_RELS_NS}/metadata/core-properties`;
const EXTENDED_PROPERTIES_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties';
const DOC_PROPS_VTYPES_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes';
const CORE_PROPERTIES_NS = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties';

const SPREADSHEET_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml';
const APPLICATION_NAME = 'simple-excel';

/** `A1:C12` as the absolute `$A$1:$C$12` a defined name needs. */
const CELL_REF = /^\$?([A-Za-z]+)\$?([0-9]+)$/;

export interface SheetPartInfo {
  readonly name: string;
  /** 1-based, stable id in workbook.xml. */
  readonly sheetId: number;
  /** Relationship id, `rId<n>`. */
  readonly relId: string;
  /** Part name relative to `xl/` (`worksheets/sheet1.xml`). */
  readonly path: string;
  readonly hidden: boolean;
  /** Present when the sheet has an autofilter (emits the `_xlnm._FilterDatabase` defined name). */
  readonly autoFilterRange?: string;
}

// ---- writers (all sync, all return complete XML documents) ---------------------------------------------------------

/**
 * Every SpreadsheetML part needs an `Override`: the `xml` `Default` maps them to generic `application/xml`, and Excel
 * then fails to find the sheets (primer 2.3).
 */
export function contentTypesXml(sheets: readonly SheetPartInfo[], hasSharedStrings: boolean): string {
  const parts: string[] = [
    XML_DECLARATION,
    `<Types xmlns="${CONTENT_TYPES_NS}">`,
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
    '<Default Extension="xml" ContentType="application/xml"/>',
    override('/xl/workbook.xml', `${SPREADSHEET_CONTENT_TYPE}.sheet.main+xml`),
  ];
  for (const sheet of sheets) {
    parts.push(override(`/xl/${sheet.path}`, `${SPREADSHEET_CONTENT_TYPE}.worksheet+xml`));
  }
  parts.push(override('/xl/styles.xml', `${SPREADSHEET_CONTENT_TYPE}.styles+xml`));
  if (hasSharedStrings) {
    parts.push(override('/xl/sharedStrings.xml', `${SPREADSHEET_CONTENT_TYPE}.sharedStrings+xml`));
  }
  parts.push(override('/docProps/core.xml', 'application/vnd.openxmlformats-package.core-properties+xml'));
  parts.push(override('/docProps/app.xml', 'application/vnd.openxmlformats-officedocument.extended-properties+xml'));
  parts.push('</Types>');
  return parts.join('');
}

export function rootRelsXml(): string {
  return [
    XML_DECLARATION,
    `<Relationships xmlns="${PACKAGE_RELS_NS}">`,
    relationship('rId1', `${OFFICE_REL_NS}/officeDocument`, 'xl/workbook.xml'),
    relationship('rId2', CORE_PROPERTIES_REL_TYPE, 'docProps/core.xml'),
    relationship('rId3', `${OFFICE_REL_NS}/extended-properties`, 'docProps/app.xml'),
    '</Relationships>',
  ].join('');
}

export function workbookXml(sheets: readonly SheetPartInfo[], date1904: boolean): string {
  const parts: string[] = [XML_DECLARATION, `<workbook xmlns="${SPREADSHEET_NS}" xmlns:r="${OFFICE_REL_NS}">`];
  if (date1904) {
    parts.push('<workbookPr date1904="1"/>');
  }
  // `activeTab` defaults to 0, which Excel repairs when sheet 0 is hidden, so point it at the first visible sheet.
  const activeTab = sheets.findIndex(sheet => !sheet.hidden);
  const activeTabAttribute = activeTab > 0 ? ` activeTab="${activeTab}"` : '';
  parts.push(
    `<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="20000" windowHeight="10000"${activeTabAttribute}/></bookViews>`,
  );

  parts.push('<sheets>');
  for (const sheet of sheets) {
    const state = sheet.hidden ? ' state="hidden"' : '';
    parts.push(`<sheet name="${escapeAttr(sheet.name)}" sheetId="${sheet.sheetId}"${state} r:id="${escapeAttr(sheet.relId)}"/>`);
  }
  parts.push('</sheets>');

  const definedNames = filterDatabaseNames(sheets);
  if (definedNames !== '') {
    parts.push(`<definedNames>${definedNames}</definedNames>`);
  }
  parts.push('</workbook>');
  return parts.join('');
}

/**
 * Sheet relationship ids come from the caller; styles and shared strings take the two ids after the highest sheet id,
 * so the facade must reserve `rId1..rIdN` for the sheets and leave `rId(N+1)` (styles) and `rId(N+2)` (sst) free.
 */
export function workbookRelsXml(sheets: readonly SheetPartInfo[], hasSharedStrings: boolean): string {
  let highestRelNumber = 0;
  const parts: string[] = [XML_DECLARATION, `<Relationships xmlns="${PACKAGE_RELS_NS}">`];
  for (const sheet of sheets) {
    highestRelNumber = Math.max(highestRelNumber, relIdNumber(sheet.relId));
    parts.push(relationship(sheet.relId, `${OFFICE_REL_NS}/worksheet`, sheet.path));
  }
  parts.push(relationship(`rId${highestRelNumber + 1}`, `${OFFICE_REL_NS}/styles`, 'styles.xml'));
  if (hasSharedStrings) {
    parts.push(relationship(`rId${highestRelNumber + 2}`, `${OFFICE_REL_NS}/sharedStrings`, 'sharedStrings.xml'));
  }
  parts.push('</Relationships>');
  return parts.join('');
}

/** `docProps/core.xml`; `now` is the created/modified timestamp (fixed when deterministic). */
export function coreXml(properties: WorkbookProperties, now: Date): string {
  const parts: string[] = [
    XML_DECLARATION,
    `<cp:coreProperties xmlns:cp="${CORE_PROPERTIES_NS}" xmlns:dc="http://purl.org/dc/elements/1.1/"` +
      ' xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/"' +
      ' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">',
  ];
  if (properties.title !== undefined) {
    parts.push(`<dc:title>${escapeText(properties.title)}</dc:title>`);
  }
  parts.push(`<dc:creator>${escapeText(properties.creator ?? APPLICATION_NAME)}</dc:creator>`);
  parts.push(`<dcterms:created xsi:type="dcterms:W3CDTF">${w3cdtf(properties.created ?? now)}</dcterms:created>`);
  parts.push(`<dcterms:modified xsi:type="dcterms:W3CDTF">${w3cdtf(now)}</dcterms:modified>`);
  parts.push('</cp:coreProperties>');
  return parts.join('');
}

export function appXml(sheets: readonly SheetPartInfo[]): string {
  const parts: string[] = [
    XML_DECLARATION,
    `<Properties xmlns="${EXTENDED_PROPERTIES_NS}" xmlns:vt="${DOC_PROPS_VTYPES_NS}">`,
    `<Application>${APPLICATION_NAME}</Application>`,
    '<HeadingPairs><vt:vector size="2" baseType="variant">' +
      `<vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>${sheets.length}</vt:i4></vt:variant>` +
      '</vt:vector></HeadingPairs>',
    `<TitlesOfParts><vt:vector size="${sheets.length}" baseType="lpstr">`,
  ];
  for (const sheet of sheets) {
    parts.push(`<vt:lpstr>${escapeText(sheet.name)}</vt:lpstr>`);
  }
  parts.push('</vt:vector></TitlesOfParts>', '</Properties>');
  return parts.join('');
}

function override(partName: string, contentType: string): string {
  return `<Override PartName="${escapeAttr(partName)}" ContentType="${contentType}"/>`;
}

function relationship(id: string, type: string, target: string): string {
  return `<Relationship Id="${escapeAttr(id)}" Type="${type}" Target="${escapeAttr(target)}"/>`;
}

function relIdNumber(relId: string): number {
  const parsed = Number.parseInt(relId.replace(/^rId/i, ''), 10);
  return Number.isFinite(parsed) ? parsed : 0;
}

/** ISO 8601 to the second, the way every producer writes `dcterms:created` (fractions make some readers stumble). */
function w3cdtf(date: Date): string {
  return `${date.toISOString().slice(0, 19)}Z`;
}

/** `'Create Accounts'` - a defined name always quotes the sheet name, with embedded apostrophes doubled. */
function quoteSheetName(name: string): string {
  return `'${name.replaceAll("'", "''")}'`;
}

function absoluteRef(ref: string): string {
  const match = CELL_REF.exec(ref.trim());
  const column = match?.[1];
  const row = match?.[2];
  if (column === undefined || row === undefined) {
    throw new XlsxError('WRITER_STATE', `"${ref}" is not a cell reference like A1.`, { ref });
  }
  return `$${column.toUpperCase()}$${row}`;
}

function absoluteRange(range: string): string {
  const [start, end] = range.split(':');
  if (start === undefined || end === undefined) {
    throw new XlsxError('WRITER_STATE', `"${range}" is not a range like A1:C12.`, { range });
  }
  return `${absoluteRef(start)}:${absoluteRef(end)}`;
}

/** Excel records an autofilter as a hidden `_xlnm._FilterDatabase` name scoped to the sheet's position. */
function filterDatabaseNames(sheets: readonly SheetPartInfo[]): string {
  let definedNames = '';
  for (const [index, sheet] of sheets.entries()) {
    if (sheet.autoFilterRange === undefined) {
      continue;
    }
    const target = `${quoteSheetName(sheet.name)}!${absoluteRange(sheet.autoFilterRange)}`;
    definedNames += `<definedName name="_xlnm._FilterDatabase" localSheetId="${index}" hidden="1">${escapeText(target)}</definedName>`;
  }
  return definedNames;
}

// ---- readers -------------------------------------------------------------------------------------------------------

export interface Relationship {
  readonly id: string;
  /** Type URI; compare with `relTypeIs` so Strict and Transitional URIs both match. */
  readonly type: string;
  /** Resolved zip entry name (relative to the package root, no leading slash, slashes normalized). */
  readonly target: string;
  readonly external: boolean;
}

/** Parse a `.rels` part; `basePath` is the directory of the source part (`'xl/'` for workbook rels, `''` for root). */
export function parseRels(xml: string, basePath: string): Relationship[] {
  void xml;
  void basePath;
  throw notImplemented('sml/package-parts');
}

/** True when a relationship type ends with `/relationships/<suffix>` (works for Strict and Transitional). */
export function relTypeIs(type: string, suffix: string): boolean {
  return type.endsWith(`/relationships/${suffix}`);
}

export interface ContentTypes {
  readonly defaults: ReadonlyMap<string, string>;
  readonly overrides: ReadonlyMap<string, string>;
  /** Content type for a part name, via override then default extension; undefined when unknown. */
  typeOf(partName: string): string | undefined;
}

export function parseContentTypes(xml: string): ContentTypes {
  void xml;
  throw notImplemented('sml/package-parts');
}

export interface WorkbookSheetEntry {
  readonly name: string;
  readonly sheetId: number;
  readonly relId: string;
  readonly state: 'visible' | 'hidden' | 'veryHidden';
}

export interface ParsedWorkbook {
  readonly sheets: readonly WorkbookSheetEntry[];
  readonly date1904: boolean;
}

export function parseWorkbook(xml: string): ParsedWorkbook {
  void xml;
  throw notImplemented('sml/package-parts');
}
