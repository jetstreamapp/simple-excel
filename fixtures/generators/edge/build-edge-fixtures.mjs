#!/usr/bin/env node
// Hand-builds the structural edge-case and hostile fixtures from raw XML parts (see minizip.mjs).
// Every fixture is a variation of the same 3-row workbook so reader dumps are directly comparable:
//   A1 "Name"  B1 "Amount"  C1 "When"
//   A2 "Zoë"   B2 12.5      C2 2024-02-29 (serial 45351, numFmt 14)
//   A3 "x"     B3 true      C3 =B2*2 (cached 25)
//
//   node build-edge-fixtures.mjs [outDir]            (default: ../../edge and ../../hostile)
import { copyFileSync, existsSync, mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { buildZip } from './minizip.mjs';

const HERE = new URL('.', import.meta.url).pathname.replace(/\/$/, '');
const FIXTURES = join(HERE, '../..');
const EDGE = join(FIXTURES, 'edge');
const HOSTILE = join(FIXTURES, 'hostile');
mkdirSync(EDGE, { recursive: true });
mkdirSync(HOSTILE, { recursive: true });

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_STRICT = 'http://purl.oclc.org/ooxml/spreadsheetml/main';
const NS_R_STRICT = 'http://purl.oclc.org/ooxml/officeDocument/relationships';

const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

function contentTypes(extra = '') {
  return `${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>${extra}</Types>`;
}
const rootRels = `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`;
function workbook(ns = NS, nsR = NS_R) {
  return `${XML}<workbook xmlns="${ns}" xmlns:r="${nsR}"><workbookPr/><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>`;
}
function workbookRels(sheetTarget = 'worksheets/sheet1.xml', stylesTarget = 'styles.xml', sstTarget = 'sharedStrings.xml') {
  return `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="${sheetTarget}"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="${stylesTarget}"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="${sstTarget}"/></Relationships>`;
}
const styles = `${XML}<styleSheet xmlns="${NS}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;
function sst(items = ['Name', 'Amount', 'When', 'Zoë', 'x']) {
  return `${XML}<sst xmlns="${NS}" count="${items.length}" uniqueCount="${items.length}">${items.map(text => `<si><t>${text}</t></si>`).join('')}</sst>`;
}
function sheet({ ns = NS, prefix = '', rows = null, dimension = '<dimension ref="A1:C3"/>' } = {}) {
  const p = prefix ? `${prefix}:` : '';
  const xmlns = prefix ? `xmlns:${prefix}="${ns}"` : `xmlns="${ns}"`;
  const body =
    rows ??
    `<${p}row r="1"><${p}c r="A1" t="s"><${p}v>0</${p}v></${p}c><${p}c r="B1" t="s"><${p}v>1</${p}v></${p}c><${p}c r="C1" t="s"><${p}v>2</${p}v></${p}c></${p}row><${p}row r="2"><${p}c r="A2" t="s"><${p}v>3</${p}v></${p}c><${p}c r="B2"><${p}v>12.5</${p}v></${p}c><${p}c r="C2" s="1"><${p}v>45351</${p}v></${p}c></${p}row><${p}row r="3"><${p}c r="A3" t="s"><${p}v>4</${p}v></${p}c><${p}c r="B3" t="b"><${p}v>1</${p}v></${p}c><${p}c r="C3"><${p}f>B2*2</${p}f><${p}v>25</${p}v></${p}c></${p}row>`;
  return `${XML}<${p}worksheet ${xmlns}>${dimension}<${p}sheetData>${body}</${p}sheetData></${p}worksheet>`;
}

function parts(overrides = {}) {
  return {
    '[Content_Types].xml': contentTypes(),
    '_rels/.rels': rootRels,
    'xl/workbook.xml': workbook(),
    'xl/_rels/workbook.xml.rels': workbookRels(),
    'xl/styles.xml': styles,
    'xl/sharedStrings.xml': sst(),
    'xl/worksheets/sheet1.xml': sheet(),
    ...overrides,
  };
}

function write(dir, name, entriesOrParts, zipOptions) {
  const entries = Array.isArray(entriesOrParts)
    ? entriesOrParts
    : Object.entries(entriesOrParts).map(([entryName, data]) => ({ name: entryName, data }));
  const archive = buildZip(entries, zipOptions);
  const path = join(dir, name);
  writeFileSync(path, archive);
  console.log(`${name.padEnd(46)} ${String(archive.length).padStart(9)} bytes`);
  return path;
}

// ---- edge (valid-but-unusual) ----
write(EDGE, 'baseline-minimal.xlsx', parts());
write(EDGE, 'no-dimension.xlsx', parts({ 'xl/worksheets/sheet1.xml': sheet({ dimension: '' }) }));
write(
  EDGE,
  'missing-r-attributes.xlsx',
  parts({
    'xl/worksheets/sheet1.xml': sheet({
      rows: '<row><c t="s"><v>0</v></c><c t="s"><v>1</v></c><c t="s"><v>2</v></c></row><row><c t="s"><v>3</v></c><c><v>12.5</v></c><c s="1"><v>45351</v></c></row><row><c t="s"><v>4</v></c><c t="b"><v>1</v></c><c><f>B2*2</f><v>25</v></c></row>',
    }),
  }),
);
write(
  EDGE,
  'empty-v-element.xlsx',
  parts({
    'xl/worksheets/sheet1.xml': sheet({
      rows: '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c></row><row r="2"><c r="A2" t="s"><v>3</v></c><c r="B2"><v/></c><c r="C2" s="1"/></row><row r="3"><c r="A3" t="s"><v>4</v></c><c r="B3" t="b"><v>1</v></c><c r="C3"><f>B2*2</f><v>25</v></c></row>',
    }),
  }),
);
write(
  EDGE,
  'inline-strings-cdata.xlsx',
  parts({
    'xl/worksheets/sheet1.xml': sheet({
      rows: '<row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t><![CDATA[Amount & <more>]]></t></is></c><c r="C1" t="inlineStr"><is><r><t>Wh</t></r><r><rPr><b/></rPr><t>en</t></r></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t xml:space="preserve"> Zoë </t></is></c><c r="B2"><v>12.5</v></c><c r="C2" s="1"><v>45351</v></c></row>',
    }),
  }),
);
write(EDGE, 'prefixed-elements.xlsx', parts({ 'xl/worksheets/sheet1.xml': sheet({ prefix: 'x' }) }));
write(EDGE, 'strict-namespaces.xlsx', {
  '[Content_Types].xml': contentTypes(),
  '_rels/.rels': rootRels.replace(
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
    'http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument',
  ),
  'xl/workbook.xml': workbook(NS_STRICT, NS_R_STRICT),
  'xl/_rels/workbook.xml.rels': workbookRels().replaceAll(
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/',
    'http://purl.oclc.org/ooxml/officeDocument/relationships/',
  ),
  'xl/styles.xml': styles.replace(NS, NS_STRICT),
  'xl/sharedStrings.xml': sst().replace(NS, NS_STRICT),
  'xl/worksheets/sheet1.xml': sheet({
    ns: NS_STRICT,
    rows: '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c></row><row r="2"><c r="A2" t="s"><v>3</v></c><c r="B2"><v>12.5</v></c><c r="C2" t="d" s="1"><v>2024-02-29</v></c></row><row r="3"><c r="A3" t="s"><v>4</v></c><c r="B3" t="b"><v>1</v></c><c r="C3"><f>B2*2</f><v>25</v></c></row>',
  }),
});
write(
  EDGE,
  'absolute-rel-targets.xlsx',
  parts({ 'xl/_rels/workbook.xml.rels': workbookRels('/xl/worksheets/sheet1.xml', '/xl/styles.xml', '/xl/sharedStrings.xml') }),
);
write(EDGE, 'backslash-rel-targets.xlsx', parts({ 'xl/_rels/workbook.xml.rels': workbookRels('worksheets\\sheet1.xml') }));
write(EDGE, 'sst-after-sheet-data-descriptors.xlsx', [
  { name: 'xl/worksheets/sheet1.xml', data: sheet(), dataDescriptor: true },
  { name: 'xl/sharedStrings.xml', data: sst(), dataDescriptor: true },
  { name: 'xl/styles.xml', data: styles, dataDescriptor: true },
  { name: 'xl/workbook.xml', data: workbook(), dataDescriptor: true },
  { name: 'xl/_rels/workbook.xml.rels', data: workbookRels(), dataDescriptor: true },
  { name: '_rels/.rels', data: rootRels, dataDescriptor: true },
  { name: '[Content_Types].xml', data: contentTypes(), dataDescriptor: true },
]);
write(
  EDGE,
  'stored-entries.xlsx',
  Object.entries(parts()).map(([name, data]) => ({ name, data, method: 'store' })),
);
write(EDGE, 'no-shared-strings-part.xlsx', {
  '[Content_Types].xml': contentTypes().replace(
    '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>',
    '',
  ),
  '_rels/.rels': rootRels,
  'xl/workbook.xml': workbook(),
  'xl/_rels/workbook.xml.rels': workbookRels().replace(/<Relationship Id="rId3"[^>]*\/>/, ''),
  'xl/styles.xml': styles,
  'xl/worksheets/sheet1.xml': sheet({
    rows: '<row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Amount</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>Zoë</t></is></c><c r="B2"><v>12.5</v></c></row>',
  }),
});
write(EDGE, 'sst-index-out-of-range.xlsx', parts({ 'xl/sharedStrings.xml': sst(['Name', 'Amount']) }));
write(EDGE, 'date1904.xlsx', parts({ 'xl/workbook.xml': workbook().replace('<workbookPr/>', '<workbookPr date1904="1"/>') }));
write(EDGE, 'hidden-and-veryhidden-sheets.xlsx', {
  ...parts(),
  '[Content_Types].xml': contentTypes(
    '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>',
  ),
  'xl/workbook.xml': `${XML}<workbook xmlns="${NS}" xmlns:r="${NS_R}"><workbookPr/><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Hidden" sheetId="2" state="hidden" r:id="rId4"/><sheet name="VeryHidden" sheetId="3" state="veryHidden" r:id="rId5"/></sheets></workbook>`,
  'xl/_rels/workbook.xml.rels': workbookRels().replace(
    '</Relationships>',
    `<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/></Relationships>`,
  ),
  'xl/worksheets/sheet2.xml': sheet({ rows: '<row r="1"><c r="A1" t="inlineStr"><is><t>hidden</t></is></c></row>', dimension: '' }),
  'xl/worksheets/sheet3.xml': sheet({ rows: '<row r="1"><c r="A1" t="inlineStr"><is><t>veryHidden</t></is></c></row>', dimension: '' }),
});

// ---- hostile ----
write(
  HOSTILE,
  'xxe-doctype-in-sharedstrings.xlsx',
  parts({
    'xl/sharedStrings.xml': `<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE sst [<!ENTITY xxe SYSTEM "file:///etc/passwd"><!ENTITY lol "lol"><!ENTITY lol2 "&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;">]>\n<sst xmlns="${NS}" count="5" uniqueCount="5"><si><t>Name</t></si><si><t>&xxe;</t></si><si><t>&lol2;</t></si><si><t>Zoë</t></si><si><t>x</t></si></sst>`,
  }),
);
write(HOSTILE, 'truncated-central-directory.xlsx', parts(), { truncateAt: 1400 });
write(HOSTILE, 'crc-mismatch.xlsx', parts(), { corruptCrc: true });
write(HOSTILE, 'duplicate-sheet-entries.xlsx', [
  ...Object.entries(parts()).map(([name, data]) => ({ name, data })),
  {
    name: 'xl/worksheets/sheet1.xml',
    data: sheet({ rows: '<row r="1"><c r="A1" t="inlineStr"><is><t>DUPLICATE ENTRY WINS?</t></is></c></row>', dimension: '' }),
  },
]);
write(
  HOSTILE,
  'deeply-nested-rich-text.xlsx',
  parts({
    'xl/sharedStrings.xml': `${XML}<sst xmlns="${NS}" count="1" uniqueCount="1"><si>${'<r>'.repeat(50000)}<t>deep</t>${'</r>'.repeat(50000)}</si></sst>`,
  }),
);
// zip bomb: a 30 MB sheet of repeated inline cells deflates ~190:1 (~160 KB, under the 200 KB commit cap);
// a multi-GB variant belongs in generate.mjs --big (generated, never committed)
const bombRow = index => `<row r="${index}">${('<c t="inlineStr"><is><t>' + 'A'.repeat(200) + '</t></is></c>').repeat(20)}</row>`;
const bombRows = [];
for (let i = 1; i <= 6000; i++) {
  bombRows.push(bombRow(i));
}
write(HOSTILE, 'zip-bomb-30mb-sheet.xlsx', parts({ 'xl/worksheets/sheet1.xml': sheet({ rows: bombRows.join(''), dimension: '' }) }));
// renamed formats (bytes of another format under an .xlsx name)
copyFileSync(join(FIXTURES, 'canonical/canonical.csv'), join(HOSTILE, 'csv-bytes-renamed.xlsx'));
for (const [source, target] of [
  ['.generated/hostile-src/canonical.ods', 'ods-renamed.xlsx'],
  ['.generated/hostile-src/canonical.xls', 'biff8-xls-renamed.xlsx'],
  ['.generated/hostile-src/canonical.encrypted.xlsx', 'encrypted-password-test.xlsx'],
]) {
  const from = join(FIXTURES, '..', source);
  if (existsSync(from)) {
    copyFileSync(from, join(HOSTILE, target));
    console.log(`${target.padEnd(46)} copied`);
  } else {
    console.log(`${target.padEnd(46)} SKIPPED (missing ${source})`);
  }
}
writeFileSync(
  join(HOSTILE, 'not-a-zip.xlsx'),
  Buffer.from(
    '<?xml version="1.0"?><Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"><Worksheet ss:Name="Sheet1"/></Workbook>',
    'utf8',
  ),
);
console.log('not-a-zip.xlsx (SpreadsheetML 2003 text) written');
