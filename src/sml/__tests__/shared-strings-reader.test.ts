import { describe, expect, it } from 'vitest';
import { allFixtures, fixtureById, readFixture, type Fixture } from '../../../test/helpers/fixtures';
import { sourceFrom } from '../../zip/source';
import { ZipReader } from '../../zip/zip-reader';
import { parseRels, relTypeIs } from '../package-parts';
import { parseSharedStrings, type ParseSharedStringsOptions } from '../shared-strings-reader';

const SPREADSHEET_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const OPTIONS: ParseSharedStringsOptions = { maxChars: 1 << 20, maxDepth: 256, maxTextLength: 1 << 20 };

function sst(body: string, attributes = ''): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="${SPREADSHEET_NS}"${attributes}>${body}</sst>`;
}

async function* chunksOf(text: string, chunkSize = 64 * 1024): AsyncGenerator<Uint8Array> {
  const bytes = new TextEncoder().encode(text);
  for (let offset = 0; offset < bytes.length; offset += chunkSize) {
    yield bytes.subarray(offset, offset + chunkSize);
  }
}

function parse(xml: string, options: ParseSharedStringsOptions = OPTIONS): Promise<string[]> {
  return parseSharedStrings(chunksOf(xml), options);
}

describe('parseSharedStrings', () => {
  it('reads one entry per si, in order', async () => {
    await expect(parse(sst('<si><t>Name</t></si><si><t>Amount</t></si>', ' count="2" uniqueCount="2"'))).resolves.toEqual([
      'Name',
      'Amount',
    ]);
  });

  it('EC-RICH-TEXT-RUNS: concatenates runs and ignores their formatting', async () => {
    const xml = sst(
      '<si><r><rPr><rFont val="Arial"/><b/><color theme="1"/></rPr><t xml:space="preserve">1. </t></r>' +
        '<r><rPr><rFont val="Arial"/></rPr><t>Create one worksheet</t></r></si>',
    );
    await expect(parse(xml)).resolves.toEqual(['1. Create one worksheet']);
  });

  it('skips rPh phonetic runs so furigana never leaks into the value', async () => {
    const xml = sst(
      '<si><t>関西</t><rPh sb="0" eb="2"><t>カンサイ</t></rPh><phoneticPr fontId="1"/></si>' +
        '<si><r><t>東京</t></r><rPh sb="0" eb="2"><t>トウキョウ</t></rPh></si>',
    );
    await expect(parse(xml)).resolves.toEqual(['関西', '東京']);
  });

  it('decodes _xHHHH_ escapes, including the defused literal', async () => {
    const xml = sst('<si><t>cr_x000D_here</t></si><si><t>_x005F_x0041_</t></si><si><t>_X0041_</t></si><si><t>plain_0041_</t></si>');
    // `_X0041_` is upper case, which Excel does not treat as an escape, and `_0041_` is not one at all.
    await expect(parse(xml)).resolves.toEqual(['cr\rhere', '_x0041_', '_X0041_', 'plain_0041_']);
  });

  it('keeps whitespace-only text verbatim and reads an empty si as an empty string', async () => {
    const xml = sst('<si><t xml:space="preserve">  padded  </t></si><si><t>   </t></si><si/><si><t/></si><si><r><t></t></r></si>');
    await expect(parse(xml)).resolves.toEqual(['  padded  ', '   ', '', '', '']);
  });

  it('decodes entities, numeric references and CDATA', async () => {
    const xml = sst('<si><t>a &amp; b &lt;c&gt; &#38; &#x1F600;</t></si><si><t><![CDATA[raw & < text]]></t></si>');
    await expect(parse(xml)).resolves.toEqual(['a & b <c> & 😀', 'raw & < text']);
  });

  it('ignores whitespace between elements and text outside t', async () => {
    const xml = sst('\n  <si>\n    <t>Name</t>\n  </si>\n  <si>\n    <r>\n      <t>Amount</t>\n    </r>\n  </si>\n');
    await expect(parse(xml)).resolves.toEqual(['Name', 'Amount']);
  });

  it('returns the real entry count whatever uniqueCount claims', async () => {
    const body = '<si><t>a</t></si><si><t>b</t></si><si><t>c</t></si>';
    await expect(parse(sst(body, ' uniqueCount="3"'))).resolves.toEqual(['a', 'b', 'c']);
    await expect(parse(sst(body, ' uniqueCount="1"'))).resolves.toEqual(['a', 'b', 'c']);
    await expect(parse(sst(body, ' uniqueCount="900000000000"'))).resolves.toEqual(['a', 'b', 'c']);
    await expect(parse(sst(body, ' count="3"'))).resolves.toEqual(['a', 'b', 'c']);
    await expect(parse(sst('', ' count="0" uniqueCount="0"'))).resolves.toEqual([]);
  });

  it('reads the same table however the bytes are chunked', async () => {
    const xml = sst('<si><t>Zoë Ångström</t></si><si><r><t>a</t></r><r><t>_x000D_</t></r></si><si><t>😀</t></si>');
    const expected = ['Zoë Ångström', 'a\r', '😀'];
    for (const chunkSize of [1, 2, 3, 5, 7, 13, 64]) {
      await expect(parseSharedStrings(chunksOf(xml, chunkSize), OPTIONS), `chunk size ${chunkSize}`).resolves.toEqual(expected);
    }
  });

  it('EC-STRICT-NAMESPACES and EC-PREFIXED-ELEMENTS: namespace shape does not matter', async () => {
    const strict = '<?xml version="1.0"?><sst xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"><si><t>Zoë</t></si></sst>';
    const prefixed = `<x:sst xmlns:x="${SPREADSHEET_NS}"><x:si><x:r><x:t>Zoë</x:t></x:r></x:si></x:sst>`;
    await expect(parse(strict)).resolves.toEqual(['Zoë']);
    await expect(parse(prefixed)).resolves.toEqual(['Zoë']);
  });

  it('throws LIMIT_EXCEEDED past maxChars', async () => {
    const xml = sst('<si><t>abcdefghij</t></si><si><t>klmnopqrst</t></si>');
    await expect(parse(xml, { ...OPTIONS, maxChars: 20 })).resolves.toHaveLength(2);
    await expect(parse(xml, { ...OPTIONS, maxChars: 15 })).rejects.toMatchObject({ name: 'XlsxError', code: 'LIMIT_EXCEEDED' });
  });

  it('EC-XXE-DOCTYPE: refuses a document type declaration', async () => {
    const xml =
      '<?xml version="1.0"?><!DOCTYPE sst [<!ENTITY xxe SYSTEM "file:///etc/passwd">]>' +
      `<sst xmlns="${SPREADSHEET_NS}"><si><t>&xxe;</t></si></sst>`;
    await expect(parse(xml)).rejects.toMatchObject({ name: 'XlsxError', code: 'XML_DOCTYPE' });
  });

  it('EC-XML-DEEP-NESTING: refuses nesting past maxDepth', async () => {
    const depth = 300;
    const xml = sst(`<si>${'<r>'.repeat(depth)}<t>deep</t>${'</r>'.repeat(depth)}</si>`);
    await expect(parse(xml, { ...OPTIONS, maxDepth: 256 })).rejects.toMatchObject({ name: 'XlsxError', code: 'LIMIT_EXCEEDED' });
  });
});

// ---- the corpus ----------------------------------------------------------------------------------------------------

function folderOf(partName: string): string {
  const lastSlash = partName.lastIndexOf('/');
  return lastSlash < 0 ? '' : partName.slice(0, lastSlash + 1);
}

function relsPathFor(partName: string): string {
  const folder = folderOf(partName);
  return `${folder}_rels/${partName.slice(folder.length)}.rels`;
}

async function openFixture(fixture: Fixture): Promise<ZipReader> {
  return ZipReader.open(sourceFrom(readFixture(fixture)), { maxEntries: 10_000, maxInflatedBytes: 1 << 30 });
}

/** The shared-strings part named by the workbook rels, or undefined for the inline-only fixtures. */
async function sharedStringsPart(zip: ZipReader): Promise<string | undefined> {
  const rootRels = parseRels(await zip.readText('_rels/.rels'), '');
  const workbookPart = rootRels.find(relationship => relTypeIs(relationship.type, 'officeDocument'))?.target ?? 'xl/workbook.xml';
  const workbookRels = parseRels(await zip.readText(relsPathFor(workbookPart)), folderOf(workbookPart));
  return workbookRels.find(relationship => relTypeIs(relationship.type, 'sharedStrings'))?.target;
}

const corpusFixtures = allFixtures().filter(fixture => !fixture.tags.includes('kind:hostile') && fixture.path.endsWith('.xlsx'));

describe('every corpus shared-string table parses', () => {
  it.each(corpusFixtures.map(fixture => [fixture.id, fixture] as const))('%s', async (_id, fixture) => {
    const zip = await openFixture(fixture);
    const part = await sharedStringsPart(zip);
    if (part === undefined) {
      // EC-SST-ABSENT-INLINE-ONLY: a workbook whose strings are all inline has no table at all - and no part we
      // failed to find either, which is what this asserts.
      expect([...zip.entries.keys()].filter(name => name.toLowerCase().includes('sharedstrings'))).toEqual([]);
      await zip.close();
      return;
    }

    const strings = await parseSharedStrings(zip.stream(part), OPTIONS);
    expect(strings.every(value => typeof value === 'string')).toBe(true);
    if (fixture.expected === 'canonical/canonical.json') {
      expect(strings, 'the canonical Name column must survive the read').toContain('Zoë Ångström');
    }
    await zip.close();
  });
});

describe('hostile shared-string tables', () => {
  it('EC-XXE-DOCTYPE: xxe-doctype-in-sharedstrings throws XML_DOCTYPE', async () => {
    const zip = await openFixture(fixtureById('hostile-xxe-doctype-in-sharedstrings'));
    await expect(parseSharedStrings(zip.stream('xl/sharedStrings.xml'), OPTIONS)).rejects.toMatchObject({
      name: 'XlsxError',
      code: 'XML_DOCTYPE',
    });
    await zip.close();
  });

  it('EC-XML-DEEP-NESTING: deeply-nested-rich-text throws LIMIT_EXCEEDED', async () => {
    const zip = await openFixture(fixtureById('hostile-deeply-nested-rich-text'));
    await expect(parseSharedStrings(zip.stream('xl/sharedStrings.xml'), { ...OPTIONS, maxDepth: 256 })).rejects.toMatchObject({
      name: 'XlsxError',
      code: 'LIMIT_EXCEEDED',
    });
    await zip.close();
  });
});
