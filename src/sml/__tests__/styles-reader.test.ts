import { describe, expect, it } from 'vitest';
import { allFixtures, readFixture } from '../../../test/helpers/fixtures';
import { XmlTokenizer } from '../../xml/tokenizer';
import { sourceFrom } from '../../zip/source';
import { ZipReader } from '../../zip/zip-reader';
import { parseRels, parseWorkbook, relTypeIs } from '../package-parts';
import { parseStyles } from '../styles';

const SPREADSHEET_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

function styleSheet(body: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="${SPREADSHEET_NS}">${body}</styleSheet>`;
}

function cellXfs(...numFmtIds: readonly number[]): string {
  return `<cellXfs count="${numFmtIds.length}">${numFmtIds
    .map(id => `<xf numFmtId="${id}" fontId="0" fillId="0" borderId="0" xfId="0"/>`)
    .join('')}</cellXfs>`;
}

function dateFlags(xml: string): number[] {
  return [...parseStyles(xml).isDateByXf];
}

describe('parseStyles', () => {
  it('EC-DATE-DETECTION-VIA-NUMFMT: flags built-in date ids and nothing else', () => {
    expect(dateFlags(styleSheet(cellXfs(0, 1, 2, 9, 49, 14, 22, 45, 47, 56)))).toEqual([0, 0, 0, 0, 0, 1, 1, 1, 1, 1]);
  });

  it('flags a custom format whose code renders a date', () => {
    const xml = styleSheet(
      '<numFmts count="3"><numFmt numFmtId="164" formatCode="yyyy-mm-dd hh:mm:ss"/>' +
        '<numFmt numFmtId="165" formatCode="&quot;USD&quot; #,##0.00"/>' +
        '<numFmt numFmtId="166" formatCode="[h]:mm:ss"/></numFmts>' +
        cellXfs(164, 165, 166),
    );
    expect(dateFlags(xml)).toEqual([1, 0, 1]);
    expect(parseStyles(xml).numFmts.get(164)).toBe('yyyy-mm-dd hh:mm:ss');
    // The format code is an attribute, so the tokenizer has already decoded `&quot;`.
    expect(parseStyles(xml).numFmts.get(165)).toBe('"USD" #,##0.00');
  });

  it('accepts a custom format inside the reserved id range (write-excel-file writes 100)', () => {
    expect(
      dateFlags(styleSheet('<numFmts count="1"><numFmt numFmtId="100" formatCode="mm/dd/yyyy"/></numFmts>' + cellXfs(0, 100))),
    ).toEqual([0, 1]);
  });

  it('lets an explicit format code win over the built-in table', () => {
    // Id 14 is the built-in `mm-dd-yy`, but a producer that redefines it means what it wrote.
    expect(dateFlags(styleSheet('<numFmts count="1"><numFmt numFmtId="14" formatCode="0.00"/></numFmts>' + cellXfs(14)))).toEqual([0]);
  });

  it('reads numFmts declared after cellXfs (out of schema order)', () => {
    expect(dateFlags(styleSheet(cellXfs(164) + '<numFmts count="1"><numFmt numFmtId="164" formatCode="d-mmm-yy"/></numFmts>'))).toEqual([
      1,
    ]);
  });

  it('only indexes cellXfs: cellStyleXfs and dxfs entries never shift the indices', () => {
    const xml = styleSheet(
      '<cellStyleXfs count="2"><xf numFmtId="0"/><xf numFmtId="14"/></cellStyleXfs>' +
        cellXfs(0, 14) +
        '<dxfs count="1"><dxf><numFmt numFmtId="170" formatCode="yyyy"/><font><b/></font></dxf></dxfs>',
    );
    expect(dateFlags(xml)).toEqual([0, 1]);
    expect(parseStyles(xml).numFmts.has(170)).toBe(false);
  });

  it('returns an empty array when there is no cellXfs element', () => {
    const parsed = parseStyles(styleSheet('<fonts count="1"><font><sz val="11"/></font></fonts>'));
    expect(parsed.isDateByXf).toEqual(new Uint8Array(0));
    expect(parsed.numFmts.size).toBe(0);
  });

  it('treats an xf with no numFmtId as the General format (write-excel-file writes `<xf ></xf>`)', () => {
    expect(dateFlags(styleSheet('<cellXfs count="2"><xf ></xf><xf numFmtId="14"/></cellXfs>'))).toEqual([0, 1]);
  });

  it('ignores applyNumberFormat: Excel renders the xf numFmtId either way', () => {
    expect(
      dateFlags(
        styleSheet('<cellXfs count="2"><xf numFmtId="14" applyNumberFormat="0"/><xf numFmtId="0" applyNumberFormat="1"/></cellXfs>'),
      ),
    ).toEqual([1, 0]);
  });

  it('EC-POI-UNDERLINE-NONE-REJECTED, EC-POI-RGB-6-HEX: reads a Salesforce report export shape', () => {
    const xml = styleSheet(
      '<numFmts count="2"><numFmt numFmtId="164" formatCode="&quot;USD&quot; #,##0.00;&quot;USD&quot; -#,##0.00"/>' +
        '<numFmt numFmtId="165" formatCode="#,##0%"/></numFmts>' +
        '<fonts count="2"><font><sz val="11.0"/><color indexed="8"/><name val="Calibri"/></font>' +
        '<font><name val="Calibri"/><sz val="12.0"/><color rgb="FF56585B"/><b val="true"/><u val="none"/></font></fonts>' +
        '<fills count="2"><fill><patternFill patternType="none"/></fill>' +
        '<fill><patternFill patternType="solid"><fgColor rgb="E9E8E5"/></patternFill></fill></fills>' +
        '<borders count="2"><border><left/></border><border><right style="thin"><color rgb="8E9297"/></right></border></borders>' +
        cellXfs(0, 164, 165, 14),
    );
    expect(dateFlags(xml)).toEqual([0, 0, 0, 1]);
  });

  it('EC-STRICT-NAMESPACES: reads a Strict styles part', () => {
    const xml =
      '<?xml version="1.0"?><styleSheet xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main">' + cellXfs(0, 14) + '</styleSheet>';
    expect(dateFlags(xml)).toEqual([0, 1]);
  });

  it('EC-PREFIXED-ELEMENTS: reads a prefixed styles part and ignores mc:Ignorable / x14ac noise', () => {
    const xml =
      `<x:styleSheet xmlns:x="${SPREADSHEET_NS}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"` +
      ' xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac x16r2">' +
      '<x:numFmts count="1"><x:numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></x:numFmts>' +
      '<x:cellXfs count="2"><x:xf numFmtId="0" x14ac:dyDescent="0.25"/><x:xf numFmtId="164"/></x:cellXfs>' +
      '<x:extLst><x:ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}"><x14:slicerStyles xmlns:x14="http://x14" defaultSlicerStyle="S"/></x:ext></x:extLst>' +
      '</x:styleSheet>';
    expect(dateFlags(xml)).toEqual([0, 1]);
  });

  it('keeps alignment and protection children from disturbing the xf order', () => {
    const xml = styleSheet(
      '<cellXfs count="3"><xf numFmtId="0"><alignment horizontal="center"/></xf>' +
        '<xf numFmtId="14" applyNumberFormat="1"><alignment wrapText="1"/><protection locked="0"/></xf>' +
        '<xf numFmtId="0"/></cellXfs>',
    );
    expect(dateFlags(xml)).toEqual([0, 1, 0]);
  });
});

// ---- the corpus ----------------------------------------------------------------------------------------------------

const corpusFixtures = allFixtures().filter(fixture => !fixture.tags.includes('kind:hostile') && fixture.path.endsWith('.xlsx'));

/** The `s` attribute of one cell, found by streaming the sheet rather than materializing it. */
function styleIndexOfCell(sheetXml: string, cellRef: string): number {
  let styleIndex = 0;
  const tokenizer = new XmlTokenizer({
    start(name: string): void {
      if (name === 'c' && tokenizer.attr('r') === cellRef) {
        styleIndex = Number.parseInt(tokenizer.attr('s') ?? '0', 10);
      }
    },
    text(): void {},
    end(): void {},
  });
  tokenizer.push(sheetXml);
  tokenizer.end();
  return styleIndex;
}

function folderOf(partName: string): string {
  const lastSlash = partName.lastIndexOf('/');
  return lastSlash < 0 ? '' : partName.slice(0, lastSlash + 1);
}

/** A part's own relationships live in `<folder>/_rels/<file>.rels`. */
function relsPathFor(partName: string): string {
  const folder = folderOf(partName);
  return `${folder}_rels/${partName.slice(folder.length)}.rels`;
}

describe('every corpus styles part parses', () => {
  it.each(corpusFixtures.map(fixture => [fixture.id, fixture] as const))('%s', async (_id, fixture) => {
    const zip = await ZipReader.open(sourceFrom(readFixture(fixture)), { maxEntries: 10_000, maxInflatedBytes: 1 << 30 });
    const rootRels = parseRels(await zip.readText('_rels/.rels'), '');
    const workbookPart = rootRels.find(relationship => relTypeIs(relationship.type, 'officeDocument'))?.target ?? 'xl/workbook.xml';
    const workbookRels = parseRels(await zip.readText(relsPathFor(workbookPart)), folderOf(workbookPart));

    const stylesPart = workbookRels.find(relationship => relTypeIs(relationship.type, 'styles'))?.target;
    expect(stylesPart, 'every corpus fixture has a styles part').toBeDefined();
    const styles = parseStyles(await zip.readText(stylesPart ?? ''));
    expect(styles.isDateByXf.length).toBeGreaterThan(0);

    // Every hand-built edge fixture shares one styles.xml: the default xf plus a `mm-dd-yy` one for the date column.
    if (fixture.tags.includes('kind:edge')) {
      expect([...styles.isDateByXf]).toEqual([0, 1]);
    }

    // The canonical Data sheet's M column is `Date`, so whatever xf M2 points at must read as a date.
    if (fixture.expected === 'canonical/canonical.json') {
      const firstSheet = parseWorkbook(await zip.readText(workbookPart)).sheets[0];
      const sheetPart = workbookRels.find(relationship => relationship.id === firstSheet?.relId)?.target ?? '';
      const styleIndex = styleIndexOfCell(await zip.readText(sheetPart), 'M2');
      expect(styles.isDateByXf[styleIndex], `xf ${styleIndex} (cell M2) must render as a date`).toBe(1);
    }
    await zip.close();
  });
});
