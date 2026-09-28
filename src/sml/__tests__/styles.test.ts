import { execFileSync } from 'node:child_process';
import { existsSync, mkdtempSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { XlsxError } from '../../errors';
import type { CellStyle } from '../../types';
import { DEFAULT_DATE_FORMAT } from '../numfmt';
import { appXml, contentTypesXml, coreXml, rootRelsXml, workbookRelsXml, workbookXml, type SheetPartInfo } from '../package-parts';
import { MAX_CELL_STYLES, StyleRegistry } from '../styles';

const DEFAULT_STYLES_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
  '<fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts>' +
  '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
  '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
  '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>' +
  '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
  '<dxfs count="0"/>' +
  '<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>' +
  '</styleSheet>';

/** Every `CellStyle` field at once - the registry has to place each component in its own collection. */
const EVERYTHING_STYLE: CellStyle = {
  font: { bold: true, italic: true, underline: true, strike: true, size: 14, color: '#FF0000', name: 'Arial' },
  fill: { color: '#EE00AA' },
  border: { top: 'thin', bottom: 'double', left: 'dashed', right: 'hair', color: '#00FF00' },
  alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
  numFmt: 'yyyy-mm-dd',
};

describe('StyleRegistry defaults', () => {
  it('emits Excel’s minimal styles.xml when nothing is registered', () => {
    expect(new StyleRegistry().toXml()).toBe(DEFAULT_STYLES_XML);
  });

  it('starts with one cellXf and no date styles', () => {
    const registry = new StyleRegistry();
    expect(registry.count).toBe(1);
    expect(registry.isDateStyle(0)).toBe(false);
    expect(registry.isDateStyle(7)).toBe(false);
  });

  it('resolves an empty style to the default xf', () => {
    const registry = new StyleRegistry();
    expect(registry.register({})).toBe(0);
    expect(registry.register({ font: {} })).toBe(0);
    expect(registry.register({ font: { bold: false }, alignment: { wrapText: false } })).toBe(0);
    expect(registry.register({ border: {} })).toBe(0);
    expect(registry.count).toBe(1);
    expect(registry.toXml()).toBe(DEFAULT_STYLES_XML);
  });

  it('registers the date and header presets lazily and once', () => {
    const registry = new StyleRegistry();
    expect(registry.count).toBe(1);
    expect(registry.defaultDateStyle).toBe(1);
    expect(registry.defaultDateStyle).toBe(1);
    expect(registry.headerStyle).toBe(2);
    expect(registry.headerStyle).toBe(2);
    expect(registry.count).toBe(3);
    expect(registry.isDateStyle(registry.defaultDateStyle)).toBe(true);
    expect(registry.register({ numFmt: DEFAULT_DATE_FORMAT })).toBe(registry.defaultDateStyle);
    expect(registry.register({ font: { bold: true } })).toBe(registry.headerStyle);
  });
});

describe('StyleRegistry deduplication', () => {
  it('returns the same id for the same style', () => {
    const registry = new StyleRegistry();
    const first = registry.register(EVERYTHING_STYLE);
    const second = registry.register(EVERYTHING_STYLE);
    expect(second).toBe(first);
    expect(registry.count).toBe(2);
  });

  it('ignores the order the style keys were written in', () => {
    const registry = new StyleRegistry();
    const first = registry.register({ font: { bold: true, size: 12 }, fill: { color: '#112233' }, numFmt: '0.00' });
    const second = registry.register({ numFmt: '0.00', fill: { color: '#112233' }, font: { size: 12, bold: true } });
    expect(second).toBe(first);
  });

  it('accepts 6-hex, 8-hex and #-prefixed colours as the same opaque ARGB', () => {
    const registry = new StyleRegistry();
    const hash = registry.register({ fill: { color: '#EE00AA' } });
    expect(registry.register({ fill: { color: 'ee00aa' } })).toBe(hash);
    expect(registry.register({ fill: { color: '80EE00AA' } })).toBe(hash);
    expect(registry.toXml()).toContain('<fgColor rgb="FFEE00AA"/>');
  });

  it('rejects a colour that is not hex', () => {
    const registry = new StyleRegistry();
    expect(() => registry.register({ fill: { color: 'rebeccapurple' } })).toThrow(XlsxError);
    expect(() => registry.register({ font: { color: '#FFF' } })).toThrow(XlsxError);
  });

  it('reuses one font, fill and border across different xfs', () => {
    const registry = new StyleRegistry();
    registry.register({ font: { bold: true }, numFmt: '0.00' });
    registry.register({ font: { bold: true }, numFmt: '0.000' });
    const xml = registry.toXml();
    expect(xml).toContain('<fonts count="2">');
    expect(registry.count).toBe(3);
  });
});

describe('StyleRegistry fragments', () => {
  it('emits CT_Font children in schema order (EC-STYLES-FONT-ELEMENT-ORDER)', () => {
    const registry = new StyleRegistry();
    registry.register(EVERYTHING_STYLE);
    expect(registry.toXml()).toContain(
      '<font><b/><i/><strike/><u/><sz val="14"/><color rgb="FFFF0000"/><name val="Arial"/><family val="2"/></font>',
    );
  });

  it('keeps the Calibri defaults for a font that only overrides one property', () => {
    const registry = new StyleRegistry();
    registry.register({ font: { bold: true } });
    expect(registry.toXml()).toContain(
      '<font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>',
    );
  });

  it('emits a solid fill with an indexed background', () => {
    const registry = new StyleRegistry();
    registry.register({ fill: { color: '#F3F3F3' } });
    expect(registry.toXml()).toContain(
      '<fill><patternFill patternType="solid"><fgColor rgb="FFF3F3F3"/><bgColor indexed="64"/></patternFill></fill>',
    );
  });

  it('emits border sides in schema order, with the diagonal always present', () => {
    const registry = new StyleRegistry();
    registry.register({ border: 'thin' });
    registry.register({ border: { top: 'medium', color: '#00FF00' } });
    const xml = registry.toXml();
    expect(xml).toContain('<border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/><diagonal/></border>');
    expect(xml).toContain('<border><left/><right/><top style="medium"><color rgb="FF00FF00"/></top><bottom/><diagonal/></border>');
  });

  it('flags every assigned component with its apply attribute and nests the alignment', () => {
    const registry = new StyleRegistry();
    const styleId = registry.register(EVERYTHING_STYLE);
    expect(registry.toXml()).toContain(
      `<xf numFmtId="164" fontId="1" fillId="2" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"` +
        ` applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>`,
    );
    expect(styleId).toBe(1);
  });

  it('emits only the alignment attributes that were set', () => {
    const registry = new StyleRegistry();
    registry.register({ alignment: { wrapText: true } });
    expect(registry.toXml()).toContain(
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1"/></xf>',
    );
  });
});

describe('StyleRegistry number formats', () => {
  it('resolves built-in codes to their id and emits no numFmts', () => {
    const registry = new StyleRegistry();
    registry.register({ numFmt: '0.00' });
    registry.register({ numFmt: 14 });
    const xml = registry.toXml();
    expect(xml).not.toContain('<numFmts');
    expect(xml).toContain('<xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>');
    expect(xml).toContain('<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>');
  });

  it('allocates custom ids from 164 and writes them first, attribute-escaped', () => {
    const registry = new StyleRegistry();
    registry.register({ numFmt: 'yyyy-mm-dd' });
    registry.register({ numFmt: '"USD" #,##0.00' });
    registry.register({ numFmt: 'yyyy-mm-dd' });
    expect(registry.toXml()).toContain(
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="2">' +
        '<numFmt numFmtId="164" formatCode="yyyy-mm-dd"/><numFmt numFmtId="165" formatCode="&quot;USD&quot; #,##0.00"/></numFmts>',
    );
  });

  it('refuses a custom id it has no format code for', () => {
    const registry = new StyleRegistry();
    expect(() => registry.register({ numFmt: 200 })).toThrow(XlsxError);
    // ... unless the code behind that id has been registered.
    expect(registry.register({ numFmt: 'yyyy-mm-dd' })).toBe(1);
    expect(registry.register({ numFmt: 164, font: { bold: true } })).toBe(2);
  });

  it('knows which xfs render as dates (EC-DATE-DETECTION-VIA-NUMFMT)', () => {
    const registry = new StyleRegistry();
    expect(registry.isDateStyle(registry.register({ numFmt: 'yyyy-mm-dd' }))).toBe(true);
    expect(registry.isDateStyle(registry.register({ numFmt: 14 }))).toBe(true);
    expect(registry.isDateStyle(registry.register({ numFmt: 21 }))).toBe(true);
    expect(registry.isDateStyle(registry.register({ numFmt: '0%' }))).toBe(false);
    expect(registry.isDateStyle(registry.register({ numFmt: '"months" 0' }))).toBe(false);
    expect(registry.isDateStyle(registry.register({ font: { bold: true } }))).toBe(false);
  });
});

describe('StyleRegistry validation (EC-STYLE-FIELD-RANGE)', () => {
  /** The field an invalid style is refused for, or 'accepted'. */
  function refusedField(style: unknown): string {
    try {
      new StyleRegistry().register(style as CellStyle);
    } catch (error) {
      if (!(error instanceof XlsxError) || error.code !== 'WRITER_STATE') {
        return `unexpected ${String(error)}`;
      }
      return String(error.detail?.field);
    }
    return 'accepted';
  }

  it('accepts every field at the edges of its range', () => {
    expect(refusedField({ font: { size: 1, name: 'A' } })).toBe('accepted');
    expect(refusedField({ font: { size: 409, name: 'x'.repeat(31) } })).toBe('accepted');
    expect(refusedField({ font: { size: 10.5 } })).toBe('accepted');
    expect(refusedField({ numFmt: 0 })).toBe('accepted');
    expect(refusedField({ numFmt: 163 })).toBe('accepted');
    expect(refusedField({ numFmt: '0'.repeat(255) })).toBe('accepted');
    expect(refusedField({ alignment: { horizontal: 'right', vertical: 'bottom' } })).toBe('accepted');
    expect(refusedField({ border: { left: 'hair', right: 'double', top: 'dotted', bottom: 'thick' } })).toBe('accepted');
  });

  it('refuses a font size outside 1-409 points', () => {
    for (const size of [0, 0.5, 410, -3, Number.NaN, Number.POSITIVE_INFINITY, '12']) {
      expect(refusedField({ font: { size } }), String(size)).toBe('font.size');
    }
  });

  it('refuses a number format code or font name that holds a control character or would be written empty', () => {
    for (const numFmt of ['\u0007', '0.00\u0007', '\uD800', '\uDC00\uD800']) {
      expect(refusedField({ numFmt }), JSON.stringify(numFmt)).toBe('numFmt');
    }
    expect(refusedField({ font: { name: '\uD800' } })).toBe('font.name');
  });

  it('refuses a font name that is empty, longer than 31 characters, not text or holds a control character', () => {
    for (const name of ['', 'x'.repeat(32), 42, '\u0001', 'Arial\u0007', 'Tab\tName']) {
      expect(refusedField({ font: { name } }), String(name)).toBe('font.name');
    }
  });

  it('refuses colours that are not hex, including non-strings, on every coloured field', () => {
    expect(refusedField({ font: { color: 'red' } })).toBe('font.color');
    expect(refusedField({ font: { color: 0xff0000 } })).toBe('font.color');
    expect(refusedField({ fill: { color: '#12345' } })).toBe('fill.color');
    expect(refusedField({ border: { top: 'thin', color: 'blue' } })).toBe('border.color');
  });

  it('refuses a fill without a colour instead of throwing a TypeError', () => {
    expect(refusedField({ fill: {} })).toBe('fill.color');
    expect(refusedField({ fill: '#FF0000' })).toBe('fill.color');
    expect(refusedField({ fill: null })).toBe('fill.color');
  });

  it('refuses alignments and border styles outside their enums', () => {
    expect(refusedField({ alignment: { horizontal: 'justify' } })).toBe('alignment.horizontal');
    expect(refusedField({ alignment: { vertical: 'middle' } })).toBe('alignment.vertical');
    expect(refusedField({ border: 'bold' })).toBe('border');
    expect(refusedField({ border: { left: 'thin', bottom: 'wavy' } })).toBe('border.bottom');
    expect(refusedField({ border: 7 })).toBe('border');
  });

  it('refuses number formats that are empty, too long, or not a built-in or registered id', () => {
    for (const numFmt of ['', '0'.repeat(256)]) {
      expect(refusedField({ numFmt }), `code of ${numFmt.length}`).toBe('numFmt');
    }
    for (const numFmt of [-1, 1.5, Number.NaN, 164, 200, true]) {
      expect(refusedField({ numFmt }), String(numFmt)).toBe('numFmt');
    }
  });

  it('refuses a style that is not an object', () => {
    expect(refusedField(null)).toBe('style');
    expect(refusedField('bold')).toBe('style');
    expect(refusedField({ font: 'bold' })).toBe('font');
    expect(refusedField({ alignment: 'center' })).toBe('alignment');
  });

  it('names the field and the value in the message', () => {
    expect(() => new StyleRegistry().register({ font: { size: 500 } })).toThrow(/font\.size 500 .*1 to 409/);
    expect(() => new StyleRegistry().register({ alignment: { horizontal: 'justify' as 'left' } })).toThrow(
      /alignment\.horizontal "justify"/,
    );
  });

  it('leaves nothing behind when a style is refused', () => {
    const registry = new StyleRegistry();
    expect(() => registry.register({ font: { bold: true }, numFmt: '0.000', fill: {} as { color: string } })).toThrow(XlsxError);
    expect(() => registry.register({ numFmt: '0.0000', font: { italic: true, color: 'red' } })).toThrow(XlsxError);
    expect(() => registry.register({ numFmt: '0.00000', fill: { color: '#00FF00' }, border: { top: 'thin', color: 'nope' } })).toThrow(
      XlsxError,
    );
    expect(registry.toXml()).toBe(DEFAULT_STYLES_XML);
  });

  it(`stops at ${MAX_CELL_STYLES} cell styles but still resolves the ones it has`, () => {
    const registry = new StyleRegistry();
    for (let index = 1; registry.count < MAX_CELL_STYLES; index++) {
      registry.register({ numFmt: `0.${'0'.repeat(index % 200)}"${index}"` });
    }
    expect(registry.count).toBe(MAX_CELL_STYLES);
    expect(() => registry.register({ font: { bold: true } })).toThrowError(expect.objectContaining({ code: 'WRITER_STATE' }));
    expect(registry.register({ numFmt: '0.0"1"' }), 'an existing style still resolves').toBe(1);
  });
});

describe('StyleRegistry date styles (EC-DATE-STYLE-MERGE)', () => {
  it('derives a date style that keeps every other component of the caller style, once', () => {
    const registry = new StyleRegistry();
    const styleId = registry.register({ font: { bold: true }, fill: { color: '#EE00AA' }, border: 'thin', alignment: { wrapText: true } });
    const derived = registry.dateStyleFor(styleId);
    expect(registry.dateStyleFor(styleId)).toBe(derived);
    expect(registry.isDateStyle(derived)).toBe(true);
    expect(registry.count).toBe(3);
    expect(registry.toXml()).toContain(
      '<xf numFmtId="164" fontId="1" fillId="2" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"' +
        ' applyAlignment="1"><alignment wrapText="1"/></xf>',
    );
  });

  it('keeps a style that already renders as a date and maps style 0 to the default date style', () => {
    const registry = new StyleRegistry();
    const dated = registry.register({ numFmt: 'dd/mm/yyyy', font: { bold: true } });
    expect(registry.dateStyleFor(dated)).toBe(dated);
    expect(registry.dateStyleFor(0)).toBe(registry.defaultDateStyle);
  });
});

describe('metadata escaping (EC-XML-CONTROL-CHARS-METADATA)', () => {
  it('drops unpaired surrogates from number format codes and font names', () => {
    const registry = new StyleRegistry();
    registry.register({ font: { name: 'Ari\uD800al' }, numFmt: '0.00" units\uD800"' });
    const xml = registry.toXml();
    expect(xml).toContain('<name val="Arial"/>');
    expect(xml).toContain('formatCode="0.00&quot; units&quot;"');
    // eslint-disable-next-line no-control-regex -- the characters XML forbids are exactly what must not be there
    expect(xml).not.toMatch(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/);
  });
});

// ---- Open XML SDK validator ---------------------------------------------------------------------------------------

const VALIDATOR = join(process.cwd(), 'node_modules', '.bin', 'ooxml-validator');
const MINIZIP = new URL('../../../fixtures/generators/edge/minizip.mjs', import.meta.url).href;

interface ZipEntry {
  name: string;
  data: string;
}

async function buildZip(entries: ZipEntry[]): Promise<Uint8Array> {
  const { buildZip: build } = (await import(MINIZIP)) as { buildZip: (entries: ZipEntry[]) => Uint8Array };
  return build(entries);
}

function sheetXml(styleIds: readonly number[]): string {
  const cells = styleIds.map((styleId, index) => `<c r="${String.fromCodePoint(65 + index)}1" s="${styleId}"><v>1</v></c>`).join('');
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    `<sheetData><row r="1">${cells}</row></sheetData><autoFilter ref="A1:C1"/></worksheet>`
  );
}

/** A whole package built from the writers in this module pair, so the validator judges both at once. */
async function buildPackage(styles: StyleRegistry, styleIds: readonly number[]): Promise<Uint8Array> {
  const sheets: SheetPartInfo[] = [
    { name: "Data & 'more'", sheetId: 1, relId: 'rId1', path: 'worksheets/sheet1.xml', hidden: false, autoFilterRange: 'A1:C1' },
    { name: 'Hidden <sheet>', sheetId: 2, relId: 'rId2', path: 'worksheets/sheet2.xml', hidden: true },
  ];
  return buildZip([
    { name: '[Content_Types].xml', data: contentTypesXml(sheets, true) },
    { name: '_rels/.rels', data: rootRelsXml() },
    { name: 'docProps/core.xml', data: coreXml({ title: 'Validator probe' }, new Date('2026-01-01T00:00:00Z')) },
    { name: 'docProps/app.xml', data: appXml(sheets) },
    { name: 'xl/workbook.xml', data: workbookXml(sheets, false) },
    { name: 'xl/_rels/workbook.xml.rels', data: workbookRelsXml(sheets, true) },
    { name: 'xl/worksheets/sheet1.xml', data: sheetXml(styleIds) },
    { name: 'xl/worksheets/sheet2.xml', data: sheetXml([0]) },
    { name: 'xl/styles.xml', data: styles.toXml() },
    {
      name: 'xl/sharedStrings.xml',
      data:
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"/>',
    },
  ]);
}

describe('Open XML SDK validator', () => {
  it.skipIf(!existsSync(VALIDATOR))('accepts a package built from these writers', async () => {
    const styles = new StyleRegistry();
    const styleIds = [
      styles.defaultDateStyle,
      styles.headerStyle,
      styles.register(EVERYTHING_STYLE),
      styles.register({ border: 'thin' }),
      styles.register({ numFmt: 14 }),
      styles.register({ font: { italic: true, size: 8 }, alignment: { horizontal: 'right' } }),
    ];
    const file = join(mkdtempSync(join(tmpdir(), 'simple-excel-styles-')), 'styles.xlsx');
    writeFileSync(file, await buildPackage(styles, styleIds));

    const report = JSON.parse(execFileSync(VALIDATOR, [file], { encoding: 'utf8' })) as { ok: boolean; errors: unknown[] };
    expect(report.errors).toEqual([]);
    expect(report.ok).toBe(true);
  });
});
