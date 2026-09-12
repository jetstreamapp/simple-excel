import { XlsxError } from '../errors';
import { notImplemented } from '../internal/not-implemented';
import type { BorderLineStyle, CellStyle, StyleId } from '../types';
import { escapeAttr } from '../xml/escape';
import { builtinIdForCode, DEFAULT_DATE_FORMAT, FIRST_CUSTOM_NUMFMT_ID, isBuiltinDateId, isDateFormatCode } from './numfmt';

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const SPREADSHEET_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

/** Excel's default font; every registered font derives from it so a workbook without a theme still looks right. */
const DEFAULT_FONT_XML = '<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>';
const DEFAULT_FONT_NAME = 'Calibri';
const DEFAULT_FONT_SIZE = 11;
/** Fills 0 and 1 are reserved: Excel repairs a workbook whose `<fills>` does not start with none + gray125. */
const RESERVED_FILLS_XML: readonly string[] = [
  '<fill><patternFill patternType="none"/></fill>',
  '<fill><patternFill patternType="gray125"/></fill>',
];
const EMPTY_BORDER_XML = '<border><left/><right/><top/><bottom/><diagonal/></border>';
const BORDER_SIDES = ['left', 'right', 'top', 'bottom'] as const;

const COLOR_HEX = /^[0-9A-F]{6}(?:[0-9A-F]{2})?$/;

type BorderSide = (typeof BORDER_SIDES)[number];
type BorderSides = { readonly [side in BorderSide]?: BorderLineStyle };

interface CellXf {
  readonly numFmtId: number;
  readonly fontId: number;
  readonly fillId: number;
  readonly borderId: number;
  readonly horizontal: string;
  readonly vertical: string;
  readonly wrapText: boolean;
  readonly isDate: boolean;
}

/** cellXfs[0] is the style a cell with no `s` attribute gets. */
const DEFAULT_XF: CellXf = { numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, horizontal: '', vertical: '', wrapText: false, isDate: false };

function xfKey(xf: CellXf): string {
  return `${xf.numFmtId}|${xf.fontId}|${xf.fillId}|${xf.borderId}|${xf.horizontal}|${xf.vertical}|${xf.wrapText ? 1 : 0}`;
}

/**
 * `#RRGGBB`, `RRGGBB` and `AARRGGBB` all become the `FFRRGGBB` ARGB Excel wants: alpha is ignored on read, so we
 * always write it opaque (primer 7.4).
 */
function normalizeColor(color: string): string {
  const hex = (color.startsWith('#') ? color.slice(1) : color).toUpperCase();
  if (!COLOR_HEX.test(hex)) {
    throw new XlsxError('WRITER_STATE', `"${color}" is not a colour. Use #RRGGBB, RRGGBB or AARRGGBB hex.`, { color });
  }
  return `FF${hex.slice(-6)}`;
}

/**
 * Growable style registry seeded with Excel's mandatory defaults (font 0 Calibri 11, fills none + gray125, border 0
 * empty, cellXfs 0). Identical `CellStyle` specs resolve to the same id. Custom number formats get ids from 164;
 * codes that match a built-in resolve to the built-in id and emit no `<numFmt>`.
 */
export class StyleRegistry {
  private readonly fonts: string[] = [DEFAULT_FONT_XML];
  private readonly fontIdByKey = new Map<string, number>();
  private readonly fills: string[] = [...RESERVED_FILLS_XML];
  private readonly fillIdByKey = new Map<string, number>();
  private readonly borders: string[] = [EMPTY_BORDER_XML];
  private readonly borderIdByKey = new Map<string, number>();
  /** Custom formats only, in the order they were registered - the ids Excel reserves for built-ins never appear. */
  private readonly customNumFmtCodeById = new Map<number, string>();
  private readonly numFmtIdByCode = new Map<string, number>();
  private readonly cellXfs: CellXf[] = [DEFAULT_XF];
  private readonly xfIdByKey = new Map<string, number>([[xfKey(DEFAULT_XF), 0]]);
  private dateStyleId: StyleId | undefined;
  private boldHeaderStyleId: StyleId | undefined;

  /** Number of cellXfs entries (default style included). */
  get count(): number {
    return this.cellXfs.length;
  }

  register(style: CellStyle): StyleId {
    const numFmtId = this.registerNumFmt(style.numFmt);
    const xf: CellXf = {
      numFmtId,
      fontId: this.registerFont(style.font),
      fillId: this.registerFill(style.fill),
      borderId: this.registerBorder(style.border),
      horizontal: style.alignment?.horizontal ?? '',
      vertical: style.alignment?.vertical ?? '',
      wrapText: style.alignment?.wrapText === true,
      isDate: this.isDateNumFmt(numFmtId),
    };

    const key = xfKey(xf);
    const existing = this.xfIdByKey.get(key);
    if (existing !== undefined) {
      return existing;
    }
    const styleId = this.cellXfs.length;
    this.cellXfs.push(xf);
    this.xfIdByKey.set(key, styleId);
    return styleId;
  }

  /** The style id the writer uses for JS Dates when the caller did not give one: default xf + DEFAULT_DATE_FORMAT. */
  get defaultDateStyle(): StyleId {
    this.dateStyleId ??= this.register({ numFmt: DEFAULT_DATE_FORMAT });
    return this.dateStyleId;
  }

  /** The bold header style. */
  get headerStyle(): StyleId {
    this.boldHeaderStyleId ??= this.register({ font: { bold: true } });
    return this.boldHeaderStyleId;
  }

  /** True when the xf carries a date/time number format (the writer needs this to serialize Dates under a caller style). */
  isDateStyle(styleId: StyleId): boolean {
    return this.cellXfs[styleId]?.isDate ?? false;
  }

  /** Complete `xl/styles.xml`, schema-valid (CT_Font child order, apply* attributes, cellStyles, dxfs, tableStyles). */
  toXml(): string {
    const parts: string[] = [XML_DECLARATION, `<styleSheet xmlns="${SPREADSHEET_NS}">`];
    if (this.customNumFmtCodeById.size > 0) {
      parts.push(`<numFmts count="${this.customNumFmtCodeById.size}">`);
      for (const [id, code] of this.customNumFmtCodeById) {
        parts.push(`<numFmt numFmtId="${id}" formatCode="${escapeAttr(code)}"/>`);
      }
      parts.push('</numFmts>');
    }
    parts.push(`<fonts count="${this.fonts.length}">`, ...this.fonts, '</fonts>');
    parts.push(`<fills count="${this.fills.length}">`, ...this.fills, '</fills>');
    parts.push(`<borders count="${this.borders.length}">`, ...this.borders, '</borders>');
    parts.push('<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>');
    parts.push(`<cellXfs count="${this.cellXfs.length}">`);
    for (const xf of this.cellXfs) {
      parts.push(cellXfXml(xf));
    }
    parts.push('</cellXfs>');
    parts.push('<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>');
    parts.push('<dxfs count="0"/>');
    parts.push('<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>');
    parts.push('</styleSheet>');
    return parts.join('');
  }

  private registerNumFmt(numFmt: CellStyle['numFmt']): number {
    if (numFmt === undefined) {
      return 0;
    }
    if (typeof numFmt === 'number') {
      // A cell pointing at a custom id with no `<numFmt>` element is a repair-dialog error, so refuse it here.
      if (numFmt >= FIRST_CUSTOM_NUMFMT_ID && !this.customNumFmtCodeById.has(numFmt)) {
        throw new XlsxError('WRITER_STATE', `Number format id ${numFmt} is not a built-in. Register the format code instead.`, { numFmt });
      }
      return numFmt;
    }
    const builtinId = builtinIdForCode(numFmt);
    if (builtinId >= 0) {
      return builtinId;
    }
    const existing = this.numFmtIdByCode.get(numFmt);
    if (existing !== undefined) {
      return existing;
    }
    const id = FIRST_CUSTOM_NUMFMT_ID + this.customNumFmtCodeById.size;
    this.customNumFmtCodeById.set(id, numFmt);
    this.numFmtIdByCode.set(numFmt, id);
    return id;
  }

  private isDateNumFmt(numFmtId: number): boolean {
    const customCode = this.customNumFmtCodeById.get(numFmtId);
    return customCode === undefined ? isBuiltinDateId(numFmtId) : isDateFormatCode(customCode);
  }

  private registerFont(font: CellStyle['font']): number {
    if (font === undefined) {
      return 0;
    }
    const bold = font.bold === true;
    const italic = font.italic === true;
    const underline = font.underline === true;
    const strike = font.strike === true;
    const size = font.size;
    const name = font.name;
    const color = font.color === undefined ? undefined : normalizeColor(font.color);
    if (!bold && !italic && !underline && !strike && size === undefined && name === undefined && color === undefined) {
      return 0;
    }

    const key = `${bold ? 1 : 0}|${italic ? 1 : 0}|${underline ? 1 : 0}|${strike ? 1 : 0}|${size ?? ''}|${color ?? ''}|${name ?? ''}`;
    const existing = this.fontIdByKey.get(key);
    if (existing !== undefined) {
      return existing;
    }
    // CT_Font children have a fixed schema order (EC-STYLES-FONT-ELEMENT-ORDER); out-of-order children fail the
    // Open XML SDK validator even though Excel opens them.
    let xml = '<font>';
    if (bold) {
      xml += '<b/>';
    }
    if (italic) {
      xml += '<i/>';
    }
    if (strike) {
      xml += '<strike/>';
    }
    if (underline) {
      xml += '<u/>';
    }
    xml += `<sz val="${size ?? DEFAULT_FONT_SIZE}"/>`;
    xml += color === undefined ? '<color theme="1"/>' : `<color rgb="${color}"/>`;
    xml += `<name val="${escapeAttr(name ?? DEFAULT_FONT_NAME)}"/><family val="2"/>`;
    if (name === undefined) {
      // `scheme` names the theme font the name came from, so it only holds for the default Calibri.
      xml += '<scheme val="minor"/>';
    }
    xml += '</font>';

    const fontId = this.fonts.length;
    this.fonts.push(xml);
    this.fontIdByKey.set(key, fontId);
    return fontId;
  }

  private registerFill(fill: CellStyle['fill']): number {
    if (fill === undefined) {
      return 0;
    }
    const color = normalizeColor(fill.color);
    const existing = this.fillIdByKey.get(color);
    if (existing !== undefined) {
      return existing;
    }
    const fillId = this.fills.length;
    this.fills.push(`<fill><patternFill patternType="solid"><fgColor rgb="${color}"/><bgColor indexed="64"/></patternFill></fill>`);
    this.fillIdByKey.set(color, fillId);
    return fillId;
  }

  private registerBorder(border: CellStyle['border']): number {
    if (border === undefined) {
      return 0;
    }
    const sides: BorderSides =
      typeof border === 'string'
        ? { left: border, right: border, top: border, bottom: border }
        : { left: border.left, right: border.right, top: border.top, bottom: border.bottom };
    const color = typeof border === 'string' || border.color === undefined ? undefined : normalizeColor(border.color);
    if (BORDER_SIDES.every(side => sides[side] === undefined)) {
      return 0;
    }

    const key = `${BORDER_SIDES.map(side => sides[side] ?? '').join('|')}|${color ?? ''}`;
    const existing = this.borderIdByKey.get(key);
    if (existing !== undefined) {
      return existing;
    }
    let xml = '<border>';
    for (const side of BORDER_SIDES) {
      const lineStyle = sides[side];
      if (lineStyle === undefined) {
        xml += `<${side}/>`;
      } else if (color === undefined) {
        // No colour means Excel's automatic border colour, which is the absent `<color>` child.
        xml += `<${side} style="${lineStyle}"/>`;
      } else {
        xml += `<${side} style="${lineStyle}"><color rgb="${color}"/></${side}>`;
      }
    }
    xml += '<diagonal/></border>';

    const borderId = this.borders.length;
    this.borders.push(xml);
    this.borderIdByKey.set(key, borderId);
    return borderId;
  }
}

function cellXfXml(xf: CellXf): string {
  let attributes = `numFmtId="${xf.numFmtId}" fontId="${xf.fontId}" fillId="${xf.fillId}" borderId="${xf.borderId}" xfId="0"`;
  if (xf.numFmtId !== 0) {
    attributes += ' applyNumberFormat="1"';
  }
  if (xf.fontId !== 0) {
    attributes += ' applyFont="1"';
  }
  if (xf.fillId !== 0) {
    attributes += ' applyFill="1"';
  }
  if (xf.borderId !== 0) {
    attributes += ' applyBorder="1"';
  }
  const hasAlignment = xf.horizontal !== '' || xf.vertical !== '' || xf.wrapText;
  if (!hasAlignment) {
    return `<xf ${attributes}/>`;
  }
  let alignment = '<alignment';
  if (xf.horizontal !== '') {
    alignment += ` horizontal="${xf.horizontal}"`;
  }
  if (xf.vertical !== '') {
    alignment += ` vertical="${xf.vertical}"`;
  }
  if (xf.wrapText) {
    alignment += ' wrapText="1"';
  }
  alignment += '/>';
  return `<xf ${attributes} applyAlignment="1">${alignment}</xf>`;
}

export interface ParsedStyles {
  /** `isDateByXf[i] === 1` when cellXfs[i] renders as a date/time. Length = cellXfs count. */
  readonly isDateByXf: Uint8Array;
  /** Custom number formats by id. */
  readonly numFmts: ReadonlyMap<number, string>;
}

/** Read what the reader needs from `xl/styles.xml`. Tolerates POI/LibreOffice/Numbers quirks (see catalog). */
export function parseStyles(xml: string): ParsedStyles {
  void xml;
  throw notImplemented('sml/styles');
}
