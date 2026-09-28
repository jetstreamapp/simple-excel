import { XlsxError } from '../errors';
import type { BorderLineStyle, CellStyle, StyleId } from '../types';
import { escapeAttr } from '../xml/escape';
import { XmlTokenizer } from '../xml/tokenizer';
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

/**
 * Cap on `cellXfs` entries (EC-STYLE-FIELD-RANGE). Excel documents 65,490 unique cell formats; staying under 64,000
 * leaves room for the ones Excel adds itself when the file is edited.
 */
export const MAX_CELL_STYLES: number = 64_000;
/** Excel's font size range, in points (the Format Cells dialog refuses anything outside it). */
const MIN_FONT_SIZE = 1;
const MAX_FONT_SIZE = 409;
/** Excel truncates or refuses longer font names. */
const MAX_FONT_NAME_LENGTH = 31;
/**
 * A control character in a font name or number format code is dropped on write and could leave nothing behind, so
 * it is refused (EC-STYLE-FIELD-RANGE).
 */
// eslint-disable-next-line no-control-regex -- the control characters are exactly what this looks for
const CONTROL_CHARACTER = /[\u0000-\u001F\uFFFE\uFFFF]/;

/** Text that would still say something once written: no control characters, and not only unpaired surrogates. */
function isWritableText(text: string): boolean {
  return !CONTROL_CHARACTER.test(text) && escapeAttr(text) !== '';
}
/** Excel's Format Cells dialog refuses a longer custom number format code. */
const MAX_NUMFMT_CODE_LENGTH = 255;
const HORIZONTAL_ALIGNMENTS: ReadonlySet<string> = new Set(['left', 'center', 'right']);
const VERTICAL_ALIGNMENTS: ReadonlySet<string> = new Set(['top', 'center', 'bottom']);
const BORDER_LINE_STYLES: ReadonlySet<string> = new Set<BorderLineStyle>(['thin', 'medium', 'thick', 'dashed', 'dotted', 'double', 'hair']);

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
function normalizeColor(color: unknown, field: string): string {
  const hex = typeof color === 'string' ? (color.startsWith('#') ? color.slice(1) : color).toUpperCase() : '';
  if (!COLOR_HEX.test(hex)) {
    throw invalidStyle(field, color, `${field} ${formatValue(color)} is not a colour. Use #RRGGBB, RRGGBB or AARRGGBB hex.`);
  }
  return `FF${hex.slice(-6)}`;
}

/** A value quoted for an error message: strings in quotes, everything else as `String()` gives it. */
/** Strings go through `JSON.stringify`, so a control character in a refused value shows up as `\u0002`. */
function formatValue(value: unknown): string {
  return typeof value === 'string' ? JSON.stringify(value) : String(value);
}

function invalidStyle(field: string, value: unknown, message: string): XlsxError {
  return new XlsxError('WRITER_STATE', message, { field, value });
}

function isObject(value: unknown): value is object {
  return typeof value === 'object' && value !== null;
}

/**
 * Check every field of a `CellStyle` before anything is registered, so a rejected style leaves no orphan font, fill
 * or number format behind (EC-STYLE-FIELD-RANGE). Numeric number format ids are checked by `registerNumFmt`, which
 * runs first.
 */
function validateStyle(style: CellStyle): void {
  if (!isObject(style)) {
    throw invalidStyle('style', style, `A cell style must be an object; got ${formatValue(style)}.`);
  }
  const { font, fill, border, alignment, numFmt } = style;
  if (font !== undefined) {
    if (!isObject(font)) {
      throw invalidStyle('font', font, `font must be an object; got ${formatValue(font)}.`);
    }
    const { size, name, color } = font;
    if (color !== undefined) {
      normalizeColor(color, 'font.color');
    }
    if (size !== undefined && !(typeof size === 'number' && size >= MIN_FONT_SIZE && size <= MAX_FONT_SIZE)) {
      throw invalidStyle(
        'font.size',
        size,
        `font.size ${formatValue(size)} is not a size Excel accepts. Use ${MIN_FONT_SIZE} to ${MAX_FONT_SIZE} points.`,
      );
    }
    if (
      name !== undefined &&
      !(typeof name === 'string' && name.length > 0 && name.length <= MAX_FONT_NAME_LENGTH && isWritableText(name))
    ) {
      throw invalidStyle(
        'font.name',
        name,
        `font.name ${formatValue(name)} is not a font name Excel accepts. Use 1 to ${MAX_FONT_NAME_LENGTH} characters and no control characters.`,
      );
    }
  }
  if (fill !== undefined) {
    if (!(isObject(fill) && (fill as { color?: unknown }).color !== undefined)) {
      throw invalidStyle('fill.color', fill, 'fill needs a color, like { color: "#FFFF00" }.');
    }
    normalizeColor(fill.color, 'fill.color');
  }
  if (border !== undefined) {
    if (typeof border === 'string') {
      validateBorderLineStyle('border', border);
    } else if (isObject(border)) {
      for (const side of BORDER_SIDES) {
        if (border[side] !== undefined) {
          validateBorderLineStyle(`border.${side}`, border[side]);
        }
      }
      if (border.color !== undefined) {
        normalizeColor(border.color, 'border.color');
      }
    } else {
      throw invalidStyle('border', border, `border must be a line style or an object; got ${formatValue(border)}.`);
    }
  }
  if (alignment !== undefined) {
    if (!isObject(alignment)) {
      throw invalidStyle('alignment', alignment, `alignment must be an object; got ${formatValue(alignment)}.`);
    }
    const { horizontal, vertical } = alignment;
    if (horizontal !== undefined && !HORIZONTAL_ALIGNMENTS.has(horizontal)) {
      throw invalidStyle(
        'alignment.horizontal',
        horizontal,
        `alignment.horizontal ${formatValue(horizontal)} is not one of left, center, right.`,
      );
    }
    if (vertical !== undefined && !VERTICAL_ALIGNMENTS.has(vertical)) {
      throw invalidStyle('alignment.vertical', vertical, `alignment.vertical ${formatValue(vertical)} is not one of top, center, bottom.`);
    }
  }
  if (typeof numFmt === 'string' && (numFmt.length === 0 || numFmt.length > MAX_NUMFMT_CODE_LENGTH || !isWritableText(numFmt))) {
    throw invalidStyle(
      'numFmt',
      numFmt,
      `numFmt ${formatValue(numFmt)} is not a format code Excel accepts. Use 1 to ${MAX_NUMFMT_CODE_LENGTH} characters and no control characters.`,
    );
  }
  if (numFmt !== undefined && typeof numFmt !== 'string' && typeof numFmt !== 'number') {
    throw invalidStyle('numFmt', numFmt, `numFmt must be a format code or a built-in format id; got ${formatValue(numFmt)}.`);
  }
}

function validateBorderLineStyle(field: string, lineStyle: unknown): void {
  if (typeof lineStyle !== 'string' || !BORDER_LINE_STYLES.has(lineStyle)) {
    throw invalidStyle(
      field,
      lineStyle,
      `${field} ${formatValue(lineStyle)} is not a border style. Use ${[...BORDER_LINE_STYLES].join(', ')}.`,
    );
  }
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
  /**
   * The style a Date written under style `i` gets (EC-DATE-STYLE-MERGE): `i` itself when it already has a date
   * format, otherwise `i` with the default date format swapped in. Filled on first use, so a date column costs one
   * array read per cell after its first cell.
   */
  private readonly dateStyleByStyleId: StyleId[] = [];
  private dateStyleId: StyleId | undefined;
  private boldHeaderStyleId: StyleId | undefined;

  /** Number of cellXfs entries (default style included). */
  get count(): number {
    return this.cellXfs.length;
  }

  /**
   * Register a style (deduplicated) and return its id. Every field is validated first (EC-STYLE-FIELD-RANGE) and a
   * bad one throws `WRITER_STATE` naming the field; so does registering a new style past `MAX_CELL_STYLES`.
   */
  register(style: CellStyle): StyleId {
    validateStyle(style);
    const numFmtId = this.registerNumFmt(style.numFmt);
    return this.internXf({
      numFmtId,
      fontId: this.registerFont(style.font),
      fillId: this.registerFill(style.fill),
      borderId: this.registerBorder(style.border),
      horizontal: style.alignment?.horizontal ?? '',
      vertical: style.alignment?.vertical ?? '',
      wrapText: style.alignment?.wrapText === true,
      isDate: this.isDateNumFmt(numFmtId),
    });
  }

  /**
   * The style id to write a Date under when the caller asked for `styleId` (EC-DATE-STYLE-MERGE): the caller's own
   * style when it already renders as a date, otherwise a derived style with the caller's font, fill, border and
   * alignment and the default date format. Derived once per style id and cached; styles are serialized at close,
   * so registering one mid-stream is safe. `styleId` must be a registered id.
   */
  dateStyleFor(styleId: StyleId): StyleId {
    const cached = this.dateStyleByStyleId[styleId];
    return cached === undefined ? this.deriveDateStyle(styleId) : cached;
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

  private deriveDateStyle(styleId: StyleId): StyleId {
    const xf = this.cellXfs[styleId];
    let derived: StyleId;
    if (xf === undefined) {
      derived = this.defaultDateStyle;
    } else if (xf.isDate) {
      derived = styleId;
    } else {
      derived = this.internXf({ ...xf, numFmtId: this.registerNumFmt(DEFAULT_DATE_FORMAT), isDate: true });
    }
    this.dateStyleByStyleId[styleId] = derived;
    return derived;
  }

  private internXf(xf: CellXf): StyleId {
    const key = xfKey(xf);
    const existing = this.xfIdByKey.get(key);
    if (existing !== undefined) {
      return existing;
    }
    if (this.cellXfs.length >= MAX_CELL_STYLES) {
      throw new XlsxError(
        'WRITER_STATE',
        `This workbook already has ${MAX_CELL_STYLES} cell styles, the most Excel can load. Reuse the ids registerStyle returned instead of registering a style per cell.`,
        { limit: MAX_CELL_STYLES },
      );
    }
    const styleId = this.cellXfs.length;
    this.cellXfs.push(xf);
    this.xfIdByKey.set(key, styleId);
    return styleId;
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
      // Ids 0-163 are Excel's built-ins (locale-specific ones included). A cell pointing at a custom id with no
      // `<numFmt>` element is a repair-dialog error (primer section 12.5), and a negative or fractional id is no id.
      const isBuiltin = Number.isInteger(numFmt) && numFmt >= 0 && numFmt < FIRST_CUSTOM_NUMFMT_ID;
      if (!isBuiltin && !this.customNumFmtCodeById.has(numFmt)) {
        throw invalidStyle(
          'numFmt',
          numFmt,
          `numFmt ${numFmt} is not a built-in number format id (0-${FIRST_CUSTOM_NUMFMT_ID - 1}). Register the format code instead.`,
        );
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
    const color = font.color === undefined ? undefined : normalizeColor(font.color, 'font.color');
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
    const color = normalizeColor(fill.color, 'fill.color');
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
    const color = typeof border === 'string' || border.color === undefined ? undefined : normalizeColor(border.color, 'border.color');
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

/**
 * Read what the reader needs from `xl/styles.xml`: the number formats and, for every `cellXfs` entry, whether it
 * renders as a date (EC-DATE-DETECTION-VIA-NUMFMT). Only `<cellXfs>` counts - `<cellStyleXfs>` and `<dxfs>` hold
 * `<xf>` and `<numFmt>` elements of their own that `c/@s` never indexes (primer 7.2). `applyNumberFormat` is
 * deliberately ignored: it declares that the xf overrides its parent cell style, while Excel renders the xf's own
 * `numFmtId` either way, so honouring it would hide dates from producers that leave the flag off.
 *
 * Everything else in the part is skipped, which is what makes it tolerant of the producer quirks in the catalog:
 * POI's `<u val="none"/>` (EC-POI-UNDERLINE-NONE-REJECTED) and 6-hex colours (EC-POI-RGB-6-HEX), LibreOffice and
 * Numbers extras, `x:`-prefixed or Strict-namespace documents (EC-STRICT-NAMESPACES), `mc:Ignorable`/`x14ac` noise.
 */
export function parseStyles(xml: string): ParsedStyles {
  const numFmts = new Map<number, string>();
  const numFmtIdByXf: number[] = [];
  const openElements: string[] = [];
  const tokenizer = new XmlTokenizer({
    start(name: string): void {
      const parent = openElements.at(-1);
      openElements.push(name);
      if (name === 'numFmt' && parent === 'numFmts') {
        const numFmtId = Number.parseInt(tokenizer.attr('numFmtId') ?? '', 10);
        const formatCode = tokenizer.attr('formatCode');
        if (Number.isFinite(numFmtId) && formatCode !== undefined) {
          numFmts.set(numFmtId, formatCode);
        }
        return;
      }
      if (name === 'xf' && parent === 'cellXfs') {
        const numFmtId = Number.parseInt(tokenizer.attr('numFmtId') ?? '', 10);
        numFmtIdByXf.push(Number.isFinite(numFmtId) ? numFmtId : 0);
      }
    },
    text(): void {},
    end(): void {
      openElements.pop();
    },
  });
  tokenizer.push(xml);
  tokenizer.end();

  const isDateByXf = new Uint8Array(numFmtIdByXf.length);
  for (const [index, numFmtId] of numFmtIdByXf.entries()) {
    // An explicit `<numFmt>` wins over the built-in table: producers redefine ids inside the reserved range
    // (write-excel-file uses 100, SheetJS overrides 56) and the code they wrote is what Excel renders.
    const formatCode = numFmts.get(numFmtId);
    const isDate = formatCode === undefined ? isBuiltinDateId(numFmtId) : isDateFormatCode(formatCode);
    isDateByXf[index] = isDate ? 1 : 0;
  }
  return { isDateByXf, numFmts };
}
