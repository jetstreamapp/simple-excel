/**
 * The corpus audit's exemption rules, side-effect free so `test/corpus-audit-divergence.test.ts` can pin them. Each rule
 * matches only the SheetJS-side shape it names: a difference that our reader could also produce by regressing must be
 * proven to be that shape (from the SheetJS cell or the raw sheet XML), never assumed from the two values alone.
 */
import * as XLSX from 'xlsx';

/** 9999-12-31 in the 1904 date system, the lower of the two systems' last serials (2,958,465 in the 1900 system). */
const LAST_SERIAL_1904 = 2_957_003;
const MARKER_DAY_START = Date.parse('1899-12-31T00:00:00.000Z');
const MARKER_DAY_LENGTH = 86_400_000;
const ONE_HOUR = 3_600_000;
/** A character past U+FFFF. */
const ASTRAL = /[\u{10000}-\u{10FFFF}]/gu;

export const ERROR_LITERALS: ReadonlySet<string> = new Set([
  '#NULL!',
  '#DIV/0!',
  '#VALUE!',
  '#REF!',
  '#NAME?',
  '#NUM!',
  '#N/A',
  '#GETTING_DATA',
  '#SPILL!',
  '#CALC!',
  '#FIELD!',
  '#BLOCKED!',
  '#UNKNOWN!',
  '#CONNECT!',
  '#BUSY!',
]);

/**
 * Facts about the SheetJS cell behind a difference, each proven only on demand. A rule that needs one is not applied
 * when the caller cannot supply it.
 */
export interface CellFacts {
  /** The cell is a text-typed cell with no value child (`<c t="inlineStr"/>`, `<c t="str"/>`) in the sheet XML. */
  textCellWithoutValue?: () => boolean;
  /** SheetJS parsed the cell as an error (`t: 'e'`). */
  sheetjsError?: () => boolean;
  /** The text of the cell's `<v>` when it is a `t="str"` cell, XML entities decoded and `_xHHHH_` escapes not. */
  strCellText?: () => string | undefined;
}

const decodeEscape = (_match: string, hex: string): string => String.fromCharCode(Number.parseInt(hex, 16));

/** One pass of ST_Xstring decoding, the reader contract: only lowercase `_xHHHH_` is an escape (EC-XML-ESCAPE-CASE). */
function decodeEscapesOnce(text: string): string {
  return text.replace(/_x([0-9A-Fa-f]{4})_/g, decodeEscape);
}

/** SheetJS's decoding of a `t="str"` value: two case-insensitive passes. */
function sheetjsDecodeStrValue(text: string): string {
  return text.replace(/_x([0-9A-Fa-f]{4})_/gi, decodeEscape).replace(/_x([0-9A-Fa-f]{4})_/gi, decodeEscape);
}

export function temporalMs(value: unknown): number | undefined {
  if (typeof value !== 'object' || value === null) {
    return undefined;
  }
  const record = value as Record<string, unknown>;
  if (typeof record.$datetime === 'string') {
    return Date.parse(`${record.$datetime}Z`);
  }
  if (typeof record.$date === 'string') {
    return Date.parse(`${record.$date}T00:00:00.000Z`);
  }
  if (typeof record.$time === 'string') {
    return Date.parse(`1899-12-30T${record.$time}Z`);
  }
  return undefined;
}

/** Year of a `$date` or `$datetime` value, read off its text: `Date.parse` cannot read a five-digit year. */
function temporalYear(value: unknown): number | undefined {
  if (typeof value !== 'object' || value === null) {
    return undefined;
  }
  const record = value as Record<string, unknown>;
  const text = typeof record.$datetime === 'string' ? record.$datetime : typeof record.$date === 'string' ? record.$date : undefined;
  const year = text === undefined ? undefined : /^(-?\d+)-/.exec(text)?.[1];
  return year === undefined ? undefined : Number(year);
}

/**
 * SheetJS behaviour already documented in `test/corpus-policies.json` and the catalog, so a difference it explains is
 * not a finding: `theirs` is SheetJS's value, `ours` ours, `facts` what can be proven about the SheetJS cell. Returns
 * the reason, or null when the difference is unexplained.
 */
export function knownSheetjsDivergence(theirs: unknown, ours: unknown, facts: CellFacts = {}): string | null {
  const theirTime = temporalMs(theirs);
  const ourTime = temporalMs(ours);
  if (theirTime !== undefined && ourTime !== undefined) {
    // EC-TIME-ONLY-HAS-DATE-PART: SheetJS hangs a time of day on 1899-12-31, we on 1899-12-30.
    if (
      theirTime >= MARKER_DAY_START &&
      theirTime < MARKER_DAY_START + MARKER_DAY_LENGTH &&
      Math.abs(theirTime - MARKER_DAY_LENGTH - ourTime) <= 1
    ) {
      return 'sheetjs-time-only-marker';
    }
    // SheetJS truncates the serial-to-milliseconds conversion where we round.
    if (Math.abs(theirTime - ourTime) <= 1) {
      return 'sheetjs-ms-truncation';
    }
    return null;
  }
  // EC-DATE-SERIAL-OVER-9999: past 9999-12-31 (serial 2,958,465, or 2,957,003 in the 1904 system) SheetJS still makes
  // a Date, of year 10000 or later or invalid, where we return the number. The lower 1904 bound covers both systems;
  // the year SheetJS gives is what proves the serial is past the end.
  if (typeof ours === 'number' && ours > LAST_SERIAL_1904 && (theirs === '$invalidDate' || (temporalYear(theirs) ?? 0) > 9999)) {
    return 'sheetjs-date-over-9999';
  }
  // Errors become defval in sheet_to_json with cellText:false (objects mode reads them as their text), except #NULL!
  // (BIFF code 0), which sheet_to_json turns into null whatever defval says. Only for a cell SheetJS parsed as an error.
  const ourError = typeof ours === 'object' && ours !== null && '$error' in ours ? (ours as { $error: string }).$error : ours;
  if (
    (theirs === '' || (theirs === null && ourError === '#NULL!')) &&
    typeof ourError === 'string' &&
    ERROR_LITERALS.has(ourError) &&
    facts.sheetjsError?.() === true
  ) {
    return 'sheetjs-error-defval';
  }
  // EC-TEXT-CELL-NO-TEXT: a text-typed cell with no value child (`<c t="inlineStr"/>`, `<c t="str"/>`) is blank for us
  // and '' for SheetJS. An empty shared string, a formula returning "" or `<is><t/></is>` is '' to both, so the shape
  // must be proven from the sheet XML.
  if (theirs === '' && ours === null && facts.textCellWithoutValue?.() === true) {
    return 'sheetjs-empty-text-cell';
  }
  if (typeof theirs === 'string' && typeof ours === 'string') {
    // EC-CRLF-NORMALIZED: SheetJS drops the CR of a CRLF (or a decoded _x000D_).
    if (ours.replace(/\r\n/g, '\n').replace(/\r/g, '\n') === theirs.replace(/\r/g, '\n') || ours.replace(/\r/g, '') === theirs) {
      return 'sheetjs-cr-collapse';
    }
    // EC-XML-NCR-ABOVE-BMP: SheetJS keeps the low 16 bits of a character reference past U+FFFF (U+1F603 -> U+F603),
    // or reads it as U+0000 in a t="str" value.
    if (
      ours.replace(ASTRAL, character => String.fromCharCode((character.codePointAt(0) as number) & 0xffff)) === theirs ||
      ours.replace(ASTRAL, '\0') === theirs
    ) {
      return 'sheetjs-ncr-above-bmp';
    }
    // EC-XML-ESCAPE-LITERAL: SheetJS decodes the escapes of a t="str" cell's <v> twice (shared and inline strings once),
    // so a literal written as `_x005F_x0041_` is `_x0041_` to us and Excel but 'A' to SheetJS. Applies only when the raw
    // text proves both sides: ours is its one correct decoding, SheetJS's the double one.
    const strText = facts.strCellText?.();
    if (strText !== undefined && ours === decodeEscapesOnce(strText) && theirs === sheetjsDecodeStrValue(strText)) {
      return 'sheetjs-str-escape-twice';
    }
    // EC-XML-ESCAPE-CASE: SheetJS also decodes the uppercase `_X0041_`, which Excel leaves literal. The lowercase form
    // is ours to decode (EC-XML-ESCAPE-UNDECODED), so it never explains a difference.
    const decoded = ours.replace(/_X([0-9A-Fa-f]{4})_/g, decodeEscape);
    if (decoded !== ours && (decoded === theirs || decoded.replace(/\r/g, '') === theirs)) {
      return 'sheetjs-escape-case';
    }
    return null;
  }
  // EC-DATE-PRE-1900 / EC-DATE-TIME-ONLY-NEGATIVE-SERIAL: serial 60 and negative serials stay numbers for us.
  if (theirTime !== undefined && typeof ours === 'number' && (ours === 60 || ours < 0)) {
    return 'sheetjs-fake-serial';
  }
  // SheetJS refuses to date serial 60 (Excel's fake 1900-02-29) and leaves it a number; we read 1900-03-01.
  if (
    typeof theirs === 'number' &&
    theirs >= 60 &&
    theirs < 61 &&
    ourTime !== undefined &&
    ourTime >= Date.parse('1900-02-28T00:00:00.000Z') &&
    ourTime < Date.parse('1900-03-02T00:00:00.000Z')
  ) {
    return 'sheetjs-serial-60';
  }
  // A number cell holding text past the double range ('inf', '1.797693134862316e+308'): SheetJS drops it, we keep the text.
  if (theirs === null && typeof ours === 'string' && /^(-?inf|nan|-?\d(\.\d+)?e\+30[89])$/i.test(ours)) {
    return 'sheetjs-unparseable-number';
  }
  return null;
}

/**
 * True when the wall clock of a `$datetime` or `$date` value does not exist in this process's time zone: a local Date
 * built from its fields lands on another wall clock (EC-DATE-DST-GAP).
 */
export function wallClockInLocalGap(value: unknown): boolean {
  if (typeof value !== 'object' || value === null) {
    return false;
  }
  const record = value as Record<string, unknown>;
  const text = typeof record.$datetime === 'string' ? record.$datetime : typeof record.$date === 'string' ? `${record.$date}T00:00` : '';
  const match = /^(\d{4})-(\d\d)-(\d\d)T(\d\d):(\d\d)/.exec(text);
  if (match === null) {
    return false;
  }
  const [year, month, day, hours, minutes] = match.slice(1).map(Number) as [number, number, number, number, number];
  const local = new Date(2000, 0, 1);
  local.setFullYear(year, month - 1, day);
  local.setHours(hours, minutes, 0, 0);
  return (
    local.getFullYear() !== year ||
    local.getMonth() !== month - 1 ||
    local.getDate() !== day ||
    local.getHours() !== hours ||
    local.getMinutes() !== minutes
  );
}

/**
 * The writer's documented conventions, not round-trip losses:
 * - ADR-003: a date before 1900-03-01 is written as ISO text.
 * - EC-DATE-DST-GAP: a local Date built for a wall clock in a spring-forward gap already holds the wall clock one hour
 *   later, so it reads back one hour later. Only that direction, and only for a wall clock that does not exist here.
 */
export function writerConvention(expected: unknown, actual: unknown): string | null {
  const expectedTime = temporalMs(expected);
  if (expectedTime !== undefined && typeof actual === 'string' && expectedTime < Date.parse('1900-03-01T00:00:00.000Z')) {
    return 'writer-pre-1900-as-text';
  }
  const actualTime = temporalMs(actual);
  if (expectedTime !== undefined && actualTime !== undefined && actualTime - expectedTime === ONE_HOUR && wallClockInLocalGap(expected)) {
    return 'writer-dst-gap';
  }
  return null;
}

/** A key SheetJS gives a column whose header cell is blank: `__EMPTY`, `__EMPTY_1`, ... */
export const BLANK_HEADER_KEY: RegExp = /^__EMPTY(?:_\d+)?$/;
/** SheetJS's key for a header cell whose text is '' (a text-typed cell with no text): '', then `_1`, `_2`, ... */
export const EMPTY_TEXT_HEADER_KEY: RegExp = /^(?:_\d+)?$/;
/** SheetJS's dateNF text of a time-only header cell: Excel's "day 0" of January 1900, as Excel shows it (`_N` dedupes). */
const TIME_ONLY_HEADER_KEY = /^1900-01-00T(\d\d:\d\d:\d\d(?:_\d+)?)$/;

/**
 * 'same' when SheetJS's key and ours name the same column, the known reason when they differ for one, or null.
 * - Two blank-header names (`__EMPTY_N`) are the same column even when an earlier difference shifted the numbering.
 * - EC-TEXT-CELL-NO-TEXT: a text-typed header cell with no value child is '' (then `_1`, ...) to SheetJS, blank to us,
 *   and so is the stub SheetJS makes for a comment on a cell with no value. `headerCellWithoutValue` must prove one of
 *   those shapes: an empty shared string or a formula returning "" is a '' header to us too.
 * - A time-only header cell (a serial below 1) is named by its wall clock on Excel's day 0 of 1900 by SheetJS and on
 *   1899-12-30 by us, the marker day our values use (EC-TIME-ONLY-HAS-DATE-PART).
 */
export function sameHeader(theirKey: string, ourKey: string, headerCellWithoutValue?: () => boolean): string | null {
  if (theirKey === ourKey || (BLANK_HEADER_KEY.test(theirKey) && BLANK_HEADER_KEY.test(ourKey))) {
    return 'same';
  }
  if (BLANK_HEADER_KEY.test(ourKey) && EMPTY_TEXT_HEADER_KEY.test(theirKey) && headerCellWithoutValue?.() === true) {
    return 'sheetjs-empty-text-cell';
  }
  const time = TIME_ONLY_HEADER_KEY.exec(theirKey)?.[1];
  if (time !== undefined && ourKey === `1899-12-30T${time}`) {
    return 'header-time-only-marker';
  }
  return null;
}

/** Jetstream's sheet_to_json options. */
export const SHEET_TO_JSON_OPTIONS: XLSX.Sheet2JSONOpts = {
  dateNF: 'yyyy"-"mm"-"dd"T"hh:mm:ss',
  defval: '',
  blankrows: false,
  rawNumbers: true,
};

/** A key JavaScript lists before every other key of an object, whatever order it was added in (an array index). */
const INDEX_KEY = /^(?:0|[1-9]\d{0,8})$/;

/**
 * The sheet column (0-based) of each of sheet_to_json's keys over `range`. `Object.keys` lists array-index keys ('8',
 * '36251') first, so a key's position is its column only when there are none of those. Otherwise the keys are named
 * again the way sheet_to_json names them (its own `format_cell` of each header cell, `__EMPTY` for none, `_N` for a
 * repeat) and used only when they come out as exactly `keys`; a key with no proven column maps to nothing.
 */
export function sheetjsKeyColumns(worksheet: XLSX.WorkSheet, range: XLSX.Range, keys: readonly string[]): Map<string, number> {
  if (!keys.some(key => INDEX_KEY.test(key))) {
    return new Map(keys.map((key, position) => [key, range.s.c + position]));
  }
  const columns = new Map<string, number>();
  const seen: Record<string, number> = {};
  for (let column = range.s.c; column <= range.e.c; column++) {
    const cell = worksheet[XLSX.utils.encode_cell({ r: range.s.r, c: column })] as XLSX.CellObject | undefined;
    const name = cell === undefined ? '__EMPTY' : XLSX.utils.format_cell(cell, undefined, SHEET_TO_JSON_OPTIONS);
    let key = name;
    let counter = seen[name] ?? 0;
    if (counter === 0) {
      seen[name] = 1;
    } else {
      do {
        key = `${name}_${counter++}`;
      } while (seen[key]);
      seen[name] = counter;
      seen[key] = 1;
    }
    columns.set(key, column);
  }
  const same = columns.size === keys.length && keys.every(key => columns.has(key));
  return same ? columns : new Map();
}

const CELL_START_TAG = /<((?:[\w.-]+:)?)c\s([^>]*)>/g;
const ATTRIBUTE = /([\w.:-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;

function attributes(text: string): Map<string, string> {
  const found = new Map<string, string>();
  for (const match of text.matchAll(ATTRIBUTE)) {
    found.set(match[1] as string, match[2] ?? match[3] ?? '');
  }
  return found;
}

/**
 * A1 refs of the text-typed cells with no value child in a worksheet part (EC-TEXT-CELL-NO-TEXT): `<c t="inlineStr"/>`,
 * `<c t="str"/>` or the same with an immediate end tag. A cell with any child (`<is/>`, `<v/>`, `<f>`) or with no `r`
 * attribute is left out, so the set never covers a cell whose shape is not proven.
 */
export function textCellsWithoutValue(xml: string): Set<string> {
  const refs = new Set<string>();
  for (const match of xml.matchAll(CELL_START_TAG)) {
    const prefix = match[1] as string;
    let attributeText = match[2] as string;
    const selfClosing = attributeText.endsWith('/');
    if (selfClosing) {
      attributeText = attributeText.slice(0, -1);
    } else {
      const after = (match.index as number) + match[0].length;
      const rest = xml.slice(after, after + 64).trimStart();
      if (!rest.startsWith(`</${prefix}c>`)) {
        continue;
      }
    }
    const parsed = attributes(attributeText);
    const type = parsed.get('t');
    const ref = parsed.get('r');
    if ((type === 'inlineStr' || type === 'str') && ref) {
      refs.add(ref.toUpperCase());
    }
  }
  return refs;
}

/**
 * The `<v>` text of every `t="str"` cell with an `r` attribute in a worksheet part, by A1 ref: XML entities decoded,
 * `_xHHHH_` escapes left as written.
 */
export function strCellTexts(xml: string): Map<string, string> {
  const texts = new Map<string, string>();
  const valuePatterns = new Map<string, RegExp>();
  for (const match of xml.matchAll(CELL_START_TAG)) {
    const prefix = match[1] as string;
    const attributeText = match[2] as string;
    if (attributeText.endsWith('/')) {
      continue;
    }
    const parsed = attributes(attributeText);
    const ref = parsed.get('r');
    if (parsed.get('t') !== 'str' || !ref) {
      continue;
    }
    const start = (match.index as number) + match[0].length;
    const end = xml.indexOf(`</${prefix}c>`, start);
    const body = end < 0 ? '' : xml.slice(start, end);
    let valuePattern = valuePatterns.get(prefix);
    if (valuePattern === undefined) {
      const escaped = prefix.replace(/[\\^$.*+?()[\]{}|-]/g, '\\$&');
      valuePattern = new RegExp(`<${escaped}v(?:\\s[^>]*)?>([^<]*)</${escaped}v>`);
      valuePatterns.set(prefix, valuePattern);
    }
    const value = valuePattern.exec(body);
    if (value) {
      texts.set(ref.toUpperCase(), decodeEntities(value[1] as string));
    }
  }
  return texts;
}

/** True when a worksheet part declares a `<dimension>` (it precedes `<sheetData>`). */
export function hasDimensionElement(xml: string): boolean {
  const sheetData = /<(?:[\w.-]+:)?sheetData\b/.exec(xml);
  const head = sheetData === null ? xml : xml.slice(0, sheetData.index);
  return /<(?:[\w.-]+:)?dimension\b/.test(head);
}

interface ZipContainer {
  FullPaths: string[];
  FileIndex: { content?: ArrayLike<number> | null }[];
}

const CFB = XLSX.CFB as { read(data: Uint8Array, options: { type: 'buffer' }): ZipContainer };
const RELATIONSHIP = /<(?:[\w.-]+:)?Relationship\s([^>]*)>/g;
const WORKBOOK_SHEET = /<(?:[\w.-]+:)?sheet\s([^>]*)>/g;

const NAMED_ENTITIES: Record<string, string> = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };

/** XML entity and character references decoded; an unknown entity is left as written. */
function decodeEntities(text: string): string {
  return text.replace(/&(?:#x([0-9A-Fa-f]+)|#(\d+)|(amp|lt|gt|quot|apos));/g, (match, hex?: string, decimal?: string, name?: string) => {
    if (name !== undefined) {
      return NAMED_ENTITIES[name] ?? match;
    }
    const codePoint = Number.parseInt((hex ?? decimal) as string, hex === undefined ? 10 : 16);
    return codePoint <= 0x10ffff ? String.fromCodePoint(codePoint) : match;
  });
}

/** A relationship target resolved against the part that owns it, as a zip path without a leading slash. */
function resolveTarget(target: string, ownerPath: string): string {
  const decoded = decodeEntities(target);
  if (decoded.startsWith('/')) {
    return decoded.slice(1);
  }
  const parts = ownerPath.split('/').slice(0, -1);
  for (const segment of decoded.split('/')) {
    if (segment === '..') {
      parts.pop();
    } else if (segment !== '.' && segment !== '') {
      parts.push(segment);
    }
  }
  return parts.join('/');
}

/**
 * Raw worksheet XML of a workbook, by sheet index in workbook order (SheetJS's `SheetNames` order), unzipped with
 * SheetJS's own zip reader rather than ours, and only on first use. A part it cannot locate is undefined, so a rule
 * needing it is not applied.
 */
export class RawSheets {
  private container: ZipContainer | null | undefined;
  private paths: (string | undefined)[] | undefined;
  private readonly xml = new Map<number, string | undefined>();
  private readonly textCells = new Map<number, Set<string> | undefined>();
  private readonly strTexts = new Map<number, Map<string, string> | undefined>();

  constructor(private readonly bytes: Uint8Array) {}

  private part(path: string): string | undefined {
    if (this.container === undefined) {
      try {
        this.container = CFB.read(this.bytes, { type: 'buffer' });
      } catch {
        this.container = null;
      }
    }
    if (this.container === null) {
      return undefined;
    }
    const wanted = path.toLowerCase();
    const index = this.container.FullPaths.findIndex(
      full =>
        full
          .replace(/^Root Entry\//i, '')
          .replace(/^\//, '')
          .toLowerCase() === wanted,
    );
    const content = index < 0 ? undefined : this.container.FileIndex[index]?.content;
    if (!content) {
      return undefined;
    }
    return new TextDecoder().decode(content instanceof Uint8Array ? content : Uint8Array.from(content));
  }

  private relationships(ownerPath: string): Map<string, { type: string; target: string }> {
    const slash = ownerPath.lastIndexOf('/');
    const relsPath = `${ownerPath.slice(0, slash + 1)}_rels/${ownerPath.slice(slash + 1)}.rels`;
    const found = new Map<string, { type: string; target: string }>();
    for (const match of (this.part(relsPath) ?? '').matchAll(RELATIONSHIP)) {
      const parsed = attributes(match[1] as string);
      const id = parsed.get('Id');
      const target = parsed.get('Target');
      if (id !== undefined && target !== undefined && parsed.get('TargetMode') !== 'External') {
        found.set(id, { type: parsed.get('Type') ?? '', target: resolveTarget(target, ownerPath) });
      }
    }
    return found;
  }

  private sheetPaths(): (string | undefined)[] {
    if (this.paths === undefined) {
      const root = [...this.relationships('').values()].find(relationship => relationship.type.endsWith('/officeDocument'));
      const workbookPath = root?.target ?? 'xl/workbook.xml';
      const relationships = this.relationships(workbookPath);
      this.paths = [...(this.part(workbookPath) ?? '').matchAll(WORKBOOK_SHEET)].map(match => {
        const idAttribute = [...attributes(match[1] as string)].find(([name]) => /^[\w.-]+:id$/.test(name));
        return idAttribute === undefined ? undefined : relationships.get(idAttribute[1])?.target;
      });
    }
    return this.paths;
  }

  /** The worksheet part's XML, or undefined when it cannot be located. */
  sheetXml(index: number): string | undefined {
    if (!this.xml.has(index)) {
      const path = this.sheetPaths()[index];
      this.xml.set(index, path === undefined ? undefined : this.part(path));
    }
    return this.xml.get(index);
  }

  /** True only when the sheet XML proves `ref` (A1) is a text-typed cell with no value child. */
  textCellWithoutValue(index: number, ref: string): boolean {
    if (!this.textCells.has(index)) {
      const xml = this.sheetXml(index);
      this.textCells.set(index, xml === undefined ? undefined : textCellsWithoutValue(xml));
    }
    return this.textCells.get(index)?.has(ref) ?? false;
  }

  /** The `<v>` text of a `t="str"` cell (see `strCellTexts`), or undefined when the sheet XML does not prove one. */
  strCellText(index: number, ref: string): string | undefined {
    if (!this.strTexts.has(index)) {
      const xml = this.sheetXml(index);
      this.strTexts.set(index, xml === undefined ? undefined : strCellTexts(xml));
    }
    return this.strTexts.get(index)?.get(ref);
  }

  /** True when the sheet part is located and has no `<dimension>`; false when it has one or cannot be read. */
  lacksDimension(index: number): boolean {
    const xml = this.sheetXml(index);
    return xml !== undefined && !hasDimensionElement(xml);
  }
}
