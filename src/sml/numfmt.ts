/** The number format we write for JS Dates by default (custom id, registered by the style registry). */
export const DEFAULT_DATE_FORMAT: string = 'yyyy-mm-dd hh:mm:ss';
/** First id available for custom number formats. */
export const FIRST_CUSTOM_NUMFMT_ID: number = 164;

/** Built-in number formats (ids 0-22, 37-40, 45-49) with their codes. */
export const BUILTIN_NUMFMTS: ReadonlyMap<number, string> = new Map<number, string>([
  [0, 'General'],
  [1, '0'],
  [2, '0.00'],
  [3, '#,##0'],
  [4, '#,##0.00'],
  [9, '0%'],
  [10, '0.00%'],
  [11, '0.00E+00'],
  [12, '# ?/?'],
  [13, '# ??/??'],
  [14, 'mm-dd-yy'],
  [15, 'd-mmm-yy'],
  [16, 'd-mmm'],
  [17, 'mmm-yy'],
  [18, 'h:mm AM/PM'],
  [19, 'h:mm:ss AM/PM'],
  [20, 'h:mm'],
  [21, 'h:mm:ss'],
  [22, 'm/d/yy h:mm'],
  [37, '#,##0 ;(#,##0)'],
  [38, '#,##0 ;[Red](#,##0)'],
  [39, '#,##0.00;(#,##0.00)'],
  [40, '#,##0.00;[Red](#,##0.00)'],
  [45, 'mm:ss'],
  [46, '[h]:mm:ss'],
  [47, 'mmss.0'],
  [48, '##0.0E+0'],
  [49, '@'],
]);

const BUILTIN_ID_BY_CODE: ReadonlyMap<string, number> = (() => {
  const idByCode = new Map<string, number>();
  for (const [id, code] of BUILTIN_NUMFMTS) {
    idByCode.set(code, id);
  }
  return idByCode;
})();

/** Bracketed tokens that mean elapsed time rather than a colour, locale, currency or condition tag. */
const ELAPSED_TIME_TOKENS: ReadonlySet<string> = new Set(['h', 'hh', 'm', 'mm', 's', 'ss']);

/** The characters that render a date or time component once literals, escapes and bracket tags are gone. */
const DATE_TOKEN_CHARS: ReadonlySet<string> = new Set(['y', 'm', 'd', 'h', 's']);

const GENERAL = 'general';
const TWELVE_HOUR_TOKENS: readonly string[] = ['am/pm', 'a/p'];

/** True for built-in ids that Excel treats as dates or times (14-22, 45-47; plus the locale ranges 27-36 and 50-58). */
export function isBuiltinDateId(id: number): boolean {
  // 27-36 and 50-58 are the CJK locale date formats, 71-81 the Thai ones (primer section 6.4).
  return (id >= 14 && id <= 22) || (id >= 27 && id <= 36) || (id >= 45 && id <= 47) || (id >= 50 && id <= 58) || (id >= 71 && id <= 81);
}

/** Case-insensitive equivalent of `code.startsWith(lowercaseWord, index)`. */
function matchesWordAt(code: string, index: number, lowercaseWord: string): boolean {
  return code.slice(index, index + lowercaseWord.length).toLowerCase() === lowercaseWord;
}

/**
 * True when a custom format code renders a date or time: after removing quoted literals, bracketed sections
 * (except elapsed `[h]`/`[mm]`/`[ss]`), backslash escapes and `General`, any of `y m d h s` or `AM/PM` remains.
 * `*` and `_` consume the following character.
 */
export function isDateFormatCode(code: string): boolean {
  let index = 0;
  while (index < code.length) {
    const char = code.charAt(index);
    if (char === '"') {
      const closing = code.indexOf('"', index + 1);
      index = closing === -1 ? code.length : closing + 1;
      continue;
    }
    if (char === '[') {
      const closing = code.indexOf(']', index + 1);
      const token = (closing === -1 ? code.slice(index + 1) : code.slice(index + 1, closing)).toLowerCase();
      if (ELAPSED_TIME_TOKENS.has(token)) {
        return true;
      }
      index = closing === -1 ? code.length : closing + 1;
      continue;
    }
    // `\x` escapes a literal character, `*x` repeats one and `_x` pads the width of one: none of them render a token.
    if (char === '\\' || char === '*' || char === '_') {
      index += 2;
      continue;
    }
    if (matchesWordAt(code, index, GENERAL)) {
      index += GENERAL.length;
      continue;
    }
    for (const token of TWELVE_HOUR_TOKENS) {
      if (matchesWordAt(code, index, token)) {
        return true;
      }
    }
    if (DATE_TOKEN_CHARS.has(char.toLowerCase())) {
      return true;
    }
    index += 1;
  }
  return false;
}

/** Id of a built-in whose code equals `code` exactly, or -1. Lets `'0.00'` resolve to id 2 without a `<numFmt>`. */
export function builtinIdForCode(code: string): number {
  return BUILTIN_ID_BY_CODE.get(code) ?? -1;
}
