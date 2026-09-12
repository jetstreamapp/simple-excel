import { notImplemented } from '../internal/not-implemented';

/** The number format we write for JS Dates by default (custom id, registered by the style registry). */
export const DEFAULT_DATE_FORMAT: string = 'yyyy-mm-dd hh:mm:ss';
/** First id available for custom number formats. */
export const FIRST_CUSTOM_NUMFMT_ID: number = 164;

/** Built-in number formats (ids 0-22, 37-40, 45-49) with their codes. */
export const BUILTIN_NUMFMTS: ReadonlyMap<number, string> = new Map<number, string>();

/** True for built-in ids that Excel treats as dates or times (14-22, 45-47; plus the locale ranges 27-36 and 50-58). */
export function isBuiltinDateId(id: number): boolean {
  void id;
  throw notImplemented('sml/numfmt');
}

/**
 * True when a custom format code renders a date or time: after removing quoted literals, bracketed sections
 * (except elapsed `[h]`/`[mm]`/`[ss]`), backslash escapes and `General`, any of `y m d h s` or `AM/PM` remains.
 * `*` and `_` consume the following character.
 */
export function isDateFormatCode(code: string): boolean {
  void code;
  throw notImplemented('sml/numfmt');
}

/** Id of a built-in whose code equals `code` exactly, or -1. Lets `'0.00'` resolve to id 2 without a `<numFmt>`. */
export function builtinIdForCode(code: string): number {
  void code;
  throw notImplemented('sml/numfmt');
}
