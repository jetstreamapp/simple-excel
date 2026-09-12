import { notImplemented } from '../internal/not-implemented';

/** Escape `&`, `<`, `>` for element text. Returns the input unchanged when nothing needs escaping. */
export function escapeText(text: string): string {
  let out = '';
  let last = 0;
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    if (code === 38 || code === 60 || code === 62) {
      out += text.slice(last, i) + (code === 38 ? '&amp;' : code === 60 ? '&lt;' : '&gt;');
      last = i + 1;
    }
  }
  return last === 0 ? text : out + text.slice(last);
}

/** Escape `&`, `<`, `>`, `"` and the whitespace characters that attribute normalization would otherwise fold. */
export function escapeAttr(text: string): string {
  let out = '';
  let last = 0;
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    let replacement: string | undefined;
    if (code === 38) {
      replacement = '&amp;';
    } else if (code === 60) {
      replacement = '&lt;';
    } else if (code === 62) {
      replacement = '&gt;';
    } else if (code === 34) {
      replacement = '&quot;';
    } else if (code === 9) {
      replacement = '&#9;';
    } else if (code === 10) {
      replacement = '&#10;';
    } else if (code === 13) {
      replacement = '&#13;';
    }
    if (replacement !== undefined) {
      out += text.slice(last, i) + replacement;
      last = i + 1;
    }
  }
  return last === 0 ? text : out + text.slice(last);
}

/**
 * Encode cell text the way Excel expects inside `<t>`: XML-escaped, with characters XML 1.0 forbids (< 0x20 except
 * tab/LF, and 0xFFFE/0xFFFF) written as `_xHHHH_`, CR written as `_x000D_`, and any literal `_xHHHH_`-shaped run
 * protected by escaping its underscore as `_x005F_` (overlapping runs included, left to right). Lone surrogates
 * become U+FFFD. Returns the input unchanged when it is already clean (the common case; no allocation).
 */
export function encodeCellText(text: string): string {
  void text;
  throw notImplemented('xml/escape');
}

/**
 * Decode `_xHHHH_` escapes (lowercase `x`, exactly four hex digits) left to right, resuming after each match the way
 * Excel does. Callers only invoke this when the text contains `_x`.
 */
export function decodeCellText(text: string): string {
  void text;
  throw notImplemented('xml/escape');
}

/** True when the text has leading/trailing whitespace, a tab, or a newline, so `<t>` needs `xml:space="preserve"`. */
export function needsSpacePreserve(text: string): boolean {
  void text;
  throw notImplemented('xml/escape');
}
