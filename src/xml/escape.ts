const CHAR_TAB = 9;
const CHAR_LINE_FEED = 10;
const CHAR_SPACE = 32;
const CHAR_AMPERSAND = 38;
const CHAR_LESS_THAN = 60;
const CHAR_GREATER_THAN = 62;
const CHAR_UNDERSCORE = 95;
const CHAR_LOWERCASE_X = 120;

const SURROGATE_FIRST = 0xd800;
const SURROGATE_HIGH_LAST = 0xdbff;
const SURROGATE_LOW_FIRST = 0xdc00;
const SURROGATE_LAST = 0xdfff;
/** U+FFFE and U+FFFF are the two BMP code points the XML 1.0 `Char` production excludes. */
const NON_CHARACTER_FIRST = 0xfffe;

const REPLACEMENT_CHARACTER = '�';
const ESCAPED_UNDERSCORE = '_x005F_';
/** `_xHHHH_` is seven characters: `_`, `x`, four hex digits, `_`. */
const ESCAPE_RUN_LENGTH = 7;

/**
 * The ASCII code units `encodeCellText` has to look at: everything below 0x20 except tab and LF (control characters
 * and CR become `_xHHHH_`), the three XML text metacharacters, and `_` (which only matters when it starts an
 * escape-shaped run, checked when the scan gets there). Everything else passes through on one comparison.
 */
const ASCII_NEEDS_ENCODING: Uint8Array = buildAsciiEncodingTable();

/**
 * The same set as one pattern: the XML metacharacters, the control characters XML forbids (tab and LF excluded),
 * surrogates, the two non-characters, and `_x`, the only shape a literal underscore ever has to be defused in.
 * Scanning for it in the regexp engine clears a clean string several times faster than the char-code loop below.
 * Global, because `mayNeedEncoding` resumes it past every surrogate pair it has cleared.
 */
// eslint-disable-next-line no-control-regex -- the control characters XML forbids are exactly what this looks for
const NEEDS_ENCODING_PATTERN = /[\u0000-\u0008\u000B-\u001F&<>\uD800-\uDFFF\uFFFE\uFFFF]|_x/g;
/** Below this length starting the regexp engine costs more than the char-code loop it saves. */
const REGEX_SCAN_MIN_LENGTH = 24;

/** Tab, LF and CR anywhere in the text: what `<t>` loses without `xml:space="preserve"`. */
const PRESERVED_BREAK_PATTERN = /[\t\n\r]/;

function buildAsciiEncodingTable(): Uint8Array {
  const table = new Uint8Array(128);
  for (let code = 0; code < CHAR_SPACE; code++) {
    table[code] = 1;
  }
  table[CHAR_TAB] = 0;
  table[CHAR_LINE_FEED] = 0;
  table[CHAR_AMPERSAND] = 1;
  table[CHAR_LESS_THAN] = 1;
  table[CHAR_GREATER_THAN] = 1;
  table[CHAR_UNDERSCORE] = 1;
  return table;
}

function hexDigitValue(code: number): number {
  if (code >= 48 && code <= 57) {
    return code - 48;
  }
  if (code >= 97 && code <= 102) {
    return code - 87;
  }
  if (code >= 65 && code <= 70) {
    return code - 55;
  }
  return -1;
}

/**
 * The code unit an `_xHHHH_` run starting at `index` stands for, or -1 when no run starts there. The `x` must be
 * lowercase (EC-XML-ESCAPE-CASE: `_X0041_` is literal text), while the four hex digits may be either case: Excel
 * writes them uppercase but other producers emit `_x000d_`, so both are decoded and both are defused on write.
 */
function escapeRunValueAt(text: string, index: number): number {
  if (index + ESCAPE_RUN_LENGTH > text.length) {
    return -1;
  }
  if (
    text.charCodeAt(index) !== CHAR_UNDERSCORE ||
    text.charCodeAt(index + 1) !== CHAR_LOWERCASE_X ||
    text.charCodeAt(index + 6) !== CHAR_UNDERSCORE
  ) {
    return -1;
  }
  let value = 0;
  for (let offset = 2; offset < 6; offset++) {
    const digit = hexDigitValue(text.charCodeAt(index + offset));
    if (digit < 0) {
      return -1;
    }
    value = value * 16 + digit;
  }
  return value;
}

function escapeSequence(code: number): string {
  return `_x${code.toString(16).toUpperCase().padStart(4, '0')}_`;
}

/** Code units `encodeCellText` writes as `_xHHHH_`, so they contribute a leading underscore to the output. */
function isWrittenAsEscapeSequence(code: number): boolean {
  return (code < CHAR_SPACE && code !== CHAR_TAB && code !== CHAR_LINE_FEED) || code >= NON_CHARACTER_FIRST;
}

/**
 * True when the `_` at `index` would start an `_xHHHH_`-shaped run in the *encoded* output, so it has to be defused
 * as `_x005F_`. Judging the run on the output rather than on the input alone matters when an escape follows the run:
 * a literal `_x0041` before a control character encodes to `_x0041` + `_x0001_`, whose first seven characters read
 * back as the escape `_x0041_`. Only characters the encoder leaves alone can supply the `x` and the four hex digits,
 * while the closing underscore can come either from a literal `_` or from the escape sequence that follows it.
 */
function startsEncodedEscapeRun(text: string, index: number): boolean {
  if (index + ESCAPE_RUN_LENGTH > text.length || text.charCodeAt(index + 1) !== CHAR_LOWERCASE_X) {
    return false;
  }
  for (let offset = 2; offset < 6; offset++) {
    if (hexDigitValue(text.charCodeAt(index + offset)) < 0) {
      return false;
    }
  }
  const closingCode = text.charCodeAt(index + 6);
  return closingCode === CHAR_UNDERSCORE || isWrittenAsEscapeSequence(closingCode);
}

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
 * True when `text` holds anything `encodeCellText` has to act on. The pattern also matches the two code units of a
 * well-formed surrogate pair, which the encoder passes through untouched (an emoji is legal XML), so a surrogate hit
 * is confirmed against the code unit after it and the scan resumes past the pair. Only a *lone* surrogate is work.
 */
function mayNeedEncoding(text: string): boolean {
  NEEDS_ENCODING_PATTERN.lastIndex = 0;
  for (let match = NEEDS_ENCODING_PATTERN.exec(text); match !== null; match = NEEDS_ENCODING_PATTERN.exec(text)) {
    const code = text.charCodeAt(match.index);
    if (code < SURROGATE_FIRST || code > SURROGATE_HIGH_LAST) {
      return true;
    }
    const nextCode = text.charCodeAt(match.index + 1);
    if (!(nextCode >= SURROGATE_LOW_FIRST && nextCode <= SURROGATE_LAST)) {
      return true;
    }
    NEEDS_ENCODING_PATTERN.lastIndex = match.index + 2;
  }
  return false;
}

/**
 * Encode cell text the way Excel expects inside `<t>`: XML-escaped, with characters XML 1.0 forbids (< 0x20 except
 * tab/LF, and 0xFFFE/0xFFFF) written as `_xHHHH_`, CR written as `_x000D_`, and any literal `_xHHHH_`-shaped run
 * protected by escaping its underscore as `_x005F_` (overlapping runs included, left to right). Lone surrogates
 * become U+FFFD. Returns the input unchanged when it is already clean (the common case; no allocation).
 *
 * One charCode pass does both jobs, so the `_x005F_` the defusing emits is never itself rescanned and the output is
 * identical to escaping underscores first and XML entities second.
 */
export function encodeCellText(text: string): string {
  if (text.length >= REGEX_SCAN_MIN_LENGTH && !mayNeedEncoding(text)) {
    return text;
  }
  let out = '';
  let last = 0;
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    if (code < 128) {
      if (ASCII_NEEDS_ENCODING[code] === 0) {
        continue;
      }
    } else if (code < SURROGATE_FIRST || (code > SURROGATE_LAST && code < NON_CHARACTER_FIRST)) {
      continue;
    }

    let replacement: string;
    if (code === CHAR_AMPERSAND) {
      replacement = '&amp;';
    } else if (code === CHAR_LESS_THAN) {
      replacement = '&lt;';
    } else if (code === CHAR_GREATER_THAN) {
      replacement = '&gt;';
    } else if (code === CHAR_UNDERSCORE) {
      if (!startsEncodedEscapeRun(text, i)) {
        continue;
      }
      replacement = ESCAPED_UNDERSCORE;
    } else if (code < CHAR_SPACE || code >= NON_CHARACTER_FIRST) {
      replacement = escapeSequence(code);
    } else if (code <= SURROGATE_HIGH_LAST) {
      const nextCode = i + 1 < text.length ? text.charCodeAt(i + 1) : 0;
      if (nextCode >= SURROGATE_LOW_FIRST && nextCode <= SURROGATE_LAST) {
        // A complete surrogate pair is an astral character (emoji): legal XML, written as raw UTF-8.
        i++;
        continue;
      }
      replacement = REPLACEMENT_CHARACTER;
    } else {
      replacement = REPLACEMENT_CHARACTER;
    }

    out += text.slice(last, i) + replacement;
    last = i + 1;
  }
  return last === 0 ? text : out + text.slice(last);
}

/**
 * Decode `_xHHHH_` escapes (lowercase `x`, exactly four hex digits) left to right, resuming after each match the way
 * Excel does. Callers only invoke this when the text contains `_x`.
 *
 * Resuming after the match is what makes the writer's defusing reversible: `_x005F_x0041_` decodes to the literal
 * `_x0041_`, never to `A`.
 */
export function decodeCellText(text: string): string {
  let out = '';
  let last = 0;
  let index = text.indexOf('_');
  while (index >= 0 && index + ESCAPE_RUN_LENGTH <= text.length) {
    const value = escapeRunValueAt(text, index);
    if (value < 0) {
      index = text.indexOf('_', index + 1);
      continue;
    }
    out += text.slice(last, index) + String.fromCharCode(value);
    last = index + ESCAPE_RUN_LENGTH;
    index = text.indexOf('_', last);
  }
  return last === 0 ? text : out + text.slice(last);
}

/**
 * True when the text has leading/trailing whitespace, a tab, or a newline, so `<t>` needs `xml:space="preserve"`.
 *
 * A leading or trailing tab, LF or CR is caught by the break pattern like any other, so only a space has to be
 * looked for at the ends.
 */
export function needsSpacePreserve(text: string): boolean {
  if (text.length === 0) {
    return false;
  }
  if (text.charCodeAt(0) === CHAR_SPACE || text.charCodeAt(text.length - 1) === CHAR_SPACE) {
    return true;
  }
  // Interior tabs and line breaks matter too: Excel drops them from a <t> without xml:space="preserve".
  return PRESERVED_BREAK_PATTERN.test(text);
}
