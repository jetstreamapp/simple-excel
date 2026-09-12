import { XlsxError } from '../errors';

export interface XmlTokenizerHandler {
  /** Element start. Read attributes through `tokenizer.attr(name)` during this callback only. */
  start(name: string): void;
  /** Character data between tags, entities decoded, delivered once per text run (never split across chunks). */
  text(text: string): void;
  end(name: string): void;
}

export interface XmlTokenizerLimits {
  /** Default 256. */
  readonly maxDepth?: number;
  /** Default 64 Mi UTF-16 units. */
  readonly maxTextLength?: number;
}

const CHAR_TAB = 9;
const CHAR_LINE_FEED = 10;
const CHAR_CARRIAGE_RETURN = 13;
const CHAR_SPACE = 32;
const CHAR_DOUBLE_QUOTE = 34;
const CHAR_HASH = 35;
const CHAR_SINGLE_QUOTE = 39;
const CHAR_SLASH = 47;
const CHAR_COLON = 58;
const CHAR_SEMICOLON = 59;
const CHAR_EQUALS = 61;
const CHAR_GREATER_THAN = 62;
const CHAR_QUESTION = 63;
const CHAR_EXCLAMATION = 33;
const CHAR_HYPHEN = 45;
const CHAR_OPEN_BRACKET = 91;
const CHAR_LOWERCASE_X = 120;
const CHAR_BYTE_ORDER_MARK = 0xfeff;

const MODE_TEXT = 0;
const MODE_CDATA = 1;
const MODE_COMMENT = 2;
const MODE_PROCESSING_INSTRUCTION = 3;

const CDATA_OPEN = '<![CDATA[';
const CDATA_CLOSE = ']]>';
const COMMENT_OPEN = '<!--';
const COMMENT_CLOSE = '-->';
const PROCESSING_INSTRUCTION_CLOSE = '?>';

/** A carriage return, alone or leading a line feed: what XML line-end normalization collapses. */
const CARRIAGE_RETURN_PATTERN = /\r\n?/g;

const DEFAULT_MAX_DEPTH = 256;
const DEFAULT_MAX_TEXT_LENGTH = 64 * 1024 * 1024;
/** A single tag this long is a broken (or hostile) file, not markup any spreadsheet produced. */
const MAX_TAG_LENGTH = 1024 * 1024;
/** `&#x10FFFF;` is the longest reference worth waiting for; anything longer is literal text. */
const MAX_ENTITY_LENGTH = 12;

const CORRUPT_FILE_ADVICE = 'The file may be damaged; try opening it in Excel and saving a fresh copy.';

function malformed(what: string): XlsxError {
  return new XlsxError('XML_MALFORMED', `This spreadsheet contains XML that could not be read (${what}). ${CORRUPT_FILE_ADVICE}`, {
    reason: what,
  });
}

function isWhitespaceCode(code: number): boolean {
  return code === CHAR_SPACE || code === CHAR_LINE_FEED || code === CHAR_TAB || code === CHAR_CARRIAGE_RETURN;
}

/**
 * The index of the `>` that closes the tag starting at `start`, or -1 when the buffer does not hold it yet. A `>` is
 * legal inside a quoted attribute value, so quoted runs are skipped rather than searched.
 */
function findTagEnd(text: string, start: number): number {
  let openQuote = 0;
  for (let index = start; index < text.length; index++) {
    const code = text.charCodeAt(index);
    if (openQuote !== 0) {
      if (code === openQuote) {
        openQuote = 0;
      }
    } else if (code === CHAR_DOUBLE_QUOTE || code === CHAR_SINGLE_QUOTE) {
      openQuote = code;
    } else if (code === CHAR_GREATER_THAN) {
      return index;
    }
  }
  return -1;
}

/**
 * `xmlns` and `xmlns:r` declare namespaces rather than data, and stripping prefixes would otherwise turn
 * `xmlns:r="...spreadsheetml..."` into an attribute named `r` on the sheet root.
 */
function isNamespaceDeclaration(text: string, start: number, end: number): boolean {
  if (end - start < 5 || !text.startsWith('xmlns', start)) {
    return false;
  }
  return end - start === 5 || text.charCodeAt(start + 5) === CHAR_COLON;
}

/**
 * The local name within `[start, end)`: everything after the last namespace prefix (`x:c` -> `c`). The search walks
 * the name itself rather than calling `lastIndexOf`, which would scan back through the whole chunk when a document
 * happens to contain no colons at all.
 */
function localNameStart(text: string, start: number, end: number): number {
  for (let index = end - 1; index > start; index--) {
    if (text.charCodeAt(index) === CHAR_COLON) {
      return index + 1;
    }
  }
  return start;
}

/** The text a character reference between `&` and `;` stands for, or undefined when it is not one we decode. */
function decodeCharacterReference(text: string, start: number, end: number): string | undefined {
  const isHex = text.charCodeAt(start + 1) === CHAR_LOWERCASE_X;
  const digitsStart = isHex ? start + 2 : start + 1;
  if (digitsStart >= end) {
    return undefined;
  }
  let value = 0;
  for (let index = digitsStart; index < end; index++) {
    const code = text.charCodeAt(index);
    let digit: number;
    if (code >= 48 && code <= 57) {
      digit = code - 48;
    } else if (isHex && code >= 97 && code <= 102) {
      digit = code - 87;
    } else if (isHex && code >= 65 && code <= 70) {
      digit = code - 55;
    } else {
      return undefined;
    }
    value = value * (isHex ? 16 : 10) + digit;
    if (value > 0x10ffff) {
      return undefined;
    }
  }
  // Code points that cannot appear in XML (NUL, lone surrogates) stay literal rather than poisoning a cell value.
  if (value === 0 || (value >= 0xd800 && value <= 0xdfff)) {
    return undefined;
  }
  return String.fromCodePoint(value);
}

/** The text an entity reference between `&` and `;` stands for, or undefined for an unknown name (left verbatim). */
function decodeEntityReference(text: string, start: number, end: number): string | undefined {
  if (text.charCodeAt(start) === CHAR_HASH) {
    return decodeCharacterReference(text, start, end);
  }
  switch (text.slice(start, end)) {
    case 'amp':
      return '&';
    case 'lt':
      return '<';
    case 'gt':
      return '>';
    case 'quot':
      return '"';
    case 'apos':
      return "'";
    default:
      return undefined;
  }
}

/**
 * The `;` that could terminate a reference starting at `start`, or -1. Only the next few characters are examined, so
 * text made of nothing but ampersands cannot turn the scan quadratic.
 */
function findReferenceEnd(text: string, start: number): number {
  const limit = Math.min(text.length, start + MAX_ENTITY_LENGTH);
  for (let index = start; index < limit; index++) {
    if (text.charCodeAt(index) === CHAR_SEMICOLON) {
      return index;
    }
  }
  return -1;
}

/** Replace every reference we recognize; unknown ones (`&foo;`, `&#zz;`) stay as written, the way Excel leaves them. */
function decodeEntities(text: string): string {
  let out = '';
  let last = 0;
  let ampersandIndex = text.indexOf('&');
  while (ampersandIndex >= 0) {
    const semicolonIndex = findReferenceEnd(text, ampersandIndex + 1);
    const decoded = semicolonIndex > 0 ? decodeEntityReference(text, ampersandIndex + 1, semicolonIndex) : undefined;
    if (decoded === undefined) {
      ampersandIndex = text.indexOf('&', ampersandIndex + 1);
      continue;
    }
    out += text.slice(last, ampersandIndex) + decoded;
    last = semicolonIndex + 1;
    ampersandIndex = text.indexOf('&', last);
  }
  return last === 0 ? text : out + text.slice(last);
}

/**
 * How much of `[start, end)` can be delivered as text when more input is still coming: everything up to a trailing
 * reference that has not been terminated yet, so `&am` + `p;` across two chunks still decodes to `&`.
 */
function textEndBeforePartialEntity(text: string, start: number, end: number): number {
  const scanFrom = Math.max(start, end - MAX_ENTITY_LENGTH);
  let ampersandIndex = text.indexOf('&', scanFrom);
  while (ampersandIndex >= 0 && ampersandIndex < end) {
    if (findReferenceEnd(text, ampersandIndex + 1) < 0) {
      return ampersandIndex;
    }
    ampersandIndex = text.indexOf('&', ampersandIndex + 1);
  }
  return end;
}

/**
 * Push tokenizer for the subset of XML that SpreadsheetML uses. Namespace prefixes are stripped from element and
 * attribute names (`x:c` -> `c`, `r:id` -> `id`) so Transitional and Strict parts look the same. Handles the five
 * predefined entities, numeric character references, CDATA, comments and processing instructions; rejects
 * `<!DOCTYPE` with `XlsxError('XML_DOCTYPE')`; reports structural problems as `XlsxError('XML_MALFORMED')`. Chunks
 * may split anywhere, including inside a tag, attribute value or entity.
 */
export class XmlTokenizer {
  private readonly handler: XmlTokenizerHandler;
  private readonly maxDepth: number;
  private readonly maxTextLength: number;
  private readonly openElements: string[] = [];
  /** Input that could not be consumed yet: a partial tag, reference or `]]>` marker. */
  private buffer = '';
  /** Character data seen so far in the current run, delivered when the next tag arrives or at `end()`. */
  private pendingText = '';
  private mode: number = MODE_TEXT;
  private needMoreInput = false;
  private sawFirstChunk = false;
  /** The previous chunk ended with a carriage return, so a line feed opening this one belongs to that line end. */
  private sawTrailingCarriageReturn = false;
  /** The chunk the running `start` callback's tag came from, with the bounds of that tag's attribute list. */
  private attributeSource = '';
  private attributeStart = 0;
  private attributeEnd = 0;

  constructor(handler: XmlTokenizerHandler, limits?: XmlTokenizerLimits) {
    this.handler = handler;
    this.maxDepth = limits?.maxDepth ?? DEFAULT_MAX_DEPTH;
    this.maxTextLength = limits?.maxTextLength ?? DEFAULT_MAX_TEXT_LENGTH;
  }

  push(chunk: string): void {
    if (chunk.length === 0) {
      return;
    }
    let text = chunk;
    if (!this.sawFirstChunk) {
      this.sawFirstChunk = true;
      if (text.charCodeAt(0) === CHAR_BYTE_ORDER_MARK) {
        text = text.slice(1);
      }
      if (text.length === 0) {
        return;
      }
    }
    text = this.normalizeLineEndings(text);
    if (text.length === 0) {
      return;
    }
    this.buffer = this.buffer.length === 0 ? text : this.buffer + text;
    this.scan(false);
  }

  /**
   * XML line-end normalization (XML 1.0 2.11): a CRLF pair and a lone CR are both one LF. It runs on the raw input,
   * before references are decoded, so a carriage return written deliberately as `&#13;` or `_x000D_` survives - which
   * is exactly the difference between a file that meant a CRLF and one that merely contains raw CRs
   * (EC-CRLF-NORMALIZED).
   */
  private normalizeLineEndings(chunk: string): string {
    const continuesLineEnd = this.sawTrailingCarriageReturn && chunk.charCodeAt(0) === CHAR_LINE_FEED;
    this.sawTrailingCarriageReturn = chunk.charCodeAt(chunk.length - 1) === CHAR_CARRIAGE_RETURN;
    const text = continuesLineEnd ? chunk.slice(1) : chunk;
    return text.indexOf('\r') < 0 ? text : text.replace(CARRIAGE_RETURN_PATTERN, '\n');
  }

  /** Flush trailing text and verify every element was closed. */
  end(): void {
    this.scan(true);
    if (this.mode !== MODE_TEXT) {
      throw malformed('a comment, CDATA section or processing instruction that never ended');
    }
    this.flushText();
    const unclosed = this.openElements.at(-1);
    if (unclosed !== undefined) {
      throw malformed(`<${unclosed}> was never closed`);
    }
  }

  /** Attribute value by local name for the element whose `start` callback is running; entities decoded. */
  attr(name: string): string | undefined {
    const source = this.attributeSource;
    const limit = this.attributeEnd;
    let index = this.attributeStart;
    while (index < limit) {
      while (index < limit && isWhitespaceCode(source.charCodeAt(index))) {
        index++;
      }
      const nameStart = index;
      while (index < limit) {
        const code = source.charCodeAt(index);
        if (code === CHAR_EQUALS || isWhitespaceCode(code)) {
          break;
        }
        index++;
      }
      const nameEnd = index;
      while (index < limit && isWhitespaceCode(source.charCodeAt(index))) {
        index++;
      }
      if (nameEnd === nameStart || index >= limit || source.charCodeAt(index) !== CHAR_EQUALS) {
        return undefined;
      }
      index++;
      while (index < limit && isWhitespaceCode(source.charCodeAt(index))) {
        index++;
      }
      const quoteCode = source.charCodeAt(index);
      if (quoteCode !== CHAR_DOUBLE_QUOTE && quoteCode !== CHAR_SINGLE_QUOTE) {
        return undefined;
      }
      const valueStart = index + 1;
      const valueEnd = source.indexOf(quoteCode === CHAR_DOUBLE_QUOTE ? '"' : "'", valueStart);
      if (valueEnd < 0 || valueEnd >= limit) {
        return undefined;
      }
      index = valueEnd + 1;
      if (isNamespaceDeclaration(source, nameStart, nameEnd)) {
        continue;
      }
      const localStart = localNameStart(source, nameStart, nameEnd);
      if (nameEnd - localStart === name.length && source.startsWith(name, localStart)) {
        const value = source.slice(valueStart, valueEnd);
        return value.indexOf('&') < 0 ? value : decodeEntities(value);
      }
    }
    return undefined;
  }

  /** Current nesting depth (elements opened and not yet closed). */
  get depth(): number {
    return this.openElements.length;
  }

  private scan(isFinal: boolean): void {
    const buffer = this.buffer;
    let position = 0;
    this.needMoreInput = false;
    while (position < buffer.length && !this.needMoreInput) {
      position = this.mode === MODE_TEXT ? this.stepText(buffer, position, isFinal) : this.stepSkipped(buffer, position, isFinal);
    }
    this.buffer = position === 0 ? buffer : buffer.slice(position);
  }

  /** Character data and the start of any markup construct. */
  private stepText(buffer: string, position: number, isFinal: boolean): number {
    const lessThanIndex = buffer.indexOf('<', position);
    if (lessThanIndex < 0) {
      return this.consumeText(buffer, position, buffer.length, isFinal);
    }
    if (lessThanIndex > position) {
      return this.consumeText(buffer, position, lessThanIndex, true);
    }
    return this.stepMarkup(buffer, position, isFinal);
  }

  /** Inside a CDATA section, comment or processing instruction: consume up to the closing marker. */
  private stepSkipped(buffer: string, position: number, isFinal: boolean): number {
    const isCdata = this.mode === MODE_CDATA;
    const terminator = isCdata ? CDATA_CLOSE : this.mode === MODE_COMMENT ? COMMENT_CLOSE : PROCESSING_INSTRUCTION_CLOSE;
    const closeIndex = buffer.indexOf(terminator, position);
    if (closeIndex < 0) {
      if (isFinal) {
        throw malformed(`a section closed by ${terminator} never ended`);
      }
      this.needMoreInput = true;
      // The tail may hold the first characters of the terminator, so keep it and consume everything before it.
      const keepFrom = Math.max(position, buffer.length - (terminator.length - 1));
      if (isCdata) {
        this.appendText(buffer, position, keepFrom, false);
      }
      return keepFrom;
    }
    if (isCdata) {
      this.appendText(buffer, position, closeIndex, false);
    }
    this.mode = MODE_TEXT;
    return closeIndex + terminator.length;
  }

  private stepMarkup(buffer: string, position: number, isFinal: boolean): number {
    const available = buffer.length - position;
    if (available < 2) {
      return this.wantMore(position, isFinal, 'a tag was never closed');
    }
    const secondCode = buffer.charCodeAt(position + 1);
    if (secondCode === CHAR_EXCLAMATION) {
      return this.stepDeclaration(buffer, position, isFinal, available);
    }
    if (secondCode === CHAR_QUESTION) {
      this.mode = MODE_PROCESSING_INSTRUCTION;
      return position + 2;
    }

    const tagEnd = findTagEnd(buffer, position + 1);
    if (tagEnd < 0) {
      if (available > MAX_TAG_LENGTH) {
        throw malformed(`a tag longer than ${MAX_TAG_LENGTH} characters`);
      }
      return this.wantMore(position, isFinal, 'a tag was never closed');
    }
    if (tagEnd - position > MAX_TAG_LENGTH) {
      throw malformed(`a tag longer than ${MAX_TAG_LENGTH} characters`);
    }
    this.flushText();
    if (secondCode === CHAR_SLASH) {
      this.handleEndTag(buffer, position + 2, tagEnd);
    } else {
      this.handleStartTag(buffer, position + 1, tagEnd);
    }
    return tagEnd + 1;
  }

  /** `<!` markup: comments and CDATA are data, everything else is a DTD declaration and is refused outright. */
  private stepDeclaration(buffer: string, position: number, isFinal: boolean, available: number): number {
    if (available < 3) {
      return this.wantMore(position, isFinal, 'a declaration was never closed');
    }
    const thirdCode = buffer.charCodeAt(position + 2);
    if (thirdCode === CHAR_HYPHEN) {
      if (available < COMMENT_OPEN.length) {
        return this.wantMore(position, isFinal, 'a comment was never closed');
      }
      if (!buffer.startsWith(COMMENT_OPEN, position)) {
        throw malformed('a comment that does not start with <!--');
      }
      this.mode = MODE_COMMENT;
      return position + COMMENT_OPEN.length;
    }
    if (thirdCode === CHAR_OPEN_BRACKET) {
      if (available < CDATA_OPEN.length) {
        return this.wantMore(position, isFinal, 'a CDATA section was never closed');
      }
      if (!buffer.startsWith(CDATA_OPEN, position)) {
        throw malformed('a section that does not start with <![CDATA[');
      }
      this.mode = MODE_CDATA;
      return position + CDATA_OPEN.length;
    }
    throw new XlsxError(
      'XML_DOCTYPE',
      'This file contains an XML document type or entity declaration, which spreadsheet files never use. ' +
        'It was not opened because such declarations can be used to attack the reader.',
    );
  }

  /** Stop consuming until the next chunk arrives; at the end of the input there is no next chunk, so it is an error. */
  private wantMore(position: number, isFinal: boolean, what: string): number {
    if (isFinal) {
      throw malformed(what);
    }
    this.needMoreInput = true;
    return position;
  }

  private handleStartTag(buffer: string, nameStart: number, tagEnd: number): void {
    let contentEnd = tagEnd;
    const selfClosing = buffer.charCodeAt(tagEnd - 1) === CHAR_SLASH;
    if (selfClosing) {
      contentEnd = tagEnd - 1;
    }
    let nameEnd = nameStart;
    while (nameEnd < contentEnd && !isWhitespaceCode(buffer.charCodeAt(nameEnd)) && buffer.charCodeAt(nameEnd) !== CHAR_SLASH) {
      nameEnd++;
    }
    if (nameEnd === nameStart) {
      throw malformed('a tag with no element name');
    }
    const name = buffer.slice(localNameStart(buffer, nameStart, nameEnd), nameEnd);

    this.openElements.push(name);
    if (this.openElements.length > this.maxDepth) {
      throw new XlsxError(
        'LIMIT_EXCEEDED',
        `This file nests XML elements more than ${this.maxDepth} levels deep, which no spreadsheet needs, so it was not read further.`,
        { maxDepth: this.maxDepth },
      );
    }

    this.attributeSource = buffer;
    this.attributeStart = nameEnd;
    this.attributeEnd = contentEnd;
    this.handler.start(name);
    // Attributes are only readable during the callback, and holding the chunk here would pin it in memory.
    this.attributeSource = '';
    this.attributeStart = 0;
    this.attributeEnd = 0;

    if (selfClosing) {
      this.openElements.pop();
      this.handler.end(name);
    }
  }

  private handleEndTag(buffer: string, nameStart: number, tagEnd: number): void {
    let nameEnd = nameStart;
    while (nameEnd < tagEnd && !isWhitespaceCode(buffer.charCodeAt(nameEnd))) {
      nameEnd++;
    }
    const name = buffer.slice(localNameStart(buffer, nameStart, nameEnd), nameEnd);
    const open = this.openElements.pop();
    if (open === undefined) {
      throw malformed(`</${name}> with no matching open element`);
    }
    if (open !== name) {
      throw malformed(`</${name}> where </${open}> was expected`);
    }
    this.handler.end(name);
  }

  /**
   * Consume character data. A run that another tag terminates is complete and decodes in full; the tail of the
   * buffer is only complete at the end of the input, otherwise a partial reference stays for the next chunk.
   */
  private consumeText(buffer: string, start: number, end: number, isComplete: boolean): number {
    let textEnd = end;
    if (!isComplete) {
      this.needMoreInput = true;
      textEnd = textEndBeforePartialEntity(buffer, start, end);
    }
    this.appendText(buffer, start, textEnd, true);
    return textEnd;
  }

  private appendText(buffer: string, start: number, end: number, decodeReferences: boolean): void {
    if (end <= start) {
      return;
    }
    const raw = buffer.slice(start, end);
    const value = decodeReferences && raw.indexOf('&') >= 0 ? decodeEntities(raw) : raw;
    if (this.pendingText.length + value.length > this.maxTextLength) {
      throw new XlsxError(
        'LIMIT_EXCEEDED',
        `This file contains a single piece of text longer than the ${this.maxTextLength} character limit and was not read further.`,
        { maxTextLength: this.maxTextLength },
      );
    }
    this.pendingText = this.pendingText.length === 0 ? value : this.pendingText + value;
  }

  private flushText(): void {
    if (this.pendingText.length === 0) {
      return;
    }
    const text = this.pendingText;
    this.pendingText = '';
    this.handler.text(text);
  }
}
