import type { SniffResult } from '../types';

const ZIP_LOCAL_FILE_HEADER: readonly number[] = [0x50, 0x4b, 0x03, 0x04];
/** An archive with no entries at all: the end-of-central-directory record is the whole file. */
const ZIP_EMPTY_ARCHIVE: readonly number[] = [0x50, 0x4b, 0x05, 0x06];
/** Compound File Binary: legacy `.xls` and every encrypted Office document. */
const CFB_HEADER: readonly number[] = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
const UTF8_BOM: readonly number[] = [0xef, 0xbb, 0xbf];
/** Excel's "Unicode Text" save and many Windows exports: UTF-16 with a byte-order mark. */
const UTF16LE_BOM: readonly number[] = [0xff, 0xfe];
const UTF16BE_BOM: readonly number[] = [0xfe, 0xff];

const XML_MARKER = '<?xml';
const HTML_MARKERS: readonly string[] = ['<html', '<!doctype html', '<table'];

/** How many bytes decide whether an unrecognized input is text; a CSV header line is well inside it. */
const TEXT_SAMPLE_BYTES = 512;
/** Real text is rarely below this; binary formats rarely reach it. */
const TEXT_BYTE_RATIO = 0.9;
/** A multi-byte UTF-8 sequence is at most four bytes, so the sample can end mid-character by at most three. */
const MAX_UTF8_SEQUENCE_TAIL = 3;
/**
 * How many control bytes (other than tab, CR and LF) a sample that is not UTF-8 may hold and still be legacy text
 * (Windows-1252, Shift-JIS, ...). Such text has next to none; binary has about one byte in nine.
 */
const LEGACY_TEXT_CONTROL_RATIO = 0.02;
/**
 * The share of UTF-16 code units that must be ASCII. Any two bytes make some code point, so printable code points
 * alone prove nothing; delimited text always has its commas, tabs and line breaks, and binary almost never has a
 * high byte of zero (one unit in 256).
 */
const UTF16_MIN_ASCII_RATIO = 0.05;

/** CFB directory entry names are UTF-16LE, so the stream that marks an encrypted package is searched as bytes. */
const ENCRYPTED_PACKAGE_MARKER: Uint8Array = utf16leBytes('EncryptedPackage');

function utf16leBytes(text: string): Uint8Array {
  const bytes = new Uint8Array(text.length * 2);
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    bytes[i * 2] = code & 0xff;
    bytes[i * 2 + 1] = code >>> 8;
  }
  return bytes;
}

function startsWithBytes(bytes: Uint8Array, prefix: readonly number[]): boolean {
  if (bytes.length < prefix.length) {
    return false;
  }
  for (let i = 0; i < prefix.length; i++) {
    if (bytes[i] !== prefix[i]) {
      return false;
    }
  }
  return true;
}

function containsBytes(haystack: Uint8Array, needle: Uint8Array): boolean {
  const firstByte = needle[0];
  if (firstByte === undefined) {
    return true;
  }
  const lastStart = haystack.length - needle.length;
  for (let start = 0; start <= lastStart; start++) {
    if (haystack[start] !== firstByte) {
      continue;
    }
    let offset = 1;
    while (offset < needle.length && haystack[start + offset] === needle[offset]) {
      offset++;
    }
    if (offset === needle.length) {
      return true;
    }
  }
  return false;
}

function isWhitespaceByte(byte: number): boolean {
  return byte === 0x20 || byte === 0x09 || byte === 0x0a || byte === 0x0d;
}

function isPrintableAsciiOrWhitespace(byte: number): boolean {
  return (byte >= 0x20 && byte <= 0x7e) || byte === 0x09 || byte === 0x0a || byte === 0x0d;
}

function skipWhitespace(bytes: Uint8Array, start: number): number {
  let index = start;
  while (index < bytes.length) {
    const byte = bytes[index];
    if (byte === undefined || !isWhitespaceByte(byte)) {
      break;
    }
    index++;
  }
  return index;
}

/** Case-insensitive ASCII comparison; `marker` must be lower case. */
function matchesAsciiMarker(bytes: Uint8Array, start: number, marker: string): boolean {
  if (start + marker.length > bytes.length) {
    return false;
  }
  for (let i = 0; i < marker.length; i++) {
    const byte = bytes[start + i] ?? 0;
    const lowered = byte >= 0x41 && byte <= 0x5a ? byte + 0x20 : byte;
    if (lowered !== marker.charCodeAt(i)) {
      return false;
    }
  }
  return true;
}

function decodesAsUtf8(sample: Uint8Array): boolean {
  // The sample is a cut of a longer input, so a trailing partial sequence is expected rather than a decode failure.
  let end = sample.length;
  for (let trimmed = 0; trimmed < MAX_UTF8_SEQUENCE_TAIL && end > 0; trimmed++) {
    const byte = sample[end - 1] ?? 0;
    if (byte < 0x80) {
      break;
    }
    end--;
    // A lead byte ends the trim: everything after it belonged to the sequence it started.
    if (byte >= 0xc0) {
      break;
    }
  }
  try {
    new TextDecoder('utf-8', { fatal: true }).decode(sample.subarray(0, end));
    return true;
  } catch {
    return false;
  }
}

/**
 * Text is what a person could have typed: printable ASCII and the three whitespace controls, plus any non-ASCII
 * byte when the sample really is UTF-8 (a CJK or accented CSV is mostly high bytes and still text). A sample that is
 * not UTF-8 is still text in a legacy encoding (EC-SNIFF-LEGACY-TEXT: a Windows-1252 or Shift-JIS CSV) when it has
 * no NUL byte and almost no other control bytes, which is what tells it apart from binary.
 */
function isProbablyText(sample: Uint8Array): boolean {
  if (sample.length === 0) {
    return true;
  }
  if (decodesAsUtf8(sample)) {
    let textBytes = 0;
    for (const byte of sample) {
      if (byte >= 0x80 || isPrintableAsciiOrWhitespace(byte)) {
        textBytes++;
      }
    }
    return textBytes / sample.length >= TEXT_BYTE_RATIO;
  }
  let controlBytes = 0;
  for (const byte of sample) {
    if (byte === 0x00) {
      return false;
    }
    if (byte < 0x80 && !isPrintableAsciiOrWhitespace(byte)) {
      controlBytes++;
    }
  }
  return controlBytes / sample.length <= LEGACY_TEXT_CONTROL_RATIO;
}

/** The UTF-16 code unit at `index` (a byte offset), or -1 past the end. */
function utf16UnitAt(bytes: Uint8Array, index: number, littleEndian: boolean): number {
  const first = bytes[index];
  const second = bytes[index + 1];
  if (first === undefined || second === undefined) {
    return -1;
  }
  return littleEndian ? first | (second << 8) : (first << 8) | second;
}

/** `matchesAsciiMarker` for UTF-16 text: each marker character is one code unit. `marker` must be lower case. */
function matchesUtf16Marker(bytes: Uint8Array, start: number, marker: string, littleEndian: boolean): boolean {
  for (let i = 0; i < marker.length; i++) {
    const unit = utf16UnitAt(bytes, start + i * 2, littleEndian);
    const lowered = unit >= 0x41 && unit <= 0x5a ? unit + 0x20 : unit;
    if (lowered !== marker.charCodeAt(i)) {
      return false;
    }
  }
  return true;
}

/**
 * UTF-16 with a byte-order mark (EC-SNIFF-LEGACY-TEXT): Excel's "Unicode Text" export is tab-separated UTF-16LE. The
 * XML and HTML markers are looked for in code units, and the rest is text when its code units are overwhelmingly
 * printable (no controls, unpaired surrogates or noncharacters) and include the ASCII punctuation delimited text
 * always has: a BOM alone is two bytes that binary can start with too.
 */
function sniffUtf16(head: Uint8Array, littleEndian: boolean): SniffResult {
  const contentStart = UTF16LE_BOM.length;
  if (contentStart >= head.length) {
    return 'empty';
  }
  let markerStart = contentStart;
  for (;;) {
    const unit = utf16UnitAt(head, markerStart, littleEndian);
    if (unit === -1 || !isWhitespaceByte(unit)) {
      break;
    }
    markerStart += 2;
  }
  if (matchesUtf16Marker(head, markerStart, XML_MARKER, littleEndian)) {
    return 'xml';
  }
  for (const marker of HTML_MARKERS) {
    if (matchesUtf16Marker(head, markerStart, marker, littleEndian)) {
      return 'html';
    }
  }
  const sampleEnd = Math.min(head.length, contentStart + TEXT_SAMPLE_BYTES);
  let units = 0;
  let textUnits = 0;
  let asciiUnits = 0;
  let pendingHighSurrogate = false;
  for (let index = contentStart; index + 1 < sampleEnd; index += 2) {
    const unit = utf16UnitAt(head, index, littleEndian);
    units++;
    const isLowSurrogate = unit >= 0xdc00 && unit <= 0xdfff;
    if (pendingHighSurrogate) {
      // The high surrogate is text only when its low half follows.
      textUnits += isLowSurrogate ? 2 : 0;
      pendingHighSurrogate = false;
      if (isLowSurrogate) {
        continue;
      }
    }
    if (unit >= 0xd800 && unit <= 0xdbff) {
      pendingHighSurrogate = true;
      continue;
    }
    // C0 and C1 controls (other than the three whitespace ones), lone low surrogates and the two noncharacters are
    // what binary looks like.
    const isControl =
      (unit < 0x20 && !isWhitespaceByte(unit)) || (unit >= 0x7f && unit < 0xa0) || isLowSurrogate || unit === 0xfffe || unit === 0xffff;
    if (!isControl) {
      textUnits++;
    }
    if (unit < 0x80 && !isControl) {
      asciiUnits++;
    }
  }
  // A sample cut in the middle of a surrogate pair still counts the half it has.
  if (pendingHighSurrogate) {
    textUnits++;
  }
  if (units === 0) {
    return 'unknown';
  }
  return textUnits / units >= TEXT_BYTE_RATIO && asciiUnits / units >= UTF16_MIN_ASCII_RATIO ? 'text' : 'unknown';
}

/**
 * Classify the first bytes of an input: `PK\x03\x04` -> `'zip'` (the caller looks at the content types to decide
 * xlsx/xlsb/ods), `PK\x05\x06` -> `'empty'`, CFB magic -> `'cfb-encrypted'` when the first 64 KiB contain the
 * UTF-16LE directory name `EncryptedPackage`, else `'cfb-legacy'`; `<?xml`/`<html`/`<table` -> `'xml'`/`'html'`;
 * zero length -> `'empty'`; mostly-printable text -> `'text'` (UTF-8, UTF-16 with a BOM, or a legacy single- or
 * multi-byte encoding such as Windows-1252 or Shift-JIS); anything else -> `'unknown'`.
 *
 * `SniffResult` also has `'xlsx' | 'xlsb' | 'ods'`, which this function never returns: every one of them is a zip,
 * and only `[Content_Types].xml` (or an `mimetype` entry) tells them apart, so the facade decides that after opening
 * the archive. Pass at least `SNIFF_BYTES` bytes - the encrypted-vs-legacy decision scans all of them (ADR-007).
 */
export function sniff(head: Uint8Array): SniffResult {
  if (head.length === 0) {
    return 'empty';
  }
  if (startsWithBytes(head, ZIP_LOCAL_FILE_HEADER)) {
    return 'zip';
  }
  if (startsWithBytes(head, ZIP_EMPTY_ARCHIVE)) {
    return 'empty';
  }
  if (startsWithBytes(head, CFB_HEADER)) {
    return containsBytes(head, ENCRYPTED_PACKAGE_MARKER) ? 'cfb-encrypted' : 'cfb-legacy';
  }
  if (startsWithBytes(head, UTF16LE_BOM) || startsWithBytes(head, UTF16BE_BOM)) {
    return sniffUtf16(head, head[0] === UTF16LE_BOM[0]);
  }

  const contentStart = startsWithBytes(head, UTF8_BOM) ? UTF8_BOM.length : 0;
  if (contentStart >= head.length) {
    return 'empty';
  }
  const markerStart = skipWhitespace(head, contentStart);
  if (matchesAsciiMarker(head, markerStart, XML_MARKER)) {
    return 'xml';
  }
  for (const marker of HTML_MARKERS) {
    if (matchesAsciiMarker(head, markerStart, marker)) {
      return 'html';
    }
  }
  return isProbablyText(head.subarray(contentStart, contentStart + TEXT_SAMPLE_BYTES)) ? 'text' : 'unknown';
}

/** How many leading bytes `sniff` wants (64 KiB, for the CFB directory scan). */
export const SNIFF_BYTES: number = 65_536;
