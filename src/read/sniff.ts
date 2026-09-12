import type { SniffResult } from '../types';

const ZIP_LOCAL_FILE_HEADER: readonly number[] = [0x50, 0x4b, 0x03, 0x04];
/** An archive with no entries at all: the end-of-central-directory record is the whole file. */
const ZIP_EMPTY_ARCHIVE: readonly number[] = [0x50, 0x4b, 0x05, 0x06];
/** Compound File Binary: legacy `.xls` and every encrypted Office document. */
const CFB_HEADER: readonly number[] = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
const UTF8_BOM: readonly number[] = [0xef, 0xbb, 0xbf];

const XML_MARKER = '<?xml';
const HTML_MARKERS: readonly string[] = ['<html', '<!doctype html', '<table'];

/** How many bytes decide whether an unrecognized input is text; a CSV header line is well inside it. */
const TEXT_SAMPLE_BYTES = 512;
/** Real text is rarely below this; binary formats rarely reach it. */
const TEXT_BYTE_RATIO = 0.9;
/** A multi-byte UTF-8 sequence is at most four bytes, so the sample can end mid-character by at most three. */
const MAX_UTF8_SEQUENCE_TAIL = 3;

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
 * byte when the sample really is UTF-8 (a CJK or accented CSV is mostly high bytes and still text).
 */
function isProbablyText(sample: Uint8Array): boolean {
  if (sample.length === 0) {
    return true;
  }
  const highBytesAreText = decodesAsUtf8(sample);
  let textBytes = 0;
  for (const byte of sample) {
    if ((byte >= 0x20 && byte <= 0x7e) || byte === 0x09 || byte === 0x0a || byte === 0x0d) {
      textBytes++;
    } else if (byte >= 0x80 && highBytesAreText) {
      textBytes++;
    }
  }
  return textBytes / sample.length >= TEXT_BYTE_RATIO;
}

/**
 * Classify the first bytes of an input: `PK\x03\x04` -> `'zip'` (the caller looks at the content types to decide
 * xlsx/xlsb/ods), `PK\x05\x06` -> `'empty'`, CFB magic -> `'cfb-encrypted'` when the first 64 KiB contain the
 * UTF-16LE directory name `EncryptedPackage`, else `'cfb-legacy'`; `<?xml`/`<html`/`<table` -> `'xml'`/`'html'`;
 * zero length -> `'empty'`; mostly-printable text -> `'text'`; anything else -> `'unknown'`.
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
