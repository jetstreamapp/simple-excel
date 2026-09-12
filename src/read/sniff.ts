import { notImplemented } from '../internal/not-implemented';
import type { SniffResult } from '../types';

/**
 * Classify the first bytes of an input: `PK\x03\x04` -> `'zip'` (the caller looks at the content types to decide
 * xlsx/xlsb/ods), `PK\x05\x06` -> `'empty'`, CFB magic -> `'cfb-encrypted'` when the first 64 KiB contain the
 * UTF-16LE directory name `EncryptedPackage`, else `'cfb-legacy'`; `<?xml`/`<html`/`<table` -> `'xml'`/`'html'`;
 * zero length -> `'empty'`; mostly-printable text -> `'text'`; anything else -> `'unknown'`.
 */
export function sniff(head: Uint8Array): SniffResult {
  void head;
  throw notImplemented('read/sniff');
}

/** How many leading bytes `sniff` wants (64 KiB, for the CFB directory scan). */
export const SNIFF_BYTES: number = 65_536;
