import { notImplemented } from '../internal/not-implemented';
import type { SharedStringBudget } from '../types';
import type { ZipEntryWriter } from '../zip/zip-writer';

export interface SharedStringStats {
  /** Total references (`count` attribute). */
  readonly count: number;
  readonly uniqueCount: number;
  /** True once the budget stopped new strings from being interned. */
  readonly frozen: boolean;
}

/**
 * Bounded shared-string table (ADR-001). `intern` returns the index for strings it accepts and -1 for strings the
 * caller must write inline (too long, or the budget is exhausted and the string is new). Serialization is chunked
 * (64 Ki-char string builder flushed to the entry), never one giant string.
 */
export class SharedStringWriter {
  constructor(budget: Required<SharedStringBudget>) {
    void budget;
    throw notImplemented('sml/shared-strings');
  }

  intern(text: string): number {
    void text;
    throw notImplemented('sml/shared-strings');
  }

  get stats(): SharedStringStats {
    throw notImplemented('sml/shared-strings');
  }

  /** Write `xl/sharedStrings.xml` into an open zip entry. */
  writeTo(entry: ZipEntryWriter): Promise<void> {
    void entry;
    throw notImplemented('sml/shared-strings');
  }
}

export interface ParseSharedStringsOptions {
  readonly maxChars: number;
  readonly maxDepth: number;
  readonly maxTextLength: number;
}

/**
 * Read `xl/sharedStrings.xml` into an array: `<r>` runs concatenated, `<rPh>` phonetic runs skipped, `_xHHHH_`
 * decoded. Whitespace-only `<t>` without `xml:space` is kept verbatim (Excel preserves it; calamine trims - catalog).
 */
export function parseSharedStrings(chunks: AsyncIterable<Uint8Array>, options: ParseSharedStringsOptions): Promise<string[]> {
  void chunks;
  void options;
  throw notImplemented('sml/shared-strings');
}
