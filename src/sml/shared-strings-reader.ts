import { notImplemented } from '../internal/not-implemented';

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
