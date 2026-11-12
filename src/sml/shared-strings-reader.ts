import { XlsxError } from '../errors';
import { isStringLengthError } from '../internal/string-length';
import { decodeCellText } from '../xml/escape';
import { XmlTokenizer } from '../xml/tokenizer';

export interface ParseSharedStringsOptions {
  /** The table's budget: every item's text plus `SHARED_STRING_ITEM_CHARS` for each item (EC-SST-EMPTY-ITEM-FLOOD). */
  readonly maxChars: number;
  readonly maxDepth: number;
  readonly maxTextLength: number;
}

/**
 * A `uniqueCount` past this is either a lie or a hostile file, and preallocating from it is exactly what would hurt;
 * the array still grows on its own past the hint.
 */
const MAX_PREALLOCATED_STRINGS = 4_000_000;

/**
 * What one `<si>` costs on top of its text, in UTF-16 units counted toward `maxChars`. Measured in Node 24: an empty
 * item holds about 12 bytes (an 8-byte array slot plus growth slack) and a short one about 28 bytes beyond its
 * characters (the slot plus a string header). 16 units is 16 to 32 bytes, so the budget stays a bound on memory
 * however the table splits its text, and 100M empty `<si/>` (a 713 KB upload that reached 2 to 3.6 GB) stop after
 * about 16.7M items under the default 256 Mi (EC-SST-EMPTY-ITEM-FLOOD).
 */
export const SHARED_STRING_ITEM_CHARS: number = 16;

/**
 * Read `xl/sharedStrings.xml` into an array: `<r>` runs concatenated, `<rPh>` phonetic runs skipped, `_xHHHH_`
 * decoded. Whitespace-only `<t>` without `xml:space` is kept verbatim (Excel preserves it; calamine trims - catalog).
 *
 * The part is streamed: each chunk is decoded and pushed straight into the tokenizer, so nothing beyond the finished
 * table is held (EC-RICH-TEXT-RUNS). Hostile parts surface as classified errors: `XML_DOCTYPE` for a DTD
 * (EC-XXE-DOCTYPE), `LIMIT_EXCEEDED` for nesting past `maxDepth` (EC-XML-DEEP-NESTING) or a table past `maxChars`,
 * where every item counts its text plus `SHARED_STRING_ITEM_CHARS` (EC-SST-EMPTY-ITEM-FLOOD).
 */
export async function parseSharedStrings(chunks: AsyncIterable<Uint8Array>, options: ParseSharedStringsOptions): Promise<string[]> {
  let strings: string[] = [];
  let stringCount = 0;
  let insideString = false;
  let phoneticDepth = 0;
  let insideText = false;
  let current = '';
  let totalChars = 0;

  /** Count `units` toward the table's budget, failing once it is spent. */
  function charge(units: number): void {
    totalChars += units;
    if (totalChars > options.maxChars) {
      throw new XlsxError(
        'LIMIT_EXCEEDED',
        `This workbook's shared-string table is larger than the ${options.maxChars} character limit (each item counts ${SHARED_STRING_ITEM_CHARS} on top of its text), so it was not read further.`,
        { maxChars: options.maxChars, items: stringCount },
      );
    }
  }

  const tokenizer = new XmlTokenizer(
    {
      start(name: string): void {
        switch (name) {
          case 'sst': {
            const uniqueCount = Number.parseInt(tokenizer.attr('uniqueCount') ?? '', 10);
            if (Number.isFinite(uniqueCount) && uniqueCount > 0) {
              strings = Array.from<string>({ length: Math.min(uniqueCount, MAX_PREALLOCATED_STRINGS) });
            }
            break;
          }
          case 'si':
            // Charged on the way in, so a flood of empty items is held to the budget as well as long text is.
            charge(SHARED_STRING_ITEM_CHARS);
            insideString = true;
            current = '';
            break;
          // Phonetic runs are the furigana Excel shows above the text, never part of the value.
          case 'rPh':
            phoneticDepth++;
            break;
          case 't':
            insideText = insideString && phoneticDepth === 0;
            break;
          default:
            break;
        }
      },
      text(text: string): void {
        if (!insideText) {
          return;
        }
        charge(text.length);
        // The tokenizer caps each run; an item built from many runs is held to the same cap (EC-CELL-TEXT-RUN-SPLIT).
        if (current.length + text.length > options.maxTextLength) {
          throw new XlsxError(
            'LIMIT_EXCEEDED',
            `Shared string ${stringCount} of this workbook is longer than the ${options.maxTextLength} character limit, so it was not read further.`,
            { index: stringCount, maxTextLength: options.maxTextLength },
          );
        }
        current = current.length === 0 ? text : current + text;
      },
      end(name: string): void {
        switch (name) {
          case 't':
            insideText = false;
            break;
          case 'rPh':
            phoneticDepth--;
            break;
          case 'si':
            insideString = false;
            // An `<si>` with no `<t>` at all is a legal empty string. Scanning for `_x` is far cheaper than the
            // decode, and escapes are rare, so only the strings that could hold one pay for it.
            strings[stringCount++] = current.includes('_x') ? decodeCellText(current) : current;
            current = '';
            break;
          default:
            break;
        }
      },
    },
    { maxDepth: options.maxDepth, maxTextLength: options.maxTextLength },
  );

  const decoder = new TextDecoder('utf-8');
  try {
    for await (const chunk of chunks) {
      tokenizer.push(decoder.decode(chunk, { stream: true }));
    }
    // Flush whatever the decoder held back from a multi-byte character split across the last chunk boundary.
    tokenizer.push(decoder.decode());
    tokenizer.end();
  } catch (error) {
    // Only reachable with limits raised past what one string can hold; still a limit, not a crash.
    if (isStringLengthError(error)) {
      throw new XlsxError(
        'LIMIT_EXCEEDED',
        "This workbook's shared text holds more than a string can, so it was not read further.",
        undefined,
        { cause: error },
      );
    }
    throw error;
  }

  strings.length = stringCount;
  return strings;
}
