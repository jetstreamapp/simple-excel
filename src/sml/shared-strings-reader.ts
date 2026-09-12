import { XlsxError } from '../errors';
import { decodeCellText } from '../xml/escape';
import { XmlTokenizer } from '../xml/tokenizer';

export interface ParseSharedStringsOptions {
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
 * Read `xl/sharedStrings.xml` into an array: `<r>` runs concatenated, `<rPh>` phonetic runs skipped, `_xHHHH_`
 * decoded. Whitespace-only `<t>` without `xml:space` is kept verbatim (Excel preserves it; calamine trims - catalog).
 *
 * The part is streamed: each chunk is decoded and pushed straight into the tokenizer, so nothing beyond the finished
 * table is held (EC-RICH-TEXT-RUNS). Hostile parts surface as classified errors: `XML_DOCTYPE` for a DTD
 * (EC-XXE-DOCTYPE), `LIMIT_EXCEEDED` for nesting past `maxDepth` (EC-XML-DEEP-NESTING) or more than `maxChars`
 * characters of text.
 */
export async function parseSharedStrings(chunks: AsyncIterable<Uint8Array>, options: ParseSharedStringsOptions): Promise<string[]> {
  let strings: string[] = [];
  let stringCount = 0;
  let insideString = false;
  let phoneticDepth = 0;
  let insideText = false;
  let current = '';
  let totalChars = 0;

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
        totalChars += text.length;
        if (totalChars > options.maxChars) {
          throw new XlsxError(
            'LIMIT_EXCEEDED',
            `This workbook's shared text is longer than the ${options.maxChars} character limit, so it was not read further.`,
            { maxChars: options.maxChars },
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
  for await (const chunk of chunks) {
    tokenizer.push(decoder.decode(chunk, { stream: true }));
  }
  // Flush whatever the decoder held back from a multi-byte character split across the last chunk boundary.
  tokenizer.push(decoder.decode());
  tokenizer.end();

  strings.length = stringCount;
  return strings;
}
