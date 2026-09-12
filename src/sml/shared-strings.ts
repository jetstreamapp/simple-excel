import type { SharedStringBudget } from '../types';
import { encodeCellText, needsSpacePreserve } from '../xml/escape';
import { encodeXmlChunk } from '../xml/utf8';
import type { ZipEntryWriter } from '../zip/zip-writer';

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const SPREADSHEET_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

/** Serialization joins and encodes at this many UTF-16 units, so a million-entry table is never one giant string. */
const FLUSH_THRESHOLD_CHARS = 64 * 1024;

/** What `intern` returns for a string the caller has to write inline. */
const WRITE_INLINE = -1;

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
 *
 * Freezing is one-way and only stops *new* strings: a string already in the table keeps resolving to its index, so a
 * sheet that crosses the budget mid-way mixes `t="s"` and `t="inlineStr"` cells, which is legal and which every
 * reader in the oracle accepts.
 */
export class SharedStringWriter {
  private readonly maxUnique: number;
  private readonly maxChars: number;
  private readonly maxLength: number;
  private readonly indexByText = new Map<string, number>();
  private readonly texts: string[] = [];
  private references = 0;
  private internedChars = 0;
  private frozen = false;

  constructor(budget: Required<SharedStringBudget>) {
    this.maxUnique = budget.maxUnique;
    this.maxChars = budget.maxChars;
    this.maxLength = budget.maxLength;
  }

  /** Index of `text` in the table, or -1 when the caller must write it inline. */
  intern(text: string): number {
    if (text.length > this.maxLength) {
      return WRITE_INLINE;
    }
    const existing = this.indexByText.get(text);
    if (existing !== undefined) {
      this.references++;
      return existing;
    }
    if (this.frozen) {
      return WRITE_INLINE;
    }
    if (this.texts.length >= this.maxUnique || this.internedChars + text.length > this.maxChars) {
      this.frozen = true;
      return WRITE_INLINE;
    }
    const index = this.texts.length;
    this.texts.push(text);
    this.indexByText.set(text, index);
    this.internedChars += text.length;
    this.references++;
    return index;
  }

  get stats(): SharedStringStats {
    return { count: this.references, uniqueCount: this.texts.length, frozen: this.frozen };
  }

  /** Write `xl/sharedStrings.xml` into an open zip entry. */
  async writeTo(entry: ZipEntryWriter): Promise<void> {
    const pending: string[] = [];
    let pendingChars = 0;

    const flush = async (): Promise<void> => {
      if (pendingChars === 0) {
        return;
      }
      const text = pending.join('');
      pending.length = 0;
      pendingChars = 0;
      await entry.write(encodeXmlChunk(text));
    };

    const append = (fragment: string): void => {
      pending.push(fragment);
      pendingChars += fragment.length;
    };

    append(`${XML_DECLARATION}<sst xmlns="${SPREADSHEET_NS}" count="${this.references}" uniqueCount="${this.texts.length}">`);
    for (const text of this.texts) {
      const encoded = encodeCellText(text);
      append(needsSpacePreserve(text) ? `<si><t xml:space="preserve">${encoded}</t></si>` : `<si><t>${encoded}</t></si>`);
      if (pendingChars >= FLUSH_THRESHOLD_CHARS) {
        await flush();
      }
    }
    append('</sst>');
    await flush();
  }
}
