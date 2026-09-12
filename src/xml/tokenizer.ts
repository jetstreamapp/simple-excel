import { notImplemented } from '../internal/not-implemented';

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

/**
 * Push tokenizer for the subset of XML that SpreadsheetML uses. Namespace prefixes are stripped from element and
 * attribute names (`x:c` -> `c`, `r:id` -> `id`) so Transitional and Strict parts look the same. Handles the five
 * predefined entities, numeric character references, CDATA, comments and processing instructions; rejects
 * `<!DOCTYPE` with `XlsxError('XML_DOCTYPE')`; reports structural problems as `XlsxError('XML_MALFORMED')`. Chunks
 * may split anywhere, including inside a tag, attribute value or entity.
 */
export class XmlTokenizer {
  constructor(handler: XmlTokenizerHandler, limits?: XmlTokenizerLimits) {
    void handler;
    void limits;
    throw notImplemented('xml/tokenizer');
  }

  push(chunk: string): void {
    void chunk;
    throw notImplemented('xml/tokenizer');
  }

  /** Flush trailing text and verify every element was closed. */
  end(): void {
    throw notImplemented('xml/tokenizer');
  }

  /** Attribute value by local name for the element whose `start` callback is running; entities decoded. */
  attr(name: string): string | undefined {
    void name;
    throw notImplemented('xml/tokenizer');
  }

  /** Current nesting depth (elements opened and not yet closed). */
  get depth(): number {
    throw notImplemented('xml/tokenizer');
  }
}
