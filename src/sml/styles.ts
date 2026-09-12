import { notImplemented } from '../internal/not-implemented';
import type { CellStyle, StyleId } from '../types';

/**
 * Growable style registry seeded with Excel's mandatory defaults (font 0 Calibri 11, fills none + gray125, border 0
 * empty, cellXfs 0). Identical `CellStyle` specs resolve to the same id. Custom number formats get ids from 164;
 * codes that match a built-in resolve to the built-in id and emit no `<numFmt>`.
 */
export class StyleRegistry {
  constructor() {
    throw notImplemented('sml/styles');
  }

  /** Number of cellXfs entries (default style included). */
  get count(): number {
    throw notImplemented('sml/styles');
  }

  register(style: CellStyle): StyleId {
    void style;
    throw notImplemented('sml/styles');
  }

  /** The style id the writer uses for JS Dates when the caller did not give one: default xf + DEFAULT_DATE_FORMAT. */
  get defaultDateStyle(): StyleId {
    throw notImplemented('sml/styles');
  }

  /** The bold header style. */
  get headerStyle(): StyleId {
    throw notImplemented('sml/styles');
  }

  /** True when the xf carries a date/time number format (the writer needs this to serialize Dates under a caller style). */
  isDateStyle(styleId: StyleId): boolean {
    void styleId;
    throw notImplemented('sml/styles');
  }

  /** Complete `xl/styles.xml`, schema-valid (CT_Font child order, apply* attributes, cellStyles, dxfs, tableStyles). */
  toXml(): string {
    throw notImplemented('sml/styles');
  }
}

export interface ParsedStyles {
  /** `isDateByXf[i] === 1` when cellXfs[i] renders as a date/time. Length = cellXfs count. */
  readonly isDateByXf: Uint8Array;
  /** Custom number formats by id. */
  readonly numFmts: ReadonlyMap<number, string>;
}

/** Read what the reader needs from `xl/styles.xml`. Tolerates POI/LibreOffice/Numbers quirks (see catalog). */
export function parseStyles(xml: string): ParsedStyles {
  void xml;
  throw notImplemented('sml/styles');
}
