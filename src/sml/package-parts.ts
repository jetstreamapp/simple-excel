import { notImplemented } from '../internal/not-implemented';
import type { WorkbookProperties } from '../types';

export interface SheetPartInfo {
  readonly name: string;
  /** 1-based, stable id in workbook.xml. */
  readonly sheetId: number;
  /** Relationship id, `rId<n>`. */
  readonly relId: string;
  /** Part name relative to `xl/` (`worksheets/sheet1.xml`). */
  readonly path: string;
  readonly hidden: boolean;
  /** Present when the sheet has an autofilter (emits the `_xlnm._FilterDatabase` defined name). */
  readonly autoFilterRange?: string;
}

// ---- writers (all sync, all return complete XML documents) ---------------------------------------------------------

export function contentTypesXml(sheets: readonly SheetPartInfo[], hasSharedStrings: boolean): string {
  void sheets;
  void hasSharedStrings;
  throw notImplemented('sml/package-parts');
}

export function rootRelsXml(): string {
  throw notImplemented('sml/package-parts');
}

export function workbookXml(sheets: readonly SheetPartInfo[], date1904: boolean): string {
  void sheets;
  void date1904;
  throw notImplemented('sml/package-parts');
}

export function workbookRelsXml(sheets: readonly SheetPartInfo[], hasSharedStrings: boolean): string {
  void sheets;
  void hasSharedStrings;
  throw notImplemented('sml/package-parts');
}

/** `docProps/core.xml`; `now` is the created/modified timestamp (fixed when deterministic). */
export function coreXml(properties: WorkbookProperties, now: Date): string {
  void properties;
  void now;
  throw notImplemented('sml/package-parts');
}

export function appXml(sheets: readonly SheetPartInfo[]): string {
  void sheets;
  throw notImplemented('sml/package-parts');
}

// ---- readers -------------------------------------------------------------------------------------------------------

export interface Relationship {
  readonly id: string;
  /** Type URI; compare with `relTypeIs` so Strict and Transitional URIs both match. */
  readonly type: string;
  /** Resolved zip entry name (relative to the package root, no leading slash, slashes normalized). */
  readonly target: string;
  readonly external: boolean;
}

/** Parse a `.rels` part; `basePath` is the directory of the source part (`'xl/'` for workbook rels, `''` for root). */
export function parseRels(xml: string, basePath: string): Relationship[] {
  void xml;
  void basePath;
  throw notImplemented('sml/package-parts');
}

/** True when a relationship type ends with `/relationships/<suffix>` (works for Strict and Transitional). */
export function relTypeIs(type: string, suffix: string): boolean {
  void type;
  void suffix;
  throw notImplemented('sml/package-parts');
}

export interface ContentTypes {
  readonly defaults: ReadonlyMap<string, string>;
  readonly overrides: ReadonlyMap<string, string>;
  /** Content type for a part name, via override then default extension; undefined when unknown. */
  typeOf(partName: string): string | undefined;
}

export function parseContentTypes(xml: string): ContentTypes {
  void xml;
  throw notImplemented('sml/package-parts');
}

export interface WorkbookSheetEntry {
  readonly name: string;
  readonly sheetId: number;
  readonly relId: string;
  readonly state: 'visible' | 'hidden' | 'veryHidden';
}

export interface ParsedWorkbook {
  readonly sheets: readonly WorkbookSheetEntry[];
  readonly date1904: boolean;
}

export function parseWorkbook(xml: string): ParsedWorkbook {
  void xml;
  throw notImplemented('sml/package-parts');
}
