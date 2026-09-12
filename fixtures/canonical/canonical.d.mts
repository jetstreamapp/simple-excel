/** Hand-written declarations for canonical.mjs so the TypeScript test harness can import it. */
export type TypedValue =
  | null
  | string
  | number
  | boolean
  | { $date: string }
  | { $datetime: string }
  | { $time: string }
  | { $error: string }
  | { $formula: string; cached: TypedValue };

export interface CanonicalColumn {
  name: string;
  numFmt: string | null;
  value: (index: number) => TypedValue;
}

export interface CanonicalFeatures {
  merges: { s: { r: number; c: number }; e: { r: number; c: number } }[];
  freeze: { rows: number; cols: number };
  autoFilter: string;
  hyperlink: { cell: string; url: string };
  richText: { cell: string; runs: { text: string; bold?: boolean }[] };
  note: { cell: string; text: string };
  validation: { range: string; list: string[] };
  conditionalFormat: { range: string; operator: string; value: number; fillColor: string };
  hiddenRows: number[];
  hiddenColumns: string[];
  columnWidths: Record<string, number>;
}

export interface CanonicalSheet {
  name: string;
  hidden: boolean;
  rows: TypedValue[][];
  numFmts?: Record<string, string>;
  features?: CanonicalFeatures;
}

export interface CanonicalWorkbook {
  version: number;
  generatedBy: string;
  sheets: CanonicalSheet[];
}

export interface DateParts {
  y: number;
  m: number;
  d: number;
  hh: number;
  mm: number;
  ss: number;
  ms: number;
}

export const SHEET_DATA: string;
export const SHEET_FEATURES: string;
export const SHEET_LONG_NAME: string;
export const SHEET_HIDDEN: string;
export const ROW_COUNT: number;
export const COLUMNS: CanonicalColumn[];
export const EXCEL_MAX_CELL_CHARS: number;
export const TRUNCATION_SUFFIX: string;
export const FEATURES: CanonicalFeatures;
export function truncateForExcel<T>(value: T): T;
export function date(iso: string): { $date: string };
export function datetime(iso: string): { $datetime: string };
export function time(iso: string): { $time: string };
export function error(code: string): { $error: string };
export function formula(text: string, cached: TypedValue): { $formula: string; cached: TypedValue };
export function isTyped(value: unknown, key: string): boolean;
export function isTemporal(value: unknown): boolean;
export function dataRows(): TypedValue[][];
export function featureRows(): TypedValue[][];
export function longNameRows(): TypedValue[][];
export function hiddenSheetRows(): TypedValue[][];
export function canonicalWorkbook(): CanonicalWorkbook;
export function components(typed: { $date?: string; $datetime?: string; $time?: string }): DateParts;
export function toSerial(typed: { $date?: string; $datetime?: string; $time?: string }): number;
export function toLocalDate(typed: { $date?: string; $datetime?: string; $time?: string }): Date;
export function toUtcDate(typed: { $date?: string; $datetime?: string; $time?: string }): Date;
export function fromLocalDate(value: Date): TypedValue;
export function fromUtcDate(value: Date): TypedValue;
