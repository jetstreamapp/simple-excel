import { notImplemented } from '../internal/not-implemented';

/** Wall-clock components; `month` is 1-12. */
export interface DateComponents {
  readonly year: number;
  readonly month: number;
  readonly day: number;
  readonly hour: number;
  readonly minute: number;
  readonly second: number;
  readonly millisecond: number;
}

/** Serial for 1900-03-01 in the 1900 system: the first serial after Excel's fake 1900-02-29 (serial 60). */
export const SERIAL_1900_03_01: number = 61;
/** Difference between the two date systems: 1904-01-01 is serial 0 in the 1904 system and 1462 in the 1900 system. */
export const SERIAL_1904_OFFSET: number = 1462;

/**
 * Excel serial (days since the epoch, fraction = time of day) for a wall-clock instant. Rounded to the nearest
 * millisecond. Uses the 1900 system unless `date1904`; dates before the epoch (1900-01-01 or 1904-01-01) and the
 * non-existent 1900-02-29 return null. Time-only inputs (year 1899, month 12, day 30) yield a fraction in [0, 1).
 */
export function serialFromComponents(components: DateComponents, date1904: boolean): number | null {
  void components;
  void date1904;
  throw notImplemented('sml/date');
}

/**
 * Inverse of `serialFromComponents`. Serial 60 in the 1900 system (Excel's fake leap day) maps to 1900-03-01 and
 * serials below 61 shift by one day, matching what every other reader does. Returns null for negative serials
 * (Excel displays those as `#####`) and non-finite input. Fractions are rounded to the nearest millisecond.
 */
export function componentsFromSerial(serial: number, date1904: boolean): DateComponents | null {
  void serial;
  void date1904;
  throw notImplemented('sml/date');
}

/** Build a JS Date whose local (`'local'`) or UTC (`'utc'`) fields equal the components. */
export function dateFromComponents(components: DateComponents, fields: 'local' | 'utc'): Date {
  void components;
  void fields;
  throw notImplemented('sml/date');
}

/** Read the local or UTC fields of a Date as components. Invalid dates return null. */
export function componentsFromDate(date: Date, fields: 'local' | 'utc'): DateComponents | null {
  void date;
  void fields;
  throw notImplemented('sml/date');
}

/**
 * Parse an ISO 8601 date/time from a Strict `t="d"` cell: `YYYY-MM-DD`, `YYYY-MM-DDThh:mm:ss` with any number of
 * fractional digits, or a bare `hh:mm:ss` time (day 1899-12-30). Returns null when malformed.
 */
export function componentsFromIso(text: string): DateComponents | null {
  void text;
  throw notImplemented('sml/date');
}
