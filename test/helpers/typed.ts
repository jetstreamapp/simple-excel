/**
 * The typed-dump encoding shared with `oracle/lib/typed.mjs` and `fixtures/canonical/canonical.mjs`:
 *   null | string | number | boolean | {$date} | {$datetime} | {$time} | {$error} | {$formula, cached}
 */
import type { CellError, CellValue } from '../../src/types';
import type { TypedValue } from '../../fixtures/canonical/canonical.mjs';

export type { TypedValue };

export interface TypedSheet {
  name: string;
  hidden: boolean;
  rows: TypedValue[][];
  merges?: string[];
  autoFilter?: string | null;
  hiddenRows?: number[];
  hiddenColumns?: string[];
  freeze?: { rows: number; cols: number } | null;
  hyperlinks?: { cell: string; url: string }[];
  comments?: { cell: string; text: string }[];
  features?: unknown;
}

export interface TypedDump {
  reader: string;
  sheets: TypedSheet[];
}

export function isTyped(value: unknown, key: string): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && key in value;
}

const pad = (value: number, width = 2): string => String(value).padStart(width, '0');

/** Typed encoding from a JS Date, reading the wall clock from local or UTC getters. */
export function fromJsDate(value: Date, clock: 'local' | 'utc'): TypedValue {
  if (Number.isNaN(value.getTime())) {
    return '$invalidDate';
  }
  const parts =
    clock === 'utc'
      ? [
          value.getUTCFullYear(),
          value.getUTCMonth() + 1,
          value.getUTCDate(),
          value.getUTCHours(),
          value.getUTCMinutes(),
          value.getUTCSeconds(),
          value.getUTCMilliseconds(),
        ]
      : [
          value.getFullYear(),
          value.getMonth() + 1,
          value.getDate(),
          value.getHours(),
          value.getMinutes(),
          value.getSeconds(),
          value.getMilliseconds(),
        ];
  const [year, month, day, hour, minute, second, millisecond] = parts as [number, number, number, number, number, number, number];
  const timeText = `${pad(hour)}:${pad(minute)}:${pad(second)}.${pad(millisecond, 3)}`;
  if (year === 1899 && month === 12 && day === 30) {
    return { $time: timeText };
  }
  const dateText = `${pad(year, 4)}-${pad(month)}-${pad(day)}`;
  if (hour === 0 && minute === 0 && second === 0 && millisecond === 0) {
    return { $date: dateText };
  }
  return { $datetime: `${dateText}T${timeText}` };
}

/** Typed encoding of a value read through the public API (`errors: 'object'`). */
export function toTyped(value: CellValue | CellError, clock: 'local' | 'utc'): TypedValue {
  if (value instanceof Date) {
    return fromJsDate(value, clock);
  }
  if (typeof value === 'object' && value !== null) {
    return { $error: value.error };
  }
  return value;
}

/** Trim trailing null cells and trailing empty rows so dumps from sparse and dense readers compare equal. */
export function normalizeRows(rows: TypedValue[][]): TypedValue[][] {
  const trimmed = rows.map(row => {
    let end = row.length;
    while (end > 0 && row[end - 1] === null) {
      end--;
    }
    return row.slice(0, end);
  });
  let lastRow = trimmed.length;
  while (lastRow > 0 && trimmed[lastRow - 1]?.length === 0) {
    lastRow--;
  }
  return trimmed.slice(0, lastRow);
}

export function columnLetter(index0: number): string {
  let n = index0 + 1;
  let out = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

export function a1(row0: number, col0: number): string {
  return `${columnLetter(col0)}${row0 + 1}`;
}
