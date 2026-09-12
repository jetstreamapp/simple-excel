/**
 * Shared "typed dump" format for every reader in the oracle. A dump is:
 *   {
 *     reader: 'sheetjs 0.20.3', file: '<path>', sheets: [
 *       { name, hidden, rows: [[typed, ...], ...],           // rows[0] is spreadsheet row 1; null = blank cell
 *         merges?: ['A1:C1'], hiddenRows?: [12], hiddenColumns?: ['D'], autoFilter?: 'A2:C12',
 *         freeze?: { rows, cols }, hyperlinks?: [{ cell, url }], comments?: [{ cell, text }],
 *         validations?: [{ range, type }], conditionalFormats?: [{ range }] }
 *     ]
 *   }
 * Cell values use the same typed encoding as fixtures/canonical/canonical.mjs.
 */
import { date, datetime, error, formula, isTyped, time } from '../../fixtures/canonical/canonical.mjs';

export { date, datetime, error, formula, isTyped, time };

const pad = (n, width = 2) => String(n).padStart(width, '0');

/** Typed encoding from a JS Date, reading wall-clock digits with the given getters ('local' | 'utc'). */
export function fromJsDate(value, clock = 'local') {
  const get =
    clock === 'utc'
      ? {
          y: value.getUTCFullYear(),
          m: value.getUTCMonth() + 1,
          d: value.getUTCDate(),
          hh: value.getUTCHours(),
          mm: value.getUTCMinutes(),
          ss: value.getUTCSeconds(),
          ms: value.getUTCMilliseconds(),
        }
      : {
          y: value.getFullYear(),
          m: value.getMonth() + 1,
          d: value.getDate(),
          hh: value.getHours(),
          mm: value.getMinutes(),
          ss: value.getSeconds(),
          ms: value.getMilliseconds(),
        };
  if (Number.isNaN(value.getTime())) {
    return { $invalidDate: true };
  }
  if (get.y === 1899 && get.m === 12 && get.d === 30) {
    return time(`${pad(get.hh)}:${pad(get.mm)}:${pad(get.ss)}.${pad(get.ms, 3)}`);
  }
  if (get.hh === 0 && get.mm === 0 && get.ss === 0 && get.ms === 0) {
    return date(`${pad(get.y, 4)}-${pad(get.m)}-${pad(get.d)}`);
  }
  return datetime(`${pad(get.y, 4)}-${pad(get.m)}-${pad(get.d)}T${pad(get.hh)}:${pad(get.mm)}:${pad(get.ss)}.${pad(get.ms, 3)}`);
}

const EXCEL_EPOCH_MS = Date.UTC(1899, 11, 30);

/** Typed encoding from an Excel serial number (1900 system incl. the leap-year bug). */
export function fromSerial(serial, date1904 = false) {
  let days = serial;
  if (date1904) {
    days += 1462;
  } else if (days < 61 && days >= 0) {
    days += 1; // no real 1900-02-29: serials 1..60 map one day later in a proleptic calendar
  }
  const ms = Math.round(days * 86400000);
  const asDate = new Date(EXCEL_EPOCH_MS + ms);
  return fromJsDate(asDate, 'utc');
}

/** Trim trailing null cells and trailing empty rows so dumps from sparse and dense readers compare equal. */
export function normalizeRows(rows) {
  const trimmed = rows.map(row => {
    let end = row.length;
    while (end > 0 && row[end - 1] === null) {
      end--;
    }
    return row.slice(0, end);
  });
  let lastRow = trimmed.length;
  while (lastRow > 0 && trimmed[lastRow - 1].length === 0) {
    lastRow--;
  }
  return trimmed.slice(0, lastRow);
}

/** A1 helpers */
export function columnLetter(index0) {
  let n = index0 + 1;
  let out = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

export function a1(row0, col0) {
  return `${columnLetter(col0)}${row0 + 1}`;
}

export function parseA1(ref) {
  const match = /^([A-Z]+)(\d+)$/.exec(ref);
  if (!match) {
    throw new Error(`bad A1 ref ${ref}`);
  }
  let col = 0;
  for (const ch of match[1]) {
    col = col * 26 + (ch.charCodeAt(0) - 64);
  }
  return { row0: Number(match[2]) - 1, col0: col - 1 };
}

export function stableStringify(value) {
  return JSON.stringify(value, null, 2) + '\n';
}
