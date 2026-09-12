import { XlsxError } from '../errors';

export const MAX_SHEET_NAME_LENGTH: number = 31;

const FORBIDDEN_CHARACTERS = /[:\\/?*[\]]/g;
/** Excel keeps `History` for the change-tracking sheet and refuses it in any casing. */
const RESERVED_NAME = 'history';
/** A name made only of these can go into a formula unquoted. Anything else (spaces, accents, `'`) is quoted. */
const UNQUOTED_NAME = /^[A-Za-z_][A-Za-z0-9_]*$/;

/**
 * Trim whitespace and strip leading and trailing apostrophes, repeatedly: stripping an apostrophe can expose
 * whitespace and vice versa, so `" 'Name' "` and `"' Name '"` both end up as `Name`. Internal apostrophes stay
 * (EC-SHEET-NAME-31-APOSTROPHE).
 */
function stripEdges(name: string): string {
  let result = name;
  let previous = '';
  while (result !== previous) {
    previous = result;
    result = result.trim();
    if (result.startsWith("'")) {
      result = result.slice(1);
    }
    if (result.endsWith("'")) {
      result = result.slice(0, -1);
    }
  }
  return result;
}

/** Cut to `maxLength` UTF-16 units without leaving a lone high surrogate behind (an emoji is dropped whole). */
function truncate(name: string, maxLength: number): string {
  if (name.length <= maxLength) {
    return name;
  }
  const lastCode = name.charCodeAt(maxLength - 1);
  const isHighSurrogate = lastCode >= 0xd800 && lastCode <= 0xdbff;
  return name.slice(0, isHighSurrogate ? maxLength - 1 : maxLength);
}

function firstFreeDefaultName(taken: Set<string>): string {
  let index = taken.size + 1;
  while (taken.has(`sheet${index}`)) {
    index++;
  }
  return `Sheet${index}`;
}

function withCopySuffix(base: string, attempt: number): string {
  const suffix = ` (${attempt})`;
  return `${truncate(base, MAX_SHEET_NAME_LENGTH - suffix.length).trimEnd()}${suffix}`;
}

/**
 * Make a sheet name Excel accepts and unique within the workbook: `: \ / ? * [ ]` become `_`, leading and trailing
 * apostrophes are stripped, the name is trimmed to 31 characters, `History` (reserved, case-insensitive) becomes
 * `History_`, empty input becomes `Sheet<n>`. Collisions (case-insensitive against `taken`) get ` (2)`, ` (3)`, ...
 * fitted inside the 31-character budget. The returned name is added to `taken`.
 */
export function sanitizeSheetName(name: string, taken: Set<string>): string {
  if (typeof name !== 'string') {
    throw new XlsxError('INVALID_SHEET_NAME', 'A sheet name must be a string.', { name });
  }
  let candidate = stripEdges(name.replace(FORBIDDEN_CHARACTERS, '_'));
  if (candidate.toLowerCase() === RESERVED_NAME) {
    candidate = `${candidate}_`;
  }
  // Truncating can expose a trailing apostrophe or space, so clean the edges again afterwards.
  candidate = stripEdges(truncate(candidate, MAX_SHEET_NAME_LENGTH));
  if (candidate.length === 0) {
    candidate = firstFreeDefaultName(taken);
  }
  let unique = candidate;
  let attempt = 1;
  while (taken.has(unique.toLowerCase())) {
    attempt++;
    unique = withCopySuffix(candidate, attempt);
  }
  taken.add(unique.toLowerCase());
  return unique;
}

/** Whether a name is already valid (no sanitizing needed) - used to warn callers. */
export function isValidSheetName(name: string): boolean {
  return typeof name === 'string' && sanitizeSheetName(name, new Set<string>()) === name;
}

/** Quote a sheet name for a formula/defined-name reference (`'My Sheet'!A1`, apostrophes doubled). */
export function quoteSheetName(name: string): string {
  if (UNQUOTED_NAME.test(name)) {
    return name;
  }
  return `'${name.replaceAll("'", "''")}'`;
}
