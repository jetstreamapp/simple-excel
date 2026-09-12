import { notImplemented } from '../internal/not-implemented';

export const MAX_SHEET_NAME_LENGTH: number = 31;

/**
 * Make a sheet name Excel accepts and unique within the workbook: `: \ / ? * [ ]` become `_`, leading and trailing
 * apostrophes are stripped, the name is trimmed to 31 characters, `History` (reserved, case-insensitive) becomes
 * `History_`, empty input becomes `Sheet<n>`. Collisions (case-insensitive against `taken`) get ` (2)`, ` (3)`, ...
 * fitted inside the 31-character budget. The returned name is added to `taken`.
 */
export function sanitizeSheetName(name: string, taken: Set<string>): string {
  void name;
  void taken;
  throw notImplemented('sml/sheet-name');
}

/** Whether a name is already valid (no sanitizing needed) - used to warn callers. */
export function isValidSheetName(name: string): boolean {
  void name;
  throw notImplemented('sml/sheet-name');
}

/** Quote a sheet name for a formula/defined-name reference (`'My Sheet'!A1`, apostrophes doubled). */
export function quoteSheetName(name: string): string {
  void name;
  throw notImplemented('sml/sheet-name');
}
