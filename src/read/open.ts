import { notImplemented } from '../internal/not-implemented';
import type { CellError, CellValue, OpenOptions, Workbook } from '../types';
import type { SourceInput } from '../zip/source';

/**
 * Open an xlsx from bytes (`ArrayBuffer`, `Uint8Array`, `Blob`/`File`, or a `RandomAccessSource`). Sniffs the format
 * first and throws a classified `XlsxError` (`ENCRYPTED` with a `password-protected` message, `LEGACY_XLS`, `XLSB`,
 * `ODS`, `NOT_XLSX`) for anything that is not a modern workbook. The workbook lists sheets without touching sheet
 * XML; sheet contents stream on demand. With `errors: 'object'` the row values include `CellError` objects.
 */
export function openWorkbook(
  input: SourceInput,
  options: OpenOptions & { readonly errors: 'object' },
): Promise<Workbook<CellValue | CellError>>;
export function openWorkbook(input: SourceInput, options?: OpenOptions): Promise<Workbook>;
export function openWorkbook(input: SourceInput, options?: OpenOptions): Promise<Workbook<CellValue | CellError>> {
  void input;
  void options;
  throw notImplemented('read/open');
}
