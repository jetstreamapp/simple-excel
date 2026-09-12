/**
 * Every failure the library raises is an `XlsxError` with a stable `code`, so callers can branch on the
 * code and show their own copy. Messages are written for end users (they name the fix), not for developers.
 */
export type XlsxErrorCode =
  /** The bytes are not an OOXML spreadsheet (csv, html, docx, empty file, ...). `detail.format` says what was sniffed. */
  | 'NOT_XLSX'
  /** CFB container with an EncryptedPackage stream. The message always contains the words `password-protected`. */
  | 'ENCRYPTED'
  /** CFB container without encryption: a legacy BIFF .xls workbook. */
  | 'LEGACY_XLS'
  /** A zip container whose workbook part is binary (.xlsb). */
  | 'XLSB'
  /** An OpenDocument spreadsheet (.ods). */
  | 'ODS'
  /** Central directory or an entry runs past the end of the input. */
  | 'ZIP_TRUNCATED'
  /** Inflated size exceeded a cap or the compression ratio guard. */
  | 'ZIP_BOMB'
  /** Two central-directory entries share a normalized name. */
  | 'ZIP_DUPLICATE_ENTRY'
  /** Compression method other than stored/deflate, or an encrypted entry. */
  | 'ZIP_UNSUPPORTED'
  /** An entry's bytes do not match the CRC-32 recorded for it. */
  | 'ZIP_CRC_MISMATCH'
  /** `<!DOCTYPE` (or an internal entity declaration) in any part. */
  | 'XML_DOCTYPE'
  /** Structurally broken XML in a part we had to parse. */
  | 'XML_MALFORMED'
  /** A configurable limit was hit (rows, columns, shared strings, nesting depth, text length, entries). */
  | 'LIMIT_EXCEEDED'
  /** `workbook.sheet(x)` for a name or index that does not exist. */
  | 'SHEET_NOT_FOUND'
  /** A sheet name that cannot be sanitized into something Excel accepts. */
  | 'INVALID_SHEET_NAME'
  /** A row or column index past 1,048,576 x 16,384, or a non-monotonic write. */
  | 'ROW_OUT_OF_RANGE'
  /** A cell longer than 32,767 characters with `cellOverflow: 'throw'`. */
  | 'CELL_TOO_LONG'
  /** A zip entry or offset would overflow 32 bits while zip64 is disabled. */
  | 'ENTRY_TOO_LARGE'
  /** Writer used out of order (two open sheets, write after close, ...). */
  | 'WRITER_STATE'
  /** The caller's AbortSignal fired or `abort()` was called. */
  | 'ABORTED'
  /** A required platform feature (CompressionStream, Blob, ...) is missing. */
  | 'UNSUPPORTED_ENVIRONMENT';

export class XlsxError extends Error {
  readonly code: XlsxErrorCode;
  readonly detail: Readonly<Record<string, unknown>> | undefined;

  constructor(code: XlsxErrorCode, message: string, detail?: Readonly<Record<string, unknown>>) {
    super(message);
    this.name = 'XlsxError';
    this.code = code;
    this.detail = detail;
  }
}

export function isXlsxError(value: unknown): value is XlsxError {
  return value instanceof XlsxError || (typeof value === 'object' && value !== null && (value as { name?: unknown }).name === 'XlsxError');
}
