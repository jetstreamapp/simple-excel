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
  /**
   * The caller cancelled: their AbortSignal fired or they called `abort()`. Reading from a workbook or source the
   * caller already closed reports it too. Never used for a failure the caller did not ask for.
   */
  | 'ABORTED'
  /**
   * Something underneath the library failed: a sink's write or close, a Node stream or file (including an HTTP
   * response whose client disconnected), a `RandomAccessSource` read, a compressor, or a callback the caller passed
   * in. The original error is the standard `cause`.
   */
  | 'IO_ERROR'
  /** A required platform feature (CompressionStream, Blob, ...) is missing. */
  | 'UNSUPPORTED_ENVIRONMENT';

/**
 * Marks every `XlsxError` from every copy of the library (the core and node entries, the ESM and CommonJS builds, two
 * installed versions), so `instanceof XlsxError` and `isXlsxError` hold whichever copy threw it.
 */
const XLSX_ERROR_BRAND: unique symbol = Symbol.for('@jetstreamapp/simple-excel/XlsxError');

function hasBrand(value: unknown): boolean {
  return typeof value === 'object' && value !== null && (value as { [XLSX_ERROR_BRAND]?: unknown })[XLSX_ERROR_BRAND] === true;
}

export class XlsxError extends Error {
  readonly code: XlsxErrorCode;
  readonly detail: Readonly<Record<string, unknown>> | undefined;

  /**
   * `options.cause` becomes the standard `Error.cause`: the platform, stream or caller error this one classifies.
   * `detail` carries structured facts about the failure (limits, part names), never the cause.
   */
  constructor(code: XlsxErrorCode, message: string, detail?: Readonly<Record<string, unknown>>, options?: ErrorOptions) {
    super(message, options);
    this.name = 'XlsxError';
    this.code = code;
    this.detail = detail;
  }

  /**
   * `instanceof XlsxError` holds for an error from any copy of the library, not only this one: the node entry is a
   * separate bundle, and a dual package can load the ESM and CommonJS builds side by side. A subclass keeps the
   * ordinary prototype check.
   */
  static override [Symbol.hasInstance](value: unknown): value is XlsxError {
    if (this !== XlsxError) {
      return Function.prototype[Symbol.hasInstance].call(this, value);
    }
    return hasBrand(value) || Function.prototype[Symbol.hasInstance].call(this, value);
  }

  static {
    Object.defineProperty(this.prototype, XLSX_ERROR_BRAND, { value: true });
  }
}

/** True for an `XlsxError` from any copy of the library (see `XlsxError[Symbol.hasInstance]`). */
export function isXlsxError(value: unknown): value is XlsxError {
  return hasBrand(value);
}

/**
 * Classify a failure the library did not raise itself: an `XlsxError` passes through unchanged, anything else
 * becomes a new `XlsxError` with `code` and `message` and the original as its `cause`.
 */
export function classifyError(reason: unknown, code: XlsxErrorCode, message: string): XlsxError {
  if (hasBrand(reason)) {
    return reason as XlsxError;
  }
  return new XlsxError(code, message, undefined, { cause: reason });
}
