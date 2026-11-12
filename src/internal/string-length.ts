/**
 * True for the error an engine throws when a string would pass its maximum length: `RangeError` from concatenation
 * or `String.prototype.repeat`, and Node's `TextDecoder`, which throws a plain `Error` with code
 * `ERR_STRING_TOO_LONG`. Duck-typed so the library stays free of Node globals. Callers turn it into
 * `XlsxError('LIMIT_EXCEEDED')`, so a limit raised past what one string can hold still fails classified.
 */
export function isStringLengthError(error: unknown): boolean {
  if (error instanceof RangeError) {
    return true;
  }
  return typeof error === 'object' && error !== null && (error as { readonly code?: unknown }).code === 'ERR_STRING_TOO_LONG';
}
