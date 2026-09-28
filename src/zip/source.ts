import { XlsxError } from '../errors';
import type { RandomAccessSource } from '../types';

/**
 * The bytes of a workbook: an `ArrayBuffer` or `SharedArrayBuffer`, any `ArrayBufferView` (`Uint8Array`, Node's
 * `Buffer`, `DataView`, any other typed array), a `Blob`/`File`, or a `RandomAccessSource`.
 */
export type SourceInput = ArrayBuffer | SharedArrayBuffer | ArrayBufferView | Uint8Array | Blob | RandomAccessSource;

const ACCEPTED_INPUTS =
  'an ArrayBuffer, a Uint8Array (or Buffer, or any other typed array or DataView), a Blob or File, or a RandomAccessSource';

/** Objects and functions can carry the duck-typed members below; every other value is rejected before a lookup. */
function isObjectLike(input: unknown): input is object {
  return (typeof input === 'object' && input !== null) || typeof input === 'function';
}

function isRandomAccessSource(input: object): input is RandomAccessSource {
  const candidate = input as Partial<RandomAccessSource>;
  return typeof candidate.size === 'number' && typeof candidate.read === 'function';
}

/** Duck-typed so any `Blob`/`File` lookalike works: we only ever call `size`, `slice` and `arrayBuffer`. */
function isBlobLike(input: object): input is Blob {
  const candidate = input as Partial<Blob>;
  return typeof candidate.size === 'number' && typeof candidate.slice === 'function' && typeof candidate.arrayBuffer === 'function';
}

/**
 * An `ArrayBuffer` or `SharedArrayBuffer`, including one from another realm (an iframe, a worker, an Electron bridge),
 * where `instanceof` fails. `SharedArrayBuffer` may not even exist as a global (no cross-origin isolation).
 */
function isBinaryBuffer(input: object): input is ArrayBufferLike {
  const tag = Object.prototype.toString.call(input);
  return tag === '[object ArrayBuffer]' || tag === '[object SharedArrayBuffer]';
}

function isSharedBuffer(buffer: ArrayBufferLike): boolean {
  return Object.prototype.toString.call(buffer) === '[object SharedArrayBuffer]';
}

/** What the caller passed, for the error message. */
function describeInput(input: unknown): string {
  if (input === null) {
    return 'null';
  }
  if (Array.isArray(input)) {
    return 'an array';
  }
  if (typeof input === 'object') {
    return 'an object that is none of those';
  }
  return typeof input === 'undefined' ? 'undefined' : `a ${typeof input}`;
}

function unsupportedInput(input: unknown): XlsxError {
  const received = describeInput(input);
  const hint =
    typeof input === 'string'
      ? ' A string cannot be read as a workbook: pass the file itself, or convert a binary string to bytes first (Uint8Array.from(text, character => character.charCodeAt(0))).'
      : '';
  return new XlsxError('NOT_XLSX', `A workbook is read from ${ACCEPTED_INPUTS}, but this call passed ${received}.${hint}`, {
    received: input === null ? 'null' : typeof input,
  });
}

/** Clamp a requested range to the source so `read` past the end returns what exists instead of throwing. */
function clampRange(size: number, offset: number, length: number): { start: number; end: number } {
  const start = Math.min(Math.max(offset, 0), size);
  const end = Math.min(start + Math.max(length, 0), size);
  return { start, end };
}

/**
 * Zero-copy range reads, except over shared memory: `DecompressionStream` and `TextDecoder` refuse a view of a
 * `SharedArrayBuffer`, so each range read from one is copied (one read at a time, never the whole input).
 */
function sourceFromBytes(bytes: Uint8Array): RandomAccessSource {
  const shared = isSharedBuffer(bytes.buffer);
  return {
    size: bytes.byteLength,
    read(offset: number, length: number): Promise<Uint8Array> {
      const { start, end } = clampRange(bytes.byteLength, offset, length);
      return Promise.resolve(shared ? bytes.slice(start, end) : bytes.subarray(start, end));
    },
  };
}

function sourceFromBlob(blob: Blob): RandomAccessSource {
  return {
    size: blob.size,
    async read(offset: number, length: number): Promise<Uint8Array> {
      const { start, end } = clampRange(blob.size, offset, length);
      if (start === end) {
        return new Uint8Array(0);
      }
      return new Uint8Array(await blob.slice(start, end).arrayBuffer());
    },
  };
}

/**
 * Wrap bytes as a `RandomAccessSource`: zero-copy views over an `ArrayBuffer` or any `ArrayBufferView` (from this
 * realm or another; shared memory is copied one range read at a time), `slice().arrayBuffer()` over a `Blob`/`File`,
 * pass-through for an existing source. Anything else throws `XlsxError('NOT_XLSX')` naming what is accepted.
 */
export function sourceFrom(input: SourceInput): RandomAccessSource {
  if (input instanceof Uint8Array) {
    return sourceFromBytes(input);
  }
  const candidate: unknown = input;
  if (!isObjectLike(candidate)) {
    throw unsupportedInput(candidate);
  }
  if (ArrayBuffer.isView(candidate)) {
    return sourceFromBytes(new Uint8Array(candidate.buffer, candidate.byteOffset, candidate.byteLength));
  }
  if (isBinaryBuffer(candidate)) {
    return sourceFromBytes(new Uint8Array(candidate));
  }
  if (isRandomAccessSource(candidate)) {
    return candidate;
  }
  if (isBlobLike(candidate)) {
    return sourceFromBlob(candidate);
  }
  throw unsupportedInput(candidate);
}
