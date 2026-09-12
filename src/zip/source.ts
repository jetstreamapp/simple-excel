import type { RandomAccessSource } from '../types';

export type SourceInput = ArrayBuffer | Uint8Array | Blob | RandomAccessSource;

function isRandomAccessSource(input: SourceInput): input is RandomAccessSource {
  const candidate = input as Partial<RandomAccessSource>;
  return typeof candidate.size === 'number' && typeof candidate.read === 'function';
}

/** Duck-typed so any `Blob`/`File` lookalike works: we only ever call `size`, `slice` and `arrayBuffer`. */
function isBlobLike(input: SourceInput): input is Blob {
  const candidate = input as Partial<Blob>;
  return typeof candidate.size === 'number' && typeof candidate.slice === 'function' && typeof candidate.arrayBuffer === 'function';
}

/** Clamp a requested range to the source so `read` past the end returns what exists instead of throwing. */
function clampRange(size: number, offset: number, length: number): { start: number; end: number } {
  const start = Math.min(Math.max(offset, 0), size);
  const end = Math.min(start + Math.max(length, 0), size);
  return { start, end };
}

function sourceFromBytes(bytes: Uint8Array): RandomAccessSource {
  return {
    size: bytes.byteLength,
    read(offset: number, length: number): Promise<Uint8Array> {
      const { start, end } = clampRange(bytes.byteLength, offset, length);
      return Promise.resolve(bytes.subarray(start, end));
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
 * Wrap bytes as a `RandomAccessSource`: zero-copy subarray views over `ArrayBuffer`/`Uint8Array`,
 * `slice().arrayBuffer()` over `Blob`/`File`, pass-through for an existing source.
 */
export function sourceFrom(input: SourceInput): RandomAccessSource {
  if (input instanceof Uint8Array) {
    return sourceFromBytes(input);
  }
  if (input instanceof ArrayBuffer) {
    return sourceFromBytes(new Uint8Array(input));
  }
  if (isRandomAccessSource(input)) {
    return input;
  }
  if (isBlobLike(input)) {
    return sourceFromBlob(input);
  }
  throw new TypeError('sourceFrom expects an ArrayBuffer, a Uint8Array, a Blob or a RandomAccessSource');
}
