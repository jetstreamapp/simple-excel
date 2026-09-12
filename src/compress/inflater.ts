import { notImplemented } from '../internal/not-implemented';

export interface InflaterOptions {
  /** Throw `XlsxError('ZIP_BOMB')` once inflated output exceeds this many bytes. */
  readonly maxBytes: number;
  /** `'store'` passes bytes through unchanged. */
  readonly method: 'deflate' | 'store';
}

export interface Inflater {
  /** Push compressed bytes; resolves after every resulting inflated chunk was handed to `onChunk`. */
  push(chunk: Uint8Array): Promise<void>;
  /** Signal end of compressed input; resolves after the final inflated bytes were delivered. */
  finish(): Promise<void>;
  abort(reason?: unknown): Promise<void>;
  readonly bytesIn: number;
  readonly bytesOut: number;
}

/**
 * Streaming raw-deflate decoder over `DecompressionStream('deflate-raw')`, mirroring `createDeflater`. Enforces
 * `maxBytes` on the actual inflated byte count and aborts the stream when exceeded.
 */
export function createInflater(onChunk: (chunk: Uint8Array) => Promise<void>, options: InflaterOptions): Inflater {
  void onChunk;
  void options;
  throw notImplemented('compress/inflater');
}
