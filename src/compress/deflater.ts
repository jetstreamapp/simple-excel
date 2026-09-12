import { notImplemented } from '../internal/not-implemented';
import type { ChunkHandler, Deflater, DeflaterFactory, DeflaterOptions } from '../types';

/** True when `CompressionStream('deflate-raw')` is available in this environment. */
export function hasNativeDeflate(): boolean {
  throw notImplemented('compress/deflater');
}

/**
 * Deflater over `CompressionStream('deflate-raw')` with a writer/reader pump: every `push` awaits the stream writer,
 * compressed output is drained to `onChunk` as it appears, errors are sticky and re-thrown by later calls. When the
 * method is `'store'` (or the platform lacks CompressionStream) bytes pass straight through.
 */
export const createDeflater: DeflaterFactory = (onChunk: ChunkHandler, options: DeflaterOptions): Deflater => {
  void onChunk;
  void options;
  throw notImplemented('compress/deflater');
};

/** Pass-through "compressor" for `method: 'store'` entries; `bytesOut === bytesIn`. */
export function createStoredDeflater(onChunk: ChunkHandler): Deflater {
  void onChunk;
  throw notImplemented('compress/deflater');
}
