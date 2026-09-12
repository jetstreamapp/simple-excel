import { notImplemented } from '../internal/not-implemented';
import type { RandomAccessSource } from '../types';

export type SourceInput = ArrayBuffer | Uint8Array | Blob | RandomAccessSource;

/**
 * Wrap bytes as a `RandomAccessSource`: zero-copy subarray views over `ArrayBuffer`/`Uint8Array`,
 * `slice().arrayBuffer()` over `Blob`/`File`, pass-through for an existing source.
 */
export function sourceFrom(input: SourceInput): RandomAccessSource {
  void input;
  throw notImplemented('zip/source');
}
