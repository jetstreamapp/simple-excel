import { notImplemented } from './internal/not-implemented';
import type { ByteSink } from './types';

export interface BlobSink extends ByteSink {
  /** The finished Blob; rejects until `close()` resolved. */
  result(): Promise<Blob>;
  /** Bytes received so far. */
  readonly bytesWritten: number;
}

/**
 * Collects chunks into a Blob. Chunks are folded into sub-Blobs every 32 MiB and the arrays dropped, so Chromium can
 * page the data to disk and the JS heap stays flat for multi-hundred-megabyte files.
 */
export function collectToBlob(type: string = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'): BlobSink {
  void type;
  throw notImplemented('sinks');
}

export interface BytesSink extends ByteSink {
  /** The concatenated bytes; throws until `close()` resolved. */
  result(): Uint8Array;
  readonly bytesWritten: number;
}

/** Collects chunks in memory (tests, Node, small files). Throws `LIMIT_EXCEEDED` past `maxBytes` (default 1 GiB). */
export function collectToBytes(options?: { readonly maxBytes?: number }): BytesSink {
  void options;
  throw notImplemented('sinks');
}

/** Adapt a `WritableStream<Uint8Array>` (File System Access, Node `Writable.toWeb`, ...) as a sink. */
export function fromWritableStream(stream: WritableStream<Uint8Array>): ByteSink {
  void stream;
  throw notImplemented('sinks');
}

/** Expose a sink as a `WritableStream<Uint8Array>` for `pipeTo` consumers. */
export function toWritableStream(sink: ByteSink): WritableStream<Uint8Array> {
  void sink;
  throw notImplemented('sinks');
}

/** Wrap a caller-supplied sink so that every method call after a failure rejects with the original reason. */
export function guardSink(sink: ByteSink): ByteSink {
  void sink;
  throw notImplemented('sinks');
}
