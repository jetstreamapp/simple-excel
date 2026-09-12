import { isXlsxError, XlsxError } from './errors';
import type { ByteSink } from './types';

/**
 * Pending chunks are folded into one sub-Blob once they reach this size and the array is dropped, so Chromium can page
 * the data out and the JS heap stays flat for multi-hundred-megabyte workbooks.
 */
const BLOB_FOLD_THRESHOLD_BYTES = 32 * 1024 * 1024;

/** Past this the caller wants a streaming sink, not one contiguous allocation. */
const DEFAULT_MAX_BYTES = 1024 * 1024 * 1024;

interface Deferred<T> {
  readonly promise: Promise<T>;
  resolve(value: T): void;
  reject(reason: unknown): void;
}

function createDeferred<T>(): Deferred<T> {
  let resolve!: (value: T) => void;
  let reject!: (reason: unknown) => void;
  const promise = new Promise<T>((resolveResult, rejectResult) => {
    resolve = resolveResult;
    reject = rejectResult;
  });
  // `result()` is often never called on the abort path; marking the promise handled keeps an abort from surfacing as
  // an unhandled rejection. Callers still see the rejection because `result()` hands back this same promise.
  void promise.catch(() => undefined);
  return { promise, resolve, reject };
}

/**
 * `BlobPart` only admits views over a non-shared `ArrayBuffer`, and a plain `Uint8Array` might be backed by a
 * `SharedArrayBuffer` as far as the type system knows. Every chunk the writer produces is a regular array, and the
 * `Blob` constructor accepts any view at runtime.
 */
function asBlobParts(chunks: Uint8Array[]): BlobPart[] {
  return chunks as BlobPart[];
}

/** Errors keep their classification; anything else becomes an `ABORTED` error carrying the original as `detail.cause`. */
function toAbortError(reason: unknown): XlsxError {
  if (isXlsxError(reason)) {
    return reason;
  }
  return new XlsxError('ABORTED', 'The write was aborted.', { cause: reason });
}

export interface BlobSink extends ByteSink {
  /** The finished Blob; rejects until `close()` resolved. */
  result(): Promise<Blob>;
  /** Bytes received so far. */
  readonly bytesWritten: number;
}

/**
 * Collects chunks into a Blob. Chunks are folded into sub-Blobs every 32 MiB and the arrays dropped, so Chromium can
 * page the data to disk and the JS heap stays flat for multi-hundred-megabyte files.
 *
 * The sink takes ownership of every chunk it is handed and never copies it: the writer encodes a fresh chunk per
 * flush, so a defensive `chunk.slice()` would double the cost of every export. A caller that reuses a scratch buffer
 * must hand over a copy.
 */
export function collectToBlob(type: string = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'): BlobSink {
  if (typeof Blob === 'undefined') {
    throw new XlsxError(
      'UNSUPPORTED_ENVIRONMENT',
      'This environment has no Blob. Collect the bytes with collectToBytes() or write to a stream instead.',
    );
  }

  const foldedParts: Blob[] = [];
  let pendingChunks: Uint8Array[] = [];
  let pendingBytes = 0;
  let bytesWritten = 0;
  let closed = false;
  let failure: XlsxError | undefined;
  const result = createDeferred<Blob>();

  function foldPendingChunks(): void {
    if (pendingChunks.length === 0) {
      return;
    }
    foldedParts.push(new Blob(asBlobParts(pendingChunks)));
    pendingChunks = [];
    pendingBytes = 0;
  }

  return {
    get bytesWritten(): number {
      return bytesWritten;
    },
    async write(chunk: Uint8Array): Promise<void> {
      if (failure) {
        throw failure;
      }
      if (closed) {
        throw new XlsxError('WRITER_STATE', 'This sink is already closed; no more bytes can be written to it.');
      }
      pendingChunks.push(chunk);
      pendingBytes += chunk.byteLength;
      bytesWritten += chunk.byteLength;
      if (pendingBytes >= BLOB_FOLD_THRESHOLD_BYTES) {
        foldPendingChunks();
      }
    },
    async close(): Promise<void> {
      if (failure) {
        throw failure;
      }
      if (closed) {
        return;
      }
      closed = true;
      foldPendingChunks();
      result.resolve(new Blob(foldedParts, { type }));
    },
    async abort(reason?: unknown): Promise<void> {
      if (closed || failure) {
        return;
      }
      failure = toAbortError(reason);
      pendingChunks = [];
      pendingBytes = 0;
      foldedParts.length = 0;
      result.reject(failure);
    },
    result(): Promise<Blob> {
      return result.promise;
    },
  };
}

export interface BytesSink extends ByteSink {
  /** The concatenated bytes; throws until `close()` resolved. */
  result(): Uint8Array;
  readonly bytesWritten: number;
}

/** Collects chunks in memory (tests, Node, small files). Throws `LIMIT_EXCEEDED` past `maxBytes` (default 1 GiB). */
export function collectToBytes(options?: { readonly maxBytes?: number }): BytesSink {
  const maxBytes = options?.maxBytes ?? DEFAULT_MAX_BYTES;
  const chunks: Uint8Array[] = [];
  let bytesWritten = 0;
  let failure: XlsxError | undefined;
  let collected: Uint8Array | undefined;

  /** One allocation at the end; a lone chunk is handed back as-is. */
  function concatenate(): Uint8Array {
    const [first] = chunks;
    if (chunks.length === 1 && first !== undefined) {
      return first;
    }
    const merged = new Uint8Array(bytesWritten);
    let offset = 0;
    for (const chunk of chunks) {
      merged.set(chunk, offset);
      offset += chunk.byteLength;
    }
    return merged;
  }

  return {
    get bytesWritten(): number {
      return bytesWritten;
    },
    async write(chunk: Uint8Array): Promise<void> {
      if (failure) {
        throw failure;
      }
      if (collected !== undefined) {
        throw new XlsxError('WRITER_STATE', 'This sink is already closed; no more bytes can be written to it.');
      }
      if (bytesWritten + chunk.byteLength > maxBytes) {
        throw new XlsxError(
          'LIMIT_EXCEEDED',
          `This file is larger than the ${maxBytes} byte limit for collecting it in memory. Write it to a stream or a file instead.`,
          { maxBytes, bytesWritten },
        );
      }
      chunks.push(chunk);
      bytesWritten += chunk.byteLength;
    },
    async close(): Promise<void> {
      if (failure) {
        throw failure;
      }
      if (collected !== undefined) {
        return;
      }
      collected = concatenate();
      chunks.length = 0;
    },
    async abort(reason?: unknown): Promise<void> {
      if (collected !== undefined || failure) {
        return;
      }
      failure = toAbortError(reason);
      chunks.length = 0;
    },
    result(): Uint8Array {
      if (failure) {
        throw failure;
      }
      if (collected === undefined) {
        throw new XlsxError('WRITER_STATE', 'The bytes are only available after close() resolves.');
      }
      return collected;
    },
  };
}

/** Adapt a `WritableStream<Uint8Array>` (File System Access, Node `Writable.toWeb`, ...) as a sink. */
export function fromWritableStream(stream: WritableStream<Uint8Array>): ByteSink {
  const writer = stream.getWriter();
  let released = false;
  // A stream that errors rejects `closed` whether or not anyone looks at it, and `releaseLock()` rejects `ready`;
  // every failure the caller cares about is reported by write/close/abort instead.
  void writer.closed.catch(() => undefined);

  function releaseLock(): void {
    if (released) {
      return;
    }
    released = true;
    const pendingReady = writer.ready;
    void pendingReady.catch(() => undefined);
    writer.releaseLock();
  }

  return {
    async write(chunk: Uint8Array): Promise<void> {
      await writer.ready;
      await writer.write(chunk);
    },
    async close(): Promise<void> {
      try {
        await writer.close();
      } finally {
        releaseLock();
      }
    },
    async abort(reason?: unknown): Promise<void> {
      try {
        await writer.abort(reason);
      } finally {
        releaseLock();
      }
    },
  };
}

/** Expose a sink as a `WritableStream<Uint8Array>` for `pipeTo` consumers. */
export function toWritableStream(sink: ByteSink): WritableStream<Uint8Array> {
  return new WritableStream<Uint8Array>({
    write(chunk: Uint8Array): Promise<void> {
      return sink.write(chunk);
    },
    close(): Promise<void> {
      return sink.close();
    },
    abort(reason: unknown): Promise<void> {
      return sink.abort(reason);
    },
  });
}

/**
 * Wrap a caller-supplied sink so that every method call after a failure rejects with the original reason.
 *
 * Calls are also serialized: a write issued while an earlier call is still pending waits for it, so the writer may
 * hand over a chunk without awaiting it and still know the sink sees the chunks in order. A second `close()`, and an
 * `abort()` after a close or another abort, are no-ops; `abort()` never rejects with the sticky reason, so an error
 * path can always tear the sink down.
 */
export function guardSink(sink: ByteSink): ByteSink {
  let failure: XlsxError | undefined;
  let closed = false;
  let aborted = false;
  /** Tail of the serialized call chain. It never rejects, so one failure is reported to one caller only. */
  let tail: Promise<void> = Promise.resolve();

  function remember(reason: unknown): XlsxError {
    failure ??= toAbortError(reason);
    return failure;
  }

  function enqueue(operation: () => Promise<void>): Promise<void> {
    const running = tail.then(operation);
    tail = running.then(
      () => undefined,
      () => undefined,
    );
    return running;
  }

  return {
    write(chunk: Uint8Array): Promise<void> {
      return enqueue(async () => {
        if (failure) {
          throw failure;
        }
        if (closed) {
          throw new XlsxError('WRITER_STATE', 'This sink is already closed; no more bytes can be written to it.');
        }
        try {
          await sink.write(chunk);
        } catch (error) {
          throw remember(error);
        }
      });
    },
    close(): Promise<void> {
      return enqueue(async () => {
        if (failure) {
          throw failure;
        }
        if (closed) {
          return;
        }
        closed = true;
        try {
          await sink.close();
        } catch (error) {
          throw remember(error);
        }
      });
    },
    abort(reason?: unknown): Promise<void> {
      return enqueue(async () => {
        if (closed || aborted) {
          return;
        }
        aborted = true;
        remember(reason);
        try {
          await sink.abort(reason);
        } catch (error) {
          remember(error);
        }
      });
    },
  };
}
