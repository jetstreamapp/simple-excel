import { XlsxError } from '../errors';
import type { ChunkHandler, Deflater, DeflaterFactory, DeflaterOptions } from '../types';

/** True when `CompressionStream('deflate-raw')` is available in this environment. */
export function hasNativeDeflate(): boolean {
  // `typeof` on an undeclared global is the one probe that cannot throw, which is what callers rely on.
  return typeof CompressionStream === 'function';
}

/**
 * Remembers the first failure seen by a deflater. Everything after a broken compression stream is meaningless, so
 * every later call rejects with that same reason rather than with a second, derived one.
 */
class FailureLatch {
  private latched = false;
  private reason: unknown;

  record(reason: unknown): void {
    if (!this.latched) {
      this.latched = true;
      this.reason = reason;
    }
  }

  throwIfFailed(): void {
    if (this.latched) {
      throw this.reason;
    }
  }
}

/**
 * `CompressionStream('deflate-raw')` driven by the writer/reader pump pattern: `push` awaits the stream writer, which
 * is where back-pressure from a slow `onChunk` surfaces, while a single drain loop hands compressed chunks to
 * `onChunk` strictly in order. `finish` resolves only after the last handler call settled.
 */
function createNativeDeflater(onChunk: ChunkHandler): Deflater {
  const compressionStream = new CompressionStream('deflate-raw');
  const writer = compressionStream.writable.getWriter();
  const reader = compressionStream.readable.getReader();
  const failure = new FailureLatch();
  let bytesIn = 0;
  let bytesOut = 0;
  let finishing = false;

  const pump = (async (): Promise<void> => {
    for (;;) {
      const { done, value } = await reader.read();
      if (done) {
        return;
      }
      bytesOut += value.byteLength;
      await onChunk(value);
    }
  })();

  // `finish` and `abort` wait on this rather than on the raw pump, so a failure that happens between calls is
  // handled the moment it occurs and never surfaces as an unhandled rejection.
  const pumpSettled: Promise<void> = pump.then(undefined, async (reason: unknown) => {
    failure.record(reason);
    // Nobody will drain the readable again: release both ends so a `push` parked on back-pressure fails fast
    // instead of waiting forever for a reader that is gone.
    await Promise.allSettled([reader.cancel(reason), writer.abort(reason)]);
  });
  writer.closed.catch((reason: unknown) => failure.record(reason));

  return {
    async push(chunk: Uint8Array): Promise<void> {
      failure.throwIfFailed();
      if (finishing) {
        throw new XlsxError('WRITER_STATE', 'Bytes were pushed into a compressor that is already finished.');
      }
      try {
        // The platform types the writable as `BufferSource`, which excludes a SharedArrayBuffer-backed view; the
        // chunks we are handed are ordinary Uint8Arrays.
        await writer.write(chunk as Uint8Array<ArrayBuffer>);
        bytesIn += chunk.byteLength;
      } catch (reason) {
        failure.record(reason);
      }
      failure.throwIfFailed();
    },

    async finish(): Promise<void> {
      failure.throwIfFailed();
      if (finishing) {
        throw new XlsxError('WRITER_STATE', 'The compressor was finished twice.');
      }
      finishing = true;
      try {
        await writer.close();
      } catch (reason) {
        failure.record(reason);
      }
      // Wait for the drain loop even when closing failed: callers rely on every compressed chunk having reached
      // `onChunk` by the time `finish` resolves.
      await pumpSettled;
      failure.throwIfFailed();
    },

    async abort(reason?: unknown): Promise<void> {
      failure.record(reason ?? new XlsxError('ABORTED', 'The compressor was aborted.'));
      await Promise.allSettled([writer.abort(reason), reader.cancel(reason)]);
    },

    get bytesIn(): number {
      return bytesIn;
    },

    get bytesOut(): number {
      return bytesOut;
    },
  };
}

/** Pass-through "compressor" for `method: 'store'` entries; `bytesOut === bytesIn`. */
export function createStoredDeflater(onChunk: ChunkHandler): Deflater {
  const failure = new FailureLatch();
  let bytes = 0;
  let finished = false;

  return {
    async push(chunk: Uint8Array): Promise<void> {
      failure.throwIfFailed();
      if (finished) {
        throw new XlsxError('WRITER_STATE', 'Bytes were pushed into a compressor that is already finished.');
      }
      try {
        await onChunk(chunk);
        bytes += chunk.byteLength;
      } catch (reason) {
        failure.record(reason);
      }
      failure.throwIfFailed();
    },

    async finish(): Promise<void> {
      failure.throwIfFailed();
      if (finished) {
        throw new XlsxError('WRITER_STATE', 'The compressor was finished twice.');
      }
      finished = true;
    },

    async abort(reason?: unknown): Promise<void> {
      failure.record(reason ?? new XlsxError('ABORTED', 'The compressor was aborted.'));
    },

    get bytesIn(): number {
      return bytes;
    },

    get bytesOut(): number {
      return bytes;
    },
  };
}

/**
 * Deflater over `CompressionStream('deflate-raw')` with a writer/reader pump: every `push` awaits the stream writer,
 * compressed output is drained to `onChunk` as it appears, errors are sticky and re-thrown by later calls. When the
 * method is `'store'` (or the platform lacks CompressionStream) bytes pass straight through.
 *
 * `options.level` is accepted for parity with the node entry's zlib deflater and ignored here: the platform API has
 * no compression level (ADR-005).
 */
export const createDeflater: DeflaterFactory = (onChunk: ChunkHandler, options: DeflaterOptions): Deflater => {
  if (options.method === 'store' || !hasNativeDeflate()) {
    return createStoredDeflater(onChunk);
  }
  return createNativeDeflater(onChunk);
};
