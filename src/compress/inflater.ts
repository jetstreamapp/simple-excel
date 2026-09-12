import { XlsxError } from '../errors';

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
 * Whatever the decompressor rejects with (corrupt or truncated deflate data) becomes `ZIP_TRUNCATED` with the
 * platform error in `detail.cause`; errors we classified ourselves pass through unchanged.
 */
function asXlsxError(error: unknown): XlsxError {
  if (error instanceof XlsxError) {
    return error;
  }
  return new XlsxError('ZIP_TRUNCATED', 'This file contains compressed data that could not be decoded; it is probably damaged.', {
    cause: error,
  });
}

function bombError(maxBytes: number): XlsxError {
  return new XlsxError('ZIP_BOMB', `This file expands to more than the ${maxBytes} byte limit and was not read further.`, { maxBytes });
}

function abortError(reason: unknown): XlsxError {
  return reason instanceof XlsxError ? reason : new XlsxError('ABORTED', 'The read was aborted.', { cause: reason });
}

/** `method: 'store'` entries are already the bytes we want: pass them straight through, still counting the cap. */
function createStoredInflater(onChunk: (chunk: Uint8Array) => Promise<void>, maxBytes: number): Inflater {
  let bytesIn = 0;
  let bytesOut = 0;
  let failure: XlsxError | undefined;
  let finished = false;

  return {
    async push(chunk: Uint8Array): Promise<void> {
      if (failure) {
        throw failure;
      }
      bytesIn += chunk.byteLength;
      bytesOut += chunk.byteLength;
      if (bytesOut > maxBytes) {
        failure = bombError(maxBytes);
        throw failure;
      }
      try {
        await onChunk(chunk);
      } catch (error) {
        failure = asXlsxError(error);
        throw failure;
      }
    },
    async finish(): Promise<void> {
      if (failure) {
        throw failure;
      }
      finished = true;
    },
    async abort(reason?: unknown): Promise<void> {
      if (!finished) {
        failure ??= abortError(reason);
      }
    },
    get bytesIn(): number {
      return bytesIn;
    },
    get bytesOut(): number {
      return bytesOut;
    },
  };
}

/**
 * Streaming raw-deflate decoder over `DecompressionStream('deflate-raw')`, mirroring `createDeflater`. Enforces
 * `maxBytes` on the actual inflated byte count and aborts the stream when exceeded.
 */
export function createInflater(onChunk: (chunk: Uint8Array) => Promise<void>, options: InflaterOptions): Inflater {
  const { maxBytes, method } = options;
  if (method === 'store') {
    return createStoredInflater(onChunk, maxBytes);
  }
  if (typeof DecompressionStream === 'undefined') {
    throw new XlsxError('UNSUPPORTED_ENVIRONMENT', 'This environment has no DecompressionStream, which is required to read xlsx files.');
  }

  const stream = new DecompressionStream('deflate-raw');
  const writer = (stream.writable as WritableStream<Uint8Array>).getWriter();
  const reader = (stream.readable as ReadableStream<Uint8Array>).getReader();

  let bytesIn = 0;
  let bytesOut = 0;
  let failure: XlsxError | undefined;
  let state: 'open' | 'finished' | 'aborted' = 'open';

  /** The first failure wins and is sticky, so callers see the real cause rather than the follow-on stream error. */
  function recordFailure(error: unknown): XlsxError {
    failure ??= asXlsxError(error);
    return failure;
  }

  // One drain loop delivers inflated chunks to `onChunk` in stream order, and its back-pressure reaches
  // `writer.write` through the transform stream. It never rejects: failures are recorded and re-thrown by
  // whichever call the caller is waiting on.
  const drainLoop: Promise<void> = (async (): Promise<void> => {
    for (;;) {
      const { done, value } = await reader.read();
      if (done) {
        return;
      }
      bytesOut += value.byteLength;
      if (bytesOut > maxBytes) {
        throw bombError(maxBytes);
      }
      await onChunk(value);
    }
  })().catch(async (error: unknown) => {
    const recorded = recordFailure(error);
    await Promise.all([writer.abort(recorded).catch(() => {}), reader.cancel(recorded).catch(() => {})]);
  });

  return {
    async push(chunk: Uint8Array): Promise<void> {
      if (failure) {
        throw failure;
      }
      bytesIn += chunk.byteLength;
      try {
        await writer.write(chunk);
      } catch (error) {
        throw recordFailure(error);
      }
      if (failure) {
        throw failure;
      }
    },
    async finish(): Promise<void> {
      if (state === 'open') {
        state = 'finished';
        try {
          await writer.close();
        } catch (error) {
          recordFailure(error);
        }
      }
      await drainLoop;
      if (failure) {
        throw failure;
      }
    },
    async abort(reason?: unknown): Promise<void> {
      if (state === 'aborted') {
        return;
      }
      const wasOpen = state === 'open';
      state = 'aborted';
      if (wasOpen) {
        const aborted = abortError(reason);
        failure ??= aborted;
        await Promise.all([writer.abort(aborted).catch(() => {}), reader.cancel(aborted).catch(() => {})]);
      }
      await drainLoop;
    },
    get bytesIn(): number {
      return bytesIn;
    },
    get bytesOut(): number {
      return bytesOut;
    },
  };
}
