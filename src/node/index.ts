/// <reference types="node" />
/**
 * Node-only helpers: file-backed sources and sinks and a zlib-backed deflater. Everything in the core entry works in
 * Node too; this entry only adds what needs `node:` modules.
 *
 * The `node:` types come from the triple-slash reference above: the declaration build does not include `@types/node`,
 * so nothing Node-typed may appear in an exported signature (hence `NodeWritableLike`), while the implementations
 * stay fully typed in both builds.
 */
import { createWriteStream } from 'node:fs';
import { open as openFile } from 'node:fs/promises';
import { createDeflateRaw } from 'node:zlib';
import { XlsxError } from '../errors';
import type { ByteSink, ChunkHandler, Deflater, DeflaterFactory, DeflaterOptions, RandomAccessSource } from '../types';

export * from '../index';

/** Whatever a caller aborts with has to become an `Error` before a stream can be destroyed with it. */
function toError(reason: unknown): Error {
  if (reason instanceof Error) {
    return reason;
  }
  return new XlsxError('ABORTED', 'The write was aborted.', { cause: reason });
}

export interface FileSource extends RandomAccessSource {
  close(): Promise<void>;
}

/**
 * Random-access source over a file (positional reads on a `FileHandle`); `close()` releases the handle, and a read
 * after it throws `XlsxError('ABORTED')` rather than the handle's bare `EBADF`.
 */
export async function fromFile(path: string): Promise<FileSource> {
  const handle = await openFile(path, 'r');
  const { size } = await handle.stat();
  let closed = false;

  return {
    size,
    /** Positional reads share no cursor, so reads may overlap; short reads only happen at end of file. */
    async read(offset: number, length: number): Promise<Uint8Array> {
      if (closed) {
        throw new XlsxError('ABORTED', `The file ${path} was closed. Open it again to read from it.`, { path });
      }
      const buffer = new Uint8Array(length);
      let filled = 0;
      while (filled < length) {
        const { bytesRead } = await handle.read(buffer, filled, length - filled, offset + filled);
        if (bytesRead === 0) {
          return buffer.subarray(0, filled);
        }
        filled += bytesRead;
      }
      return buffer;
    },
    async close(): Promise<void> {
      if (closed) {
        return;
      }
      closed = true;
      await handle.close();
    },
  };
}

/** Sink that streams to a file through `fs.createWriteStream`, honoring back-pressure. */
export function toFile(path: string): ByteSink {
  return toWritable(createWriteStream(path));
}

/** The subset of Node's `Writable` that `toWritable` needs (structural, so the core types stay Node-free). */
export interface NodeWritableLike {
  write(chunk: Uint8Array, callback?: (error?: Error | null) => void): boolean;
  end(callback?: () => void): unknown;
  destroy(error?: Error): unknown;
  /** `'drain'`, `'finish'` and `'end'` call the listener with no arguments; `'error'` calls it with the error. */
  once(event: string, listener: (error?: Error) => void): unknown;
  off(event: string, listener: (error?: Error) => void): unknown;
  readonly writableNeedDrain?: boolean;
}

/** Resolves when `event` fires, rejects if the stream errors first; both listeners are removed either way. */
function waitForEvent(stream: NodeWritableLike, event: string): Promise<void> {
  return new Promise<void>((resolve, reject) => {
    function detach(): void {
      stream.off(event, onEvent);
      stream.off('error', onError);
    }
    function onEvent(): void {
      detach();
      resolve();
    }
    function onError(error?: Error): void {
      detach();
      reject(error ?? new Error(`The stream failed while waiting for '${event}'.`));
    }
    stream.once(event, onEvent);
    stream.once('error', onError);
  });
}

/** Sink over any Node `Writable` (a response, a socket, a `PassThrough`). */
export function toWritable(writable: NodeWritableLike): ByteSink {
  let streamError: Error | undefined;
  let destroyed = false;

  // Subscribed for the life of the sink so the failure is sticky, and re-subscribed after every event because an
  // 'error' with no listener takes the process down.
  function captureError(error?: Error): void {
    streamError ??= error ?? new Error('The stream failed.');
    writable.once('error', captureError);
  }
  writable.once('error', captureError);

  return {
    async write(chunk: Uint8Array): Promise<void> {
      if (streamError) {
        throw streamError;
      }
      if (!writable.write(chunk)) {
        await waitForEvent(writable, 'drain');
      }
    },
    async close(): Promise<void> {
      if (streamError) {
        throw streamError;
      }
      const finished = waitForEvent(writable, 'finish');
      writable.end();
      await finished;
    },
    async abort(reason?: unknown): Promise<void> {
      if (destroyed) {
        return;
      }
      destroyed = true;
      const error = reason === undefined ? undefined : toError(reason);
      streamError ??= error;
      writable.destroy(error);
    },
  };
}

/** Pass-through "compressor" for `method: 'store'` entries; `bytesOut === bytesIn`. */
function createPassThroughDeflater(onChunk: ChunkHandler): Deflater {
  let bytesIn = 0;
  let bytesOut = 0;
  return {
    async push(chunk: Uint8Array): Promise<void> {
      bytesIn += chunk.byteLength;
      bytesOut += chunk.byteLength;
      await onChunk(chunk);
    },
    async finish(): Promise<void> {
      // Nothing is buffered: every chunk was handed on by `push`.
    },
    async abort(): Promise<void> {
      // Nothing to tear down.
    },
    get bytesIn(): number {
      return bytesIn;
    },
    get bytesOut(): number {
      return bytesOut;
    },
  };
}

/** Deflater over `zlib.createDeflateRaw` (`level` 1-9; 1 is 2-3x faster than the default for xlsx-shaped XML). */
export function nodeDeflater(level?: number): DeflaterFactory {
  return (onChunk: ChunkHandler, options: DeflaterOptions): Deflater => {
    if (options.method === 'store') {
      return createPassThroughDeflater(onChunk);
    }

    const resolvedLevel = options.level ?? level;
    const deflate = createDeflateRaw(resolvedLevel === undefined ? {} : { level: resolvedLevel });
    const outputQueue: Uint8Array[] = [];
    let bytesIn = 0;
    let bytesOut = 0;
    let failure: Error | undefined;
    let pumping = false;
    let pumpFinished: Promise<void> = Promise.resolve();

    function recordFailure(reason: unknown): Error {
      failure ??= toError(reason);
      return failure;
    }

    /**
     * Hands queued output to `onChunk` one chunk at a time, in order. The stream stays paused while anything is
     * queued, so a slow consumer back-pressures zlib's readable side, which back-pressures `push` in turn.
     */
    function startPump(): void {
      if (pumping) {
        return;
      }
      pumping = true;
      pumpFinished = (async () => {
        try {
          for (let chunk = outputQueue.shift(); chunk !== undefined; chunk = outputQueue.shift()) {
            await onChunk(chunk);
          }
          deflate.resume();
        } catch (error) {
          outputQueue.length = 0;
          deflate.destroy(recordFailure(error));
        } finally {
          pumping = false;
        }
      })();
    }

    deflate.on('data', (chunk: Uint8Array) => {
      bytesOut += chunk.byteLength;
      // zlib hands out Buffers; the view keeps the same bytes without a copy and without Buffer's divergent `slice`.
      outputQueue.push(new Uint8Array(chunk.buffer, chunk.byteOffset, chunk.byteLength));
      deflate.pause();
      startPump();
    });
    deflate.on('error', (error: Error) => {
      recordFailure(error);
    });

    return {
      async push(chunk: Uint8Array): Promise<void> {
        if (failure) {
          throw failure;
        }
        bytesIn += chunk.byteLength;
        if (!deflate.write(chunk)) {
          await waitForEvent(deflate, 'drain');
        }
      },
      async finish(): Promise<void> {
        if (failure) {
          throw failure;
        }
        await new Promise<void>((resolve, reject) => {
          function detach(): void {
            deflate.off('end', onEnd);
            deflate.off('error', onError);
            deflate.off('close', onClose);
          }
          function onEnd(): void {
            detach();
            resolve();
          }
          function onError(error: Error): void {
            detach();
            reject(recordFailure(error));
          }
          function onClose(): void {
            detach();
            reject(recordFailure(failure ?? new Error('The deflate stream closed before it finished.')));
          }
          deflate.once('end', onEnd);
          deflate.once('error', onError);
          deflate.once('close', onClose);
          deflate.end();
        });
        // 'end' means zlib emitted its last chunk, not that the consumer has taken it: let the pump run dry.
        let drained = false;
        while (!drained) {
          await pumpFinished;
          drained = failure !== undefined || (outputQueue.length === 0 && !pumping);
        }
        if (failure) {
          throw failure;
        }
      },
      async abort(reason?: unknown): Promise<void> {
        recordFailure(reason ?? new XlsxError('ABORTED', 'The write was aborted.'));
        outputQueue.length = 0;
        deflate.destroy();
      },
      get bytesIn(): number {
        return bytesIn;
      },
      get bytesOut(): number {
        return bytesOut;
      },
    };
  };
}
