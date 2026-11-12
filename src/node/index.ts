/// <reference types="node" />
/**
 * Node-only helpers: file-backed sources and sinks and a zlib-backed deflater. Everything in the core entry works in
 * Node too; this entry only adds what needs `node:` modules.
 *
 * The `node:` types come from the triple-slash reference above: the declaration build does not include `@types/node`,
 * so nothing Node-typed may appear in an exported signature (hence `NodeWritableLike`), while the implementations
 * stay fully typed in both builds.
 *
 * The core is imported by its package name, not by path: the build keeps it external, so the node bundle loads the
 * published core bundle instead of carrying a second copy of it (one `XlsxError` class, one set of module state per
 * module format). `tsconfig.json` and `vitest.config.mts` map the name back to `src/index.ts`.
 */
import { randomBytes } from 'node:crypto';
import { createWriteStream, type WriteStream } from 'node:fs';
import { type FileHandle, open as openFile, rename, unlink } from 'node:fs/promises';
import { basename, dirname, join } from 'node:path';
import { createDeflateRaw } from 'node:zlib';
import {
  isXlsxError,
  XlsxError,
  type ByteSink,
  type ChunkHandler,
  type Deflater,
  type DeflaterFactory,
  type DeflaterOptions,
  type RandomAccessSource,
} from '@jetstreamapp/simple-excel';

export * from '@jetstreamapp/simple-excel';

/** A cancellation: an `XlsxError` reason keeps its classification, anything else becomes `ABORTED` with it as `cause`. */
function toAbortError(reason: unknown): XlsxError {
  if (isXlsxError(reason)) {
    return reason;
  }
  return new XlsxError('ABORTED', 'The write was aborted.', undefined, reason === undefined ? undefined : { cause: reason });
}

/** A failure of a stream, file or zlib: `IO_ERROR` with the original as `cause`, unless it already is classified. */
function toIoError(reason: unknown, message: string, detail?: Readonly<Record<string, unknown>>): XlsxError {
  if (isXlsxError(reason)) {
    return reason;
  }
  return new XlsxError('IO_ERROR', message, detail, { cause: reason });
}

/** A stream that closed under a write, or before `close()` finished it: the peer went away (or someone destroyed it). */
function closedEarlyError(): XlsxError {
  return new XlsxError(
    'IO_ERROR',
    'The stream closed before the spreadsheet was completely written (for an HTTP response: the client disconnected).',
  );
}

export interface FileSource extends RandomAccessSource {
  close(): Promise<void>;
}

/**
 * Random-access source over a file (positional reads on a `FileHandle`); `close()` releases the handle. A file that
 * cannot be opened or read is `IO_ERROR` (the `fs` error, with its `code`, is the `cause`); a read after `close()`
 * is `ABORTED` rather than the handle's bare `EBADF`.
 */
export async function fromFile(path: string): Promise<FileSource> {
  let handle: FileHandle;
  try {
    handle = await openFile(path, 'r');
  } catch (error) {
    throw toIoError(error, `The file ${path} could not be opened.`, { path });
  }
  let size: number;
  try {
    ({ size } = await handle.stat());
  } catch (error) {
    await handle.close().catch(() => undefined);
    throw toIoError(error, `The file ${path} could not be read.`, { path });
  }
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
      try {
        while (filled < length) {
          const { bytesRead } = await handle.read(buffer, filled, length - filled, offset + filled);
          if (bytesRead === 0) {
            return buffer.subarray(0, filled);
          }
          filled += bytesRead;
        }
      } catch (error) {
        if (closed) {
          throw new XlsxError('ABORTED', `The file ${path} was closed. Open it again to read from it.`, { path });
        }
        throw toIoError(error, `The file ${path} could not be read.`, { path });
      }
      return buffer;
    },
    async close(): Promise<void> {
      if (closed) {
        return;
      }
      closed = true;
      try {
        await handle.close();
      } catch (error) {
        throw toIoError(error, `The file ${path} could not be closed.`, { path });
      }
    },
  };
}

/** Resolves once the stream released its file descriptor (`'close'`), at once when it already has. */
function streamClosed(stream: WriteStream): Promise<void> {
  if (stream.closed) {
    return Promise.resolve();
  }
  return new Promise<void>(resolve => {
    stream.once('close', () => resolve());
  });
}

/**
 * Sink that streams to a file through `fs.createWriteStream`, honoring back-pressure. The bytes go to a hidden
 * temporary file next to the target (`.<name>.<random>.partial`), which `close()` renames over `path` once every byte
 * is on disk and the descriptor is closed; the target never holds a partial file, and an existing file there is
 * replaced only by a complete one. On `abort()` or any failure the temporary file is removed. Writing needs
 * permission to create files in the target's directory.
 */
export function toFile(path: string): ByteSink {
  const temporaryPath = join(dirname(path), `.${basename(path)}.${randomBytes(6).toString('hex')}.partial`);
  let opened: { readonly stream: WriteStream; readonly sink: ByteSink } | undefined;
  /** An abort that came before anything was written: there is no file to remove, only later calls to refuse. */
  let abortedEarly: XlsxError | undefined;
  let discarded: Promise<void> | undefined;

  /** The temporary file is created by the first write or close, so a sink that is never used leaves nothing behind. */
  function open(): { readonly stream: WriteStream; readonly sink: ByteSink } {
    if (abortedEarly) {
      throw abortedEarly;
    }
    if (opened === undefined) {
      const stream = createWriteStream(temporaryPath, { flags: 'wx' });
      opened = { stream, sink: toWritable(stream) };
    }
    return opened;
  }

  /** Tear the stream down and remove the temporary file; once, however many paths fail. */
  function discard(): Promise<void> {
    discarded ??= (async (): Promise<void> => {
      if (opened === undefined) {
        return;
      }
      opened.stream.destroy();
      await streamClosed(opened.stream);
      await unlink(temporaryPath).catch(() => undefined);
    })();
    return discarded;
  }

  return {
    async write(chunk: Uint8Array): Promise<void> {
      const { sink } = open();
      try {
        await sink.write(chunk);
      } catch (error) {
        await discard();
        throw error;
      }
    },
    async close(): Promise<void> {
      try {
        const { stream, sink } = open();
        await sink.close();
        await streamClosed(stream);
        await rename(temporaryPath, path);
      } catch (error) {
        await discard();
        throw toIoError(error, `The file ${path} could not be written.`, { path });
      }
    },
    async abort(reason?: unknown): Promise<void> {
      if (opened === undefined) {
        abortedEarly ??= toAbortError(reason);
        return;
      }
      await opened.sink.abort(reason);
      await discard();
    },
  };
}

/** The subset of Node's `Writable` that `toWritable` needs (structural, so the core types stay Node-free). */
export interface NodeWritableLike {
  write(chunk: Uint8Array, callback?: (error?: Error | null) => void): boolean;
  end(callback?: () => void): unknown;
  destroy(error?: Error): unknown;
  /** `'drain'`, `'finish'` and `'close'` call the listener with no arguments; `'error'` calls it with the error. */
  once(event: string, listener: (error?: Error) => void): unknown;
  off(event: string, listener: (error?: Error) => void): unknown;
  readonly writableNeedDrain?: boolean;
  /** True once the stream was destroyed (by us, by a failure, or by a peer that went away). */
  readonly destroyed?: boolean;
  /** True once `end()` was called. */
  readonly writableEnded?: boolean;
  /** True once every byte was flushed after `end()`. */
  readonly writableFinished?: boolean;
}

/**
 * Resolves when `event` fires. Rejects with the stream's error if it errors first, and with `closedEarlyError()` if
 * it closes first: a destroyed stream (a disconnected HTTP client) emits only `'close'`, and waiting for `'drain'` or
 * `'finish'` would wait forever. Every listener is removed either way.
 */
function waitForEvent(stream: NodeWritableLike, event: string): Promise<void> {
  return new Promise<void>((resolve, reject) => {
    function detach(): void {
      stream.off(event, onEvent);
      stream.off('error', onError);
      stream.off('close', onClose);
    }
    function onEvent(): void {
      detach();
      resolve();
    }
    function onError(error?: Error): void {
      detach();
      reject(error ?? closedEarlyError());
    }
    function onClose(): void {
      detach();
      reject(closedEarlyError());
    }
    stream.once(event, onEvent);
    stream.once('error', onError);
    stream.once('close', onClose);
  });
}

/**
 * Sink over any Node `Writable` (an HTTP response, a socket, a `PassThrough`). A stream that errors, or closes before
 * the file is complete (a client that disconnects mid-download), fails the pending and every later call with
 * `IO_ERROR`; `abort()` destroys the stream.
 */
export function toWritable(writable: NodeWritableLike): ByteSink {
  let failure: XlsxError | undefined;
  let closedEarly = false;
  let aborted = false;

  function fail(reason: unknown): XlsxError {
    failure ??= toIoError(reason, 'Writing the spreadsheet to the stream failed.');
    return failure;
  }

  // Subscribed for the life of the sink so the failure is sticky, and re-subscribed after every event because an
  // 'error' with no listener takes the process down.
  function captureError(error?: Error): void {
    fail(error ?? closedEarlyError());
    writable.once('error', captureError);
  }
  writable.once('error', captureError);
  // A destroyed stream says so only with 'close' (and `destroyed`, which not every Writable-alike keeps).
  writable.once('close', () => {
    closedEarly = writable.writableFinished !== true;
  });

  function assertWritable(): void {
    if (failure) {
      throw failure;
    }
    if (closedEarly || writable.destroyed === true || writable.writableEnded === true) {
      throw fail(closedEarlyError());
    }
  }

  return {
    async write(chunk: Uint8Array): Promise<void> {
      assertWritable();
      if (!writable.write(chunk)) {
        try {
          await waitForEvent(writable, 'drain');
        } catch (error) {
          throw fail(error);
        }
      }
    },
    async close(): Promise<void> {
      if (failure) {
        throw failure;
      }
      if (writable.writableFinished === true) {
        return;
      }
      if (closedEarly || writable.destroyed === true) {
        throw fail(closedEarlyError());
      }
      const finished = waitForEvent(writable, 'finish');
      if (writable.writableEnded !== true) {
        writable.end();
      }
      try {
        await finished;
      } catch (error) {
        throw fail(error);
      }
    },
    async abort(reason?: unknown): Promise<void> {
      if (aborted) {
        return;
      }
      aborted = true;
      const error = toAbortError(reason);
      failure ??= error;
      // Destroyed with an error only when the caller gave a reason, as `stream.destroy(reason)` would be.
      writable.destroy(reason === undefined ? undefined : error);
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
    let failure: XlsxError | undefined;
    let pumping = false;
    let pumpFinished: Promise<void> = Promise.resolve();

    /** zlib's own errors are `IO_ERROR` (the zlib error is the `cause`); the consumer's `XlsxError`s pass through. */
    function recordFailure(reason: unknown): XlsxError {
      failure ??= toIoError(reason, 'The spreadsheet could not be compressed.');
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
          try {
            await waitForEvent(deflate, 'drain');
          } catch (error) {
            throw recordFailure(error);
          }
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
            reject(
              recordFailure(
                new XlsxError('IO_ERROR', 'The spreadsheet could not be compressed: the compressor closed before it finished.'),
              ),
            );
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
        failure ??= toAbortError(reason);
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
