/**
 * Node-only helpers: file-backed sources and sinks and a zlib-backed deflater. Everything in the core entry works in
 * Node too; this entry only adds what needs `node:` modules.
 */
import { notImplemented } from '../internal/not-implemented';
import type { ByteSink, DeflaterFactory, RandomAccessSource } from '../types';

export * from '../index';

export interface FileSource extends RandomAccessSource {
  close(): Promise<void>;
}

/** Random-access source over a file (positional reads on a `FileHandle`); `close()` releases the handle. */
export function fromFile(path: string): Promise<FileSource> {
  void path;
  throw notImplemented('node/fromFile');
}

/** Sink that streams to a file through `fs.createWriteStream`, honoring back-pressure. */
export function toFile(path: string): ByteSink {
  void path;
  throw notImplemented('node/toFile');
}

/** The subset of Node's `Writable` that `toWritable` needs (structural, so the core types stay Node-free). */
export interface NodeWritableLike {
  write(chunk: Uint8Array, callback?: (error?: Error | null) => void): boolean;
  end(callback?: () => void): unknown;
  destroy(error?: Error): unknown;
  once(event: string, listener: (...args: never[]) => void): unknown;
  off(event: string, listener: (...args: never[]) => void): unknown;
  readonly writableNeedDrain?: boolean;
}

/** Sink over any Node `Writable` (a response, a socket, a `PassThrough`). */
export function toWritable(writable: NodeWritableLike): ByteSink {
  void writable;
  throw notImplemented('node/toWritable');
}

/** Deflater over `zlib.createDeflateRaw` (`level` 1-9; 1 is 2-3x faster than the default for xlsx-shaped XML). */
export function nodeDeflater(level?: number): DeflaterFactory {
  void level;
  throw notImplemented('node/nodeDeflater');
}
