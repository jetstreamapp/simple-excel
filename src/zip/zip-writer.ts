import { notImplemented } from '../internal/not-implemented';
import type { ByteSink, DeflaterFactory } from '../types';

export interface ZipWriterOptions {
  /** Emit zip64 extra fields in every local header and central entry (decided up-front, never retroactively). */
  readonly zip64: boolean;
  /** Fixed DOS timestamp (2026-01-01 00:00:00) for reproducible bytes. */
  readonly deterministic: boolean;
  readonly deflater: DeflaterFactory;
  readonly compression: 'deflate' | 'store';
}

export interface ZipEntryOptions {
  /** Override the writer-level compression for this entry. */
  readonly method?: 'deflate' | 'store';
}

export interface ZipEntrySummary {
  readonly name: string;
  readonly crc32: number;
  readonly compressedSize: number;
  readonly uncompressedSize: number;
  readonly offset: number;
}

export interface ZipEntryWriter {
  write(chunk: Uint8Array): Promise<void>;
  close(): Promise<ZipEntrySummary>;
}

export interface ZipWriteResult {
  readonly bytes: number;
  readonly entries: readonly ZipEntrySummary[];
}

/**
 * Streaming zip writer. Each entry is a local file header with bit 3 (sizes and CRC in a trailing data descriptor)
 * and bit 11 (UTF-8 names), followed by the compressed bytes as the deflater produces them, then the descriptor.
 * The central directory is buffered in memory and written by `close()`. Only one entry may be open at a time.
 */
export class ZipWriter {
  constructor(sink: ByteSink, options: ZipWriterOptions) {
    void sink;
    void options;
    throw notImplemented('zip/zip-writer');
  }

  /** Bytes written to the sink so far. */
  get bytesWritten(): number {
    throw notImplemented('zip/zip-writer');
  }

  beginEntry(name: string, options?: ZipEntryOptions): Promise<ZipEntryWriter> {
    void name;
    void options;
    throw notImplemented('zip/zip-writer');
  }

  /** Convenience for small parts: one entry, one chunk. */
  writeEntry(name: string, bytes: Uint8Array, options?: ZipEntryOptions): Promise<ZipEntrySummary> {
    void name;
    void bytes;
    void options;
    throw notImplemented('zip/zip-writer');
  }

  /** Write the central directory (and zip64 records when enabled) and close the sink. */
  close(): Promise<ZipWriteResult> {
    throw notImplemented('zip/zip-writer');
  }

  abort(reason?: unknown): Promise<void> {
    void reason;
    throw notImplemented('zip/zip-writer');
  }
}
