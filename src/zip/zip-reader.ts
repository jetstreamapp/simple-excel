import { notImplemented } from '../internal/not-implemented';
import type { RandomAccessSource } from '../types';

export interface ZipReaderLimits {
  /** Default 10,000. */
  readonly maxEntries: number;
  /** Per entry inflated cap. Default 1 GiB. */
  readonly maxInflatedBytes: number;
}

export interface ZipEntry {
  /** Normalized name: backslashes to slashes, no leading slash. */
  readonly name: string;
  readonly method: 'store' | 'deflate';
  readonly compressedSize: number;
  readonly uncompressedSize: number;
  readonly crc32: number;
  /** Offset of the local file header. */
  readonly localHeaderOffset: number;
}

/**
 * Random-access zip reader: finds the end-of-central-directory record (and zip64 locator) in the tail, parses the
 * central directory into `entries`, and streams entries through the inflater on demand. The local header is read
 * for its own name/extra lengths (they may differ from the central copy).
 */
export class ZipReader {
  readonly entries: ReadonlyMap<string, ZipEntry>;

  private constructor(entries: ReadonlyMap<string, ZipEntry>) {
    this.entries = entries;
  }

  static open(source: RandomAccessSource, limits: ZipReaderLimits): Promise<ZipReader> {
    void source;
    void limits;
    throw notImplemented('zip/zip-reader');
  }

  has(name: string): boolean {
    void name;
    throw notImplemented('zip/zip-reader');
  }

  /**
   * Inflated chunks of one entry in order. Breaking out of the loop early cancels the inflate. The CRC is checked
   * once the entry is fully consumed and a mismatch throws `ZIP_CRC_MISMATCH` after the last chunk.
   */
  stream(name: string): AsyncIterable<Uint8Array> {
    void name;
    throw notImplemented('zip/zip-reader');
  }

  /** Whole entry in memory (small parts only); the CRC is verified. */
  read(name: string): Promise<Uint8Array> {
    void name;
    throw notImplemented('zip/zip-reader');
  }

  /** Whole entry decoded as UTF-8 text (BOM stripped). */
  readText(name: string): Promise<string> {
    void name;
    throw notImplemented('zip/zip-reader');
  }

  close(): Promise<void> {
    throw notImplemented('zip/zip-reader');
  }
}
