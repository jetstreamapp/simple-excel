import { XlsxError } from '../errors';
import type { ByteSink, Deflater, DeflaterFactory } from '../types';
import { crc32, CRC32_INITIAL } from './crc32';

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
  /**
   * Override the writer-level zip64 decision for this entry. Zip64 has to be declared in the local file header
   * before any data is written (ADR-002), so the facade decides per worksheet: a sheet whose size is unknown gets
   * it, small fixed parts do not. Defaults to the writer-level flag for `beginEntry` and to false for `writeEntry`,
   * whose one-chunk payloads never approach 4 GiB.
   */
  readonly zip64?: boolean;
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

const LOCAL_HEADER_SIGNATURE = 0x0403_4b50;
const DATA_DESCRIPTOR_SIGNATURE = 0x0807_4b50;
const CENTRAL_HEADER_SIGNATURE = 0x0201_4b50;
const ZIP64_END_SIGNATURE = 0x0606_4b50;
const ZIP64_LOCATOR_SIGNATURE = 0x0706_4b50;
const END_OF_CENTRAL_DIRECTORY_SIGNATURE = 0x0605_4b50;

const ZIP64_EXTRA_FIELD_ID = 0x0001;
/** id + size + 8-byte uncompressed + 8-byte compressed placeholders. */
const ZIP64_LOCAL_EXTRA_LENGTH = 20;
/** id + size + 8-byte uncompressed + compressed + local-header offset. */
const ZIP64_CENTRAL_EXTRA_LENGTH = 28;

/** Bit 3: CRC and sizes follow the data. Bit 11: the name is UTF-8. */
const GENERAL_PURPOSE_FLAGS = 0x0008 | 0x0800;
const VERSION_BASE = 20;
const VERSION_ZIP64 = 45;
const METHOD_STORE = 0;
const METHOD_DEFLATE = 8;
const MAX_UINT16 = 0xffff;
const MAX_UINT32 = 0xffff_ffff;

const NAME_ENCODER = new TextEncoder();

interface DosTimestamp {
  readonly time: number;
  readonly date: number;
}

/** 2026-01-01 00:00:00, the fixed stamp deterministic mode writes into every entry. */
const DETERMINISTIC_TIMESTAMP: DosTimestamp = { time: 0, date: ((2026 - 1980) << 9) | (1 << 5) | 1 };

/** DOS epoch, used for clocks the 1980-2107 DOS range cannot express. */
const DOS_EPOCH_TIMESTAMP: DosTimestamp = { time: 0, date: (1 << 5) | 1 };

function dosTimestampFrom(when: Date): DosTimestamp {
  const year = when.getFullYear();
  if (year < 1980 || year > 2107) {
    return DOS_EPOCH_TIMESTAMP;
  }
  return {
    time: (when.getHours() << 11) | (when.getMinutes() << 5) | (when.getSeconds() >> 1),
    date: ((year - 1980) << 9) | ((when.getMonth() + 1) << 5) | when.getDate(),
  };
}

/**
 * Guards a field that has to fit in 32 bits because zip64 is off. Called before every write that could push a
 * counter past the limit, so the archive fails while its bytes are still consistent rather than silently wrapping.
 * Exported so the limit can be tested without producing four gigabytes of data.
 */
export function assertFits(value: number, what: string): void {
  if (value > MAX_UINT32) {
    throw new XlsxError(
      'ENTRY_TOO_LARGE',
      `This spreadsheet is too large for a standard zip file: ${what} passes 4 GB. Write it with zip64 enabled or split it across sheets.`,
      { value, what },
    );
  }
}

/** Little-endian record builder: one growable buffer per writer, fields written through a DataView. */
class ByteBuilder {
  private bytes: Uint8Array;
  private view: DataView;
  private position = 0;

  constructor(capacity: number) {
    this.bytes = new Uint8Array(capacity);
    this.view = new DataView(this.bytes.buffer);
  }

  uint16(value: number): void {
    this.ensure(2);
    this.view.setUint16(this.position, value, true);
    this.position += 2;
  }

  uint32(value: number): void {
    this.ensure(4);
    this.view.setUint32(this.position, value, true);
    this.position += 4;
  }

  uint64(value: number): void {
    this.ensure(8);
    this.view.setUint32(this.position, value % 0x1_0000_0000, true);
    this.view.setUint32(this.position + 4, Math.floor(value / 0x1_0000_0000), true);
    this.position += 8;
  }

  raw(chunk: Uint8Array): void {
    this.ensure(chunk.length);
    this.bytes.set(chunk, this.position);
    this.position += chunk.length;
  }

  /** A copy of what was built, so the sink owns its bytes while the builder is reused for the next record. */
  take(): Uint8Array {
    const record = this.bytes.slice(0, this.position);
    this.position = 0;
    return record;
  }

  private ensure(extra: number): void {
    if (this.position + extra <= this.bytes.length) {
      return;
    }
    const grown = new Uint8Array(Math.max(this.bytes.length * 2, this.position + extra));
    grown.set(this.bytes.subarray(0, this.position));
    this.bytes = grown;
    this.view = new DataView(grown.buffer);
  }
}

/** What `close()` needs to emit one central directory entry. */
interface CentralDirectoryRecord {
  readonly nameBytes: Uint8Array;
  readonly method: number;
  readonly timestamp: DosTimestamp;
  readonly zip64: boolean;
  readonly summary: ZipEntrySummary;
}

/**
 * Streaming zip writer. Each entry is a local file header with bit 3 (sizes and CRC in a trailing data descriptor)
 * and bit 11 (UTF-8 names), followed by the compressed bytes as the deflater produces them, then the descriptor.
 * The central directory is buffered in memory and written by `close()`. Only one entry may be open at a time.
 */
export class ZipWriter {
  /** Set once any entry was opened as zip64: the archive then finishes with the zip64 end records. */
  private anyEntryZip64 = false;
  private readonly sink: ByteSink;
  private readonly options: ZipWriterOptions;
  private readonly builder = new ByteBuilder(512);
  private readonly records: CentralDirectoryRecord[] = [];
  private offset = 0;
  private openEntryName: string | undefined;
  private openDeflater: Deflater | undefined;
  private failure: unknown;
  private aborted = false;
  private closed = false;

  constructor(sink: ByteSink, options: ZipWriterOptions) {
    this.sink = sink;
    this.options = options;
  }

  /** Bytes written to the sink so far. */
  get bytesWritten(): number {
    return this.offset;
  }

  async beginEntry(name: string, options?: ZipEntryOptions): Promise<ZipEntryWriter> {
    return this.openEntry(name, options, options?.zip64 ?? this.options.zip64);
  }

  /** Convenience for small parts: one entry, one chunk. 32-bit unless the caller asks otherwise. */
  async writeEntry(name: string, bytes: Uint8Array, options?: ZipEntryOptions): Promise<ZipEntrySummary> {
    const entry = await this.openEntry(name, options, options?.zip64 ?? false);
    await entry.write(bytes);
    return entry.close();
  }

  /** Write the central directory (and zip64 records when enabled) and close the sink. */
  async close(): Promise<ZipWriteResult> {
    this.assertUsable();
    if (this.openEntryName !== undefined) {
      throw new XlsxError('WRITER_STATE', `The zip entry "${this.openEntryName}" is still open; close it before closing the file.`);
    }
    this.closed = true;

    const centralDirectoryOffset = this.offset;
    const centralDirectory = this.buildCentralDirectory();
    await this.writeToSink(centralDirectory);
    await this.writeToSink(this.buildEndRecords(centralDirectoryOffset, centralDirectory.length));

    try {
      await this.sink.close();
    } catch (reason) {
      await this.fail(reason);
    }
    return { bytes: this.offset, entries: this.records.map(record => record.summary) };
  }

  async abort(reason?: unknown): Promise<void> {
    this.failure ??= new XlsxError('ABORTED', 'Writing the spreadsheet was cancelled.', reason === undefined ? undefined : { reason });
    if (this.aborted) {
      return;
    }
    await this.abortEverything(reason);
  }

  private async openEntry(name: string, options: ZipEntryOptions | undefined, zip64: boolean): Promise<ZipEntryWriter> {
    if (zip64) {
      this.anyEntryZip64 = true;
    }
    this.assertUsable();
    if (this.openEntryName !== undefined) {
      throw new XlsxError('WRITER_STATE', `Cannot start the zip entry "${name}" while "${this.openEntryName}" is open.`);
    }

    const method = (options?.method ?? this.options.compression) === 'store' ? METHOD_STORE : METHOD_DEFLATE;
    const timestamp = this.options.deterministic ? DETERMINISTIC_TIMESTAMP : dosTimestampFrom(new Date());
    const nameBytes = NAME_ENCODER.encode(name);
    const localHeaderOffset = this.offset;
    this.openEntryName = name;

    let crc = CRC32_INITIAL;
    let uncompressedSize = 0;
    let compressedSize = 0;
    let entryClosed = false;

    const deflater = this.options.deflater(
      async (chunk: Uint8Array): Promise<void> => {
        if (!zip64) {
          await this.guardFits(compressedSize + chunk.length, `the compressed size of "${name}"`);
        }
        await this.writeToSink(chunk);
        compressedSize += chunk.length;
      },
      { method: method === METHOD_DEFLATE ? 'deflate' : 'store' },
    );
    this.openDeflater = deflater;

    await this.writeToSink(this.buildLocalHeader(nameBytes, method, timestamp, zip64));

    return {
      write: async (chunk: Uint8Array): Promise<void> => {
        this.assertUsable();
        if (entryClosed) {
          throw new XlsxError('WRITER_STATE', `The zip entry "${name}" is closed and cannot take more bytes.`);
        }
        if (chunk.length === 0) {
          return;
        }
        if (!zip64) {
          await this.guardFits(uncompressedSize + chunk.length, `the size of "${name}"`);
        }
        crc = crc32(chunk, crc);
        uncompressedSize += chunk.length;
        try {
          await deflater.push(chunk);
        } catch (reason) {
          await this.fail(reason);
        }
      },

      close: async (): Promise<ZipEntrySummary> => {
        this.assertUsable();
        if (entryClosed) {
          throw new XlsxError('WRITER_STATE', `The zip entry "${name}" is already closed.`);
        }
        entryClosed = true;
        try {
          // Every compressed chunk has reached the sink once this resolves, so the sizes below are final.
          await deflater.finish();
        } catch (reason) {
          await this.fail(reason);
        }

        const summary: ZipEntrySummary = { name, crc32: crc, compressedSize, uncompressedSize, offset: localHeaderOffset };
        await this.writeToSink(this.buildDataDescriptor(summary, zip64));
        this.records.push({ nameBytes, method, timestamp, zip64, summary });
        this.openEntryName = undefined;
        this.openDeflater = undefined;
        return summary;
      },
    };
  }

  private buildLocalHeader(nameBytes: Uint8Array, method: number, timestamp: DosTimestamp, zip64: boolean): Uint8Array {
    const builder = this.builder;
    builder.uint32(LOCAL_HEADER_SIGNATURE);
    builder.uint16(zip64 ? VERSION_ZIP64 : VERSION_BASE);
    builder.uint16(GENERAL_PURPOSE_FLAGS);
    builder.uint16(method);
    builder.uint16(timestamp.time);
    builder.uint16(timestamp.date);
    // CRC and both sizes live in the data descriptor, so APPNOTE 4.4.9 puts literal zeros here - for a zip64 entry
    // too, whose zip64 extra field below carries the matching 8-byte zero placeholders. Writing the 0xFFFFFFFF
    // "see the extra field" sentinel instead makes Excel silently repair the archive (oracle 2026-09-12
    // phase-b-writer: sentinel REPAIRED, zeros PASS).
    builder.uint32(0);
    builder.uint32(0);
    builder.uint32(0);
    builder.uint16(nameBytes.length);
    builder.uint16(zip64 ? ZIP64_LOCAL_EXTRA_LENGTH : 0);
    builder.raw(nameBytes);
    if (zip64) {
      builder.uint16(ZIP64_EXTRA_FIELD_ID);
      builder.uint16(ZIP64_LOCAL_EXTRA_LENGTH - 4);
      builder.uint64(0);
      builder.uint64(0);
    }
    return builder.take();
  }

  private buildDataDescriptor(summary: ZipEntrySummary, zip64: boolean): Uint8Array {
    const builder = this.builder;
    builder.uint32(DATA_DESCRIPTOR_SIGNATURE);
    builder.uint32(summary.crc32);
    if (zip64) {
      builder.uint64(summary.compressedSize);
      builder.uint64(summary.uncompressedSize);
    } else {
      builder.uint32(summary.compressedSize);
      builder.uint32(summary.uncompressedSize);
    }
    return builder.take();
  }

  private buildCentralDirectory(): Uint8Array {
    const builder = new ByteBuilder(Math.max(64, this.records.length * 80));
    for (const { nameBytes, method, timestamp, zip64, summary } of this.records) {
      // A 32-bit entry in a zip64 archive stays 32-bit unless one of its own values actually overflows.
      const needsZip64 =
        zip64 || summary.compressedSize > MAX_UINT32 || summary.uncompressedSize > MAX_UINT32 || summary.offset > MAX_UINT32;
      const version = needsZip64 ? VERSION_ZIP64 : VERSION_BASE;
      builder.uint32(CENTRAL_HEADER_SIGNATURE);
      builder.uint16(version); // version made by
      builder.uint16(version); // version needed
      builder.uint16(GENERAL_PURPOSE_FLAGS);
      builder.uint16(method);
      builder.uint16(timestamp.time);
      builder.uint16(timestamp.date);
      builder.uint32(summary.crc32);
      builder.uint32(needsZip64 ? MAX_UINT32 : summary.compressedSize);
      builder.uint32(needsZip64 ? MAX_UINT32 : summary.uncompressedSize);
      builder.uint16(nameBytes.length);
      builder.uint16(needsZip64 ? ZIP64_CENTRAL_EXTRA_LENGTH : 0);
      builder.uint16(0); // comment length
      builder.uint16(0); // disk number
      builder.uint16(0); // internal attributes
      builder.uint32(0); // external attributes
      builder.uint32(needsZip64 ? MAX_UINT32 : summary.offset);
      builder.raw(nameBytes);
      if (needsZip64) {
        builder.uint16(ZIP64_EXTRA_FIELD_ID);
        builder.uint16(ZIP64_CENTRAL_EXTRA_LENGTH - 4);
        builder.uint64(summary.uncompressedSize);
        builder.uint64(summary.compressedSize);
        builder.uint64(summary.offset);
      }
    }
    return builder.take();
  }

  private buildEndRecords(centralDirectoryOffset: number, centralDirectorySize: number): Uint8Array {
    const builder = this.builder;
    const entryCount = this.records.length;
    const needsZip64 =
      this.options.zip64 ||
      this.anyEntryZip64 ||
      entryCount > MAX_UINT16 ||
      centralDirectoryOffset > MAX_UINT32 ||
      centralDirectorySize > MAX_UINT32;

    if (needsZip64) {
      const zip64EndOffset = this.offset;
      builder.uint32(ZIP64_END_SIGNATURE);
      builder.uint64(44); // size of this record past this field
      builder.uint16(VERSION_ZIP64); // version made by
      builder.uint16(VERSION_ZIP64); // version needed
      builder.uint32(0); // this disk
      builder.uint32(0); // disk holding the central directory
      builder.uint64(entryCount);
      builder.uint64(entryCount);
      builder.uint64(centralDirectorySize);
      builder.uint64(centralDirectoryOffset);

      builder.uint32(ZIP64_LOCATOR_SIGNATURE);
      builder.uint32(0); // disk holding the zip64 end record
      builder.uint64(zip64EndOffset);
      builder.uint32(1); // total disks
    }

    // Once the zip64 pair is written the 32-bit end record must point at it with sentinels in every field, even where
    // the real value would still fit: Excel decides the archive is zip64 from these and repairs it when they read as
    // a complete 32-bit record (oracle 2026-09-12 phase-b-writer: real values REPAIRED, sentinels PASS).
    builder.uint32(END_OF_CENTRAL_DIRECTORY_SIGNATURE);
    builder.uint16(0); // this disk
    builder.uint16(0); // disk holding the central directory
    builder.uint16(needsZip64 ? MAX_UINT16 : entryCount);
    builder.uint16(needsZip64 ? MAX_UINT16 : entryCount);
    builder.uint32(needsZip64 ? MAX_UINT32 : centralDirectorySize);
    builder.uint32(needsZip64 ? MAX_UINT32 : centralDirectoryOffset);
    builder.uint16(0); // comment length
    return builder.take();
  }

  private async writeToSink(chunk: Uint8Array): Promise<void> {
    if (!this.options.zip64 && !this.anyEntryZip64) {
      await this.guardFits(this.offset + chunk.length, 'the size of the file');
    }
    try {
      await this.sink.write(chunk);
    } catch (reason) {
      await this.fail(reason);
    }
    this.offset += chunk.length;
  }

  private async guardFits(value: number, what: string): Promise<void> {
    try {
      assertFits(value, what);
    } catch (reason) {
      await this.fail(reason);
    }
  }

  private assertUsable(): void {
    if (this.failure !== undefined) {
      throw this.failure;
    }
    if (this.closed) {
      throw new XlsxError('WRITER_STATE', 'The zip file is already closed.');
    }
  }

  /** Latch the first failure, tear the archive down and rethrow it; every later call sees the same reason. */
  private async fail(reason: unknown): Promise<never> {
    if (this.failure === undefined) {
      this.failure = reason;
      await this.abortEverything(reason);
    }
    throw this.failure;
  }

  private async abortEverything(reason: unknown): Promise<void> {
    this.aborted = true;
    const deflater = this.openDeflater;
    this.openDeflater = undefined;
    this.openEntryName = undefined;
    if (deflater !== undefined) {
      // Deliberately not awaited: the failure may have come from inside this deflater's own chunk handler, and
      // waiting there for its drain loop to settle would wait on ourselves. Aborting releases the stream either way.
      void deflater.abort(reason).catch(() => {});
    }
    try {
      await this.sink.abort(reason);
    } catch {
      // The sink is already broken; the failure that got us here is the one worth reporting.
    }
  }
}
