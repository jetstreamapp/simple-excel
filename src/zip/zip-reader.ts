import { createInflater } from '../compress/inflater';
import { XlsxError } from '../errors';
import type { RandomAccessSource } from '../types';
import { crc32, CRC32_INITIAL } from './crc32';

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

const EOCD_SIGNATURE = 0x0605_4b50;
const ZIP64_EOCD_SIGNATURE = 0x0606_4b50;
const ZIP64_LOCATOR_SIGNATURE = 0x0706_4b50;
const CENTRAL_HEADER_SIGNATURE = 0x0201_4b50;
const LOCAL_HEADER_SIGNATURE = 0x0403_4b50;

const EOCD_SIZE = 22;
const ZIP64_LOCATOR_SIZE = 20;
const ZIP64_EOCD_SIZE = 56;
const CENTRAL_HEADER_SIZE = 46;
const LOCAL_HEADER_SIZE = 30;
const MAX_COMMENT_SIZE = 65_535;
/** The EOCD plus a maximal comment, the zip64 locator and the zip64 record that may precede it. */
const TAIL_SIZE = EOCD_SIZE + MAX_COMMENT_SIZE + ZIP64_LOCATOR_SIZE + ZIP64_EOCD_SIZE;

const UINT16_MARKER = 0xffff;
const UINT32_MARKER = 0xffff_ffff;
const ZIP64_EXTRA_HEADER_ID = 0x0001;

const FLAG_ENCRYPTED = 0x0001;
const METHOD_STORE = 0;
const METHOD_DEFLATE = 8;

const READ_CHUNK_SIZE = 256 * 1024;
const LARGE_READ_CHUNK_SIZE = 1024 * 1024;
const LARGE_SOURCE_SIZE = 64 * 1024 * 1024;
/** How far an entry may inflate past the size its central-directory record declares before we call it a bomb. */
const DECLARED_SIZE_SLACK_BYTES = 4096;

const nameDecoder = new TextDecoder('utf-8');
/** `ignoreBOM` keeps a leading U+FEFF in the output so `readText` strips exactly one, never two. */
const textDecoder = new TextDecoder('utf-8', { ignoreBOM: true });

function truncated(what: string): XlsxError {
  return new XlsxError('ZIP_TRUNCATED', `This file is not a readable zip archive (${what}); it is probably damaged or incomplete.`);
}

/** OPC part names use forward slashes and no leading slash; zip writers on Windows sometimes disagree. */
function normalizeEntryName(name: string): string {
  let normalized = name.includes('\\') ? name.split('\\').join('/') : name;
  let start = 0;
  while (start < normalized.length && normalized.charCodeAt(start) === 0x2f) {
    start++;
  }
  if (start > 0) {
    normalized = normalized.slice(start);
  }
  return normalized;
}

/**
 * OPC part names compare ASCII case-insensitively (ECMA-376 Part 2, 6.2.2.3), so `xl/SharedStrings.xml` answers to a
 * relationship that says `sharedStrings.xml`. Only A-Z fold: a Unicode-aware `toLowerCase` would equate names the
 * standard keeps apart.
 */
function asciiLowerCase(name: string): string {
  let index = 0;
  while (index < name.length) {
    const code = name.charCodeAt(index);
    if (code >= 0x41 && code <= 0x5a) {
      break;
    }
    index++;
  }
  if (index === name.length) {
    return name;
  }
  let lowered = name.slice(0, index);
  for (; index < name.length; index++) {
    const code = name.charCodeAt(index);
    lowered += code >= 0x41 && code <= 0x5a ? String.fromCharCode(code + 0x20) : name.charAt(index);
  }
  return lowered;
}

/** First entry wins for each folded name; the exact name, which `entries` keys by, is always tried first. */
function foldedIndex(entries: ReadonlyMap<string, ZipEntry>): ReadonlyMap<string, ZipEntry> {
  const index = new Map<string, ZipEntry>();
  for (const [name, entry] of entries) {
    const folded = asciiLowerCase(name);
    if (!index.has(folded)) {
      index.set(folded, entry);
    }
  }
  return index;
}

function closedError(): XlsxError {
  return new XlsxError('ABORTED', 'This workbook was closed. Open it again to read from it.');
}

function readUint64(view: DataView, offset: number): number {
  const value = view.getBigUint64(offset, true);
  if (value > BigInt(Number.MAX_SAFE_INTEGER)) {
    throw new XlsxError('LIMIT_EXCEEDED', 'This file declares a size larger than this library can address.', { value: value.toString() });
  }
  return Number(value);
}

interface CentralDirectoryLocation {
  readonly offset: number;
  readonly size: number;
  readonly entryCount: number;
}

/**
 * A one-slot hand-off between the inflate pump and the async generator that yields to the caller. The pump's
 * `send` resolves only once the chunk was taken, which is what keeps a streaming read flat in memory.
 */
interface ChunkChannel {
  send(chunk: Uint8Array): Promise<void>;
  close(): void;
  fail(error: unknown): void;
  /** Resolves with `undefined` once the producer closed the channel. */
  receive(): Promise<Uint8Array | undefined>;
  /** The receiver went away: release the producer and drop everything it sends from here on. */
  cancel(): void;
}

function createChunkChannel(): ChunkChannel {
  let pendingChunk: Uint8Array | undefined;
  let releaseSender: (() => void) | undefined;
  let deliver: ((chunk: Uint8Array | undefined) => void) | undefined;
  let rejectReceiver: ((error: unknown) => void) | undefined;
  let failure: unknown;
  let closed = false;
  let cancelled = false;

  return {
    send(chunk: Uint8Array): Promise<void> {
      if (cancelled) {
        return Promise.resolve();
      }
      if (deliver) {
        const waiting = deliver;
        deliver = undefined;
        rejectReceiver = undefined;
        waiting(chunk);
        return Promise.resolve();
      }
      pendingChunk = chunk;
      return new Promise<void>(resolve => {
        releaseSender = resolve;
      });
    },
    close(): void {
      closed = true;
      if (deliver) {
        const waiting = deliver;
        deliver = undefined;
        rejectReceiver = undefined;
        waiting(undefined);
      }
    },
    fail(error: unknown): void {
      failure ??= error;
      if (rejectReceiver) {
        const waiting = rejectReceiver;
        deliver = undefined;
        rejectReceiver = undefined;
        waiting(error);
      }
    },
    receive(): Promise<Uint8Array | undefined> {
      if (pendingChunk !== undefined) {
        const chunk = pendingChunk;
        pendingChunk = undefined;
        const release = releaseSender;
        releaseSender = undefined;
        release?.();
        return Promise.resolve(chunk);
      }
      if (failure !== undefined) {
        return Promise.reject(failure);
      }
      if (closed) {
        return Promise.resolve(undefined);
      }
      return new Promise<Uint8Array | undefined>((resolve, reject) => {
        deliver = resolve;
        rejectReceiver = reject;
      });
    },
    cancel(): void {
      cancelled = true;
      pendingChunk = undefined;
      const release = releaseSender;
      releaseSender = undefined;
      release?.();
    },
  };
}

/**
 * Random-access zip reader: finds the end-of-central-directory record (and zip64 locator) in the tail, parses the
 * central directory into `entries`, and streams entries through the inflater on demand. The local header is read
 * for its own name/extra lengths (they may differ from the central copy).
 */
export class ZipReader {
  /** Keyed by the exact (normalized) entry name. `has`, `stream`, `read` and `readText` also match without case. */
  readonly entries: ReadonlyMap<string, ZipEntry>;

  private readonly source: RandomAccessSource;
  private readonly limits: ZipReaderLimits;
  /** `entries` keyed by ASCII-lower-cased name: the fallback when a part name's case does not match (EC-PART-NAME-CASE). */
  private readonly entriesByFoldedName: ReadonlyMap<string, ZipEntry>;
  private closed = false;

  private constructor(source: RandomAccessSource, limits: ZipReaderLimits, entries: ReadonlyMap<string, ZipEntry>) {
    this.source = source;
    this.limits = limits;
    this.entries = entries;
    this.entriesByFoldedName = foldedIndex(entries);
  }

  static async open(source: RandomAccessSource, limits: ZipReaderLimits): Promise<ZipReader> {
    const location = await findCentralDirectory(source);
    if (location.entryCount > limits.maxEntries) {
      throw new XlsxError(
        'LIMIT_EXCEEDED',
        `This file contains ${location.entryCount} parts, more than the ${limits.maxEntries} allowed.`,
        {
          entryCount: location.entryCount,
          maxEntries: limits.maxEntries,
        },
      );
    }
    if (location.offset + location.size > source.size) {
      throw truncated('the central directory runs past the end of the file');
    }
    const centralDirectory = await readExactFrom(source, location.offset, location.size);
    return new ZipReader(source, limits, parseCentralDirectory(centralDirectory, location.entryCount));
  }

  has(name: string): boolean {
    return this.entry(name) !== undefined;
  }

  /**
   * Inflated chunks of one entry in order. Breaking out of the loop early cancels the inflate. The CRC is checked
   * once the entry is fully consumed and a mismatch throws `ZIP_CRC_MISMATCH` after the last chunk.
   */
  stream(name: string): AsyncIterable<Uint8Array> {
    return this.streamEntry(name);
  }

  /** Whole entry in memory (small parts only); the CRC is verified. */
  async read(name: string): Promise<Uint8Array> {
    const entry = this.entry(name);
    // The declared size is almost always exact, so one allocation up front avoids copying the part twice.
    let declared: Uint8Array | undefined = entry && entry.uncompressedSize > 0 ? new Uint8Array(entry.uncompressedSize) : undefined;
    const overflow: Uint8Array[] = [];
    let total = 0;

    for await (const chunk of this.stream(name)) {
      if (declared && total + chunk.byteLength <= declared.byteLength) {
        declared.set(chunk, total);
      } else {
        if (declared) {
          overflow.push(declared.subarray(0, total));
          declared = undefined;
        }
        overflow.push(chunk);
      }
      total += chunk.byteLength;
    }

    if (declared) {
      return total === declared.byteLength ? declared : declared.subarray(0, total);
    }
    const joined = new Uint8Array(total);
    let offset = 0;
    for (const chunk of overflow) {
      joined.set(chunk, offset);
      offset += chunk.byteLength;
    }
    return joined;
  }

  /** Whole entry decoded as UTF-8 text (BOM stripped). */
  async readText(name: string): Promise<string> {
    const text = textDecoder.decode(await this.read(name));
    return text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
  }

  /**
   * Release the source. Reading anything afterwards, including a stream that was still open, fails with
   * `XlsxError('ABORTED')` whatever kind of source it is, rather than with whatever the closed source throws.
   */
  async close(): Promise<void> {
    if (this.closed) {
      return;
    }
    this.closed = true;
    await this.source.close?.();
  }

  /** The exact name first, then the same name in any ASCII case. */
  private entry(name: string): ZipEntry | undefined {
    const normalized = normalizeEntryName(name);
    return this.entries.get(normalized) ?? this.entriesByFoldedName.get(asciiLowerCase(normalized));
  }

  private async *streamEntry(name: string): AsyncGenerator<Uint8Array> {
    if (this.closed) {
      throw closedError();
    }
    const entry = this.entry(name);
    if (!entry) {
      throw new XlsxError('NOT_XLSX', `missing part ${name}`, { part: name });
    }
    if (entry.uncompressedSize > this.limits.maxInflatedBytes) {
      throw new XlsxError(
        'ZIP_BOMB',
        `The part ${entry.name} declares ${entry.uncompressedSize} bytes, more than the ${this.limits.maxInflatedBytes} byte limit.`,
        { part: entry.name, declaredSize: entry.uncompressedSize, maxInflatedBytes: this.limits.maxInflatedBytes },
      );
    }

    let dataStart: number;
    try {
      dataStart = await this.findEntryData(entry);
    } catch (error) {
      throw this.closed ? closedError() : error;
    }
    const declaredCeiling = entry.uncompressedSize + DECLARED_SIZE_SLACK_BYTES;
    const channel = createChunkChannel();
    let inflatedBytes = 0;
    let runningCrc = CRC32_INITIAL;

    const inflater = createInflater(
      async (chunk: Uint8Array): Promise<void> => {
        inflatedBytes += chunk.byteLength;
        if (inflatedBytes > declaredCeiling) {
          throw new XlsxError('ZIP_BOMB', `The part ${entry.name} inflates well past the ${entry.uncompressedSize} bytes it declares.`, {
            part: entry.name,
            declaredSize: entry.uncompressedSize,
          });
        }
        runningCrc = crc32(chunk, runningCrc);
        await channel.send(chunk);
      },
      { maxBytes: this.limits.maxInflatedBytes, method: entry.method },
    );

    const chunkSize = this.source.size > LARGE_SOURCE_SIZE ? LARGE_READ_CHUNK_SIZE : READ_CHUNK_SIZE;
    const pump: Promise<void> = (async (): Promise<void> => {
      try {
        let position = dataStart;
        let remaining = entry.compressedSize;
        while (remaining > 0) {
          // Every source kind stops the same way, an in-memory one included.
          if (this.closed) {
            throw closedError();
          }
          const wanted = Math.min(chunkSize, remaining);
          const compressed = await readExactFrom(this.source, position, wanted);
          position += wanted;
          remaining -= wanted;
          await inflater.push(compressed);
        }
        await inflater.finish();
        if (runningCrc !== entry.crc32) {
          throw new XlsxError('ZIP_CRC_MISMATCH', `The part ${entry.name} does not match its checksum; the file is damaged.`, {
            part: entry.name,
            expected: entry.crc32,
            actual: runningCrc,
          });
        }
        channel.close();
      } catch (error) {
        // A source closed under a running read fails in its own way (a file handle says EBADF); report the cause.
        channel.fail(this.closed ? closedError() : error);
      }
    })();

    try {
      for (;;) {
        const chunk = await channel.receive();
        if (chunk === undefined) {
          return;
        }
        if (this.closed) {
          throw closedError();
        }
        yield chunk;
      }
    } finally {
      // Release the pump before aborting: it may be parked in `send` waiting for a chunk nobody will take.
      channel.cancel();
      await inflater.abort();
      await pump;
    }
  }

  /**
   * Start of the entry's compressed data. The local header's own name and extra lengths decide it - they are
   * allowed to differ from the central directory's copy (EC-ZIP catalog).
   */
  private async findEntryData(entry: ZipEntry): Promise<number> {
    if (entry.localHeaderOffset + LOCAL_HEADER_SIZE > this.source.size) {
      throw truncated(`the header for ${entry.name} is past the end of the file`);
    }
    const header = await readExactFrom(this.source, entry.localHeaderOffset, LOCAL_HEADER_SIZE);
    const view = new DataView(header.buffer, header.byteOffset, header.byteLength);
    if (view.getUint32(0, true) !== LOCAL_HEADER_SIGNATURE) {
      throw truncated(`the header for ${entry.name} is not a local file header`);
    }
    const nameLength = view.getUint16(26, true);
    const extraLength = view.getUint16(28, true);
    const dataStart = entry.localHeaderOffset + LOCAL_HEADER_SIZE + nameLength + extraLength;
    if (dataStart + entry.compressedSize > this.source.size) {
      throw truncated(`the data for ${entry.name} runs past the end of the file`);
    }
    return dataStart;
  }
}

async function readExactFrom(source: RandomAccessSource, offset: number, length: number): Promise<Uint8Array> {
  const bytes = await source.read(offset, length);
  if (bytes.byteLength < length) {
    throw truncated(`${length} bytes were needed at offset ${offset} but only ${bytes.byteLength} exist`);
  }
  return bytes;
}

/**
 * Scan the tail backwards for the end-of-central-directory record, then follow the zip64 locator when its fixed
 * fields are saturated with markers.
 */
async function findCentralDirectory(source: RandomAccessSource): Promise<CentralDirectoryLocation> {
  const { size } = source;
  if (size < EOCD_SIZE) {
    throw truncated('it is too small to hold a zip directory');
  }
  const tailLength = Math.min(size, TAIL_SIZE);
  const tailStart = size - tailLength;
  const tail = await readExactFrom(source, tailStart, tailLength);
  const tailView = new DataView(tail.buffer, tail.byteOffset, tail.byteLength);

  for (let index = tail.byteLength - EOCD_SIZE; index >= 0; index--) {
    if (tailView.getUint32(index, true) !== EOCD_SIGNATURE) {
      continue;
    }
    const commentLength = tailView.getUint16(index + 20, true);
    if (index + EOCD_SIZE + commentLength > tail.byteLength) {
      continue;
    }
    const entryCount = tailView.getUint16(index + 10, true);
    const centralSize = tailView.getUint32(index + 12, true);
    const centralOffset = tailView.getUint32(index + 16, true);
    const needsZip64 = entryCount === UINT16_MARKER || centralSize === UINT32_MARKER || centralOffset === UINT32_MARKER;
    if (needsZip64) {
      return readZip64Location(source, tailStart + index);
    }
    if (centralOffset + centralSize > size) {
      continue;
    }
    return { offset: centralOffset, size: centralSize, entryCount };
  }
  throw truncated('no end-of-central-directory record was found');
}

async function readZip64Location(source: RandomAccessSource, eocdOffset: number): Promise<CentralDirectoryLocation> {
  if (eocdOffset < ZIP64_LOCATOR_SIZE) {
    throw truncated('the zip64 locator is missing');
  }
  const locator = await readExactFrom(source, eocdOffset - ZIP64_LOCATOR_SIZE, ZIP64_LOCATOR_SIZE);
  const locatorView = new DataView(locator.buffer, locator.byteOffset, locator.byteLength);
  if (locatorView.getUint32(0, true) !== ZIP64_LOCATOR_SIGNATURE) {
    throw truncated('the zip64 locator is missing');
  }

  const recordOffset = readUint64(locatorView, 8);
  if (recordOffset + ZIP64_EOCD_SIZE > source.size) {
    throw truncated('the zip64 directory record is past the end of the file');
  }
  const record = await readExactFrom(source, recordOffset, ZIP64_EOCD_SIZE);
  const recordView = new DataView(record.buffer, record.byteOffset, record.byteLength);
  if (recordView.getUint32(0, true) !== ZIP64_EOCD_SIGNATURE) {
    throw truncated('the zip64 directory record is missing');
  }
  return {
    entryCount: readUint64(recordView, 32),
    size: readUint64(recordView, 40),
    offset: readUint64(recordView, 48),
  };
}

function parseCentralDirectory(bytes: Uint8Array, entryCount: number): ReadonlyMap<string, ZipEntry> {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const entries = new Map<string, ZipEntry>();
  let cursor = 0;

  for (let index = 0; index < entryCount; index++) {
    if (cursor + CENTRAL_HEADER_SIZE > bytes.byteLength || view.getUint32(cursor, true) !== CENTRAL_HEADER_SIGNATURE) {
      throw truncated(`the central directory ends after ${index} of ${entryCount} parts`);
    }
    const flags = view.getUint16(cursor + 8, true);
    const method = view.getUint16(cursor + 10, true);
    const crc = view.getUint32(cursor + 16, true);
    const nameLength = view.getUint16(cursor + 28, true);
    const extraLength = view.getUint16(cursor + 30, true);
    const commentLength = view.getUint16(cursor + 32, true);
    const nameStart = cursor + CENTRAL_HEADER_SIZE;
    const extraStart = nameStart + nameLength;
    const nextCursor = extraStart + extraLength + commentLength;
    if (nextCursor > bytes.byteLength) {
      throw truncated('a central directory entry runs past the directory');
    }

    // Names are decoded as UTF-8 whether or not the writer set the UTF-8 flag (bit 11): OPC part names are ASCII
    // in practice, and every writer that uses non-ASCII names encodes them as UTF-8 regardless of the flag.
    const rawName = nameDecoder.decode(bytes.subarray(nameStart, extraStart));
    if ((flags & FLAG_ENCRYPTED) !== 0) {
      throw new XlsxError('ZIP_UNSUPPORTED', `The part ${rawName} is encrypted; password-protected zip archives are not supported.`, {
        part: rawName,
      });
    }
    if (method !== METHOD_STORE && method !== METHOD_DEFLATE) {
      throw new XlsxError('ZIP_UNSUPPORTED', `The part ${rawName} uses compression method ${method}, which is not supported.`, {
        part: rawName,
        method,
      });
    }

    const zip64 = readZip64Extra(view, extraStart, extraLength, {
      uncompressedSize: view.getUint32(cursor + 24, true),
      compressedSize: view.getUint32(cursor + 20, true),
      localHeaderOffset: view.getUint32(cursor + 42, true),
    });

    const name = normalizeEntryName(rawName);
    cursor = nextCursor;
    // Directory markers carry no data and are not parts; every xlsx writer that emits them expects them ignored.
    if (name.length === 0 || name.endsWith('/')) {
      continue;
    }
    if (entries.has(name)) {
      throw new XlsxError('ZIP_DUPLICATE_ENTRY', `This file contains two parts named ${name}, which is not valid.`, { part: name });
    }
    entries.set(name, {
      name,
      method: method === METHOD_DEFLATE ? 'deflate' : 'store',
      compressedSize: zip64.compressedSize,
      uncompressedSize: zip64.uncompressedSize,
      crc32: crc,
      localHeaderOffset: zip64.localHeaderOffset,
    });
  }
  return entries;
}

interface Zip64Sizes {
  uncompressedSize: number;
  compressedSize: number;
  localHeaderOffset: number;
}

/**
 * The zip64 extended information field (0x0001) supplies whichever of uncompressed size, compressed size and local
 * header offset are saturated in the fixed record, in that order and only for the saturated ones.
 */
function readZip64Extra(view: DataView, extraStart: number, extraLength: number, sizes: Zip64Sizes): Zip64Sizes {
  const needsUncompressed = sizes.uncompressedSize === UINT32_MARKER;
  const needsCompressed = sizes.compressedSize === UINT32_MARKER;
  const needsOffset = sizes.localHeaderOffset === UINT32_MARKER;
  if (!needsUncompressed && !needsCompressed && !needsOffset) {
    return sizes;
  }

  const extraEnd = extraStart + extraLength;
  let cursor = extraStart;
  while (cursor + 4 <= extraEnd) {
    const headerId = view.getUint16(cursor, true);
    const dataSize = view.getUint16(cursor + 2, true);
    const dataStart = cursor + 4;
    if (headerId !== ZIP64_EXTRA_HEADER_ID || dataStart + dataSize > extraEnd) {
      cursor = dataStart + dataSize;
      continue;
    }
    let fieldCursor = dataStart;
    const resolved = { ...sizes };
    if (needsUncompressed && fieldCursor + 8 <= dataStart + dataSize) {
      resolved.uncompressedSize = readUint64(view, fieldCursor);
      fieldCursor += 8;
    }
    if (needsCompressed && fieldCursor + 8 <= dataStart + dataSize) {
      resolved.compressedSize = readUint64(view, fieldCursor);
      fieldCursor += 8;
    }
    if (needsOffset && fieldCursor + 8 <= dataStart + dataSize) {
      resolved.localHeaderOffset = readUint64(view, fieldCursor);
    }
    return resolved;
  }
  throw truncated('a zip64 size field is missing from the central directory');
}
