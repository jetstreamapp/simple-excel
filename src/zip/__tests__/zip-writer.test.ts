import { execFileSync } from 'node:child_process';
import { createHash, randomBytes } from 'node:crypto';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { inflateRawSync } from 'node:zlib';
import { describe, expect, it } from 'vitest';
import { createDeflater } from '../../compress/deflater';
import { isXlsxError } from '../../errors';
import type { ByteSink } from '../../types';
import { crc32 } from '../crc32';
import { assertFits, ZipWriter, type ZipWriterOptions } from '../zip-writer';

// ---------------------------------------------------------------------------------------------------------------------
// A sink and a zip parser, both written by hand here so the tests never validate the writer against itself
// ---------------------------------------------------------------------------------------------------------------------

class MemorySink implements ByteSink {
  readonly chunks: Uint8Array[] = [];
  closed = false;
  aborted = false;
  abortReason: unknown;

  async write(chunk: Uint8Array): Promise<void> {
    this.chunks.push(chunk.slice());
  }

  async close(): Promise<void> {
    this.closed = true;
  }

  async abort(reason?: unknown): Promise<void> {
    this.aborted = true;
    this.abortReason = reason;
  }

  bytes(): Uint8Array {
    const total = this.chunks.reduce((sum, chunk) => sum + chunk.length, 0);
    const joined = new Uint8Array(total);
    let position = 0;
    for (const chunk of this.chunks) {
      joined.set(chunk, position);
      position += chunk.length;
    }
    return joined;
  }
}

const LOCAL_HEADER_SIGNATURE = 0x0403_4b50;
const DATA_DESCRIPTOR_SIGNATURE = 0x0807_4b50;
const CENTRAL_HEADER_SIGNATURE = 0x0201_4b50;
const ZIP64_END_SIGNATURE = 0x0606_4b50;
const ZIP64_LOCATOR_SIGNATURE = 0x0706_4b50;
const END_OF_CENTRAL_DIRECTORY_SIGNATURE = 0x0605_4b50;
const MAX_UINT32 = 0xffff_ffff;

interface Zip64ExtraField {
  readonly uncompressedSize?: number;
  readonly compressedSize?: number;
  readonly offset?: number;
}

interface LocalHeader {
  readonly versionNeeded: number;
  readonly flags: number;
  readonly method: number;
  readonly dosTime: number;
  readonly dosDate: number;
  readonly crc: number;
  readonly compressedSizeField: number;
  readonly uncompressedSizeField: number;
  readonly name: string;
  readonly extraFieldIds: readonly number[];
  readonly zip64Extra: Zip64ExtraField | undefined;
  readonly dataOffset: number;
}

interface DataDescriptor {
  readonly crc: number;
  readonly compressedSize: number;
  readonly uncompressedSize: number;
  readonly byteLength: number;
}

interface ParsedEntry {
  readonly name: string;
  readonly versionMadeBy: number;
  readonly versionNeeded: number;
  readonly flags: number;
  readonly method: number;
  readonly dosTime: number;
  readonly dosDate: number;
  readonly crc: number;
  readonly compressedSizeField: number;
  readonly uncompressedSizeField: number;
  readonly compressedSize: number;
  readonly uncompressedSize: number;
  readonly offsetField: number;
  readonly offset: number;
  readonly zip64Extra: Zip64ExtraField | undefined;
  readonly local: LocalHeader;
  readonly descriptor: DataDescriptor;
  readonly data: Uint8Array;
}

interface ParsedZip {
  readonly entries: readonly ParsedEntry[];
  readonly eocd: {
    readonly entryCount: number;
    readonly centralDirectorySize: number;
    readonly centralDirectoryOffset: number;
    readonly offset: number;
  };
  readonly zip64End:
    | {
        readonly recordSize: number;
        readonly versionMadeBy: number;
        readonly versionNeeded: number;
        readonly entryCount: number;
        readonly centralDirectorySize: number;
        readonly centralDirectoryOffset: number;
        readonly offset: number;
      }
    | undefined;
  readonly zip64Locator: { readonly zip64EndOffset: number; readonly totalDisks: number } | undefined;
}

function parseZip64Extra(extra: Uint8Array, fieldCount: number): Zip64ExtraField | undefined {
  const view = new DataView(extra.buffer, extra.byteOffset, extra.byteLength);
  let position = 0;
  while (position + 4 <= extra.length) {
    const id = view.getUint16(position, true);
    const size = view.getUint16(position + 2, true);
    if (id === 0x0001) {
      const values: number[] = [];
      for (let index = 0; index < fieldCount && index * 8 + 8 <= size; index++) {
        values.push(Number(view.getBigUint64(position + 4 + index * 8, true)));
      }
      return { uncompressedSize: values[0], compressedSize: values[1], offset: values[2] };
    }
    position += 4 + size;
  }
  return undefined;
}

function extraFieldIds(extra: Uint8Array): number[] {
  const view = new DataView(extra.buffer, extra.byteOffset, extra.byteLength);
  const ids: number[] = [];
  let position = 0;
  while (position + 4 <= extra.length) {
    ids.push(view.getUint16(position, true));
    position += 4 + view.getUint16(position + 2, true);
  }
  return ids;
}

function parseLocalHeader(bytes: Uint8Array, offset: number): LocalHeader {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  expect(view.getUint32(offset, true), `local header signature at ${offset}`).toBe(LOCAL_HEADER_SIGNATURE);
  const nameLength = view.getUint16(offset + 26, true);
  const extraLength = view.getUint16(offset + 28, true);
  const extra = bytes.subarray(offset + 30 + nameLength, offset + 30 + nameLength + extraLength);
  return {
    versionNeeded: view.getUint16(offset + 4, true),
    flags: view.getUint16(offset + 6, true),
    method: view.getUint16(offset + 8, true),
    dosTime: view.getUint16(offset + 10, true),
    dosDate: view.getUint16(offset + 12, true),
    crc: view.getUint32(offset + 14, true),
    compressedSizeField: view.getUint32(offset + 18, true),
    uncompressedSizeField: view.getUint32(offset + 22, true),
    name: new TextDecoder().decode(bytes.subarray(offset + 30, offset + 30 + nameLength)),
    extraFieldIds: extraFieldIds(extra),
    zip64Extra: parseZip64Extra(extra, 2),
    dataOffset: offset + 30 + nameLength + extraLength,
  };
}

function parseDataDescriptor(bytes: Uint8Array, offset: number, zip64: boolean): DataDescriptor {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  expect(view.getUint32(offset, true), `data descriptor signature at ${offset}`).toBe(DATA_DESCRIPTOR_SIGNATURE);
  if (zip64) {
    return {
      crc: view.getUint32(offset + 4, true),
      compressedSize: Number(view.getBigUint64(offset + 8, true)),
      uncompressedSize: Number(view.getBigUint64(offset + 16, true)),
      byteLength: 24,
    };
  }
  return {
    crc: view.getUint32(offset + 4, true),
    compressedSize: view.getUint32(offset + 8, true),
    uncompressedSize: view.getUint32(offset + 12, true),
    byteLength: 16,
  };
}

/** Walks the end records, then the central directory, then each local header the directory points at. */
function parseZip(bytes: Uint8Array): ParsedZip {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);

  let eocdOffset = -1;
  for (let candidate = bytes.length - 22; candidate >= 0; candidate--) {
    if (view.getUint32(candidate, true) === END_OF_CENTRAL_DIRECTORY_SIGNATURE) {
      eocdOffset = candidate;
      break;
    }
  }
  expect(eocdOffset, 'end of central directory record').toBeGreaterThanOrEqual(0);

  const eocd = {
    entryCount: view.getUint16(eocdOffset + 10, true),
    centralDirectorySize: view.getUint32(eocdOffset + 12, true),
    centralDirectoryOffset: view.getUint32(eocdOffset + 16, true),
    offset: eocdOffset,
  };
  expect(view.getUint16(eocdOffset + 20, true), 'zip comment length').toBe(0);
  expect(eocdOffset + 22, 'bytes after the end record').toBe(bytes.length);

  let zip64Locator: ParsedZip['zip64Locator'];
  let zip64End: ParsedZip['zip64End'];
  if (eocdOffset >= 20 && view.getUint32(eocdOffset - 20, true) === ZIP64_LOCATOR_SIGNATURE) {
    zip64Locator = {
      zip64EndOffset: Number(view.getBigUint64(eocdOffset - 20 + 8, true)),
      totalDisks: view.getUint32(eocdOffset - 20 + 16, true),
    };
    const at = zip64Locator.zip64EndOffset;
    expect(view.getUint32(at, true), 'zip64 end of central directory signature').toBe(ZIP64_END_SIGNATURE);
    zip64End = {
      recordSize: Number(view.getBigUint64(at + 4, true)),
      versionMadeBy: view.getUint16(at + 12, true),
      versionNeeded: view.getUint16(at + 14, true),
      entryCount: Number(view.getBigUint64(at + 32, true)),
      centralDirectorySize: Number(view.getBigUint64(at + 40, true)),
      centralDirectoryOffset: Number(view.getBigUint64(at + 48, true)),
      offset: at,
    };
  }

  const entryCount = zip64End?.entryCount ?? eocd.entryCount;
  const entries: ParsedEntry[] = [];
  let position = zip64End?.centralDirectoryOffset ?? eocd.centralDirectoryOffset;
  for (let index = 0; index < entryCount; index++) {
    expect(view.getUint32(position, true), `central header signature at ${position}`).toBe(CENTRAL_HEADER_SIGNATURE);
    const nameLength = view.getUint16(position + 28, true);
    const extraLength = view.getUint16(position + 30, true);
    const commentLength = view.getUint16(position + 32, true);
    const extra = bytes.subarray(position + 46 + nameLength, position + 46 + nameLength + extraLength);
    const zip64Extra = parseZip64Extra(extra, 3);
    const compressedSizeField = view.getUint32(position + 20, true);
    const uncompressedSizeField = view.getUint32(position + 24, true);
    const offsetField = view.getUint32(position + 42, true);
    const compressedSize = zip64Extra?.compressedSize ?? compressedSizeField;
    const uncompressedSize = zip64Extra?.uncompressedSize ?? uncompressedSizeField;
    const offset = zip64Extra?.offset ?? offsetField;

    const local = parseLocalHeader(bytes, offset);
    const data = bytes.subarray(local.dataOffset, local.dataOffset + compressedSize);
    const descriptor = parseDataDescriptor(bytes, local.dataOffset + compressedSize, local.zip64Extra !== undefined);

    entries.push({
      name: new TextDecoder().decode(bytes.subarray(position + 46, position + 46 + nameLength)),
      versionMadeBy: view.getUint16(position + 4, true),
      versionNeeded: view.getUint16(position + 6, true),
      flags: view.getUint16(position + 8, true),
      method: view.getUint16(position + 10, true),
      dosTime: view.getUint16(position + 12, true),
      dosDate: view.getUint16(position + 14, true),
      crc: view.getUint32(position + 16, true),
      compressedSizeField,
      uncompressedSizeField,
      compressedSize,
      uncompressedSize,
      offsetField,
      offset,
      zip64Extra,
      local,
      descriptor,
      data,
    });
    position += 46 + nameLength + extraLength + commentLength;
  }

  return { entries, eocd, zip64End, zip64Locator };
}

/** Every invariant that must hold for any archive we produce, whatever the options were. */
function expectConsistentArchive(bytes: Uint8Array, parsed: ParsedZip): void {
  let expectedOffset = 0;
  for (const entry of parsed.entries) {
    expect(entry.offset, `local header offset of ${entry.name}`).toBe(expectedOffset);
    expect(entry.local.name).toBe(entry.name);
    expect(entry.local.method).toBe(entry.method);
    expect(entry.local.flags).toBe(entry.flags);
    expect(entry.local.crc, 'crc is deferred to the data descriptor').toBe(0);
    expect(entry.descriptor.crc).toBe(entry.crc);
    expect(entry.descriptor.compressedSize).toBe(entry.compressedSize);
    expect(entry.descriptor.uncompressedSize).toBe(entry.uncompressedSize);
    expect(entry.data.length).toBe(entry.compressedSize);

    const content = entry.method === 0 ? entry.data : new Uint8Array(inflateRawSync(entry.data));
    expect(content.length, `uncompressed size of ${entry.name}`).toBe(entry.uncompressedSize);
    expect(crc32(content), `crc of ${entry.name}`).toBe(entry.crc);

    expectedOffset = entry.local.dataOffset + entry.compressedSize + entry.descriptor.byteLength;
  }
  const centralDirectoryOffset = parsed.zip64End?.centralDirectoryOffset ?? parsed.eocd.centralDirectoryOffset;
  expect(centralDirectoryOffset, 'the central directory follows the last entry').toBe(expectedOffset);
  expect(bytes.length).toBeGreaterThan(centralDirectoryOffset);
}

function contentOf(entry: ParsedEntry): Uint8Array {
  return entry.method === 0 ? entry.data : new Uint8Array(inflateRawSync(entry.data));
}

function text(entry: ParsedEntry): string {
  return new TextDecoder().decode(contentOf(entry));
}

function digest(bytes: Uint8Array): string {
  return createHash('sha256').update(bytes).digest('hex');
}

function encode(value: string): Uint8Array {
  return new TextEncoder().encode(value);
}

function random(byteLength: number): Uint8Array {
  const buffer = randomBytes(byteLength);
  return new Uint8Array(buffer.buffer, buffer.byteOffset, buffer.byteLength);
}

function createWriter(overrides: Partial<ZipWriterOptions> = {}): { writer: ZipWriter; sink: MemorySink } {
  const sink = new MemorySink();
  const writer = new ZipWriter(sink, {
    zip64: false,
    deterministic: true,
    deflater: createDeflater,
    compression: 'deflate',
    ...overrides,
  });
  return { writer, sink };
}

/** Deterministic mode's fixed stamp: 2026-01-01 00:00:00. */
const DETERMINISTIC_DOS_DATE = ((2026 - 1980) << 9) | (1 << 5) | 1;

/** The Info-ZIP cross-check only runs where the binary exists. */
const hasUnzip = ((): boolean => {
  try {
    execFileSync('unzip', ['-v'], { stdio: 'ignore' });
    return true;
  } catch {
    return false;
  }
})();

// ---------------------------------------------------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------------------------------------------------

describe('ZipWriter', () => {
  it('writes entries that parse back and inflate to the original bytes', async () => {
    const { writer, sink } = createWriter();
    const contentTypes = encode('<?xml version="1.0"?><Types/>');
    const sheet = encode(`<worksheet>${'<row><c><v>1</v></c></row>'.repeat(500)}</worksheet>`);

    const first = await writer.writeEntry('[Content_Types].xml', contentTypes);
    const entry = await writer.beginEntry('xl/worksheets/sheet1.xml');
    await entry.write(sheet.subarray(0, 100));
    await entry.write(sheet.subarray(100));
    const second = await entry.close();
    const result = await writer.close();

    expect(sink.closed).toBe(true);
    expect(result.bytes).toBe(sink.bytes().length);
    expect(result.bytes).toBe(writer.bytesWritten);
    expect(result.entries).toEqual([first, second]);
    expect(first.offset).toBe(0);
    expect(second.uncompressedSize).toBe(sheet.length);
    expect(second.crc32).toBe(crc32(sheet));

    const bytes = sink.bytes();
    const parsed = parseZip(bytes);
    expectConsistentArchive(bytes, parsed);
    expect(parsed.entries.map(({ name }) => name)).toEqual(['[Content_Types].xml', 'xl/worksheets/sheet1.xml']);
    expect(parsed.eocd.entryCount).toBe(2);
    expect(parsed.zip64End, 'a small archive needs no zip64 records').toBeUndefined();
    expect(text(parsed.entries[0]!)).toBe('<?xml version="1.0"?><Types/>');
    expect(digest(contentOf(parsed.entries[1]!))).toBe(digest(sheet));
    for (const parsedEntry of parsed.entries) {
      expect(parsedEntry.flags, 'bit 3 (data descriptor) and bit 11 (utf-8)').toBe(0x0808);
      expect(parsedEntry.method).toBe(8);
      expect(parsedEntry.versionNeeded).toBe(20);
      expect(parsedEntry.versionMadeBy).toBe(20);
      expect(parsedEntry.local.versionNeeded).toBe(20);
      expect(parsedEntry.local.extraFieldIds).toEqual([]);
      expect(parsedEntry.local.compressedSizeField).toBe(0);
      expect(parsedEntry.local.uncompressedSizeField).toBe(0);
    }
  });

  it('writes non-ascii names as UTF-8', async () => {
    const { writer, sink } = createWriter();
    const name = 'Ünïcödé/sheet1.xml';
    const content = encode('héllo wörld');
    await writer.writeEntry(name, content);
    await writer.close();

    const bytes = sink.bytes();
    const parsed = parseZip(bytes);
    expectConsistentArchive(bytes, parsed);
    const entry = parsed.entries[0]!;
    expect(entry.name).toBe(name);
    expect(entry.local.name).toBe(name);
    expect(text(entry)).toBe('héllo wörld');
    // The name is longer in bytes than in code units, so the raw bytes really are UTF-8, not latin-1.
    const nameBytes = new TextEncoder().encode(name);
    expect(nameBytes.length).toBeGreaterThan(name.length);
    expect(bytes.subarray(30, 30 + nameBytes.length)).toEqual(nameBytes);
  });

  it('stores entries verbatim with method 0', async () => {
    const { writer, sink } = createWriter();
    const payload = random(4_096);
    const summary = await writer.writeEntry('media/image1.png', payload, { method: 'store' });
    await writer.writeEntry('xl/workbook.xml', encode('<workbook/>'));
    await writer.close();

    expect(summary.compressedSize).toBe(payload.length);
    expect(summary.uncompressedSize).toBe(payload.length);

    const bytes = sink.bytes();
    const parsed = parseZip(bytes);
    expectConsistentArchive(bytes, parsed);
    expect(parsed.entries[0]!.method).toBe(0);
    expect(digest(parsed.entries[0]!.data)).toBe(digest(payload));
    expect(parsed.entries[1]!.method, 'the writer-level method still applies to the other entries').toBe(8);
  });

  it('uses the writer-level store compression for every entry', async () => {
    const { writer, sink } = createWriter({ compression: 'store' });
    await writer.writeEntry('a.txt', encode('aaaa'));
    const entry = await writer.beginEntry('b.txt');
    await entry.write(encode('bbbb'));
    await entry.close();
    await writer.close();

    const bytes = sink.bytes();
    const parsed = parseZip(bytes);
    expectConsistentArchive(bytes, parsed);
    expect(parsed.entries.map(({ method }) => method)).toEqual([0, 0]);
    expect(text(parsed.entries[1]!)).toBe('bbbb');
  });

  it('streams a 5 MiB entry written in 64 KiB chunks', async () => {
    const { writer, sink } = createWriter();
    const payload = new Uint8Array(5 * 1024 * 1024);
    for (let index = 0; index < payload.length; index++) {
      payload[index] = (index * 31 + (index >> 8)) & 0xff;
    }

    const entry = await writer.beginEntry('xl/worksheets/sheet1.xml');
    for (let offset = 0; offset < payload.length; offset += 64 * 1024) {
      await entry.write(payload.subarray(offset, offset + 64 * 1024));
    }
    const summary = await entry.close();
    await writer.close();

    expect(summary.uncompressedSize).toBe(payload.length);
    expect(summary.compressedSize).toBeLessThan(payload.length);

    const bytes = sink.bytes();
    const parsed = parseZip(bytes);
    expectConsistentArchive(bytes, parsed);
    expect(digest(contentOf(parsed.entries[0]!))).toBe(digest(payload));
  }, 30_000);

  it('declares zip64 up front in the local header, the central entry and the end records', async () => {
    const { writer, sink } = createWriter({ zip64: true });
    const sheet = encode('<worksheet><sheetData/></worksheet>');
    const summary = await writer.writeEntry('[Content_Types].xml', encode('<Types/>'));
    const entry = await writer.beginEntry('xl/worksheets/sheet1.xml');
    await entry.write(sheet);
    await entry.close();
    await writer.close();

    const bytes = sink.bytes();
    const parsed = parseZip(bytes);
    expectConsistentArchive(bytes, parsed);

    const big = parsed.entries[1]!;
    expect(big.local.versionNeeded, 'Excel needs version 45 in the local header, not just the directory').toBe(45);
    expect(big.local.compressedSizeField).toBe(MAX_UINT32);
    expect(big.local.uncompressedSizeField).toBe(MAX_UINT32);
    expect(big.local.extraFieldIds).toEqual([0x0001]);
    expect(big.local.zip64Extra, 'the local sizes are placeholders until the descriptor').toEqual({
      uncompressedSize: 0,
      compressedSize: 0,
      offset: undefined,
    });
    expect(big.versionNeeded).toBe(45);
    expect(big.versionMadeBy).toBe(45);
    expect(big.compressedSizeField).toBe(MAX_UINT32);
    expect(big.uncompressedSizeField).toBe(MAX_UINT32);
    expect(big.offsetField).toBe(MAX_UINT32);
    expect(big.zip64Extra).toEqual({
      uncompressedSize: sheet.length,
      compressedSize: big.compressedSize,
      offset: big.offset,
    });
    // 8-byte sizes in the descriptor, which is what makes the entry 24 bytes rather than 16.
    expect(big.descriptor.byteLength).toBe(24);
    expect(text(big)).toBe('<worksheet><sheetData/></worksheet>');

    // Small static parts stay 32-bit even in a zip64 archive.
    const small = parsed.entries[0]!;
    expect(small.name).toBe('[Content_Types].xml');
    expect(small.local.versionNeeded).toBe(20);
    expect(small.local.extraFieldIds).toEqual([]);
    expect(small.versionNeeded).toBe(20);
    expect(small.zip64Extra).toBeUndefined();
    expect(small.compressedSizeField).toBe(summary.compressedSize);
    expect(small.offsetField).toBe(0);
    expect(small.descriptor.byteLength).toBe(16);

    expect(parsed.zip64Locator).toEqual({ zip64EndOffset: parsed.zip64End!.offset, totalDisks: 1 });
    expect(parsed.zip64End).toEqual({
      recordSize: 44,
      versionMadeBy: 45,
      versionNeeded: 45,
      entryCount: 2,
      centralDirectorySize: parsed.eocd.centralDirectorySize,
      centralDirectoryOffset: parsed.eocd.centralDirectoryOffset,
      offset: parsed.eocd.offset - 20 - 56,
    });
    // Values that still fit stay readable in the 32-bit end record; the sentinels only appear when they must.
    expect(parsed.eocd.entryCount).toBe(2);
  });

  it('writes the deterministic timestamp into every entry', async () => {
    const { writer, sink } = createWriter();
    await writer.writeEntry('a.txt', encode('a'));
    await writer.writeEntry('b.txt', encode('b'));
    await writer.close();

    const parsed = parseZip(sink.bytes());
    for (const entry of parsed.entries) {
      expect(entry.dosDate).toBe(DETERMINISTIC_DOS_DATE);
      expect(entry.dosTime).toBe(0);
      expect(entry.local.dosDate).toBe(DETERMINISTIC_DOS_DATE);
      expect(entry.local.dosTime).toBe(0);
      expect(1980 + (entry.dosDate >> 9), 'year').toBe(2026);
      expect((entry.dosDate >> 5) & 0x0f, 'month').toBe(1);
      expect(entry.dosDate & 0x1f, 'day').toBe(1);
    }
  });

  it('writes the current local time when not deterministic', async () => {
    const { writer, sink } = createWriter({ deterministic: false });
    await writer.writeEntry('a.txt', encode('a'));
    await writer.close();

    const entry = parseZip(sink.bytes()).entries[0]!;
    const now = new Date();
    expect(1980 + (entry.dosDate >> 9)).toBe(now.getFullYear());
    expect((entry.dosDate >> 5) & 0x0f).toBe(now.getMonth() + 1);
    expect(entry.dosDate & 0x1f).toBe(now.getDate());
    expect(entry.dosTime >> 11).toBe(now.getHours());
  });

  it('reports the byte count as entries are written', async () => {
    const { writer, sink } = createWriter();
    expect(writer.bytesWritten).toBe(0);
    const entry = await writer.beginEntry('a.txt');
    expect(writer.bytesWritten, 'the local header goes out before any data').toBe(30 + 'a.txt'.length);
    await entry.write(encode('hello'));
    await entry.close();
    const afterEntry = writer.bytesWritten;
    expect(afterEntry).toBeGreaterThan(30 + 'a.txt'.length);
    const result = await writer.close();
    expect(result.bytes).toBeGreaterThan(afterEntry);
    expect(result.bytes).toBe(sink.bytes().length);
  });

  describe('writer state', () => {
    const isWriterState = (thrown: unknown): boolean => isXlsxError(thrown) && thrown.code === 'WRITER_STATE';

    it('allows only one open entry', async () => {
      const { writer } = createWriter();
      await writer.beginEntry('a.txt');
      await expect(writer.beginEntry('b.txt')).rejects.toSatisfy(isWriterState);
      await expect(writer.writeEntry('c.txt', encode('c'))).rejects.toSatisfy(isWriterState);
    });

    it('rejects a write or a second close after the entry is closed', async () => {
      const { writer } = createWriter();
      const entry = await writer.beginEntry('a.txt');
      await entry.write(encode('a'));
      await entry.close();
      await expect(entry.write(encode('more'))).rejects.toSatisfy(isWriterState);
      await expect(entry.close()).rejects.toSatisfy(isWriterState);
    });

    it('rejects closing the writer while an entry is open', async () => {
      const { writer } = createWriter();
      await writer.beginEntry('a.txt');
      await expect(writer.close()).rejects.toSatisfy(isWriterState);
    });

    it('rejects everything after the writer is closed', async () => {
      const { writer } = createWriter();
      await writer.writeEntry('a.txt', encode('a'));
      await writer.close();
      await expect(writer.close()).rejects.toSatisfy(isWriterState);
      await expect(writer.beginEntry('b.txt')).rejects.toSatisfy(isWriterState);
      await expect(writer.writeEntry('b.txt', encode('b'))).rejects.toSatisfy(isWriterState);
    });
  });

  describe('abort', () => {
    it('aborts the sink mid-entry and rejects later calls with ABORTED', async () => {
      const { writer, sink } = createWriter();
      const entry = await writer.beginEntry('xl/worksheets/sheet1.xml');
      await entry.write(random(128 * 1024));

      const reason = new Error('the user closed the tab');
      await writer.abort(reason);

      expect(sink.aborted).toBe(true);
      expect(sink.abortReason).toBe(reason);
      expect(sink.closed).toBe(false);

      const isAborted = (thrown: unknown): boolean => isXlsxError(thrown) && thrown.code === 'ABORTED';
      await expect(entry.write(encode('more'))).rejects.toSatisfy(isAborted);
      await expect(entry.close()).rejects.toSatisfy(isAborted);
      await expect(writer.beginEntry('b.txt')).rejects.toSatisfy(isAborted);
      await expect(writer.close()).rejects.toSatisfy(isAborted);
      await writer.abort();
      expect(sink.abortReason, 'aborting twice keeps the first reason').toBe(reason);
    });

    it('aborts the sink when the sink itself fails', async () => {
      const sink = new MemorySink();
      const failure = new Error('disk full');
      sink.write = async (): Promise<void> => {
        throw failure;
      };
      const writer = new ZipWriter(sink, { zip64: false, deterministic: true, deflater: createDeflater, compression: 'deflate' });

      await expect(writer.beginEntry('a.txt')).rejects.toBe(failure);
      expect(sink.aborted).toBe(true);
      await expect(writer.close()).rejects.toBe(failure);
    });

    it('aborts when the sink fails on compressed data rather than on a header', async () => {
      const sink = new MemorySink();
      const failure = new Error('disk full');
      const originalWrite = sink.write.bind(sink);
      let writes = 0;
      // The header goes out, then the deflater's first chunk fails inside the drain loop.
      sink.write = async (chunk: Uint8Array): Promise<void> => {
        writes++;
        if (writes > 1) {
          throw failure;
        }
        await originalWrite(chunk);
      };
      const writer = new ZipWriter(sink, { zip64: false, deterministic: true, deflater: createDeflater, compression: 'deflate' });

      const entry = await writer.beginEntry('xl/worksheets/sheet1.xml');
      await entry.write(random(512 * 1024)).catch(() => {});
      await expect(entry.close()).rejects.toBe(failure);
      expect(sink.aborted).toBe(true);
      expect(sink.abortReason).toBe(failure);
    }, 10_000);
  });

  describe('32-bit limits', () => {
    it('assertFits accepts anything a 32-bit field can hold', () => {
      expect(() => assertFits(0, 'the size of "a.txt"')).not.toThrow();
      expect(() => assertFits(MAX_UINT32, 'the size of "a.txt"')).not.toThrow();
    });

    it('assertFits throws ENTRY_TOO_LARGE one byte past the limit', () => {
      let thrown: unknown;
      try {
        assertFits(MAX_UINT32 + 1, 'the size of "xl/worksheets/sheet1.xml"');
      } catch (reason) {
        thrown = reason;
      }
      expect(isXlsxError(thrown)).toBe(true);
      expect(isXlsxError(thrown) && thrown.code).toBe('ENTRY_TOO_LARGE');
      expect(isXlsxError(thrown) && thrown.detail).toEqual({
        value: MAX_UINT32 + 1,
        what: 'the size of "xl/worksheets/sheet1.xml"',
      });
      expect(String(thrown)).toContain('xl/worksheets/sheet1.xml');
    });
  });

  it.skipIf(!hasUnzip)('produces archives that Info-ZIP accepts', async () => {
    const directory = mkdtempSync(join(tmpdir(), 'simple-excel-zip-'));
    try {
      for (const zip64 of [false, true]) {
        const { writer, sink } = createWriter({ zip64 });
        await writer.writeEntry('[Content_Types].xml', encode('<?xml version="1.0"?><Types/>'));
        await writer.writeEntry('media/image1.png', random(8_192), { method: 'store' });
        const entry = await writer.beginEntry('Ünïcödé/sheet1.xml');
        for (let index = 0; index < 32; index++) {
          await entry.write(encode(`<row r="${index}"><c><v>${index}</v></c></row>`));
        }
        await entry.close();
        await writer.close();

        const file = join(directory, `zip64-${zip64}.zip`);
        writeFileSync(file, sink.bytes());
        const output = execFileSync('unzip', ['-t', file], { encoding: 'utf8' });
        expect(output, `unzip -t with zip64: ${zip64}`).toContain('No errors detected');
      }
    } finally {
      rmSync(directory, { recursive: true, force: true });
    }
  });
});
