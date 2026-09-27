import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { deflateRawSync } from 'node:zlib';
import { describe, expect, it } from 'vitest';
import type { RandomAccessSource } from '../../types';
import { crc32 } from '../crc32';
import { sourceFrom } from '../source';
import { ZipReader, type ZipReaderLimits } from '../zip-reader';

const REPO_ROOT = join(import.meta.dirname, '..', '..', '..');
const MIB = 1024 * 1024;
const DEFAULT_LIMITS: ZipReaderLimits = { maxEntries: 10_000, maxInflatedBytes: 1024 * MIB };

const encoder = new TextEncoder();
const decoder = new TextDecoder();

// ---------------------------------------------------------------------------------------------------------------------
// Fixtures
// ---------------------------------------------------------------------------------------------------------------------

interface ManifestFixture {
  id: string;
  path: string;
  generated: boolean;
}

let manifestFixtures: ManifestFixture[] | undefined;

function fixtureBytes(id: string): Uint8Array<ArrayBuffer> {
  manifestFixtures ??= (JSON.parse(readFileSync(join(REPO_ROOT, 'fixtures', 'manifest.json'), 'utf8')) as { fixtures: ManifestFixture[] })
    .fixtures;
  const fixture = manifestFixtures.find(candidate => candidate.id === id);
  if (!fixture) {
    throw new Error(`unknown fixture ${id}`);
  }
  const file = join(REPO_ROOT, fixture.generated ? '.generated' : 'fixtures', fixture.path);
  return new Uint8Array(readFileSync(file));
}

function openFixture(id: string, limits: ZipReaderLimits = DEFAULT_LIMITS): Promise<ZipReader> {
  return ZipReader.open(sourceFrom(fixtureBytes(id)), limits);
}

// ---------------------------------------------------------------------------------------------------------------------
// A hand-built zip writer, so each test can shape the exact container it needs
// ---------------------------------------------------------------------------------------------------------------------

class ByteWriter {
  private readonly parts: Uint8Array[] = [];
  private size = 0;

  get length(): number {
    return this.size;
  }

  push(bytes: Uint8Array): this {
    this.parts.push(bytes);
    this.size += bytes.byteLength;
    return this;
  }

  u16(value: number): this {
    const bytes = new Uint8Array(2);
    new DataView(bytes.buffer).setUint16(0, value, true);
    return this.push(bytes);
  }

  u32(value: number): this {
    const bytes = new Uint8Array(4);
    new DataView(bytes.buffer).setUint32(0, value >>> 0, true);
    return this.push(bytes);
  }

  u64(value: number): this {
    const bytes = new Uint8Array(8);
    new DataView(bytes.buffer).setBigUint64(0, BigInt(value), true);
    return this.push(bytes);
  }

  toBytes(): Uint8Array {
    const joined = new Uint8Array(this.size);
    let offset = 0;
    for (const part of this.parts) {
      joined.set(part, offset);
      offset += part.byteLength;
    }
    return joined;
  }
}

interface TestEntry {
  name: string;
  data: Uint8Array | string;
  method?: 'deflate' | 'store';
  /** Sizes and CRC move to a trailing data descriptor (general purpose bit 3). */
  dataDescriptor?: boolean;
  /** Saturate the 32-bit fields and carry the real values in zip64 extra fields. */
  zip64?: boolean;
  /** Extra general purpose flag bits, e.g. 0x0001 for an encrypted entry. */
  flags?: number;
  /** Raw compression method code, for methods the reader must refuse. */
  methodCode?: number;
  /** Bytes that only the local header knows about, so its extra length differs from the central copy. */
  localExtra?: Uint8Array;
  /** A different (shorter) name in the local header, so its name length differs from the central copy. */
  localName?: string;
  /** Wrong CRC in the central directory. */
  crcOverride?: number;
  /** Wrong compressed size in the central directory. */
  centralCompressedSize?: number;
}

interface BuildOptions {
  comment?: string;
  /** Cut the archive after this many bytes. */
  truncateAt?: number;
}

function zip64Extra(values: readonly number[]): Uint8Array {
  const extra = new ByteWriter();
  extra.u16(0x0001).u16(values.length * 8);
  for (const value of values) {
    extra.u64(value);
  }
  return extra.toBytes();
}

function buildZip(entries: readonly TestEntry[], options: BuildOptions = {}): Uint8Array {
  const archive = new ByteWriter();
  const centralRecords: Uint8Array[] = [];
  let anyZip64 = false;

  for (const entry of entries) {
    const raw = typeof entry.data === 'string' ? encoder.encode(entry.data) : entry.data;
    const method = entry.method ?? 'deflate';
    const compressed = method === 'deflate' ? new Uint8Array(deflateRawSync(raw, { level: 6 })) : raw;
    const crc = entry.crcOverride ?? crc32(raw);
    const nameBytes = encoder.encode(entry.name);
    const localNameBytes = encoder.encode(entry.localName ?? entry.name);
    const useZip64 = entry.zip64 === true;
    anyZip64 ||= useZip64;
    const flags = (entry.dataDescriptor ? 0x0008 : 0) | 0x0800 | (entry.flags ?? 0);
    const methodCode = entry.methodCode ?? (method === 'deflate' ? 8 : 0);
    const localExtra = useZip64 ? zip64Extra([raw.byteLength, compressed.byteLength]) : (entry.localExtra ?? new Uint8Array(0));
    const offset = archive.length;

    archive
      .u32(0x0403_4b50)
      .u16(useZip64 ? 45 : 20)
      .u16(flags)
      .u16(methodCode)
      .u16(0)
      .u16(0x5a21)
      .u32(entry.dataDescriptor ? 0 : crc)
      .u32(entry.dataDescriptor ? 0 : compressed.byteLength)
      .u32(entry.dataDescriptor ? 0 : raw.byteLength)
      .u16(localNameBytes.byteLength)
      .u16(localExtra.byteLength)
      .push(localNameBytes)
      .push(localExtra)
      .push(compressed);
    if (entry.dataDescriptor) {
      archive.u32(0x0807_4b50).u32(crc).u32(compressed.byteLength).u32(raw.byteLength);
    }

    const centralExtra = useZip64 ? zip64Extra([raw.byteLength, compressed.byteLength, offset]) : new Uint8Array(0);
    const central = new ByteWriter();
    central
      .u32(0x0201_4b50)
      .u16(useZip64 ? 45 : 20)
      .u16(useZip64 ? 45 : 20)
      .u16(flags)
      .u16(methodCode)
      .u16(0)
      .u16(0x5a21)
      .u32(crc)
      .u32(useZip64 ? 0xffff_ffff : (entry.centralCompressedSize ?? compressed.byteLength))
      .u32(useZip64 ? 0xffff_ffff : raw.byteLength)
      .u16(nameBytes.byteLength)
      .u16(centralExtra.byteLength)
      .u16(0)
      .u16(0)
      .u16(0)
      .u32(0)
      .u32(useZip64 ? 0xffff_ffff : offset)
      .push(nameBytes)
      .push(centralExtra);
    centralRecords.push(central.toBytes());
  }

  const centralStart = archive.length;
  for (const record of centralRecords) {
    archive.push(record);
  }
  const centralSize = archive.length - centralStart;
  const comment = encoder.encode(options.comment ?? '');

  if (anyZip64) {
    const zip64Start = archive.length;
    archive
      .u32(0x0606_4b50)
      .u64(ZIP64_EOCD_REMAINDER)
      .u16(45)
      .u16(45)
      .u32(0)
      .u32(0)
      .u64(entries.length)
      .u64(entries.length)
      .u64(centralSize)
      .u64(centralStart);
    archive.u32(0x0706_4b50).u32(0).u64(zip64Start).u32(1);
    archive.u32(0x0605_4b50).u16(0).u16(0).u16(0xffff).u16(0xffff).u32(0xffff_ffff).u32(0xffff_ffff);
  } else {
    archive.u32(0x0605_4b50).u16(0).u16(0).u16(entries.length).u16(entries.length).u32(centralSize).u32(centralStart);
  }
  archive.u16(comment.byteLength).push(comment);

  const bytes = archive.toBytes();
  return options.truncateAt === undefined ? bytes : bytes.subarray(0, options.truncateAt);
}

/** The zip64 end-of-central-directory record's own size field: the record minus its signature and this field. */
const ZIP64_EOCD_REMAINDER = 44;

function openBuilt(
  entries: readonly TestEntry[],
  options: BuildOptions = {},
  limits: ZipReaderLimits = DEFAULT_LIMITS,
): Promise<ZipReader> {
  return ZipReader.open(sourceFrom(buildZip(entries, options)), limits);
}

async function collect(chunks: AsyncIterable<Uint8Array>): Promise<{ total: number; largestChunk: number; count: number }> {
  let total = 0;
  let largestChunk = 0;
  let count = 0;
  for await (const chunk of chunks) {
    total += chunk.byteLength;
    largestChunk = Math.max(largestChunk, chunk.byteLength);
    count++;
  }
  return { total, largestChunk, count };
}

function repetitiveXml(byteLength: number): string {
  return `<r>${'a'.repeat(byteLength)}</r>`;
}

// ---------------------------------------------------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------------------------------------------------

describe('ZipReader.open', () => {
  it('parses the central directory of a real workbook', async () => {
    const reader = await openFixture('edge-baseline-minimal');
    expect([...reader.entries.keys()]).toEqual([
      '[Content_Types].xml',
      '_rels/.rels',
      'xl/workbook.xml',
      'xl/_rels/workbook.xml.rels',
      'xl/styles.xml',
      'xl/sharedStrings.xml',
      'xl/worksheets/sheet1.xml',
    ]);
    const sheet = reader.entries.get('xl/worksheets/sheet1.xml');
    expect(sheet).toMatchObject({ method: 'deflate', compressedSize: 252, uncompressedSize: 502 });
    expect(reader.has('xl/workbook.xml')).toBe(true);
    expect(reader.has('xl/nope.xml')).toBe(false);
  });

  it('finds parts wherever the writer put them (EC-ZIP-CONTENT-TYPES-LAST)', async () => {
    const reader = await openFixture('jetstream-multi-object-template-gsheets');
    expect(reader.entries.size).toBe(30);
    expect([...reader.entries.keys()].at(-1)).toBe('[Content_Types].xml');
    expect((await reader.readText('[Content_Types].xml')).startsWith('<?xml')).toBe(true);
  });

  it('skips directory entries', async () => {
    const reader = await openFixture('golden-canonical-exceljs');
    expect([...reader.entries.keys()].some(name => name.endsWith('/'))).toBe(false);
    expect(reader.has('xl/worksheets/sheet1.xml')).toBe(true);
  });

  it('reads from a Blob source', async () => {
    const reader = await ZipReader.open(sourceFrom(new Blob([fixtureBytes('edge-baseline-minimal')])), DEFAULT_LIMITS);
    expect(reader.entries.size).toBe(7);
    expect((await reader.readText('xl/workbook.xml')).includes('<workbook')).toBe(true);
  });

  it('normalizes backslashes and leading slashes in names', async () => {
    const reader = await openBuilt([
      { name: 'xl\\worksheets\\sheet1.xml', data: '<sheet/>' },
      { name: '/[Content_Types].xml', data: '<types/>' },
    ]);
    expect([...reader.entries.keys()]).toEqual(['xl/worksheets/sheet1.xml', '[Content_Types].xml']);
    expect(await reader.readText('xl/worksheets/sheet1.xml')).toBe('<sheet/>');
  });

  it('rejects duplicate names (EC-ZIP-DUPLICATE-ENTRIES)', async () => {
    await expect(openFixture('hostile-duplicate-sheet-entries')).rejects.toMatchObject({
      name: 'XlsxError',
      code: 'ZIP_DUPLICATE_ENTRY',
    });
  });

  it('rejects names that collide only after normalization', async () => {
    await expect(
      openBuilt([
        { name: 'xl/sheet.xml', data: 'a' },
        { name: 'xl\\sheet.xml', data: 'b' },
      ]),
    ).rejects.toMatchObject({ code: 'ZIP_DUPLICATE_ENTRY' });
  });

  it('rejects an archive with no end-of-central-directory record (EC-ZIP-TRUNCATED)', async () => {
    await expect(openFixture('hostile-truncated-central-directory')).rejects.toMatchObject({
      name: 'XlsxError',
      code: 'ZIP_TRUNCATED',
    });
  });

  it('rejects bytes that are not a zip at all', async () => {
    await expect(openFixture('hostile-not-a-zip')).rejects.toMatchObject({ code: 'ZIP_TRUNCATED' });
    await expect(openFixture('hostile-csv-bytes-renamed')).rejects.toMatchObject({ code: 'ZIP_TRUNCATED' });
    await expect(ZipReader.open(sourceFrom(new Uint8Array(0)), DEFAULT_LIMITS)).rejects.toMatchObject({ code: 'ZIP_TRUNCATED' });
  });

  it('opens hostile fixtures whose container is a valid zip', async () => {
    // These are hostile at the XML or content-type layer, which is not this module's business: the container
    // itself is sound and must parse.
    for (const id of [
      'hostile-xxe-doctype-in-sharedstrings',
      'hostile-deeply-nested-rich-text',
      'hostile-ods-renamed',
      'hostile-xlsb-renamed',
    ]) {
      const reader = await openFixture(id);
      expect(reader.entries.size, id).toBeGreaterThan(0);
    }
  });

  it('fails cleanly on containers that only look like a zip', async () => {
    // A BIFF8 workbook and a CFB-encrypted package both carry byte patterns that a naive scan mistakes for a
    // record; the facade sniffs these first, but the reader still has to classify rather than crash.
    for (const id of ['hostile-biff8-xls-renamed', 'hostile-encrypted-password-test']) {
      await expect(openFixture(id), id).rejects.toMatchObject({ name: 'XlsxError', code: 'ZIP_TRUNCATED' });
    }
  });

  it('rejects a central directory that runs past the end of the file', async () => {
    const archive = buildZip([{ name: 'a.xml', data: 'hello' }]);
    const view = new DataView(archive.buffer, archive.byteOffset, archive.byteLength);
    // Point the EOCD at a central directory size the file cannot hold.
    view.setUint32(archive.byteLength - 10, 0x0f_ff_ff_ff, true);
    await expect(ZipReader.open(sourceFrom(archive), DEFAULT_LIMITS)).rejects.toMatchObject({ code: 'ZIP_TRUNCATED' });
  });

  it('rejects a truncated central directory', async () => {
    const archive = buildZip([
      { name: 'a.xml', data: 'hello' },
      { name: 'b.xml', data: 'world' },
    ]);
    // Keep the EOCD but drop the bytes of the second central record it points at.
    const cut = new Uint8Array(archive.byteLength - 30);
    cut.set(archive.subarray(0, archive.byteLength - 52), 0);
    cut.set(archive.subarray(archive.byteLength - 22), cut.byteLength - 22);
    await expect(ZipReader.open(sourceFrom(cut), DEFAULT_LIMITS)).rejects.toMatchObject({ code: 'ZIP_TRUNCATED' });
  });

  it('finds the record behind a zip comment', async () => {
    const reader = await openBuilt([{ name: 'a.xml', data: 'hello' }], { comment: 'PK not really the record' });
    expect(await reader.readText('a.xml')).toBe('hello');
  });

  it('refuses encrypted entries', async () => {
    await expect(openBuilt([{ name: 'a.xml', data: 'secret', flags: 0x0001 }])).rejects.toMatchObject({
      code: 'ZIP_UNSUPPORTED',
    });
  });

  it('refuses compression methods other than stored and deflate', async () => {
    await expect(openBuilt([{ name: 'a.xml', data: 'hello', methodCode: 12 }])).rejects.toMatchObject({
      code: 'ZIP_UNSUPPORTED',
    });
  });

  it('refuses an archive with more entries than the limit allows', async () => {
    const entries = Array.from({ length: 12 }, (_, index) => ({ name: `part${index}.xml`, data: 'x' }));
    await expect(openBuilt(entries, {}, { ...DEFAULT_LIMITS, maxEntries: 10 })).rejects.toMatchObject({
      code: 'LIMIT_EXCEEDED',
    });
  });

  it('reads a zip64 archive', async () => {
    const reader = await openBuilt([
      { name: 'xl/workbook.xml', data: '<workbook/>', zip64: true },
      { name: 'xl/worksheets/sheet1.xml', data: repetitiveXml(50_000), zip64: true },
    ]);
    expect([...reader.entries.keys()]).toEqual(['xl/workbook.xml', 'xl/worksheets/sheet1.xml']);
    expect(reader.entries.get('xl/worksheets/sheet1.xml')?.uncompressedSize).toBe(50_007);
    expect(await reader.readText('xl/workbook.xml')).toBe('<workbook/>');
    expect((await reader.read('xl/worksheets/sheet1.xml')).byteLength).toBe(50_007);
  });

  it('closes the underlying source', async () => {
    let closed = 0;
    const bytes = fixtureBytes('edge-baseline-minimal');
    const source: RandomAccessSource = {
      size: bytes.byteLength,
      read: (offset, length) => Promise.resolve(bytes.subarray(offset, offset + length)),
      close: async () => {
        closed++;
      },
    };
    const reader = await ZipReader.open(source, DEFAULT_LIMITS);
    await reader.close();
    expect(closed).toBe(1);
  });
});

describe('ZipReader.stream', () => {
  it('streams a deflated entry and verifies its CRC', async () => {
    const reader = await openFixture('edge-baseline-minimal');
    const { total } = await collect(reader.stream('xl/worksheets/sheet1.xml'));
    expect(total).toBe(502);
  });

  it('streams stored entries (EC-ZIP-STORED-ENTRIES)', async () => {
    const reader = await openFixture('edge-stored-entries');
    expect([...reader.entries.values()].every(entry => entry.method === 'store')).toBe(true);
    const sheet = await reader.readText('xl/worksheets/sheet1.xml');
    expect(sheet.startsWith('<?xml')).toBe(true);
    expect(sheet).toBe(await (await openFixture('edge-baseline-minimal')).readText('xl/worksheets/sheet1.xml'));
  });

  it('streams entries written with data descriptors (EC-ZIP-DATA-DESCRIPTORS)', async () => {
    const reader = await openFixture('edge-sst-after-sheet-data-descriptors');
    expect([...reader.entries.keys()][0]).toBe('xl/worksheets/sheet1.xml');
    expect((await reader.readText('xl/sharedStrings.xml')).includes('<sst')).toBe(true);
    expect((await reader.read('xl/worksheets/sheet1.xml')).byteLength).toBe(502);
  });

  it('uses the local header name and extra lengths, not the central copy', async () => {
    const reader = await openBuilt([
      {
        name: 'xl/worksheets/sheet1.xml',
        localName: 'x',
        localExtra: new Uint8Array([0x99, 0x99, 0x04, 0x00, 1, 2, 3, 4]),
        data: '<sheet>local header wins</sheet>',
      },
    ]);
    expect(await reader.readText('xl/worksheets/sheet1.xml')).toBe('<sheet>local header wins</sheet>');
  });

  it('throws ZIP_CRC_MISMATCH after the last chunk (EC-ZIP-CRC-MISMATCH)', async () => {
    const reader = await openFixture('hostile-crc-mismatch');
    let delivered = 0;
    let caught: unknown;
    try {
      for await (const chunk of reader.stream('xl/worksheets/sheet1.xml')) {
        delivered += chunk.byteLength;
      }
    } catch (error) {
      caught = error;
    }
    expect(caught).toMatchObject({ name: 'XlsxError', code: 'ZIP_CRC_MISMATCH' });
    expect(delivered).toBe(502);
  });

  it('reports a CRC mismatch from read() as well', async () => {
    const reader = await openFixture('hostile-crc-mismatch');
    await expect(reader.read('[Content_Types].xml')).rejects.toMatchObject({ code: 'ZIP_CRC_MISMATCH' });
  });

  it('rejects an entry whose declared size is over the inflated cap (EC-ZIP-BOMB)', async () => {
    const reader = await openFixture('hostile-zip-bomb-30mb-sheet', { ...DEFAULT_LIMITS, maxInflatedBytes: 8 * MIB });
    await expect(collect(reader.stream('xl/worksheets/sheet1.xml'))).rejects.toMatchObject({
      name: 'XlsxError',
      code: 'ZIP_BOMB',
    });
    // The rest of the workbook still reads: the cap is per entry, checked lazily.
    expect((await reader.readText('xl/workbook.xml')).includes('<workbook')).toBe(true);
  });

  it('streams the same entry in flat memory when the cap allows it', async () => {
    const reader = await openFixture('hostile-zip-bomb-30mb-sheet', { ...DEFAULT_LIMITS, maxInflatedBytes: 64 * MIB });
    const { total, largestChunk, count } = await collect(reader.stream('xl/worksheets/sheet1.xml'));
    expect(total).toBe(28_559_061);
    expect(count).toBeGreaterThan(100);
    expect(largestChunk).toBeLessThanOrEqual(MIB);
  });

  it('rejects an entry that inflates far past the size it declares', async () => {
    const archive = buildZip([{ name: 'xl/worksheets/sheet1.xml', data: repetitiveXml(2 * MIB) }]);
    const view = new DataView(archive.buffer, archive.byteOffset, archive.byteLength);
    const centralStart = view.getUint32(archive.byteLength - 6, true);
    view.setUint32(centralStart + 24, 64, true); // claim 64 uncompressed bytes
    const reader = await ZipReader.open(sourceFrom(archive), DEFAULT_LIMITS);
    await expect(collect(reader.stream('xl/worksheets/sheet1.xml'))).rejects.toMatchObject({ code: 'ZIP_BOMB' });
  });

  it('stops cleanly when the consumer breaks out early', async () => {
    const reader = await openBuilt([{ name: 'xl/worksheets/sheet1.xml', data: repetitiveXml(4 * MIB) }]);
    let chunks = 0;
    for await (const chunk of reader.stream('xl/worksheets/sheet1.xml')) {
      expect(chunk.byteLength).toBeGreaterThan(0);
      chunks++;
      break;
    }
    expect(chunks).toBe(1);
    // The reader is still usable afterwards.
    expect((await reader.read('xl/worksheets/sheet1.xml')).byteLength).toBe(4 * MIB + 7);
  });

  it('throws NOT_XLSX for an entry that does not exist', async () => {
    const reader = await openFixture('edge-baseline-minimal');
    await expect(collect(reader.stream('xl/worksheets/sheet9.xml'))).rejects.toMatchObject({
      name: 'XlsxError',
      code: 'NOT_XLSX',
      message: expect.stringContaining('missing part'),
    });
    await expect(reader.read('nope.xml')).rejects.toMatchObject({ code: 'NOT_XLSX' });
  });

  it('throws ZIP_TRUNCATED when an entry runs past the end of the file', async () => {
    const reader = await openBuilt([{ name: 'a.xml', data: 'hello', centralCompressedSize: 5_000 }]);
    await expect(collect(reader.stream('a.xml'))).rejects.toMatchObject({ code: 'ZIP_TRUNCATED' });
  });

  it('throws ZIP_TRUNCATED when the local header is missing', async () => {
    const archive = buildZip([{ name: 'a.xml', data: 'hello' }]);
    const view = new DataView(archive.buffer, archive.byteOffset, archive.byteLength);
    view.setUint32(0, 0x0000_0000, true);
    const reader = await ZipReader.open(sourceFrom(archive), DEFAULT_LIMITS);
    await expect(collect(reader.stream('a.xml'))).rejects.toMatchObject({ code: 'ZIP_TRUNCATED' });
  });

  it('throws ZIP_TRUNCATED when the compressed data is corrupt', async () => {
    const archive = buildZip([{ name: 'a.xml', data: repetitiveXml(10_000) }]);
    const dataStart = 30 + 'a.xml'.length;
    archive.fill(0x5a, dataStart + 5, dataStart + 25);
    const reader = await ZipReader.open(sourceFrom(archive), DEFAULT_LIMITS);
    await expect(collect(reader.stream('a.xml'))).rejects.toMatchObject({ code: 'ZIP_TRUNCATED' });
  });
});

describe('ZipReader.read', () => {
  it('returns the whole entry as one buffer', async () => {
    const reader = await openBuilt([{ name: 'a.xml', data: repetitiveXml(700_000) }]);
    const bytes = await reader.read('a.xml');
    expect(bytes.byteLength).toBe(700_007);
    expect(decoder.decode(bytes.subarray(0, 3))).toBe('<r>');
  });

  it('returns an empty buffer for an empty entry', async () => {
    const reader = await openBuilt([{ name: 'empty.xml', data: new Uint8Array(0), method: 'store' }]);
    expect((await reader.read('empty.xml')).byteLength).toBe(0);
  });

  it('accepts a name written with backslashes or a leading slash', async () => {
    const reader = await openFixture('edge-baseline-minimal');
    expect((await reader.read('/xl\\workbook.xml')).byteLength).toBe(295);
  });
});

describe('ZipReader.readText', () => {
  it('decodes UTF-8', async () => {
    const reader = await openBuilt([{ name: 'a.xml', data: '<t>Ünïcödé — ✓</t>' }]);
    expect(await reader.readText('a.xml')).toBe('<t>Ünïcödé — ✓</t>');
  });

  it('strips a leading byte order mark', async () => {
    const withBom = new Uint8Array([0xef, 0xbb, 0xbf, ...encoder.encode('<?xml version="1.0"?>')]);
    const reader = await openBuilt([{ name: 'a.xml', data: withBom }]);
    expect(await reader.readText('a.xml')).toBe('<?xml version="1.0"?>');
    expect((await reader.read('a.xml')).byteLength).toBe(withBom.byteLength);
  });
});

describe('ZipReader: part names in another case', () => {
  it('EC-PART-NAME-CASE: has, stream, read and readText fall back to an ASCII case-insensitive match', async () => {
    const reader = await openBuilt([
      { name: 'xl/SharedStrings.xml', data: '<sst/>' },
      { name: 'XL/Worksheets/Sheet1.XML', data: '<worksheet/>' },
    ]);
    expect(reader.has('xl/sharedStrings.xml')).toBe(true);
    expect(await reader.readText('xl/sharedstrings.xml')).toBe('<sst/>');
    expect(await reader.readText('/xl\\worksheets\\sheet1.xml')).toBe('<worksheet/>');
    expect((await reader.read('xl/worksheets/sheet1.xml')).byteLength).toBe('<worksheet/>'.length);
    expect((await collect(reader.stream('XL/SHAREDSTRINGS.XML'))).total).toBe('<sst/>'.length);
    // `entries` keeps the names exactly as the archive spells them.
    expect([...reader.entries.keys()]).toEqual(['xl/SharedStrings.xml', 'XL/Worksheets/Sheet1.XML']);
    expect(reader.has('xl/styles.xml')).toBe(false);
  });

  it('EC-PART-NAME-CASE: only ASCII letters fold', async () => {
    const reader = await openBuilt([{ name: 'xl/Ünï.xml', data: 'a' }]);
    expect(reader.has('XL/Ünï.xml')).toBe(true);
    expect(reader.has('xl/ünï.xml')).toBe(false);
  });

  it('EC-PART-NAME-CASE: names that differ only in case are not duplicates; the exact name wins, then the first', async () => {
    const reader = await openBuilt([
      { name: 'xl/Sheet.xml', data: 'first' },
      { name: 'xl/sheet.xml', data: 'second' },
    ]);
    expect(await reader.readText('xl/Sheet.xml')).toBe('first');
    expect(await reader.readText('xl/sheet.xml')).toBe('second');
    expect(await reader.readText('XL/SHEET.XML')).toBe('first');
  });
});

describe('ZipReader.close', () => {
  it('fails any read after close with ABORTED, and closes the source once', async () => {
    let closes = 0;
    const bytes = buildZip([{ name: 'a.xml', data: '<a/>' }]);
    const source: RandomAccessSource = {
      size: bytes.byteLength,
      read: (offset, length) =>
        closes > 0
          ? Promise.reject(new Error('EBADF: bad file descriptor, read'))
          : Promise.resolve(bytes.subarray(offset, offset + length)),
      close: async () => {
        closes++;
      },
    };
    const reader = await ZipReader.open(source, DEFAULT_LIMITS);
    await reader.close();
    await reader.close();
    expect(closes).toBe(1);
    await expect(reader.readText('a.xml')).rejects.toMatchObject({ name: 'XlsxError', code: 'ABORTED' });
    await expect(collect(reader.stream('a.xml'))).rejects.toMatchObject({ code: 'ABORTED' });
    // The directory is still there to look at; only reads fail.
    expect(reader.has('a.xml')).toBe(true);
  });
});
