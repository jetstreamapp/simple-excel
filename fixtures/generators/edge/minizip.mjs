// Minimal zip writer (deflate or stored, optional data descriptors) used to hand-build edge-case and hostile
// fixtures from raw XML parts. Kept dependency-free on purpose: the format primer (docs 02) describes exactly
// these structures. Not a general-purpose zip library - no zip64, no directory entries.
import { deflateRawSync } from 'node:zlib';

const CRC_TABLE = new Uint32Array(256).map((_, n) => {
  let c = n;
  for (let k = 0; k < 8; k++) {
    c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  }
  return c >>> 0;
});

export function crc32(bytes) {
  let crc = 0xffffffff;
  for (let i = 0; i < bytes.length; i++) {
    crc = CRC_TABLE[(crc ^ bytes[i]) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function u16(value) {
  const out = Buffer.alloc(2);
  out.writeUInt16LE(value);
  return out;
}

function u32(value) {
  const out = Buffer.alloc(4);
  out.writeUInt32LE(value >>> 0);
  return out;
}

/**
 * entries: [{ name, data: Buffer|string, method?: 'deflate'|'store', dataDescriptor?: boolean }]
 * options: { truncateAt?: number (cut the archive after N bytes), corruptCrc?: boolean }
 */
export function buildZip(entries, options = {}) {
  const chunks = [];
  const central = [];
  let offset = 0;
  const DOS_TIME = u16(0);
  const DOS_DATE = u16((46 << 9) | (9 << 5) | 11); // 2026-09-11

  for (const entry of entries) {
    const name = Buffer.from(entry.name, 'utf8');
    const raw = Buffer.isBuffer(entry.data) ? entry.data : Buffer.from(entry.data, 'utf8');
    const method = entry.method ?? 'deflate';
    const compressed = method === 'deflate' ? deflateRawSync(raw, { level: 6 }) : raw;
    const crc = options.corruptCrc ? 0xdeadbeef : crc32(raw);
    const flags = (entry.dataDescriptor ? 0x0008 : 0) | 0x0800; // bit 11: UTF-8 names
    const methodCode = method === 'deflate' ? 8 : 0;
    const local = Buffer.concat([
      u32(0x04034b50),
      u16(20),
      u16(flags),
      u16(methodCode),
      DOS_TIME,
      DOS_DATE,
      u32(entry.dataDescriptor ? 0 : crc),
      u32(entry.dataDescriptor ? 0 : compressed.length),
      u32(entry.dataDescriptor ? 0 : raw.length),
      u16(name.length),
      u16(0),
      name,
    ]);
    chunks.push(local, compressed);
    let entryLength = local.length + compressed.length;
    if (entry.dataDescriptor) {
      const descriptor = Buffer.concat([u32(0x08074b50), u32(crc), u32(compressed.length), u32(raw.length)]);
      chunks.push(descriptor);
      entryLength += descriptor.length;
    }
    central.push(
      Buffer.concat([
        u32(0x02014b50),
        u16(20),
        u16(20),
        u16(flags),
        u16(methodCode),
        DOS_TIME,
        DOS_DATE,
        u32(crc),
        u32(compressed.length),
        u32(raw.length),
        u16(name.length),
        u16(0),
        u16(0),
        u16(0),
        u16(0),
        u32(0),
        u32(offset),
        name,
      ]),
    );
    offset += entryLength;
  }
  const centralStart = offset;
  const centralBytes = Buffer.concat(central);
  const eocd = Buffer.concat([
    u32(0x06054b50),
    u16(0),
    u16(0),
    u16(entries.length),
    u16(entries.length),
    u32(centralBytes.length),
    u32(centralStart),
    u16(0),
  ]);
  let archive = Buffer.concat([...chunks, centralBytes, eocd]);
  if (options.truncateAt !== undefined) {
    archive = archive.subarray(0, options.truncateAt);
  }
  return archive;
}
