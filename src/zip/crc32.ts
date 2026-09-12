/** Seed for a fresh CRC-32 computation; pass a previous result to continue across chunks. */
export const CRC32_INITIAL: number = 0;

// Table-driven IEEE CRC-32 (polynomial 0xEDB88320, the zip flavour). The table is built once on first use.
let table: Int32Array | undefined;

function buildTable(): Int32Array {
  const built = new Int32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) {
      c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    }
    built[n] = c;
  }
  return built;
}

/**
 * CRC-32 of `bytes`, continuing from `seed` (the value returned for the previous chunk). The result is an unsigned
 * 32-bit integer; `crc32(b, crc32(a)) === crc32(concat(a, b))`.
 */
export function crc32(bytes: Uint8Array, seed: number = CRC32_INITIAL): number {
  table ??= buildTable();
  let crc = ~seed;
  const length = bytes.length;
  for (let i = 0; i < length; i++) {
    crc = (table[(crc ^ bytes[i]!) & 0xff] as number) ^ (crc >>> 8);
  }
  return ~crc >>> 0;
}
