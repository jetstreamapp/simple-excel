/** Seed for a fresh CRC-32 computation; pass a previous result to continue across chunks. */
export const CRC32_INITIAL: number = 0;

/** The IEEE CRC-32 polynomial in reflected form: the zip flavour. */
const POLYNOMIAL = 0xedb88320;
/** Entries per table; the flat table holds `SLICE_WIDTH` of them end to end. */
const TABLE_SIZE = 256;
/** Bytes retired per iteration of the sliced loop, one table per byte position. */
const SLICE_WIDTH = 16;

/**
 * Slicing-by-16 tables, built once on first use (16 KB). Table 0 is the classic byte-at-a-time table and each later
 * one is the previous table advanced by another byte, which lets sixteen lookups and fifteen XORs retire sixteen
 * input bytes at a time. They live end to end in one `Int32Array`: sixteen separate arrays measured the same and
 * cost sixteen live bindings in the hot loop.
 */
let tables: Int32Array | undefined;

function buildTables(): Int32Array {
  const built = new Int32Array(TABLE_SIZE * SLICE_WIDTH);
  for (let n = 0; n < TABLE_SIZE; n++) {
    let value = n;
    for (let bit = 0; bit < 8; bit++) {
      value = value & 1 ? POLYNOMIAL ^ (value >>> 1) : value >>> 1;
    }
    built[n] = value;
  }
  for (let level = 1; level < SLICE_WIDTH; level++) {
    for (let n = 0; n < TABLE_SIZE; n++) {
      const value = built[(level - 1) * TABLE_SIZE + n] as number;
      built[level * TABLE_SIZE + n] = (value >>> 8) ^ (built[value & 0xff] as number);
    }
  }
  return built;
}

/**
 * CRC-32 of `bytes`, continuing from `seed` (the value returned for the previous chunk). The result is an unsigned
 * 32-bit integer; `crc32(b, crc32(a)) === crc32(concat(a, b))`.
 */
export function crc32(bytes: Uint8Array, seed: number = CRC32_INITIAL): number {
  const table = (tables ??= buildTables());
  const length = bytes.length;
  let crc = ~seed;
  let index = 0;

  if (length >= SLICE_WIDTH) {
    // A DataView reads the four little-endian words at any alignment, which a Uint32Array view over the chunk cannot.
    const view = new DataView(bytes.buffer, bytes.byteOffset, length);
    const limit = length - SLICE_WIDTH;
    while (index <= limit) {
      const word0 = crc ^ view.getUint32(index, true);
      const word1 = view.getUint32(index + 4, true);
      const word2 = view.getUint32(index + 8, true);
      const word3 = view.getUint32(index + 12, true);
      crc =
        (table[15 * TABLE_SIZE + (word0 & 0xff)] as number) ^
        (table[14 * TABLE_SIZE + ((word0 >>> 8) & 0xff)] as number) ^
        (table[13 * TABLE_SIZE + ((word0 >>> 16) & 0xff)] as number) ^
        (table[12 * TABLE_SIZE + (word0 >>> 24)] as number) ^
        (table[11 * TABLE_SIZE + (word1 & 0xff)] as number) ^
        (table[10 * TABLE_SIZE + ((word1 >>> 8) & 0xff)] as number) ^
        (table[9 * TABLE_SIZE + ((word1 >>> 16) & 0xff)] as number) ^
        (table[8 * TABLE_SIZE + (word1 >>> 24)] as number) ^
        (table[7 * TABLE_SIZE + (word2 & 0xff)] as number) ^
        (table[6 * TABLE_SIZE + ((word2 >>> 8) & 0xff)] as number) ^
        (table[5 * TABLE_SIZE + ((word2 >>> 16) & 0xff)] as number) ^
        (table[4 * TABLE_SIZE + (word2 >>> 24)] as number) ^
        (table[3 * TABLE_SIZE + (word3 & 0xff)] as number) ^
        (table[2 * TABLE_SIZE + ((word3 >>> 8) & 0xff)] as number) ^
        (table[1 * TABLE_SIZE + ((word3 >>> 16) & 0xff)] as number) ^
        (table[word3 >>> 24] as number);
      index += SLICE_WIDTH;
    }
  }

  for (; index < length; index++) {
    crc = (table[(crc ^ (bytes[index] as number)) & 0xff] as number) ^ (crc >>> 8);
  }
  return ~crc >>> 0;
}
