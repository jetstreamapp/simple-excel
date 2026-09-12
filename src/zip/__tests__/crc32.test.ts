import { randomBytes } from 'node:crypto';
import { describe, expect, it } from 'vitest';
import { crc32, CRC32_INITIAL } from '../crc32';

/** Bitwise reference implementation, independent of the table. */
function referenceCrc32(bytes: Uint8Array): number {
  let crc = 0xffffffff;
  for (const byte of bytes) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit++) {
      crc = crc & 1 ? (crc >>> 1) ^ 0xedb88320 : crc >>> 1;
    }
  }
  return (crc ^ 0xffffffff) >>> 0;
}

const encode = (text: string): Uint8Array => new TextEncoder().encode(text);

describe('crc32', () => {
  it('matches the known vectors', () => {
    expect(crc32(new Uint8Array(0))).toBe(0);
    expect(crc32(encode('a'))).toBe(0xe8b7be43);
    expect(crc32(encode('123456789'))).toBe(0xcbf43926);
    expect(crc32(encode('The quick brown fox jumps over the lazy dog'))).toBe(0x414fa339);
  });

  it('continues across chunks', () => {
    const whole = randomBytes(100_003);
    const bytes = new Uint8Array(whole.buffer, whole.byteOffset, whole.byteLength);
    let running = CRC32_INITIAL;
    for (let offset = 0; offset < bytes.length; offset += 7_919) {
      running = crc32(bytes.subarray(offset, offset + 7_919), running);
    }
    expect(running).toBe(crc32(bytes));
  });

  it('agrees with a bitwise reference on random input', () => {
    for (let round = 0; round < 20; round++) {
      const random = randomBytes(1 + Math.floor(Math.random() * 4_096));
      const bytes = new Uint8Array(random.buffer, random.byteOffset, random.byteLength);
      expect(crc32(bytes)).toBe(referenceCrc32(bytes));
    }
  });

  it('returns unsigned values', () => {
    expect(crc32(encode('ÿÿÿÿ'))).toBeGreaterThanOrEqual(0);
  });
});
