import { describe, expect, it } from 'vitest';
import type { RandomAccessSource } from '../../types';
import { sourceFrom } from '../source';

const sample = new Uint8Array([0, 1, 2, 3, 4, 5, 6, 7, 8, 9]);

describe('sourceFrom', () => {
  describe('Uint8Array', () => {
    it('reports the size and reads ranges', async () => {
      const source = sourceFrom(sample);
      expect(source.size).toBe(10);
      expect(await source.read(0, 4)).toEqual(new Uint8Array([0, 1, 2, 3]));
      expect(await source.read(6, 4)).toEqual(new Uint8Array([6, 7, 8, 9]));
    });

    it('returns zero-copy views', async () => {
      const source = sourceFrom(sample);
      const view = await source.read(2, 3);
      expect(view.buffer).toBe(sample.buffer);
      expect(view.byteOffset).toBe(sample.byteOffset + 2);
    });

    it('honours the offset of a view into a larger buffer', async () => {
      const backing = new Uint8Array([100, 101, 102, 103, 104, 105]);
      const window = backing.subarray(2, 5);
      const source = sourceFrom(window);
      expect(source.size).toBe(3);
      expect(await source.read(0, 3)).toEqual(new Uint8Array([102, 103, 104]));
    });

    it('returns fewer bytes only at the end of the input', async () => {
      const source = sourceFrom(sample);
      expect((await source.read(8, 100)).byteLength).toBe(2);
      expect((await source.read(10, 4)).byteLength).toBe(0);
      expect((await source.read(999, 4)).byteLength).toBe(0);
    });
  });

  describe('ArrayBuffer', () => {
    it('wraps the buffer', async () => {
      const buffer = new ArrayBuffer(4);
      new Uint8Array(buffer).set([9, 8, 7, 6]);
      const source = sourceFrom(buffer);
      expect(source.size).toBe(4);
      expect(await source.read(1, 2)).toEqual(new Uint8Array([8, 7]));
    });
  });

  describe('Blob', () => {
    it('reads slices', async () => {
      const source = sourceFrom(new Blob([sample]));
      expect(source.size).toBe(10);
      expect(await source.read(0, 3)).toEqual(new Uint8Array([0, 1, 2]));
      expect(await source.read(7, 50)).toEqual(new Uint8Array([7, 8, 9]));
      expect((await source.read(10, 5)).byteLength).toBe(0);
    });

    it('never hands back a view the caller can see change', async () => {
      const source = sourceFrom(new Blob([sample]));
      const first = await source.read(0, 4);
      first.fill(0xff);
      expect(await source.read(0, 4)).toEqual(new Uint8Array([0, 1, 2, 3]));
    });

    it('accepts anything shaped like a Blob', async () => {
      const blob = new Blob([sample]);
      const lookalike = {
        size: blob.size,
        slice: (start: number, end: number) => blob.slice(start, end),
        arrayBuffer: () => blob.arrayBuffer(),
      };
      const source = sourceFrom(lookalike as unknown as Blob);
      expect(await source.read(4, 2)).toEqual(new Uint8Array([4, 5]));
    });
  });

  describe('RandomAccessSource', () => {
    it('passes an existing source straight through', () => {
      const existing: RandomAccessSource = {
        size: 3,
        read: () => Promise.resolve(new Uint8Array([1, 2, 3])),
      };
      expect(sourceFrom(existing)).toBe(existing);
    });
  });

  it('rejects anything else', () => {
    expect(() => sourceFrom({ nope: true } as unknown as Uint8Array)).toThrow(TypeError);
  });
});
