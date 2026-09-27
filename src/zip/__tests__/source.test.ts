import { runInNewContext } from 'node:vm';
import { describe, expect, it } from 'vitest';
import { isXlsxError } from '../../errors';
import type { RandomAccessSource } from '../../types';
import { sourceFrom, type SourceInput } from '../source';

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

  describe('other binary inputs (EC-INPUT-TYPE)', () => {
    it('views any ArrayBufferView without copying', async () => {
      const backing = new Uint8Array([100, 0, 1, 2, 3, 4, 5, 6, 7, 200]);
      for (const view of [
        new DataView(backing.buffer, 1, 8),
        new Int8Array(backing.buffer, 1, 8),
        new Uint8ClampedArray(backing.buffer, 1, 8),
        new Uint16Array(backing.buffer.slice(2), 0, 4),
      ]) {
        const source = sourceFrom(view);
        expect(source.size, view.constructor.name).toBe(8);
        const read = await source.read(0, 8);
        expect(read.buffer, view.constructor.name).toBe(view.buffer);
      }
      expect(await sourceFrom(new DataView(backing.buffer, 1, 8)).read(2, 3)).toEqual(new Uint8Array([2, 3, 4]));
    });

    it('reads a SharedArrayBuffer, copying each range because streams refuse shared memory', async () => {
      const shared = new SharedArrayBuffer(4);
      new Uint8Array(shared).set([9, 8, 7, 6]);
      for (const input of [shared, new Uint8Array(shared)]) {
        const source = sourceFrom(input);
        expect(source.size).toBe(4);
        const read = await source.read(1, 2);
        expect(read).toEqual(new Uint8Array([8, 7]));
        expect(Object.prototype.toString.call(read.buffer)).toBe('[object ArrayBuffer]');
      }
    });

    it('accepts buffers and views from another realm', async () => {
      const foreignBuffer = runInNewContext('const b = new ArrayBuffer(3); new Uint8Array(b).set([1, 2, 3]); b') as ArrayBuffer;
      const foreignView = runInNewContext('new Int16Array([258, 772])') as Int16Array;
      expect(foreignBuffer instanceof ArrayBuffer).toBe(false);
      expect(await sourceFrom(foreignBuffer).read(0, 3)).toEqual(new Uint8Array([1, 2, 3]));
      expect(await sourceFrom(foreignView).read(0, 4)).toEqual(new Uint8Array([2, 1, 4, 3]));
    });
  });

  it('EC-INPUT-TYPE: rejects anything else with a classified error naming the accepted inputs', () => {
    for (const input of [{ nope: true }, null, undefined, 7, Symbol('x'), 'PK\u0003\u0004']) {
      let thrown: unknown;
      try {
        sourceFrom(input as unknown as SourceInput);
      } catch (error) {
        thrown = error;
      }
      expect(isXlsxError(thrown) && thrown.code, String(input)).toBe('NOT_XLSX');
      expect(isXlsxError(thrown) && thrown.message, String(input)).toContain('RandomAccessSource');
    }
  });
});
