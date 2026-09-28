import { Buffer } from 'node:buffer';
import { randomBytes } from 'node:crypto';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Writable } from 'node:stream';
import { inflateRawSync } from 'node:zlib';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { fixtureById, readFixture } from '../../../test/helpers/fixtures';
import type { Deflater } from '../../types';
import { fromFile, nodeDeflater, openWorkbook, toFile, toWritable } from '../index';

const KIB = 1024;
const MIB = 1024 * 1024;

let workingDirectory = '';

beforeAll(async () => {
  workingDirectory = await mkdtemp(join(tmpdir(), 'simple-excel-node-'));
});

afterAll(async () => {
  await rm(workingDirectory, { recursive: true, force: true });
});

function bytesOf(buffer: Buffer): Uint8Array {
  return new Uint8Array(buffer.buffer, buffer.byteOffset, buffer.byteLength);
}

/** A buffer of `size` bytes built from a short repeating pattern: compressible, and cheap to build. */
function repetitiveBytes(size: number): Uint8Array {
  const pattern = Buffer.from('the quick brown fox jumps over the lazy dog, again and again and again.\n');
  const bytes = Buffer.alloc(size);
  for (let offset = 0; offset < size; offset += pattern.byteLength) {
    pattern.copy(bytes, offset, 0, Math.min(pattern.byteLength, size - offset));
  }
  return bytesOf(bytes);
}

describe('fromFile', () => {
  it('reads whole and partial ranges and reports the size', async () => {
    const path = join(workingDirectory, 'random.bin');
    const contents = randomBytes(256 * KIB + 17);
    await writeFile(path, contents);

    const source = await fromFile(path);
    try {
      expect(source.size).toBe(contents.byteLength);
      expect(Buffer.compare(Buffer.from(await source.read(0, source.size)), contents)).toBe(0);
      expect(Buffer.compare(Buffer.from(await source.read(1000, 4096)), contents.subarray(1000, 5096))).toBe(0);
    } finally {
      await source.close();
    }
  });

  it('returns a short view at end of file and nothing past it', async () => {
    const path = join(workingDirectory, 'short.bin');
    const contents = randomBytes(100);
    await writeFile(path, contents);

    const source = await fromFile(path);
    try {
      const tail = await source.read(90, 4096);
      expect(tail.byteLength).toBe(10);
      expect(Buffer.compare(Buffer.from(tail), contents.subarray(90))).toBe(0);
      expect((await source.read(100, 16)).byteLength).toBe(0);
      expect((await source.read(0, 0)).byteLength).toBe(0);
    } finally {
      await source.close();
    }
  });

  it('serves concurrent reads from one handle', async () => {
    const path = join(workingDirectory, 'concurrent.bin');
    const contents = randomBytes(512 * KIB);
    await writeFile(path, contents);

    const source = await fromFile(path);
    try {
      const ranges = [
        [0, 4096],
        [4096, 64 * KIB],
        [200 * KIB, 128 * KIB],
        [511 * KIB, 4096],
      ] as const;
      const results = await Promise.all(ranges.map(([offset, length]) => source.read(offset, length)));
      results.forEach((result, index) => {
        const [offset, length] = ranges[index] ?? [0, 0];
        expect(Buffer.compare(Buffer.from(result), contents.subarray(offset, offset + length))).toBe(0);
      });
    } finally {
      await source.close();
    }
  });

  it('closes idempotently', async () => {
    const path = join(workingDirectory, 'closed.bin');
    await writeFile(path, randomBytes(8));
    const source = await fromFile(path);
    await source.close();
    await expect(source.close()).resolves.toBeUndefined();
  });

  it('refuses a read after close with ABORTED rather than a bare EBADF', async () => {
    const path = join(workingDirectory, 'read-after-close.bin');
    await writeFile(path, randomBytes(8));
    const source = await fromFile(path);
    await source.close();
    await expect(source.read(0, 4)).rejects.toMatchObject({ name: 'XlsxError', code: 'ABORTED' });
  });

  it('fails a sheet read after workbook.close() with ABORTED, as in-memory bytes do', async () => {
    const path = join(workingDirectory, 'closed-workbook.xlsx');
    await writeFile(path, readFixture(fixtureById('edge-baseline-minimal')));
    const workbook = await openWorkbook(await fromFile(path));
    const sheet = workbook.sheet(0);
    expect((await sheet.toObjects()).rows.length).toBeGreaterThan(0);
    await workbook.close();
    await expect(sheet.toObjects()).rejects.toMatchObject({
      name: 'XlsxError',
      code: 'ABORTED',
      message: expect.stringContaining('closed'),
    });
    await expect(sheet.head(1)).rejects.toMatchObject({ code: 'ABORTED' });
  });
});

describe('toFile', () => {
  it('writes 20 MiB in 64 KiB chunks byte for byte', async () => {
    const path = join(workingDirectory, 'large.bin');
    const chunk = repetitiveBytes(64 * KIB);
    const chunkCount = 320;

    const sink = toFile(path);
    for (let i = 0; i < chunkCount; i++) {
      await sink.write(chunk);
    }
    await sink.close();

    const written = await readFile(path);
    expect(written.byteLength).toBe(chunkCount * chunk.byteLength);
    expect(Buffer.compare(written, Buffer.concat(Array.from({ length: chunkCount }, () => Buffer.from(chunk))))).toBe(0);
  });
});

describe('toWritable', () => {
  it('waits for the stream to drain before resolving a write', async () => {
    const flushed: number[] = [];
    const writable = new Writable({
      highWaterMark: 1024,
      write(chunk: Uint8Array, _encoding, callback): void {
        setTimeout(() => {
          flushed.push(chunk.byteLength);
          callback();
        }, 10);
      },
    });

    const sink = toWritable(writable);
    const startedAt = Date.now();
    for (let i = 0; i < 3; i++) {
      await sink.write(new Uint8Array(4096));
      // The chunk is past the high-water mark, so the write only resolves once the slow flush completed.
      expect(flushed.length).toBeGreaterThanOrEqual(i + 1);
    }
    await sink.close();

    expect(flushed).toEqual([4096, 4096, 4096]);
    expect(Date.now() - startedAt).toBeGreaterThanOrEqual(25);
  });

  it('makes a stream error sticky', async () => {
    const failure = new Error('disk full');
    const writable = new Writable({
      write(_chunk, _encoding, callback): void {
        callback(failure);
      },
    });

    const sink = toWritable(writable);
    await sink.write(new Uint8Array(8)).catch(() => undefined);
    await expect(sink.close()).rejects.toThrow(failure);
    await expect(sink.write(new Uint8Array(8))).rejects.toThrow(failure);
  });

  it('destroys the stream on abort and refuses later writes', async () => {
    const writable = new Writable({
      write(_chunk, _encoding, callback): void {
        callback();
      },
    });

    const sink = toWritable(writable);
    await sink.write(new Uint8Array(8));
    await sink.abort('the export was cancelled');
    await sink.abort('again');

    expect(writable.destroyed).toBe(true);
    await expect(sink.write(new Uint8Array(8))).rejects.toMatchObject({ code: 'ABORTED' });
  });
});

interface DeflateRun {
  readonly deflated: Uint8Array;
  readonly chunkCount: number;
  readonly overlaps: number;
  readonly deflater: Deflater;
}

/** Pushes `input` through a zlib deflater in `chunkSize` slices and records how the output came back. */
async function deflate(input: Uint8Array, level: number | undefined, chunkSize: number, onChunkDelay = 0): Promise<DeflateRun> {
  const output: Uint8Array[] = [];
  let overlaps = 0;
  let inFlight = 0;

  const deflater = nodeDeflater(level)(
    async chunk => {
      if (inFlight > 0) {
        overlaps++;
      }
      inFlight++;
      if (onChunkDelay > 0) {
        await new Promise(resolve => setTimeout(resolve, onChunkDelay));
      }
      output.push(chunk);
      inFlight--;
    },
    { method: 'deflate' },
  );

  for (let offset = 0; offset < input.byteLength; offset += chunkSize) {
    await deflater.push(input.subarray(offset, offset + chunkSize));
  }
  await deflater.finish();

  return { deflated: bytesOf(Buffer.concat(output.map(chunk => Buffer.from(chunk)))), chunkCount: output.length, overlaps, deflater };
}

describe('nodeDeflater', () => {
  const sizes: ReadonlyArray<readonly [string, Uint8Array]> = [
    ['0 B', new Uint8Array(0)],
    ['1 MiB of random bytes', bytesOf(randomBytes(MIB))],
    ['30 MiB of repetitive text', repetitiveBytes(30 * MIB)],
  ];

  for (const level of [1, 6]) {
    for (const [label, input] of sizes) {
      it(`round trips ${label} at level ${level}`, async () => {
        const { deflated, overlaps, deflater } = await deflate(input, level, 256 * KIB);

        expect(overlaps).toBe(0);
        expect(deflater.bytesIn).toBe(input.byteLength);
        expect(deflater.bytesOut).toBe(deflated.byteLength);
        expect(Buffer.compare(inflateRawSync(deflated), Buffer.from(input))).toBe(0);
      });
    }
  }

  it('compresses harder at level 6 than at level 1', async () => {
    const input = repetitiveBytes(4 * MIB);
    const fast = await deflate(input, 1, 64 * KIB);
    const small = await deflate(input, 6, 64 * KIB);
    expect(small.deflated.byteLength).toBeLessThan(fast.deflated.byteLength);
  });

  it('keeps output in order behind a slow consumer', async () => {
    const input = bytesOf(randomBytes(2 * MIB));
    const { deflated, chunkCount, overlaps } = await deflate(input, 1, 64 * KIB, 1);
    // Random bytes barely compress, so the pump delivered many chunks: order and exclusivity both matter.
    expect(chunkCount).toBeGreaterThan(10);
    expect(overlaps).toBe(0);
    expect(Buffer.compare(inflateRawSync(deflated), Buffer.from(input))).toBe(0);
  });

  it('passes bytes through unchanged for method "store"', async () => {
    const chunks = [repetitiveBytes(1024), repetitiveBytes(2048)];
    const output: Uint8Array[] = [];
    const deflater = nodeDeflater(1)(
      async chunk => {
        output.push(chunk);
      },
      { method: 'store' },
    );

    for (const chunk of chunks) {
      await deflater.push(chunk);
    }
    await deflater.finish();

    expect(deflater.bytesIn).toBe(3072);
    expect(deflater.bytesOut).toBe(3072);
    expect(output).toEqual(chunks);
  });

  it('reports a consumer failure instead of hanging in finish()', async () => {
    const failure = new Error('the sink went away');
    const deflater = nodeDeflater(1)(
      async () => {
        throw failure;
      },
      { method: 'deflate' },
    );

    await deflater.push(bytesOf(randomBytes(MIB))).catch(() => undefined);
    await expect(deflater.finish()).rejects.toThrow(failure);
  });

  it('rejects later calls after an abort', async () => {
    const deflater = nodeDeflater(1)(async () => undefined, { method: 'deflate' });
    await deflater.push(repetitiveBytes(64 * KIB));
    await deflater.abort(new Error('cancelled'));

    await expect(deflater.push(repetitiveBytes(16))).rejects.toThrow('cancelled');
    await expect(deflater.finish()).rejects.toThrow('cancelled');
  });
});
