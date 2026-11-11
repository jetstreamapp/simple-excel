import { Buffer } from 'node:buffer';
import { randomBytes } from 'node:crypto';
import { mkdir, mkdtemp, readdir, readFile, rm, writeFile } from 'node:fs/promises';
import { createServer, get, type IncomingMessage } from 'node:http';
import type { AddressInfo } from 'node:net';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Writable } from 'node:stream';
import { inflateRawSync } from 'node:zlib';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { fixtureById, readFixture } from '../../../test/helpers/fixtures';
import type { Deflater } from '../../types';
import { createWorkbookWriter, fromFile, isXlsxError, nodeDeflater, openWorkbook, toFile, toWritable, XlsxError } from '../index';

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

  it('reports a file that cannot be opened as IO_ERROR with the fs error as the cause', async () => {
    const path = join(workingDirectory, 'does-not-exist.xlsx');
    const failure = await fromFile(path).then(
      () => undefined,
      (reason: unknown) => reason,
    );
    expect(failure).toBeInstanceOf(XlsxError);
    expect(failure).toMatchObject({ code: 'IO_ERROR', detail: { path }, cause: expect.objectContaining({ code: 'ENOENT' }) });
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

/** The files in `directory`, hidden ones included. */
async function listDirectory(directory: string): Promise<string[]> {
  return (await readdir(directory)).toSorted();
}

/** A fresh, empty directory under the suite's working directory. */
async function freshDirectory(name: string): Promise<string> {
  const directory = join(workingDirectory, name);
  await mkdir(directory);
  return directory;
}

/** What `run` rejects with, or undefined when it resolves. */
async function failureOf(run: () => Promise<unknown>): Promise<unknown> {
  try {
    await run();
  } catch (error) {
    return error;
  }
  return undefined;
}

describe('toFile', () => {
  it('writes to a hidden temporary file and only puts the complete file at the target on close', async () => {
    const directory = await freshDirectory('to-file-rename');
    const path = join(directory, 'export.xlsx');
    const sink = toFile(path);
    await sink.write(repetitiveBytes(64 * KIB));

    const during = await listDirectory(directory);
    expect(during).toHaveLength(1);
    expect(during[0]).toMatch(/^\.export\.xlsx\.[0-9a-f]+\.partial$/);

    await sink.close();
    expect(await listDirectory(directory)).toEqual(['export.xlsx']);
    expect((await readFile(path)).byteLength).toBe(64 * KIB);
  });

  it('leaves an existing file untouched until the new one is complete, and in place after an abort', async () => {
    const directory = await freshDirectory('to-file-existing');
    const path = join(directory, 'report.xlsx');
    await writeFile(path, 'previous export');

    const workbook = createWorkbookWriter(toFile(path));
    const sheet = workbook.addSheet('Data');
    for (let index = 0; index < 20_000; index++) {
      await sheet.writeRow([index, `hello world ${index}`]);
    }
    expect(await readFile(path, 'utf8')).toBe('previous export');
    await workbook.abort(new Error('the upstream query failed'));

    expect(await listDirectory(directory)).toEqual(['report.xlsx']);
    expect(await readFile(path, 'utf8')).toBe('previous export');
  });

  it('removes the temporary file and writes nothing at the target when the workbook fails', async () => {
    const directory = await freshDirectory('to-file-failure');
    const path = join(directory, 'broken.xlsx');
    const workbook = createWorkbookWriter(toFile(path));
    const sheet = workbook.addSheet('Data');
    await sheet.writeRow(['fine']);
    await expect(sheet.writeRow([{ not: 'a cell' } as never])).rejects.toMatchObject({ code: 'WRITER_STATE' });
    await expect(workbook.close()).rejects.toMatchObject({ code: 'WRITER_STATE' });
    // The row path starts the sink abort without waiting for it; give the unlink a moment to land.
    for (let attempt = 0; attempt < 50 && (await listDirectory(directory)).length > 0; attempt++) {
      await new Promise(resolve => setTimeout(resolve, 10));
    }
    expect(await listDirectory(directory)).toEqual([]);
  });

  it('creates nothing until the first byte, so a sink that is never used leaves no file behind', async () => {
    const directory = await freshDirectory('to-file-unused');
    const unused = toFile(join(directory, 'unused.xlsx'));
    const abortedFirst = toFile(join(directory, 'aborted.xlsx'));
    await abortedFirst.abort(new Error('the export was cancelled before it began'));
    await expect(abortedFirst.write(new Uint8Array(1))).rejects.toMatchObject({ code: 'ABORTED' });
    expect(unused).toBeDefined();
    expect(await listDirectory(directory)).toEqual([]);
  });

  it('reports a directory that does not exist as IO_ERROR with the fs error as the cause', async () => {
    const path = join(workingDirectory, 'no', 'such', 'directory', 'x.xlsx');
    const workbook = createWorkbookWriter(toFile(path));
    const sheet = workbook.addSheet('Data');
    const failure = await failureOf(async () => {
      for (let index = 0; index < 1000; index++) {
        await sheet.writeRow([index]);
      }
      await sheet.close();
      await workbook.close();
    });
    expect(failure).toMatchObject({ code: 'IO_ERROR', cause: expect.objectContaining({ code: 'ENOENT' }) });
  });

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

  it('makes a stream error sticky, as IO_ERROR with the stream error as the cause', async () => {
    const failure = new Error('disk full');
    const writable = new Writable({
      write(_chunk, _encoding, callback): void {
        callback(failure);
      },
    });

    const sink = toWritable(writable);
    await sink.write(new Uint8Array(8)).catch(() => undefined);
    await expect(sink.close()).rejects.toMatchObject({ code: 'IO_ERROR', cause: failure });
    await expect(sink.write(new Uint8Array(8))).rejects.toMatchObject({ code: 'IO_ERROR', cause: failure });
  });

  it('rejects a write waiting for drain when the stream is destroyed without an error', async () => {
    const writable = new Writable({
      highWaterMark: 16,
      write(): void {
        // Never calls back: the stream stays full until it is destroyed.
      },
    });
    const sink = toWritable(writable);
    const pending = sink.write(new Uint8Array(64));
    writable.destroy();
    await expect(pending).rejects.toMatchObject({ name: 'XlsxError', code: 'IO_ERROR' });
    await expect(sink.write(new Uint8Array(1))).rejects.toMatchObject({ code: 'IO_ERROR' });
    await expect(sink.close()).rejects.toMatchObject({ code: 'IO_ERROR' });
  });

  it('refuses a write to a stream that was already destroyed or ended', async () => {
    const destroyed = new Writable({ write: (_chunk, _encoding, callback): void => callback() });
    destroyed.destroy();
    await expect(toWritable(destroyed).write(new Uint8Array(1))).rejects.toMatchObject({ code: 'IO_ERROR' });

    const ended = new Writable({ write: (_chunk, _encoding, callback): void => callback() });
    ended.end();
    await expect(toWritable(ended).write(new Uint8Array(1))).rejects.toMatchObject({ code: 'IO_ERROR' });
  });

  it('resolves close() for a stream that already finished', async () => {
    const writable = new Writable({ write: (_chunk, _encoding, callback): void => callback() });
    const sink = toWritable(writable);
    await sink.write(new Uint8Array(4));
    await new Promise<void>(resolve => writable.end(resolve));
    await expect(sink.close()).resolves.toBeUndefined();
  });

  it('rejects writeRow promptly when an HTTP client disconnects mid-download', async () => {
    let serverOutcome: Promise<unknown> | undefined;
    const server = createServer((_request, response) => {
      serverOutcome = failureOf(async () => {
        const workbook = createWorkbookWriter(toWritable(response), { deflater: nodeDeflater(1) });
        const sheet = workbook.addSheet('Data', { header: ['a', 'b', 'c'] });
        for (let index = 0; index < 2_000_000; index++) {
          await sheet.writeRow([index, `${'x'.repeat(50)}${Math.random()}`, Math.random()]);
        }
        await sheet.close();
        await workbook.close();
      });
    });
    await new Promise<void>(resolve => server.listen(0, '127.0.0.1', resolve));
    try {
      const { port } = server.address() as AddressInfo;
      await new Promise<void>(resolve => {
        const request = get({ host: '127.0.0.1', port }, (response: IncomingMessage) => {
          let received = 0;
          response.on('data', (chunk: Buffer) => {
            received += chunk.byteLength;
            if (received > 200_000) {
              request.destroy();
              resolve();
            }
          });
        });
        request.on('error', () => undefined);
      });
      const startedAt = Date.now();
      const outcome = await serverOutcome;
      expect(Date.now() - startedAt).toBeLessThan(3000);
      expect(isXlsxError(outcome)).toBe(true);
      expect(outcome).toMatchObject({ code: 'IO_ERROR' });
    } finally {
      server.closeAllConnections();
      await new Promise<void>(resolve => server.close(() => resolve()));
    }
  });

  it('wb.abort() resolves and the pending writeRow rejects while a paused stream holds back-pressure', async () => {
    const writable = new Writable({
      highWaterMark: 1024,
      write(): void {
        // Never drains.
      },
    });
    const workbook = createWorkbookWriter(toWritable(writable), { compression: 'store' });
    const sheet = workbook.addSheet('Data');
    let pendingRow: Promise<void> | undefined;
    for (let index = 0; index < 20_000 && pendingRow === undefined; index++) {
      const write = sheet.writeRow([index, `row ${index} with some text to fill the flush buffer`]);
      const outcome = await Promise.race([
        write.then(() => 'settled'),
        new Promise<string>(resolve => setTimeout(() => resolve('pending'), 20)),
      ]);
      if (outcome === 'pending') {
        pendingRow = write;
      }
    }
    expect(pendingRow, 'a row should be waiting for the stream to drain').toBeDefined();

    await workbook.abort(new Error('the user cancelled'));
    await expect(pendingRow).rejects.toMatchObject({ code: 'ABORTED' });
    expect(writable.destroyed).toBe(true);
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
    await expect(deflater.finish()).rejects.toMatchObject({ code: 'IO_ERROR', cause: failure });
  });

  it('rejects later calls after an abort', async () => {
    const deflater = nodeDeflater(1)(async () => undefined, { method: 'deflate' });
    await deflater.push(repetitiveBytes(64 * KIB));
    await deflater.abort(new Error('cancelled'));

    await expect(deflater.push(repetitiveBytes(16))).rejects.toMatchObject({ code: 'ABORTED', cause: new Error('cancelled') });
    await expect(deflater.finish()).rejects.toMatchObject({ code: 'ABORTED', cause: new Error('cancelled') });
  });
});
