import { createHash, randomBytes } from 'node:crypto';
import { inflateRawSync } from 'node:zlib';
import { describe, expect, it } from 'vitest';
import { isXlsxError } from '../../errors';
import type { ChunkHandler, Deflater } from '../../types';
import { createDeflater, createStoredDeflater, hasNativeDeflate } from '../deflater';

function random(byteLength: number): Uint8Array {
  const buffer = randomBytes(byteLength);
  return new Uint8Array(buffer.buffer, buffer.byteOffset, buffer.byteLength);
}

function concat(chunks: readonly Uint8Array[]): Uint8Array {
  const total = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
  const joined = new Uint8Array(total);
  let position = 0;
  for (const chunk of chunks) {
    joined.set(chunk, position);
    position += chunk.length;
  }
  return joined;
}

function digest(bytes: Uint8Array): string {
  return createHash('sha256').update(bytes).digest('hex');
}

function inflate(bytes: Uint8Array): Uint8Array {
  const inflated = inflateRawSync(bytes);
  return new Uint8Array(inflated.buffer, inflated.byteOffset, inflated.byteLength);
}

interface Collector {
  readonly handler: ChunkHandler;
  bytes(): Uint8Array;
}

function collector(): Collector {
  const chunks: Uint8Array[] = [];
  return {
    handler: async (chunk: Uint8Array): Promise<void> => {
      chunks.push(chunk.slice());
    },
    bytes: () => concat(chunks),
  };
}

async function pushAll(deflater: Deflater, bytes: Uint8Array, chunkSize: number): Promise<void> {
  for (let offset = 0; offset < bytes.length; offset += chunkSize) {
    await deflater.push(bytes.subarray(offset, Math.min(offset + chunkSize, bytes.length)));
  }
}

function delay(milliseconds: number): Promise<void> {
  return new Promise(resolve => {
    setTimeout(resolve, milliseconds);
  });
}

/** Repetitive but not uniform: rows of sheet-shaped XML, the payload the writer actually compresses. */
function sheetXmlChunk(firstRow: number, targetBytes: number): Uint8Array {
  const rows: string[] = [];
  let bytes = 0;
  let row = firstRow;
  while (bytes < targetBytes) {
    const text =
      `<row r="${row}" spans="1:6"><c r="A${row}" t="inlineStr"><is><t>Account ${row}</t></is></c>` +
      `<c r="B${row}"><v>${(row * 37) % 100_000}.${row % 97}</v></c>` +
      `<c r="C${row}" s="2"><v>${45_000 + (row % 900)}</v></c>` +
      `<c r="D${row}" t="inlineStr"><is><t>0015000000${(row % 100_000).toString(36)}AAA</t></is></c>` +
      `<c r="E${row}" t="b"><v>${row % 2}</v></c><c r="F${row}"><v>${row}</v></c></row>`;
    rows.push(text);
    bytes += text.length;
    row++;
  }
  return new TextEncoder().encode(rows.join(''));
}

describe('hasNativeDeflate', () => {
  it('detects CompressionStream without throwing', () => {
    expect(hasNativeDeflate()).toBe(true);
  });
});

describe('createDeflater (native)', () => {
  it.each([
    { label: '0 bytes', bytes: 0 },
    { label: '1 byte', bytes: 1 },
    { label: '1 MiB of random data', bytes: 1024 * 1024 },
  ])('round trips $label through inflateRawSync', async ({ bytes }) => {
    const input = random(bytes);
    const sink = collector();
    const deflater = createDeflater(sink.handler, { method: 'deflate' });

    await pushAll(deflater, input, 64 * 1024);
    await deflater.finish();

    const compressed = sink.bytes();
    expect(digest(inflate(compressed))).toBe(digest(input));
    expect(deflater.bytesIn).toBe(input.length);
    expect(deflater.bytesOut).toBe(compressed.length);
  });

  it('round trips 30 MiB of repetitive data', async () => {
    const input = sheetXmlChunk(1, 30 * 1024 * 1024);
    const sink = collector();
    const deflater = createDeflater(sink.handler, { method: 'deflate' });

    await pushAll(deflater, input, 256 * 1024);
    await deflater.finish();

    const compressed = sink.bytes();
    expect(digest(inflate(compressed))).toBe(digest(input));
    expect(deflater.bytesIn).toBe(input.length);
    expect(compressed.length).toBeLessThan(input.length / 2);
  }, 60_000);

  it('ignores the compression level the platform API cannot express', async () => {
    const input = sheetXmlChunk(1, 256 * 1024);
    const [fast, best] = await Promise.all(
      [1, 9].map(async level => {
        const sink = collector();
        const deflater = createDeflater(sink.handler, { method: 'deflate', level });
        await deflater.push(input);
        await deflater.finish();
        return sink.bytes();
      }),
    );

    expect(digest(inflate(fast!))).toBe(digest(input));
    expect(fast!.length).toBe(best!.length);
  });

  it('applies back-pressure while the chunk handler is slow', async () => {
    const events: string[] = [];
    let handlersInFlight = 0;
    let maxHandlersInFlight = 0;
    const handler: ChunkHandler = async (): Promise<void> => {
      handlersInFlight++;
      maxHandlersInFlight = Math.max(maxHandlersInFlight, handlersInFlight);
      events.push('chunk:start');
      await delay(10);
      events.push('chunk:end');
      handlersInFlight--;
    };

    const pushCount = 6;
    const deflater = createDeflater(handler, { method: 'deflate' });
    const startedAt = Date.now();
    for (let index = 0; index < pushCount; index++) {
      await deflater.push(random(128 * 1024));
      events.push(`push:${index}`);
    }
    const elapsedWhilePushing = Date.now() - startedAt;
    await deflater.finish();

    // One handler call at a time, in order: the drain loop never starts the next read before the previous
    // chunk was handed over.
    expect(maxHandlersInFlight).toBe(1);
    expect(events.filter(event => event === 'chunk:end').length).toBeGreaterThan(0);
    // The pushes waited on the handler rather than racing ahead of it.
    expect(events.slice(0, events.lastIndexOf(`push:${pushCount - 1}`))).toContain('chunk:end');
    expect(elapsedWhilePushing).toBeGreaterThanOrEqual(10);
    expect(events.at(-1)).toBe('chunk:end');
  });

  it('latches a failure from the chunk handler and rejects every later call with it', async () => {
    const boom = new Error('the sink exploded');
    let calls = 0;
    const deflater = createDeflater(
      async (): Promise<void> => {
        calls++;
        throw boom;
      },
      { method: 'deflate' },
    );

    await deflater.push(random(512 * 1024)).catch(() => {});
    await expect(deflater.finish()).rejects.toBe(boom);
    await expect(deflater.push(random(16))).rejects.toBe(boom);
    await expect(deflater.finish()).rejects.toBe(boom);
    expect(calls).toBe(1);
  });

  it('aborts mid-stream and rejects later calls with the abort reason', async () => {
    const sink = collector();
    const deflater = createDeflater(sink.handler, { method: 'deflate' });
    await deflater.push(random(256 * 1024));

    const cancelled = new Error('the user cancelled the download');
    await deflater.abort(cancelled);

    await expect(deflater.push(random(16))).rejects.toBe(cancelled);
    await expect(deflater.finish()).rejects.toBe(cancelled);
  });

  it('aborts with a classified error when no reason is given', async () => {
    const sink = collector();
    const deflater = createDeflater(sink.handler, { method: 'deflate' });
    await deflater.push(random(1024));
    await deflater.abort();

    await expect(deflater.finish()).rejects.toSatisfy(reason => isXlsxError(reason) && reason.code === 'ABORTED');
  });

  it('rejects a push after finish', async () => {
    const sink = collector();
    const deflater = createDeflater(sink.handler, { method: 'deflate' });
    await deflater.push(random(1024));
    await deflater.finish();

    await expect(deflater.push(random(16))).rejects.toSatisfy(reason => isXlsxError(reason) && reason.code === 'WRITER_STATE');
  });

  it('counts every byte in and out', async () => {
    const input = random(300_007);
    const sink = collector();
    const deflater = createDeflater(sink.handler, { method: 'deflate' });

    expect(deflater.bytesIn).toBe(0);
    expect(deflater.bytesOut).toBe(0);
    await pushAll(deflater, input, 7_919);
    expect(deflater.bytesIn).toBe(input.length);
    await deflater.finish();

    expect(deflater.bytesOut).toBe(sink.bytes().length);
    // Random data barely compresses, so the raw deflate stream is within a few percent of the input.
    expect(deflater.bytesOut).toBeGreaterThan(input.length * 0.9);
  });

  it('measures native deflate throughput for 64 MiB of sheet XML', async () => {
    const chunkSize = 1024 * 1024;
    const chunkCount = 64;
    const chunks: Uint8Array[] = [];
    for (let index = 0; index < chunkCount; index++) {
      chunks.push(sheetXmlChunk(index * 10_000 + 1, chunkSize));
    }
    const totalBytes = chunks.reduce((sum, chunk) => sum + chunk.length, 0);

    let bytesOut = 0;
    const deflater = createDeflater(
      async (chunk: Uint8Array): Promise<void> => {
        bytesOut += chunk.length;
      },
      { method: 'deflate' },
    );

    const startedAt = performance.now();
    for (const chunk of chunks) {
      await deflater.push(chunk);
    }
    await deflater.finish();
    const elapsedSeconds = (performance.now() - startedAt) / 1000;

    const megabytesPerSecond = totalBytes / 1024 / 1024 / elapsedSeconds;
    // oxlint-disable-next-line no-console
    console.log(
      `native deflate: ${(totalBytes / 1024 / 1024).toFixed(0)} MiB in ${(elapsedSeconds * 1000).toFixed(0)} ms ` +
        `= ${megabytesPerSecond.toFixed(1)} MB/s, ratio ${(totalBytes / bytesOut).toFixed(1)}x`,
    );
    expect(deflater.bytesIn).toBe(totalBytes);
    expect(bytesOut).toBeGreaterThan(0);
  }, 120_000);
});

describe('stored deflater', () => {
  it('passes bytes straight through', async () => {
    const input = random(100_003);
    const sink = collector();
    const deflater = createStoredDeflater(sink.handler);

    await pushAll(deflater, input, 4_096);
    await deflater.finish();

    expect(digest(sink.bytes())).toBe(digest(input));
    expect(deflater.bytesIn).toBe(input.length);
    expect(deflater.bytesOut).toBe(input.length);
  });

  it('is what createDeflater returns for method store', async () => {
    const input = random(50_000);
    const sink = collector();
    const deflater = createDeflater(sink.handler, { method: 'store' });

    await deflater.push(input);
    await deflater.finish();

    expect(digest(sink.bytes())).toBe(digest(input));
    expect(deflater.bytesOut).toBe(deflater.bytesIn);
  });

  it('latches handler failures and abort reasons like the native path', async () => {
    const boom = new Error('the sink exploded');
    const failing = createStoredDeflater(async (): Promise<void> => {
      throw boom;
    });
    await expect(failing.push(random(16))).rejects.toBe(boom);
    await expect(failing.finish()).rejects.toBe(boom);

    const aborted = createStoredDeflater(async (): Promise<void> => {});
    await aborted.abort();
    await expect(aborted.push(random(16))).rejects.toSatisfy(reason => isXlsxError(reason) && reason.code === 'ABORTED');
  });
});
