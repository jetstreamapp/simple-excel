import { randomBytes } from 'node:crypto';
import { deflateRawSync } from 'node:zlib';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { isXlsxError, XlsxError } from '../../errors';
import { createInflater } from '../inflater';

const MIB = 1024 * 1024;

function bytesOf(buffer: Buffer): Uint8Array {
  return new Uint8Array(buffer.buffer, buffer.byteOffset, buffer.byteLength);
}

function deflate(bytes: Uint8Array): Uint8Array {
  return bytesOf(deflateRawSync(bytes, { level: 6 }));
}

function repetitive(length: number): Uint8Array {
  const bytes = new Uint8Array(length);
  bytes.fill(0x41);
  return bytes;
}

function concat(chunks: readonly Uint8Array[]): Uint8Array {
  const total = chunks.reduce((sum, chunk) => sum + chunk.byteLength, 0);
  const joined = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    joined.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return joined;
}

/** Feed `compressed` through an inflater in `chunkSize` slices and return everything `onChunk` received. */
async function inflateAll(
  compressed: Uint8Array,
  options: { chunkSize?: number; maxBytes?: number; method?: 'deflate' | 'store' } = {},
): Promise<Uint8Array> {
  const received: Uint8Array[] = [];
  const inflater = createInflater(
    async chunk => {
      received.push(chunk);
    },
    { maxBytes: options.maxBytes ?? 64 * MIB, method: options.method ?? 'deflate' },
  );
  const chunkSize = options.chunkSize ?? compressed.byteLength;
  for (let offset = 0; offset < compressed.byteLength; offset += chunkSize) {
    await inflater.push(compressed.subarray(offset, offset + chunkSize));
  }
  await inflater.finish();
  return concat(received);
}

async function captureUnhandledRejections(run: () => Promise<void>): Promise<unknown[]> {
  const rejections: unknown[] = [];
  const handler = (reason: unknown): void => {
    rejections.push(reason);
  };
  process.on('unhandledRejection', handler);
  try {
    await run();
    // Node reports unhandled rejections once the microtask queue has drained, so give it a turn.
    await new Promise(resolve => setTimeout(resolve, 25));
  } finally {
    process.off('unhandledRejection', handler);
  }
  return rejections;
}

describe('createInflater', () => {
  describe('deflate', () => {
    it('round trips an empty stream', async () => {
      expect(await inflateAll(deflate(new Uint8Array(0)))).toEqual(new Uint8Array(0));
    });

    it('round trips a single byte', async () => {
      expect(await inflateAll(deflate(new Uint8Array([0x7a])))).toEqual(new Uint8Array([0x7a]));
    });

    it('round trips 1 MiB of random bytes pushed in small slices', async () => {
      const original = bytesOf(randomBytes(MIB));
      const inflated = await inflateAll(deflate(original), { chunkSize: 4_096 });
      expect(inflated.byteLength).toBe(original.byteLength);
      expect(inflated).toEqual(original);
    });

    it('round trips 30 MB of repetitive bytes', async () => {
      const original = repetitive(30 * 1_000_000);
      const inflated = await inflateAll(deflate(original), { chunkSize: 64 * 1024 });
      expect(inflated.byteLength).toBe(original.byteLength);
      expect(inflated.every(byte => byte === 0x41)).toBe(true);
    });

    it('delivers chunks strictly in order, one at a time', async () => {
      const original = bytesOf(randomBytes(512 * 1024));
      const received: Uint8Array[] = [];
      let inFlight = 0;
      let maxInFlight = 0;
      const inflater = createInflater(
        async chunk => {
          inFlight++;
          maxInFlight = Math.max(maxInFlight, inFlight);
          await new Promise(resolve => setTimeout(resolve, 0));
          received.push(chunk);
          inFlight--;
        },
        { maxBytes: 64 * MIB, method: 'deflate' },
      );
      const compressed = deflate(original);
      for (let offset = 0; offset < compressed.byteLength; offset += 8_192) {
        await inflater.push(compressed.subarray(offset, offset + 8_192));
      }
      await inflater.finish();

      expect(maxInFlight).toBe(1);
      expect(concat(received)).toEqual(original);
    });

    it('has delivered every inflated byte by the time finish resolves', async () => {
      const original = repetitive(4 * MIB);
      let delivered = 0;
      const inflater = createInflater(
        async chunk => {
          await new Promise(resolve => setTimeout(resolve, 0));
          delivered += chunk.byteLength;
        },
        { maxBytes: 64 * MIB, method: 'deflate' },
      );
      await inflater.push(deflate(original));
      await inflater.finish();
      expect(delivered).toBe(original.byteLength);
      expect(inflater.bytesOut).toBe(original.byteLength);
    });

    it('counts bytes in and out', async () => {
      const original = repetitive(100_000);
      const compressed = deflate(original);
      const inflater = createInflater(async () => {}, { maxBytes: MIB, method: 'deflate' });
      await inflater.push(compressed);
      await inflater.finish();
      expect(inflater.bytesIn).toBe(compressed.byteLength);
      expect(inflater.bytesOut).toBe(original.byteLength);
    });

    it('throws ZIP_BOMB as soon as the inflated output passes maxBytes', async () => {
      const compressed = deflate(repetitive(30 * 1_000_000));
      await expect(inflateAll(compressed, { maxBytes: MIB, chunkSize: 4_096 })).rejects.toMatchObject({
        name: 'XlsxError',
        code: 'ZIP_BOMB',
      });
    });

    it('stops inflating once the cap is hit rather than draining the whole stream', async () => {
      const compressed = deflate(repetitive(30 * 1_000_000));
      const inflater = createInflater(async () => {}, { maxBytes: MIB, method: 'deflate' });
      await expect(
        (async () => {
          for (let offset = 0; offset < compressed.byteLength; offset += 4_096) {
            await inflater.push(compressed.subarray(offset, offset + 4_096));
          }
          await inflater.finish();
        })(),
      ).rejects.toMatchObject({ code: 'ZIP_BOMB' });
      expect(inflater.bytesOut).toBeLessThan(30 * 1_000_000);
    });

    it('reports malformed deflate data as ZIP_TRUNCATED with the platform error attached', async () => {
      const garbage = new Uint8Array([0x03, 0x17, 0x99, 0x42, 0x11, 0x7f]);
      let caught: unknown;
      try {
        await inflateAll(garbage);
      } catch (error) {
        caught = error;
      }
      expect(caught).toBeInstanceOf(XlsxError);
      expect((caught as XlsxError).code).toBe('ZIP_TRUNCATED');
      expect((caught as XlsxError).detail?.cause).toBeDefined();
    });

    it('reports a deflate stream that ends early as ZIP_TRUNCATED', async () => {
      const compressed = deflate(repetitive(200_000));
      await expect(inflateAll(compressed.subarray(0, compressed.byteLength - 20))).rejects.toMatchObject({ code: 'ZIP_TRUNCATED' });
    });

    it('keeps the first failure sticky', async () => {
      const inflater = createInflater(async () => {}, { maxBytes: 64 * MIB, method: 'deflate' });
      await inflater.push(new Uint8Array([0x03, 0x17, 0x99, 0x42, 0x11, 0x7f]));
      const first = await inflater.finish().catch((error: unknown) => error);
      const second = await inflater.finish().catch((error: unknown) => error);
      const third = await inflater.push(new Uint8Array([1, 2, 3])).catch((error: unknown) => error);
      expect(first).toBeInstanceOf(XlsxError);
      expect(second).toBe(first);
      expect(third).toBe(first);
    });

    it('propagates an XlsxError thrown by onChunk unchanged', async () => {
      const refusal = new XlsxError('LIMIT_EXCEEDED', 'the consumer said no');
      const inflater = createInflater(
        async () => {
          throw refusal;
        },
        { maxBytes: 64 * MIB, method: 'deflate' },
      );
      await inflater.push(deflate(repetitive(1_000)));
      await expect(inflater.finish()).rejects.toBe(refusal);
    });

    it('aborts without leaving unhandled rejections', async () => {
      const rejections = await captureUnhandledRejections(async () => {
        const inflater = createInflater(async () => {}, { maxBytes: 64 * MIB, method: 'deflate' });
        await inflater.push(deflate(repetitive(2 * MIB)).subarray(0, 64));
        await inflater.abort(new Error('caller went away'));
        await expect(inflater.push(new Uint8Array([1]))).rejects.toMatchObject({ code: 'ABORTED' });
        await expect(inflater.finish()).rejects.toMatchObject({ code: 'ABORTED' });
        await inflater.abort();
      });
      expect(rejections).toEqual([]);
    });

    it('leaves no unhandled rejection when the data is corrupt and nothing awaits push', async () => {
      const rejections = await captureUnhandledRejections(async () => {
        const inflater = createInflater(async () => {}, { maxBytes: 64 * MIB, method: 'deflate' });
        await inflater.push(new Uint8Array([0xff, 0xff, 0xff, 0xff, 0xff])).catch(() => {});
        await inflater.finish().catch(() => {});
      });
      expect(rejections).toEqual([]);
    });

    it('stops delivering chunks after abort', async () => {
      let delivered = 0;
      const inflater = createInflater(
        async chunk => {
          delivered += chunk.byteLength;
        },
        { maxBytes: 64 * MIB, method: 'deflate' },
      );
      await inflater.abort();
      const deliveredAtAbort = delivered;
      await expect(inflater.push(deflate(repetitive(MIB)))).rejects.toMatchObject({ code: 'ABORTED' });
      expect(delivered).toBe(deliveredAtAbort);
    });

    it('is a no-op to abort after a clean finish', async () => {
      const inflater = createInflater(async () => {}, { maxBytes: 64 * MIB, method: 'deflate' });
      await inflater.push(deflate(repetitive(1_000)));
      await inflater.finish();
      await expect(inflater.abort()).resolves.toBeUndefined();
    });
  });

  describe('store', () => {
    it('passes bytes through unchanged', async () => {
      const original = bytesOf(randomBytes(100_000));
      const received: Uint8Array[] = [];
      const inflater = createInflater(
        async chunk => {
          received.push(chunk);
        },
        { maxBytes: MIB, method: 'store' },
      );
      for (let offset = 0; offset < original.byteLength; offset += 7_919) {
        await inflater.push(original.subarray(offset, offset + 7_919));
      }
      await inflater.finish();
      expect(concat(received)).toEqual(original);
      expect(inflater.bytesIn).toBe(original.byteLength);
      expect(inflater.bytesOut).toBe(original.byteLength);
    });

    it('enforces maxBytes', async () => {
      const inflater = createInflater(async () => {}, { maxBytes: 1_000, method: 'store' });
      await inflater.push(new Uint8Array(600));
      await expect(inflater.push(new Uint8Array(600))).rejects.toMatchObject({ code: 'ZIP_BOMB' });
      await expect(inflater.finish()).rejects.toMatchObject({ code: 'ZIP_BOMB' });
    });

    it('keeps an abort sticky', async () => {
      const inflater = createInflater(async () => {}, { maxBytes: 1_000, method: 'store' });
      await inflater.abort();
      await expect(inflater.push(new Uint8Array(1))).rejects.toMatchObject({ code: 'ABORTED' });
    });
  });
});

describe('createInflater: environments without deflate-raw', () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it('reports a DecompressionStream that refuses deflate-raw (Node 20.0-20.11) as UNSUPPORTED_ENVIRONMENT', () => {
    const Native = DecompressionStream;
    vi.stubGlobal(
      'DecompressionStream',
      class extends Native {
        constructor(format: CompressionFormat) {
          if (format === 'deflate-raw') {
            throw new TypeError(`The argument 'format' is invalid. Received '${format}'`);
          }
          super(format);
        }
      },
    );
    let thrown: unknown;
    try {
      createInflater(() => Promise.resolve(), { maxBytes: MIB, method: 'deflate' });
    } catch (error) {
      thrown = error;
    }
    expect(isXlsxError(thrown) && thrown.code).toBe('UNSUPPORTED_ENVIRONMENT');
    expect(isXlsxError(thrown) && thrown.message).toContain('deflate-raw');
    expect(isXlsxError(thrown) && thrown.message).toContain('Node.js 20.12');
    expect(isXlsxError(thrown) && thrown.detail?.cause).toBeInstanceOf(TypeError);
    // Stored entries need no decompressor at all.
    expect(() => createInflater(() => Promise.resolve(), { maxBytes: MIB, method: 'store' })).not.toThrow();
  });

  it('reports a missing DecompressionStream as UNSUPPORTED_ENVIRONMENT', () => {
    vi.stubGlobal('DecompressionStream', undefined);
    expect(() => createInflater(() => Promise.resolve(), { maxBytes: MIB, method: 'deflate' })).toThrow(
      expect.objectContaining({ code: 'UNSUPPORTED_ENVIRONMENT' }) as unknown as Error,
    );
  });
});
