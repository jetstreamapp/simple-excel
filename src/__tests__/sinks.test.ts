import { Buffer } from 'node:buffer';
import { randomBytes } from 'node:crypto';
import { describe, expect, it, vi } from 'vitest';
import { isXlsxError, XlsxError } from '../errors';
import { collectToBlob, collectToBytes, fromWritableStream, guardSink, toWritableStream } from '../sinks';
import type { ByteSink } from '../types';

const MIB = 1024 * 1024;

function randomChunkSequence(chunkCount: number, maxChunkSize: number): Uint8Array[] {
  const chunks: Uint8Array[] = [];
  for (let i = 0; i < chunkCount; i++) {
    const size = 1 + Math.floor(Math.random() * maxChunkSize);
    const random = randomBytes(size);
    chunks.push(new Uint8Array(random.buffer, random.byteOffset, random.byteLength));
  }
  return chunks;
}

function reference(chunks: readonly Uint8Array[]): Buffer {
  return Buffer.concat(chunks.map(chunk => Buffer.from(chunk)));
}

const delay = (milliseconds: number): Promise<void> => new Promise(resolve => setTimeout(resolve, milliseconds));

/** True when the promise has not settled after a macrotask; the promise is marked handled so a rejection is not global. */
async function isPending(promise: Promise<unknown>): Promise<boolean> {
  let settled = false;
  const watched = promise.then(
    () => {
      settled = true;
    },
    () => {
      settled = true;
    },
  );
  await Promise.race([watched, delay(10)]);
  return !settled;
}

interface RecordingSink extends ByteSink {
  readonly calls: string[];
  readonly overlaps: number;
  release(): void;
}

/** A sink whose writes only complete when `release()` is called, and that records call order and overlap. */
function createGatedSink(): RecordingSink {
  const calls: string[] = [];
  let overlaps = 0;
  let inFlight = 0;
  let releaseGate: () => void = () => undefined;
  let gate = new Promise<void>(resolve => {
    releaseGate = resolve;
  });

  async function run(label: string): Promise<void> {
    calls.push(label);
    if (inFlight > 0) {
      overlaps++;
    }
    inFlight++;
    await gate;
    inFlight--;
  }

  return {
    calls,
    get overlaps(): number {
      return overlaps;
    },
    release(): void {
      releaseGate();
      gate = Promise.resolve();
    },
    write(chunk: Uint8Array): Promise<void> {
      return run(`write:${chunk.byteLength}`);
    },
    close(): Promise<void> {
      return run('close');
    },
    abort(): Promise<void> {
      return run('abort');
    },
  };
}

describe('collectToBytes', () => {
  it('reassembles a random chunk sequence exactly', async () => {
    const chunks = randomChunkSequence(64, 8 * 1024);
    const sink = collectToBytes();
    for (const chunk of chunks) {
      await sink.write(chunk);
    }
    await sink.close();

    const expected = reference(chunks);
    expect(sink.bytesWritten).toBe(expected.byteLength);
    expect(Buffer.compare(Buffer.from(sink.result()), expected)).toBe(0);
  });

  it('hands back a lone chunk without copying it', async () => {
    const chunk = new Uint8Array([1, 2, 3]);
    const sink = collectToBytes();
    await sink.write(chunk);
    await sink.close();
    expect(sink.result()).toBe(chunk);
  });

  it('collects nothing into an empty array', async () => {
    const sink = collectToBytes();
    await sink.close();
    expect(sink.result()).toEqual(new Uint8Array(0));
  });

  it('throws LIMIT_EXCEEDED past maxBytes without recording the chunk', async () => {
    const sink = collectToBytes({ maxBytes: 1024 });
    await sink.write(new Uint8Array(600));
    await expect(sink.write(new Uint8Array(600))).rejects.toMatchObject({ code: 'LIMIT_EXCEEDED' });
    expect(sink.bytesWritten).toBe(600);
  });

  it('refuses result() before close and writes after it', async () => {
    const sink = collectToBytes();
    expect(() => sink.result()).toThrow(XlsxError);
    expect(() => sink.result()).toThrowError(expect.objectContaining({ code: 'WRITER_STATE' }));
    await sink.close();
    await expect(sink.write(new Uint8Array(1))).rejects.toMatchObject({ code: 'WRITER_STATE' });
  });

  it('wraps a non-XlsxError abort reason and keeps an XlsxError as-is', async () => {
    const plain = collectToBytes();
    const cause = new Error('the tab was closed');
    await plain.abort(cause);
    try {
      plain.result();
      expect.unreachable('result() must throw after an abort');
    } catch (error) {
      expect(isXlsxError(error)).toBe(true);
      expect(error).toMatchObject({ code: 'ABORTED', detail: { cause } });
    }

    const classified = collectToBytes();
    const limit = new XlsxError('LIMIT_EXCEEDED', 'too big');
    await classified.abort(limit);
    expect(() => classified.result()).toThrow(limit);
    await expect(classified.write(new Uint8Array(1))).rejects.toThrow(limit);
  });
});

describe('collectToBlob', () => {
  it('reassembles a random chunk sequence exactly', async () => {
    const chunks = randomChunkSequence(64, 8 * 1024);
    const sink = collectToBlob();
    for (const chunk of chunks) {
      await sink.write(chunk);
    }
    await sink.close();

    const blob = await sink.result();
    const expected = reference(chunks);
    expect(blob.size).toBe(expected.byteLength);
    expect(blob.type).toBe('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    expect(Buffer.compare(Buffer.from(await blob.arrayBuffer()), expected)).toBe(0);
    expect(sink.bytesWritten).toBe(expected.byteLength);
  });

  it('folds pending chunks into sub-Blobs every 32 MiB', async () => {
    const sink = collectToBlob('application/octet-stream');
    const chunkCount = 70;
    for (let i = 0; i < chunkCount; i++) {
      await sink.write(new Uint8Array(MIB).fill(i));
    }
    await sink.close();

    const blob = await sink.result();
    expect(blob.size).toBe(chunkCount * MIB);
    expect(blob.type).toBe('application/octet-stream');

    // Inside the second fold, and across the 32 MiB fold boundary: the seam must be invisible.
    const sample = new Uint8Array(await blob.slice(33 * MIB, 33 * MIB + 4).arrayBuffer());
    expect([...sample]).toEqual([33, 33, 33, 33]);
    const seam = new Uint8Array(await blob.slice(32 * MIB - 2, 32 * MIB + 2).arrayBuffer());
    expect([...seam]).toEqual([31, 31, 32, 32]);
  });

  it('resolves a result() requested before close', async () => {
    const sink = collectToBlob();
    const pending = sink.result();
    expect(await isPending(pending)).toBe(true);
    await sink.write(new Uint8Array([7, 7]));
    await sink.close();
    expect((await pending).size).toBe(2);
  });

  it('rejects (not hangs) after an abort, with the XlsxError it was aborted with', async () => {
    const failure = new XlsxError('WRITER_STATE', 'the writer failed');
    const blobSink = collectToBlob();
    const bytesSink = collectToBytes();
    await blobSink.abort(failure);
    await bytesSink.abort(failure);
    const pending = blobSink.result();
    expect(await isPending(pending)).toBe(false);
    await expect(pending).rejects.toBe(failure);
    expect(() => bytesSink.result()).toThrow(failure);
  });

  it('rejects the result after an abort and refuses later writes', async () => {
    const sink = collectToBlob();
    await sink.write(new Uint8Array(4));
    await sink.abort(new Error('navigated away'));

    await expect(sink.result()).rejects.toMatchObject({ code: 'ABORTED' });
    await expect(sink.write(new Uint8Array(1))).rejects.toMatchObject({ code: 'ABORTED' });
    await expect(sink.close()).rejects.toMatchObject({ code: 'ABORTED' });
    expect(sink.bytesWritten).toBe(4);
  });

  it('reports an environment without Blob instead of failing later', () => {
    vi.stubGlobal('Blob', undefined);
    try {
      expect(() => collectToBlob()).toThrowError(expect.objectContaining({ code: 'UNSUPPORTED_ENVIRONMENT' }));
    } finally {
      vi.unstubAllGlobals();
    }
  });

  it('ignores a second close and refuses writes after it', async () => {
    const sink = collectToBlob();
    await sink.write(new Uint8Array(2));
    await sink.close();
    await sink.close();
    await expect(sink.write(new Uint8Array(1))).rejects.toMatchObject({ code: 'WRITER_STATE' });
    expect((await sink.result()).size).toBe(2);
  });
});

describe('fromWritableStream', () => {
  it('writes every chunk through and releases the lock on close', async () => {
    const received: Uint8Array[] = [];
    const stream = new WritableStream<Uint8Array>({
      write(chunk): void {
        received.push(chunk);
      },
    });
    const sink = fromWritableStream(stream);
    const chunks = randomChunkSequence(8, 1024);
    for (const chunk of chunks) {
      await sink.write(chunk);
    }
    await sink.close();

    expect(Buffer.compare(reference(received), reference(chunks))).toBe(0);
    expect(stream.locked).toBe(false);
  });

  it('applies back-pressure from a slow underlying sink', async () => {
    let releaseWrite: () => void = () => undefined;
    const started: number[] = [];
    const stream = new WritableStream<Uint8Array>(
      {
        async write(chunk): Promise<void> {
          started.push(chunk.byteLength);
          await new Promise<void>(resolve => {
            releaseWrite = resolve;
          });
        },
      },
      new CountQueuingStrategy({ highWaterMark: 1 }),
    );
    const sink = fromWritableStream(stream);

    const first = sink.write(new Uint8Array(1));
    const second = sink.write(new Uint8Array(2));
    expect(await isPending(second)).toBe(true);
    expect(started).toEqual([1]);

    releaseWrite();
    await first;
    expect(await isPending(second)).toBe(true);
    releaseWrite();
    await second;
    expect(started).toEqual([1, 2]);
    await sink.abort('done');
  });

  it('surfaces a failure from the underlying stream on every later call', async () => {
    const failure = new Error('quota exceeded');
    const stream = new WritableStream<Uint8Array>({
      write(): void {
        throw failure;
      },
    });
    const sink = fromWritableStream(stream);

    await expect(sink.write(new Uint8Array(1))).rejects.toThrow(failure);
    await expect(sink.write(new Uint8Array(1))).rejects.toThrow(failure);
    // A broken stream answers `close()` with its own state error; `guardSink` is what makes one reason stick.
    await expect(sink.close()).rejects.toThrow();
    await expect(guardSink(fromWritableStream(stream)).write(new Uint8Array(1))).rejects.toMatchObject({ code: 'ABORTED' });
  });

  it('forwards the abort reason and releases the lock', async () => {
    const reasons: unknown[] = [];
    const stream = new WritableStream<Uint8Array>({
      abort(reason): void {
        reasons.push(reason);
      },
    });
    const sink = fromWritableStream(stream);
    await sink.abort('user cancelled');
    expect(reasons).toEqual(['user cancelled']);
    expect(stream.locked).toBe(false);
  });
});

describe('toWritableStream', () => {
  it('round trips through a TransformStream pipeline', async () => {
    const chunks = randomChunkSequence(16, 4 * 1024);
    const collector = collectToBytes();
    const passThrough = new TransformStream<Uint8Array, Uint8Array>();

    const source = new ReadableStream<Uint8Array>({
      start(controller): void {
        for (const chunk of chunks) {
          controller.enqueue(chunk);
        }
        controller.close();
      },
    });

    await source.pipeThrough(passThrough).pipeTo(toWritableStream(collector));

    expect(Buffer.compare(Buffer.from(collector.result()), reference(chunks))).toBe(0);
  });

  it('aborts the sink when the pipeline fails', async () => {
    const collector = collectToBytes();
    const failure = new Error('source exploded');
    const source = new ReadableStream<Uint8Array>({
      start(controller): void {
        controller.error(failure);
      },
    });

    await expect(source.pipeTo(toWritableStream(collector))).rejects.toThrow(failure);
    expect(() => collector.result()).toThrowError(expect.objectContaining({ code: 'ABORTED', detail: { cause: failure } }));
  });
});

describe('guardSink', () => {
  it('repeats the original failure on every later call', async () => {
    const failure = new XlsxError('ENTRY_TOO_LARGE', 'no room left');
    const guarded = guardSink({
      write(): Promise<void> {
        return Promise.reject(failure);
      },
      close(): Promise<void> {
        return Promise.resolve();
      },
      abort(): Promise<void> {
        return Promise.resolve();
      },
    });

    await expect(guarded.write(new Uint8Array(1))).rejects.toThrow(failure);
    await expect(guarded.write(new Uint8Array(1))).rejects.toThrow(failure);
    await expect(guarded.close()).rejects.toThrow(failure);
  });

  it('classifies a non-XlsxError failure as ABORTED with the original as the cause', async () => {
    const cause = new Error('quota exceeded');
    const guarded = guardSink({
      write(): Promise<void> {
        return Promise.reject(cause);
      },
      close(): Promise<void> {
        return Promise.resolve();
      },
      abort(): Promise<void> {
        return Promise.resolve();
      },
    });

    await expect(guarded.write(new Uint8Array(1))).rejects.toMatchObject({ code: 'ABORTED', detail: { cause } });
    await expect(guarded.close()).rejects.toMatchObject({ code: 'ABORTED', detail: { cause } });
  });

  it('makes every call after an abort reject with the abort reason', async () => {
    const sink = createGatedSink();
    sink.release();
    const guarded = guardSink(sink);

    await guarded.abort(new Error('cancelled'));
    await expect(guarded.write(new Uint8Array(1))).rejects.toMatchObject({ code: 'ABORTED' });
    await guarded.abort(new Error('cancelled twice'));
    expect(sink.calls).toEqual(['abort']);
  });

  it('treats a second close and an abort after close as no-ops', async () => {
    const sink = createGatedSink();
    sink.release();
    const guarded = guardSink(sink);

    await guarded.write(new Uint8Array(3));
    await guarded.close();
    await guarded.close();
    await guarded.abort(new Error('too late'));

    expect(sink.calls).toEqual(['write:3', 'close']);
    await expect(guarded.write(new Uint8Array(1))).rejects.toMatchObject({ code: 'WRITER_STATE' });
  });

  it('serializes overlapping calls in the order they were issued', async () => {
    const sink = createGatedSink();
    const guarded = guardSink(sink);

    const writes = [guarded.write(new Uint8Array(1)), guarded.write(new Uint8Array(2)), guarded.write(new Uint8Array(3))];
    const closed = guarded.close();
    // Every call is queued behind the chain's tail, so only the first one has reached the sink.
    await delay(0);
    expect(sink.calls).toEqual(['write:1']);

    sink.release();
    await Promise.all([...writes, closed]);

    expect(sink.calls).toEqual(['write:1', 'write:2', 'write:3', 'close']);
    expect(sink.overlaps).toBe(0);
  });
});
