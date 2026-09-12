/**
 * Flat-memory guarantees, measured with `--expose-gc`. Opt-in (slow):
 *   RUN_PERF=1 NODE_OPTIONS=--expose-gc npx vitest run --project corpus test/memory.test.ts
 *
 * `NODE_OPTIONS`, not a bare `node --expose-gc`: vitest runs the suite in a worker it spawns itself, and
 * only `NODE_OPTIONS` reaches that worker. Without it `globalThis.gc` is undefined in the test, which is
 * why asking for `RUN_PERF=1` without it fails loudly here instead of skipping in silence.
 */
import { describe, expect, it } from 'vitest';
import { collectToBytes, createWorkbookWriter, openWorkbook } from '../src/index';
import { READER_READY, WRITER_READY } from './helpers/ready';

const perfRequested = process.env.RUN_PERF === '1';
const canCollectGarbage = typeof globalThis.gc === 'function';

if (perfRequested && !canCollectGarbage) {
  throw new Error(
    'RUN_PERF=1 needs a garbage collector this test can call. Re-run with ' +
      'RUN_PERF=1 NODE_OPTIONS=--expose-gc npx vitest run --project corpus test/memory.test.ts',
  );
}

const enabled = perfRequested && canCollectGarbage;
const MB = 1024 * 1024;
const toMB = (bytes: number): string => `${(bytes / MB).toFixed(1)} MB`;

function heapUsed(): number {
  globalThis.gc?.();
  return process.memoryUsage().heapUsed;
}

/**
 * Peak `heapUsed` while an operation runs. The writer and reader both await their sink or their inflate
 * stream every few chunks, so the event loop turns often enough for a 20 ms poller to see the shape.
 */
async function withHeapPeak<T>(operation: () => Promise<T>): Promise<{ result: T; peakBytes: number }> {
  let peakBytes = process.memoryUsage().heapUsed;
  const timer = setInterval(() => {
    peakBytes = Math.max(peakBytes, process.memoryUsage().heapUsed);
  }, 20);
  timer.unref();
  try {
    const result = await operation();
    return { result, peakBytes: Math.max(peakBytes, process.memoryUsage().heapUsed) };
  } finally {
    clearInterval(timer);
  }
}

function* mixedRows(count: number): Iterable<(string | number | boolean | Date)[]> {
  for (let i = 0; i < count; i++) {
    yield [
      `001Xx${String(i).padStart(12, '0')}`,
      `Account ${i % 5000}`,
      i * 1.5,
      i % 3 === 0,
      new Date(2024, 0, 1 + (i % 365)),
      `note ${i}`,
    ];
  }
}

describe.skipIf(!enabled || !WRITER_READY || !READER_READY)('memory: writer and reader stay flat', () => {
  it('writing 200k rows grows the heap by less than 64 MB (excluding the output buffer)', async () => {
    const before = heapUsed();
    const sink = collectToBytes();
    const workbook = createWorkbookWriter(sink);
    const sheet = workbook.addSheet('Data', { header: ['Id', 'Name', 'Amount', 'Flag', 'When', 'Note'] });
    const { result, peakBytes } = await withHeapPeak(async () => {
      await sheet.writeRows(mixedRows(200_000));
      await sheet.close();
      return workbook.close();
    });
    const after = heapUsed();
    const growth = after - before - sink.bytesWritten;
    console.log(
      `[memory] write 200k x 6: growth ${toMB(growth)} (heap ${toMB(before)} -> ${toMB(after)}, output ${toMB(sink.bytesWritten)}), ` +
        `peak heap during write ${toMB(peakBytes)}, limit 64.0 MB`,
    );
    expect(result.sheets[0]?.rows).toBe(200_001);
    expect(growth).toBeLessThan(64 * MB);
  }, 120_000);

  it('streaming 200k rows without retaining them stays under 150 MB', async () => {
    const sink = collectToBytes();
    const workbook = createWorkbookWriter(sink);
    const sheet = workbook.addSheet('Data');
    await sheet.writeRows(mixedRows(200_000));
    await sheet.close();
    await workbook.close();
    const bytes = sink.result();
    const before = heapUsed();
    const { result: count, peakBytes } = await withHeapPeak(async () => {
      const opened = await openWorkbook(bytes);
      let rowsSeen = 0;
      for await (const row of opened.sheet(0).rows()) {
        rowsSeen += row.length > 0 ? 1 : 0;
      }
      await opened.close();
      return rowsSeen;
    });
    const growth = heapUsed() - before;
    console.log(
      `[memory] read 200k x 6 (${toMB(bytes.byteLength)} file, resident): growth ${toMB(growth)}, ` +
        `peak heap during read ${toMB(peakBytes)} (baseline ${toMB(before)}), limit 150.0 MB`,
    );
    expect(count).toBe(200_000);
    expect(growth).toBeLessThan(150 * MB);
  }, 120_000);
});
