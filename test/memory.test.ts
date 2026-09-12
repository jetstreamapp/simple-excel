/**
 * Flat-memory guarantees, measured with `--expose-gc`. Opt-in (slow):
 *   RUN_PERF=1 node --expose-gc node_modules/vitest/vitest.mjs run --project corpus test/memory.test.ts
 */
import { describe, expect, it } from 'vitest';
import { collectToBytes, createWorkbookWriter, openWorkbook } from '../src/index';
import { READER_READY, WRITER_READY } from './helpers/ready';

const enabled = process.env.RUN_PERF === '1' && typeof globalThis.gc === 'function';

function heapUsed(): number {
  globalThis.gc?.();
  return process.memoryUsage().heapUsed;
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
    await sheet.writeRows(mixedRows(200_000));
    await sheet.close();
    const result = await workbook.close();
    const after = heapUsed();
    expect(result.sheets[0]?.rows).toBe(200_001);
    expect(after - before - sink.bytesWritten).toBeLessThan(64 * 1024 * 1024);
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
    const opened = await openWorkbook(bytes);
    let count = 0;
    for await (const row of opened.sheet(0).rows()) {
      count += row.length > 0 ? 1 : 0;
    }
    await opened.close();
    expect(count).toBe(200_000);
    expect(heapUsed() - before).toBeLessThan(150 * 1024 * 1024);
  }, 120_000);
});
