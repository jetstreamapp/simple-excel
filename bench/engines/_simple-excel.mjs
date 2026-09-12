import { existsSync, readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

/**
 * Shared adapter body for the two `@jetstreamapp/simple-excel` engines: `simple-excel` (the platform
 * `CompressionStream('deflate-raw')`, which is what a browser actually runs) and `simple-excel-zlib`
 * (the node entry's `nodeDeflater(1)`). Everything else about them is identical, so the difference the
 * table shows is exactly the cost of the compressor.
 *
 * Reads the built bundles, not `src/`: `npm run build` has to have run. The library truncates cells over
 * 32,767 characters itself, with the same rule and suffix as `lib/excel-limits.mjs`, so the adapter does
 * not pre-truncate rows — the bytes are the same and the cost stays inside the engine where it belongs.
 */
const CORE_ENTRY = new URL('../../dist/esm/index.mjs', import.meta.url);
const NODE_ENTRY = new URL('../../dist/esm/node.mjs', import.meta.url);
const PACKAGE_JSON = new URL('../../package.json', import.meta.url);

const BUILD_HINT = 'dist/ is missing — run `npm run build` before `npm run bench`';

const SHEET_NAME = 'Sheet1';

function distIsBuilt() {
  return existsSync(fileURLToPath(CORE_ENTRY)) && existsSync(fileURLToPath(NODE_ENTRY));
}

/** Wraps a sink so the harness can see when the first compressed chunk reaches the destination. */
function probeSink(sink, onFirstByte) {
  let sawFirstChunk = false;
  return {
    write(chunk) {
      if (!sawFirstChunk) {
        sawFirstChunk = true;
        onFirstByte();
      }
      return sink.write(chunk);
    },
    close: () => sink.close(),
    abort: reason => sink.abort(reason),
  };
}

/**
 * Builds the engine interface. With `useZlibDeflater` the writer compresses through
 * `nodeDeflater(1)`; without it (the default) through the platform `CompressionStream`.
 */
export async function createEngine({ useZlibDeflater = false } = {}) {
  if (!distIsBuilt()) {
    return { version: 'not built', supportsStreaming: true, readInput: 'bytes', skipReason: BUILD_HINT };
  }
  const { createWorkbookWriter, openWorkbook } = await import(CORE_ENTRY.href);
  const { nodeDeflater, toFile } = await import(NODE_ENTRY.href);
  const version = JSON.parse(readFileSync(fileURLToPath(PACKAGE_JSON), 'utf8')).version;
  const writeOptions = useZlibDeflater ? { deflater: nodeDeflater(1) } : {};

  return {
    version,
    supportsStreaming: true,
    /** `openWorkbook` takes bytes in the browser (a File/ArrayBuffer), so the harness hands it bytes. */
    readInput: 'bytes',

    async write(rowIterable, { columns, rowCount, outPath }) {
      const startedAt = performance.now();
      let firstByteMs = null;
      const sink = probeSink(toFile(outPath), () => {
        firstByteMs = performance.now() - startedAt;
      });
      const workbook = createWorkbookWriter(sink, writeOptions);
      try {
        const sheet = workbook.addSheet(SHEET_NAME, { header: columns, rowCount });
        await sheet.writeRows(rowIterable);
        await sheet.close();
        const result = await workbook.close();
        return { bytes: result.bytes, firstByteMs };
      } catch (error) {
        await workbook.abort(error);
        throw error;
      }
    },

    /** Jetstream's read shape: every record as an object, dates as Dates, the whole array materialized. */
    async readTyped(bytes) {
      const workbook = await openWorkbook(bytes);
      const { rows, headers } = await workbook.sheet(0).toObjects();
      await workbook.close();
      return { rows: rows.length, cells: rows.length * headers.length };
    },

    /** Array mode, streamed: rows are counted and dropped, so nothing but one row is ever resident. */
    async readRaw(bytes) {
      const workbook = await openWorkbook(bytes);
      let rowsSeen = 0;
      let columnCount = 0;
      for await (const row of workbook.sheet(0).rows()) {
        if (rowsSeen === 0) {
          columnCount = row.length;
        }
        rowsSeen++;
      }
      await workbook.close();
      const dataRows = Math.max(0, rowsSeen - 1);
      return { rows: dataRows, cells: dataRows * columnCount };
    },
  };
}
