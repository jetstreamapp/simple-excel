import { createEngine } from './_simple-excel.mjs';

/**
 * `@jetstreamapp/simple-excel` on the path a browser takes: streaming writer over the Node file sink,
 * compression through the platform `CompressionStream('deflate-raw')`, rows pulled from the iterable by
 * `writeRows`, header row and `rowCount` declared up front. Reads go through `openWorkbook` over bytes.
 *
 * `simple-excel-zlib` is the same adapter with the node entry's `nodeDeflater(1)`.
 */
export const name = 'simple-excel';

export async function load() {
  return createEngine();
}
