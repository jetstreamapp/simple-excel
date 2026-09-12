import { createEngine } from './_simple-excel.mjs';

/**
 * `@jetstreamapp/simple-excel` compressing through the node entry's `nodeDeflater(1)` instead of the
 * platform `CompressionStream`. Identical in every other respect to `simple-excel`, so the pair measures
 * what the compressor costs; only a Node caller can choose it.
 */
export const name = 'simple-excel-zlib';

export async function load() {
  return createEngine({ useZlibDeflater: true });
}
