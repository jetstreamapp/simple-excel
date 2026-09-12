import { pathToFileURL } from 'node:url';
import { isAbsolute, resolve } from 'node:path';

/**
 * Placeholder for the in-house engine. Point `XLSX_ENGINE_OURS` at a module exporting the same
 * interface as the other adapters (`write`, `readTyped`, `readRaw`, optional `version`,
 * `supportsStreaming`, `readInput`) and it runs under the name `ours`; unset, the cell is skipped.
 */
export const name = 'ours';

const modulePath = process.env.XLSX_ENGINE_OURS;

export const skipReason = modulePath ? undefined : 'XLSX_ENGINE_OURS is not set (path to a module implementing the engine interface)';

export async function load() {
  if (!modulePath) {
    return { name, version: 'n/a', supportsStreaming: false, skipReason };
  }
  const absolutePath = isAbsolute(modulePath) ? modulePath : resolve(process.cwd(), modulePath);
  const implementation = await import(pathToFileURL(absolutePath).href);
  return {
    version: implementation.version ?? 'unversioned',
    supportsStreaming: implementation.supportsStreaming ?? false,
    readInput: implementation.readInput ?? 'path',
    write: implementation.write,
    readTyped: implementation.readTyped,
    readRaw: implementation.readRaw,
    name,
  };
}
