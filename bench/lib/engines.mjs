import { readdirSync } from 'node:fs';
import { pathToFileURL } from 'node:url';

/**
 * Engine registry. Every `engines/<name>.mjs` exports the engine interface documented in the README
 * (`name`, `version`, `supportsStreaming`, `write`, `readTyped`, `readRaw`). An adapter may instead
 * export `load()` returning that interface (used by `simple-excel.mjs`, which imports the built bundle
 * lazily) and/or `skipReason` (string) when it cannot run in this environment. Files starting with `_`
 * hold shared adapter code and are not engines.
 */
const ENGINES_DIR = new URL('../engines/', import.meta.url);

export const OPS = ['write', 'read-typed', 'read-raw'];

export function listEngineNames() {
  return readdirSync(ENGINES_DIR)
    .filter(file => file.endsWith('.mjs') && !file.startsWith('_'))
    .map(file => file.replace(/\.mjs$/, ''))
    .sort();
}

export async function loadEngine(name) {
  if (!listEngineNames().includes(name)) {
    throw new Error(`Unknown engine "${name}". Known: ${listEngineNames().join(', ')}`);
  }
  const module = await import(pathToFileURL(new URL(`${name}.mjs`, ENGINES_DIR).pathname).href);
  const engine = typeof module.load === 'function' ? await module.load() : module;
  return { name, ...engine, skipReason: module.skipReason ?? engine.skipReason };
}

/** Maps an op name to the engine method that implements it (null when the engine marks it unsupported). */
export function getOpImplementation(engine, op) {
  const method = { write: engine.write, 'read-typed': engine.readTyped, 'read-raw': engine.readRaw }[op];
  if (method === undefined) {
    throw new Error(`Unknown op "${op}". Known: ${OPS.join(', ')}`);
  }
  return typeof method === 'function' ? method : null;
}
