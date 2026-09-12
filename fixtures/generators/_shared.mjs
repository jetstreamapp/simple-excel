import { mkdirSync, writeFileSync } from 'node:fs';
import { dirname } from 'node:path';
import { canonicalWorkbook, isTyped, isTemporal, truncateForExcel } from '../canonical/canonical.mjs';

export { canonicalWorkbook, isTyped, isTemporal, truncateForExcel };

/** Every golden writer applies Jetstream's over-limit policy so the file is writable by Excel-strict libraries. */
export function truncateRows(rows) {
  return rows.map(row => row.map(truncateForExcel));
}

export function ensureDir(outPath) {
  mkdirSync(dirname(outPath), { recursive: true });
}

/**
 * Sidecar next to a golden recording which structural features the generator managed to write, so the
 * compatibility matrix can distinguish "feature absent from the file" from "reader dropped it".
 */
export function writeFeatureSidecar(outPath, generator, applied, skipped) {
  const sidecar = outPath.replace(/\.xlsx$/, '.features.json');
  writeFileSync(sidecar, JSON.stringify({ generator, applied, skipped }, null, 2) + '\n');
  return sidecar;
}

/** Try a feature; record success or the failure reason without aborting the whole golden. */
const SKIP = new Set((process.env.GOLDEN_SKIP ?? '').split(',').filter(Boolean));

export async function attempt(tracker, name, fn) {
  if (SKIP.has(name) || SKIP.has('*')) {
    tracker.skipped.push({ feature: name, reason: 'skipped via GOLDEN_SKIP' });
    return;
  }
  try {
    await fn();
    tracker.applied.push(name);
  } catch (error) {
    tracker.skipped.push({ feature: name, reason: error?.message ?? String(error) });
  }
}

export function newTracker() {
  return { applied: [], skipped: [] };
}

/** Runs a generator module directly: `node canonical-<gen>.mjs [outPath]`. */
export async function runStandalone(importMetaUrl, generate, defaultOut) {
  if (process.argv[1] && new URL(importMetaUrl).pathname === process.argv[1]) {
    const outPath = process.argv[2] ?? defaultOut;
    ensureDir(outPath);
    const result = await generate(outPath);
    console.log(`wrote ${outPath}`, result ? JSON.stringify(result) : '');
  }
}
