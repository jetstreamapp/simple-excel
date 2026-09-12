#!/usr/bin/env node
/**
 * Bundle-size gate for the browser entry (`research/11-build-plan.md` §1: "core bundle <= 40 KB min+brotli").
 * The build already minifies, so this only compresses `dist/esm/index.mjs` at brotli quality 11 — the same
 * number `brotli -c -q 11 dist/esm/index.mjs | wc -c` prints — and fails when it passes the budget.
 *
 *   npm run size            # check every entry, fail if the core entry is over budget
 *   npm run size -- --json  # machine-readable, for the docs
 *
 * The node entry is reported for information only: it is never shipped to a browser.
 */
import { constants, brotliCompressSync, gzipSync } from 'node:zlib';
import { existsSync, readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

const CORE_BUDGET_BYTES = 40 * 1024;

const ENTRIES = [
  { name: 'core (dist/esm/index.mjs)', path: '../dist/esm/index.mjs', budgetBytes: CORE_BUDGET_BYTES },
  { name: 'node (dist/esm/node.mjs)', path: '../dist/esm/node.mjs', budgetBytes: null },
];

const asJson = process.argv.includes('--json');
const formatKB = bytes => `${(bytes / 1024).toFixed(1)} KB`;

const measurements = ENTRIES.map(entry => {
  const file = fileURLToPath(new URL(entry.path, import.meta.url));
  if (!existsSync(file)) {
    return { ...entry, missing: true };
  }
  const source = readFileSync(file);
  return {
    ...entry,
    missing: false,
    minifiedBytes: source.byteLength,
    gzipBytes: gzipSync(source, { level: 9 }).byteLength,
    brotliBytes: brotliCompressSync(source, { params: { [constants.BROTLI_PARAM_QUALITY]: 11 } }).byteLength,
  };
});

const missing = measurements.filter(measurement => measurement.missing);
if (missing.length > 0) {
  console.error(`bundle size: ${missing.map(measurement => measurement.path).join(', ')} not built — run \`npm run build\` first`);
  process.exit(1);
}

if (asJson) {
  console.log(JSON.stringify(measurements, null, 2));
} else {
  for (const { name, minifiedBytes, gzipBytes, brotliBytes, budgetBytes } of measurements) {
    const budget = budgetBytes === null ? '' : ` (budget ${formatKB(budgetBytes)} brotli)`;
    console.log(`${name}: ${formatKB(minifiedBytes)} minified, ${formatKB(gzipBytes)} gzip, ${formatKB(brotliBytes)} brotli${budget}`);
  }
}

const overBudget = measurements.filter(
  measurement => measurement.budgetBytes !== null && measurement.brotliBytes > measurement.budgetBytes,
);
if (overBudget.length > 0) {
  for (const { name, brotliBytes, budgetBytes } of overBudget) {
    console.error(`bundle size FAILED: ${name} is ${formatKB(brotliBytes)} brotli, over the ${formatKB(budgetBytes)} budget`);
  }
  process.exit(1);
}
