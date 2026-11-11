/**
 * Summary of a corpus-audit results file: status counts, rejection codes, unexplained diff categories with example
 * files, slowest files, highest memory, throughput.
 *
 *   tsx scripts/corpus-audit/report.ts [tmp/corpus-audit/results.jsonl] [--json summary.json]
 */
import { readFileSync, writeFileSync } from 'node:fs';
import { basename } from 'node:path';
import { parseArgs } from 'node:util';

const { values: options, positionals } = parseArgs({
  allowPositionals: true,
  options: { json: { type: 'string' }, baseline: { type: 'string' } },
});
const path = positionals[0] ?? 'tmp/corpus-audit/results.jsonl';

function load(file: string): Record<string, any>[] {
  return readFileSync(file, 'utf8')
    .split('\n')
    .filter(line => line.trim())
    .map(line => JSON.parse(line) as Record<string, any>);
}
const entries = load(path);

function bump(map: Map<string, { count: number; examples: string[] }>, key: string, file: string): void {
  const slot = map.get(key) ?? { count: 0, examples: [] };
  slot.count++;
  if (slot.examples.length < 5) {
    slot.examples.push(basename(file));
  }
  map.set(key, slot);
}

const status = new Map<string, { count: number; examples: string[] }>();
const rejections = new Map<string, { count: number; examples: string[] }>();
const unclassified = new Map<string, { count: number; examples: string[] }>();
const parityCategories = new Map<string, { count: number; examples: string[] }>();
const objectProblems = new Map<string, { count: number; examples: string[] }>();
const roundtripCategories = new Map<string, { count: number; examples: string[] }>();
const known = new Map<string, number>();
const sheetjsFailures = new Map<string, { count: number; examples: string[] }>();
let totalCells = 0;
let totalMs = 0;

for (const entry of entries) {
  bump(status, entry.status, entry.file);
  if (entry.status === 'rejected') {
    bump(rejections, entry.ours.error.code, entry.file);
  }
  if (entry.status === 'unclassified-error') {
    bump(unclassified, `${entry.ours.error.name}: ${entry.ours.error.message.slice(0, 80)}`, entry.file);
  }
  if (entry.status === 'timeout' || entry.status === 'crash' || entry.status === 'harness-error') {
    bump(unclassified, `${entry.status} in ${entry.phase}`, entry.file);
  }
  if (entry.sheetjs && !entry.sheetjs.ok) {
    bump(sheetjsFailures, `${entry.ours?.ok ? 'we read it' : 'we fail too'}: ${String(entry.sheetjs.error).slice(0, 60)}`, entry.file);
  }
  for (const category of Object.keys(entry.parity?.byCategory ?? {})) {
    bump(parityCategories, category, entry.file);
  }
  for (const [reason, count] of Object.entries(entry.parity?.known ?? {})) {
    known.set(reason, (known.get(reason) ?? 0) + (count as number));
  }
  if (entry.objects && !entry.objects.pass) {
    bump(
      objectProblems,
      entry.objects.error ? `error ${entry.objects.error.code}` : entry.objects.headersMatch ? 'values' : 'headers',
      entry.file,
    );
  }
  if (entry.roundtrip?.error) {
    bump(roundtripCategories, `error ${entry.roundtrip.error.code}: ${entry.roundtrip.error.message.slice(0, 60)}`, entry.file);
  }
  for (const category of Object.keys(entry.roundtrip?.byCategory ?? {})) {
    bump(roundtripCategories, category, entry.file);
  }
  if (entry.ours?.ok) {
    totalCells += entry.ours.cells;
    totalMs += entry.ours.ms;
  }
}

function table(title: string, map: Map<string, { count: number; examples: string[] }>): void {
  if (map.size === 0) {
    return;
  }
  console.log(`\n## ${title}`);
  for (const [key, { count, examples }] of [...map].sort((a, b) => b[1].count - a[1].count)) {
    console.log(`  ${String(count).padStart(6)}  ${key}  e.g. ${examples.slice(0, 3).join(', ')}`);
  }
}

const read = entries.filter(entry => entry.ours?.ok);
console.log(`# Corpus audit: ${entries.length} files (${path})`);
console.log(
  `read ${read.length}, ${totalCells.toLocaleString()} cells in ${(totalMs / 1000).toFixed(1)}s of reader time (${Math.round(totalCells / (totalMs / 1000)).toLocaleString()} cells/s incl. open)`,
);
table('Status', status);
table('Rejected (classified XlsxError)', rejections);
table('Unclassified errors, timeouts, crashes', unclassified);
table('SheetJS failures', sheetjsFailures);
table('Parity vs SheetJS: unexplained categories (files)', parityCategories);
console.log(`\n## Known SheetJS divergences (cells)\n  ${JSON.stringify(Object.fromEntries(known))}`);
table('Objects (toObjects vs sheet_to_json) failures (files)', objectProblems);
table('Round trip categories (files)', roundtripCategories);

console.log('\n## Slowest reads (stream pass)');
for (const entry of [...read].sort((a, b) => b.ours.ms - a.ours.ms).slice(0, 10)) {
  console.log(
    `  ${String(entry.ours.ms).padStart(7)} ms  ${(entry.bytes / 1048576).toFixed(1)} MB  ${entry.ours.cells.toLocaleString()} cells  sheetjs ${entry.sheetjs?.ms ?? '-'} ms  ${basename(entry.file)}`,
  );
}
console.log('\n## Highest peak RSS growth during the stream pass');
for (const entry of [...read]
  .sort((a, b) => b.ours.peakRssMb - b.ours.baselineRssMb - (a.ours.peakRssMb - a.ours.baselineRssMb))
  .slice(0, 10)) {
  console.log(
    `  +${entry.ours.peakRssMb - entry.ours.baselineRssMb} MB (peak ${entry.ours.peakRssMb})  ${(entry.bytes / 1048576).toFixed(1)} MB file  ${entry.ours.cells.toLocaleString()} cells  sheetjs peak ${entry.sheetjs?.peakRssMb ?? '-'} MB  ${basename(entry.file)}`,
  );
}

/**
 * `--baseline earlier.jsonl`: what changed per file since that run - status (a new LIMIT_EXCEEDED on a file that read
 * before is a limit set too tight), cell count, and the reader's time and peak-RSS growth in aggregate.
 */
if (options.baseline) {
  const before = new Map(load(options.baseline).map(entry => [entry.file as string, entry]));
  const transitions = new Map<string, { count: number; examples: string[] }>();
  let cellsChanged = 0;
  let timeBefore = 0;
  let timeAfter = 0;
  const rssGrowth: { file: string; before: number; after: number }[] = [];
  for (const entry of entries) {
    const old = before.get(entry.file);
    if (!old) {
      continue;
    }
    const describe = (value: Record<string, any>): string =>
      value.status === 'rejected' || value.status === 'unclassified-error' ? `${value.status}:${value.ours?.error?.code}` : value.status;
    if (describe(old) !== describe(entry)) {
      bump(transitions, `${describe(old)} -> ${describe(entry)}`, entry.file);
    }
    if (old.ours?.ok && entry.ours?.ok) {
      if (old.ours.cells !== entry.ours.cells) {
        cellsChanged++;
      }
      timeBefore += old.ours.ms;
      timeAfter += entry.ours.ms;
      rssGrowth.push({
        file: entry.file,
        before: old.ours.peakRssMb - old.ours.baselineRssMb,
        after: entry.ours.peakRssMb - entry.ours.baselineRssMb,
      });
    }
  }
  console.log(`\n# Against baseline ${options.baseline}`);
  console.log(transitions.size === 0 ? '  no status changes' : '');
  table('Status changes', transitions);
  console.log(`  files whose cell count changed: ${cellsChanged}`);
  console.log(`  reader time on files read both times: ${(timeBefore / 1000).toFixed(1)}s -> ${(timeAfter / 1000).toFixed(1)}s`);
  const worst = rssGrowth.sort((a, b) => b.after - b.before - (a.after - a.before)).slice(0, 5);
  console.log('  largest RSS-growth increases:');
  for (const item of worst) {
    console.log(`    +${item.before} MB -> +${item.after} MB  ${basename(item.file)}`);
  }
}

if (options.json) {
  const plain = (map: Map<string, unknown>): Record<string, unknown> => Object.fromEntries(map);
  writeFileSync(
    options.json,
    JSON.stringify(
      {
        files: entries.length,
        status: plain(status),
        rejections: plain(rejections),
        unclassified: plain(unclassified),
        sheetjsFailures: plain(sheetjsFailures),
        parityCategories: plain(parityCategories),
        known: plain(known),
        objectProblems: plain(objectProblems),
        roundtripCategories: plain(roundtripCategories),
        totalCells,
        totalMs,
      },
      null,
      1,
    ),
  );
}
