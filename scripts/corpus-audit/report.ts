/**
 * Summary of a corpus-audit results file: status counts, rejection codes, unexplained diff categories with example
 * files, slowest files, highest memory, throughput.
 *
 *   tsx scripts/corpus-audit/report.ts [tmp/corpus-audit/results.jsonl] [--json summary.json]
 */
import { readFileSync, writeFileSync } from 'node:fs';
import { basename, dirname } from 'node:path';
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
const knownObjects = new Map<string, number>();
const sheetjsFailures = new Map<string, { count: number; examples: string[] }>();
const sheetjsDropped = new Map<string, { count: number; examples: string[] }>();
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
  if (entry.sheetjs?.dropped) {
    bump(
      sheetjsDropped,
      `${entry.sheetjs.dropped.reason ?? 'unexplained'}: ${String(entry.sheetjs.dropped.error).slice(0, 90)}`,
      entry.file,
    );
  }
  for (const category of Object.keys(entry.parity?.byCategory ?? {})) {
    bump(parityCategories, category, entry.file);
  }
  for (const [reason, count] of Object.entries(entry.parity?.known ?? {})) {
    known.set(reason, (known.get(reason) ?? 0) + (count as number));
  }
  for (const [reason, count] of Object.entries(entry.objects?.known ?? {})) {
    knownObjects.set(reason, (knownObjects.get(reason) ?? 0) + (count as number));
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

// A corpus laid out one folder per writer (select-by-writer.py): the same status counts per folder.
const byFolder = new Map<string, Map<string, number>>();
for (const entry of entries) {
  const folder = basename(dirname(entry.file));
  const counts = byFolder.get(folder) ?? new Map<string, number>();
  counts.set(entry.status, (counts.get(entry.status) ?? 0) + 1);
  byFolder.set(folder, counts);
}
if (byFolder.size > 1) {
  console.log('\n## Status by folder (writer)');
  for (const [folder, counts] of [...byFolder].sort((a, b) => b[1].size - a[1].size || a[0].localeCompare(b[0]))) {
    const total = [...counts.values()].reduce((sum, count) => sum + count, 0);
    const parts = [...counts].sort((a, b) => b[1] - a[1]).map(([state, count]) => `${state} ${count}`);
    console.log(`  ${String(total).padStart(5)}  ${folder.padEnd(36)} ${parts.join(', ')}`);
  }
}

table('Rejected (classified XlsxError)', rejections);
table('Unclassified errors, timeouts, crashes', unclassified);
table('SheetJS failures', sheetjsFailures);
table('Sheets SheetJS dropped without an error (files; cause from a WTF re-read)', sheetjsDropped);
table('Parity vs SheetJS: unexplained categories (files)', parityCategories);
console.log(`\n## Known SheetJS divergences (cells)\n  ${JSON.stringify(Object.fromEntries(known))}`);
console.log(`\n## Known SheetJS divergences (objects: cells, records or files)\n  ${JSON.stringify(Object.fromEntries(knownObjects))}`);
table('Objects (toObjects vs sheet_to_json) failures (files)', objectProblems);
table('Round trip categories (files)', roundtripCategories);

// A stream pass over 5x SheetJS's read was timed twice more by the worker; the best of the three ranks it here, so one
// noisy pass of a small file does not make the list. The first pass is shown alongside.
const streamMs = (entry: Record<string, any>): number => entry.ours.bestMs ?? entry.ours.ms;
console.log('\n## Slowest reads (stream pass; best of 3 where the first was over 5x SheetJS)');
for (const entry of [...read].sort((a, b) => streamMs(b) - streamMs(a)).slice(0, 10)) {
  const firstPass = entry.ours.bestMs === undefined ? '' : ` (first pass ${entry.ours.ms} ms)`;
  console.log(
    `  ${String(streamMs(entry)).padStart(7)} ms${firstPass}  ${(entry.bytes / 1048576).toFixed(1)} MB  ${entry.ours.cells.toLocaleString()} cells  sheetjs ${entry.sheetjs?.ms ?? '-'} ms  ${basename(entry.file)}`,
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
 * before is a limit set too tight), cell count, and the reader's time and peak-RSS growth in aggregate. Files are
 * matched by name, which is the content's sha256 in both corpora, so a corpus copied or regrouped into other folders
 * still lines up with its baseline.
 */
if (options.baseline) {
  const before = new Map(load(options.baseline).map(entry => [basename(entry.file as string), entry]));
  const transitions = new Map<string, { count: number; examples: string[] }>();
  let cellsChanged = 0;
  let timeBefore = 0;
  let timeAfter = 0;
  let sheetjsBefore = 0;
  let sheetjsAfter = 0;
  const rssGrowth: { file: string; before: number; after: number }[] = [];
  let matched = 0;
  for (const entry of entries) {
    const old = before.get(basename(entry.file));
    if (!old) {
      continue;
    }
    matched++;
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
      // SheetJS on the same files is the yardstick for how loaded the machine was in each run.
      sheetjsBefore += old.sheetjs?.ms ?? 0;
      sheetjsAfter += entry.sheetjs?.ms ?? 0;
      rssGrowth.push({
        file: entry.file,
        before: old.ours.peakRssMb - old.ours.baselineRssMb,
        after: entry.ours.peakRssMb - entry.ours.baselineRssMb,
      });
    }
  }
  console.log(`\n# Against baseline ${options.baseline}`);
  console.log(`  ${matched} of ${entries.length} files are in the baseline's ${before.size}`);
  console.log(transitions.size === 0 ? '  no status changes' : '');
  table('Status changes', transitions);
  console.log(`  files whose cell count changed: ${cellsChanged}`);
  console.log(`  reader time on files read both times: ${(timeBefore / 1000).toFixed(1)}s -> ${(timeAfter / 1000).toFixed(1)}s`);
  console.log(
    `  SheetJS on the same files: ${(sheetjsBefore / 1000).toFixed(1)}s -> ${(sheetjsAfter / 1000).toFixed(1)}s; ours / SheetJS ${(timeBefore / sheetjsBefore).toFixed(2)} -> ${(timeAfter / sheetjsAfter).toFixed(2)}`,
  );
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
        sheetjsDropped: plain(sheetjsDropped),
        parityCategories: plain(parityCategories),
        known: plain(known),
        knownObjects: plain(knownObjects),
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
