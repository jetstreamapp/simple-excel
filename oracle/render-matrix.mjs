#!/usr/bin/env node
// Renders the generated part of research/05-compatibility-matrix.md from the newest oracle results
// folder (or --results <dir>). Output replaces everything between the generated markers in the doc.
//
//   node render-matrix.mjs [--results oracle/results/<dir>] [--doc ../research/05-compatibility-matrix.md]
import { existsSync, readdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { parseArgs } from 'node:util';

const HERE = new URL('.', import.meta.url).pathname.replace(/\/$/, '');
const { values: options } = parseArgs({
  options: {
    results: { type: 'string' },
    doc: { type: 'string', default: join(HERE, '../research/05-compatibility-matrix.md') },
  },
});
const resultsRoot = join(HERE, 'results');
const manifest = JSON.parse(readFileSync(join(HERE, '../fixtures/manifest.json'), 'utf8'));
const fixtureById = new Map(manifest.fixtures.map(fixture => [fixture.id, fixture]));
const runFolders = options.results
  ? [options.results]
  : readdirSync(resultsRoot)
      .filter(name => existsSync(join(resultsRoot, name, 'results.json')))
      .sort()
      .map(name => join(resultsRoot, name));
// keep only the newest run per label (folder names are <date>-<label>)
const newestByLabel = new Map();
for (const folder of runFolders) {
  const label = folder
    .split('/')
    .pop()
    .replace(/^\d{4}-\d{2}-\d{2}-/, '');
  newestByLabel.set(label, folder);
}

function cell(record) {
  if (!record) {
    return '';
  }
  switch (record.status) {
    case 'PASS':
      return record.reader === 'excel' ? `PASS${record.repairLogs ? ' (repair log)' : ''}` : 'PASS';
    case 'DIFF': {
      const pct = record.cells ? Math.round((record.matches / record.cells) * 100) : 0;
      const top = Object.entries(record.byCategory ?? {})
        .sort((a, b) => b[1] - a[1])
        .slice(0, 2)
        .map(([category, count]) => `${category} ${count}`)
        .join(', ');
      return `${pct}% (${top})`;
    }
    case 'FAIL':
      return record.schemaErrors !== undefined && record.schemaErrors !== null
        ? `FAIL: ${record.schemaErrors} schema error${record.schemaErrors === 1 ? '' : 's'}`
        : 'FAIL';
    case 'REPAIRED':
      return `REPAIRED (${(record.repairLogs ?? []).flatMap(log => log.removed).join('; ') || 'see log'})`;
    case 'REJECTED':
      return `REJECTED: ${
        String(record.error ?? '')
          .split('|')
          .map(part => part.trim())
          .filter(part => part && !part.startsWith('Node.js') && part !== '^' && part !== '}')
          .pop()
          ?.replace(/^\[cause\]: /, '')
          .slice(0, 90) ?? ''
      }`;
    case 'ACCEPTED':
      return `ACCEPTED (${(record.sheets ?? []).map(sheet => `${sheet.name}: ${sheet.rows} rows`).join(', ')})`;
    default:
      return record.status;
  }
}

const escapeCell = text =>
  String(text)
    .replace(/\|/g, '\\|')
    // oxlint-disable-next-line no-control-regex -- escaping control characters for the markdown table is the point
    .replace(/[\u0000-\u001F]/g, ch => `\\u${ch.charCodeAt(0).toString(16).padStart(4, '0')}`);
const lines = ['<!-- generated:start (node oracle/render-matrix.mjs) -->', ''];
for (const [label, folder] of newestByLabel) {
  const { meta, results } = JSON.parse(readFileSync(join(folder, 'results.json'), 'utf8'));
  const readers = meta.readers;
  lines.push(`### Run \`${label}\` (${folder.split('/').pop()})`, '', `${meta.date}, ${meta.host}, Node ${meta.node}.`, '');
  lines.push(`| fixture (generator) | ${readers.join(' | ')} |`, `|---|${readers.map(() => '---').join('|')}|`);
  for (const id of meta.fixtures) {
    const fixture = fixtureById.get(id);
    const generator = fixture
      ? `${fixture.generator.name} ${fixture.generator.version ?? ''}`
          .trim()
          .replace(/\(.*\)/, '')
          .trim()
      : '';
    const expectation = fixture?.expectedError ? `<br>expect \`${fixture.expectedError}\`` : '';
    const label2 = `\`${id}\`<br>${generator}${expectation}`;
    const cells = readers.map(reader => escapeCell(cell(results.find(record => record.fixture === id && record.reader === reader))));
    lines.push(`| ${label2} | ${cells.join(' | ')} |`);
  }
  lines.push('');
  const excelRecords = results.filter(record => record.reader === 'excel' && record.cells && label === 'goldens');
  if (excelRecords.length > 0) {
    const refs = ['E2', 'F2', 'F3', 'F7', 'G2', 'I2', 'J11', 'M2', 'M3', 'M4', 'N2', 'O2', 'O6'];
    lines.push(
      '#### What Excel holds in the sentinel cells (raw `value`; dates shown as Excel reports them)',
      '',
      'Column `E` control chars, `F` `_x` escape literals, `G` formula-like text, `I` 17-digit integer, `J11` max double, `M` dates (1899-12-31 / 1900-02-28 / 1900-03-01), `N` time-only 12:34:56.789, `O` datetimes (2024-03-10 02:30 DST gap / 1900-01-01 12:00). Caveat: AppleScript date values pass through macOS local-time normalization, so 02:30 in a DST gap reads as 03:30.',
      '',
    );
    lines.push(`| fixture | ${refs.join(' | ')} |`, `|---|${refs.map(() => '---').join('|')}|`);
    for (const record of excelRecords) {
      const byRef = new Map(record.cells.map(entry => [entry.ref, entry]));
      const values = refs.map(ref => {
        const entry = byRef.get(ref);
        if (!entry) {
          return '';
        }
        const shown = /^[MNO]\d+$/.test(ref) ? entry.value : entry.text || entry.value || '';
        return `\`${escapeCell(String(shown).replace(/day, /, '').replace(' at ', ' ').slice(0, 30))}\``;
      });
      lines.push(`| \`${record.fixture.replace('golden-canonical-', '')}\` | ${values.join(' | ')} |`);
    }
    lines.push('');
  }
}
lines.push(
  'Legend: **PASS** every compared cell equal after fixture policies; **n%** share of matching cells with the two largest mismatch categories (04 maps categories to catalog entries); **FAIL: n schema errors** Open XML SDK validator; **REPAIRED** Excel wrote a recovery log while opening (silent repair under automation); **OPENED** no ground truth; **REJECTED/ACCEPTED** hostile fixture outcome (a classified rejection is the goal); **ERROR** the reader threw on a fixture it should read.',
  '',
);
lines.push('<!-- generated:end -->');

const doc = readFileSync(options.doc, 'utf8');
const start = doc.indexOf('<!-- generated:start');
const end = doc.indexOf('<!-- generated:end -->');
if (start < 0 || end < 0) {
  throw new Error(`markers not found in ${options.doc}`);
}
writeFileSync(options.doc, doc.slice(0, start) + lines.join('\n') + doc.slice(end + '<!-- generated:end -->'.length));
console.log(`rendered ${newestByLabel.size} run(s) into ${options.doc}`);
