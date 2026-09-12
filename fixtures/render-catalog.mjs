#!/usr/bin/env node
// Renders the generated part of research/04-edge-case-catalog.md from fixtures/manifest.json +
// fixtures/edge-cases.json (the hand-maintained catalog entries). Output replaces everything between the
// <!-- generated:start --> and <!-- generated:end --> markers in the doc.
//
//   node render-catalog.mjs [--doc ../research/04-edge-case-catalog.md]
import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { parseArgs } from 'node:util';

const HERE = new URL('.', import.meta.url).pathname.replace(/\/$/, '');
const { values: options } = parseArgs({
  options: { doc: { type: 'string', default: join(HERE, '../research/04-edge-case-catalog.md') } },
});

const manifest = JSON.parse(readFileSync(join(HERE, 'manifest.json'), 'utf8'));
const catalog = JSON.parse(readFileSync(join(HERE, 'edge-cases.json'), 'utf8'));

const fixturesByTag = new Map();
for (const fixture of manifest.fixtures) {
  for (const tag of fixture.tags) {
    if (tag.startsWith('EC-')) {
      fixturesByTag.set(tag, [...(fixturesByTag.get(tag) ?? []), fixture.id]);
    }
  }
}

const escape = text =>
  String(text ?? '')
    .replace(/\|/g, '\\|')
    .replace(/\n/g, ' ');
const families = [...new Set(catalog.map(entry => entry.family))];
const lines = ['<!-- generated:start (node fixtures/render-catalog.mjs) -->', ''];
for (const family of families) {
  lines.push(`### ${family}`, '', '| Id | Symptom | Seen in | Expected behaviour | Fixtures | Status |', '|---|---|---|---|---|---|');
  for (const entry of catalog.filter(item => item.family === family)) {
    const fixtures = [...new Set([...(entry.fixtures ?? []), ...(fixturesByTag.get(entry.id) ?? [])])];
    lines.push(
      `| \`${entry.id}\` | ${escape(entry.symptom)} | ${escape(entry.seenIn)} | ${escape(entry.expected)} | ${fixtures.map(id => `\`${id}\``).join(', ') || '_none yet_'} | ${escape(entry.status ?? 'observed')} |`,
    );
  }
  lines.push('');
}
const untagged = [...fixturesByTag.keys()].filter(tag => !catalog.some(entry => entry.id === tag));
if (untagged.length > 0) {
  lines.push(
    '### Tags on fixtures without a catalog entry',
    '',
    ...untagged.map(tag => `- \`${tag}\`: ${fixturesByTag.get(tag).join(', ')}`),
    '',
  );
}
lines.push(
  `_${catalog.length} catalog entries, ${manifest.fixtures.length} fixtures, rendered ${new Date().toISOString().slice(0, 10)}._`,
  '',
  '<!-- generated:end -->',
);

const doc = readFileSync(options.doc, 'utf8');
const start = doc.indexOf('<!-- generated:start');
const end = doc.indexOf('<!-- generated:end -->');
if (start < 0 || end < 0) {
  throw new Error(`markers not found in ${options.doc}`);
}
writeFileSync(options.doc, doc.slice(0, start) + lines.join('\n') + doc.slice(end + '<!-- generated:end -->'.length));
console.log(`rendered ${catalog.length} entries into ${options.doc}`);
