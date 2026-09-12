#!/usr/bin/env node
// Loads canonical-shaped Opportunity rows into the dev org so a Salesforce report over them can be exported
// by hand (Formatted Report + Details Only). Idempotent: existing rows named 'xlsx-engine canonical …' are
// deleted first. Uses the repo's `pnpm sf:api` (JWT bearer flow) - run from the repo root.
//
//   JETSTREAM_REPO=~/dev/jetstream node fixtures/sfdc/load-report-rows.mjs [--delete-only]
import { execFileSync } from 'node:child_process';
import { writeFileSync } from 'node:fs';
import { COLUMNS, ROW_COUNT } from '../canonical/canonical.mjs';

// Runs `pnpm sf:api` from a Jetstream checkout (the only Jetstream-coupled helper in this repo)
const ROOT = process.env.JETSTREAM_REPO ? process.env.JETSTREAM_REPO.replace(/\/?$/, '/') : null;
if (!ROOT) {
  console.error('Set JETSTREAM_REPO to a Jetstream checkout (it provides `pnpm sf:api` and the .env credentials).');
  process.exit(2);
}
const NAME_PREFIX = 'xlsx-engine canonical';

function sfApi(...args) {
  const output = execFileSync('pnpm', ['-s', 'sf:api', ...args], { cwd: ROOT, encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 });
  const jsonStart =
    output.indexOf('{') >= 0 && (output.indexOf('[') < 0 || output.indexOf('{') < output.indexOf('['))
      ? output.indexOf('{')
      : output.indexOf('[');
  return JSON.parse(output.slice(jsonStart));
}

function column(name) {
  return COLUMNS.find(entry => entry.name === name).value;
}

function clampDate(typed) {
  // Salesforce dates must be within 1700-01-01 .. 4000-12-31
  const iso = typed.$date;
  const year = Number(iso.slice(0, 4));
  return year > 4000 ? '4000-12-31' : year < 1700 ? '1700-01-01' : iso;
}

// 1. delete previous rows
const existing = sfApi('query', `SELECT Id FROM Opportunity WHERE Name LIKE '${NAME_PREFIX}%'`, '--all', '--records');
if (existing.length > 0) {
  const ids = existing.map(record => record.Id);
  for (let i = 0; i < ids.length; i += 200) {
    sfApi('delete', `composite/sobjects?ids=${ids.slice(i, i + 200).join(',')}&allOrNone=false`);
  }
  console.log(`deleted ${ids.length} existing rows`);
}
if (process.argv.includes('--delete-only')) {
  process.exit(0);
}

// 2. build 30 rows from the canonical columns
const records = [];
for (let i = 0; i < ROW_COUNT; i++) {
  const percent = column('Percent')(i) * 100;
  const overLimit = column('Over32767')(i);
  const description = [
    `[${NAME_PREFIX}] row ${i}`,
    column('Json')(i),
    column('Spaces')(i),
    column('EscapeLiteral')(i),
    overLimit.length > 30000 ? overLimit.slice(0, 30000) : overLimit,
  ].join('\n');
  records.push({
    attributes: { type: 'Opportunity' },
    Name: `${NAME_PREFIX} ${String(i).padStart(2, '0')} ${column('Name')(i)}`.slice(0, 120),
    StageName: i % 4 === 0 ? 'Closed Won' : 'Prospecting',
    CloseDate: clampDate(column('Date')(i)),
    Amount: column('Currency')(i),
    Probability: Math.max(0, Math.min(100, Math.round(percent))),
    IsPrivate: column('Bool')(i),
    NextStep: `${column('FormulaLikeText')(i)} | ${column('LeadingZeros')(i)} | ${column('RTL')(i)} | ${column('Control')(i)}`.slice(
      0,
      255,
    ),
    Description: description.slice(0, 32000),
    TotalOpportunityQuantity: Math.abs(column('Float')(i)) > 1e15 ? 0 : column('Float')(i),
  });
}
const bodyPath = new URL('../../.generated/sfdc-opportunities.json', import.meta.url).pathname;
writeFileSync(bodyPath, JSON.stringify({ allOrNone: false, records }));
const results = sfApi('post', 'composite/sobjects', `@${bodyPath}`);
const failures = results.filter(result => !result.success);
console.log(`inserted ${results.length - failures.length}/${results.length} Opportunity rows`);
for (const failure of failures) {
  console.log('  FAILED', JSON.stringify(failure.errors?.map(error => `${error.statusCode}: ${error.message}`)));
}
console.log(`Report filter: Opportunity Name starts with "${NAME_PREFIX}". See fixtures/golden/sfdc-report/STEPS.md for the export steps.`);
