#!/usr/bin/env node
import { createHash } from 'node:crypto';
import { existsSync, mkdirSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';

/**
 * Fixture manifest tool.
 *
 *   node generate.mjs --check      verify every committed fixture matches its manifest bytes/sha256
 *   node generate.mjs              build every `generated: true` fixture into .generated/ and verify it
 *   node generate.mjs --update     rewrite bytes/sha256 in manifest.json from the files on disk (after
 *                                  deliberately changing a fixture)
 *
 * Generated fixtures are produced by `generators/<id>.mjs` modules exporting `default async (outPath) => void`.
 * Committed fixtures must stay under 200 KB each (policy in README.md).
 */
const FIXTURES_DIR = new URL('.', import.meta.url).pathname.replace(/\/$/, '');
const MANIFEST_PATH = join(FIXTURES_DIR, 'manifest.json');
const MAX_COMMITTED_BYTES = 200 * 1024;

const args = new Set(process.argv.slice(2));
const manifest = JSON.parse(readFileSync(MANIFEST_PATH, 'utf8'));

function sha256(path) {
  return createHash('sha256').update(readFileSync(path)).digest('hex');
}

async function buildGenerated(fixture) {
  const generatorPath = join(FIXTURES_DIR, 'generators', `${fixture.id}.mjs`);
  if (!existsSync(generatorPath)) {
    throw new Error(`${fixture.id}: generated fixture has no generators/${fixture.id}.mjs`);
  }
  const outPath = join(FIXTURES_DIR, fixture.path);
  mkdirSync(dirname(outPath), { recursive: true });
  const { default: generate } = await import(generatorPath);
  await generate(outPath);
}

const problems = [];
const ids = new Set();
for (const fixture of manifest.fixtures) {
  if (ids.has(fixture.id)) {
    problems.push(`${fixture.id}: duplicate id`);
  }
  ids.add(fixture.id);

  const path = join(FIXTURES_DIR, fixture.path);
  if (fixture.generated) {
    if (!fixture.path.startsWith('.generated/')) {
      problems.push(`${fixture.id}: generated fixtures must live under .generated/`);
      continue;
    }
    if (args.has('--check')) {
      continue;
    }
    try {
      await buildGenerated(fixture);
    } catch (error) {
      problems.push(`${fixture.id}: generation failed - ${error.message}`);
      continue;
    }
  }

  if (!existsSync(path)) {
    problems.push(`${fixture.id}: missing file ${fixture.path}`);
    continue;
  }
  const bytes = statSync(path).size;
  const digest = sha256(path);
  if (args.has('--update')) {
    fixture.bytes = bytes;
    fixture.sha256 = digest;
    continue;
  }
  if (!fixture.generated && bytes > MAX_COMMITTED_BYTES) {
    problems.push(`${fixture.id}: ${bytes} bytes exceeds the ${MAX_COMMITTED_BYTES}-byte committed-fixture limit; mark it generated`);
  }
  if (fixture.bytes !== bytes) {
    problems.push(`${fixture.id}: size ${bytes} != manifest ${fixture.bytes}`);
  }
  if (fixture.sha256 !== digest) {
    problems.push(`${fixture.id}: sha256 ${digest} != manifest ${fixture.sha256}`);
  }
}

if (args.has('--update')) {
  writeFileSync(MANIFEST_PATH, JSON.stringify(manifest, null, 2) + '\n');
  console.log(`manifest updated for ${manifest.fixtures.length} fixtures`);
} else if (problems.length > 0) {
  console.error(`fixture manifest problems:\n  ${problems.join('\n  ')}`);
  process.exit(1);
} else {
  console.log(`${manifest.fixtures.length} fixtures verified`);
}
