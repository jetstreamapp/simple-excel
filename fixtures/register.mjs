#!/usr/bin/env node
import { createHash } from 'node:crypto';
import { existsSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { parseArgs } from 'node:util';

/**
 * Register (or refresh) a committed fixture in manifest.json.
 *
 *   node register.mjs <path-relative-to-fixtures> --id <id> --generator "<name> <version>" \
 *     --provenance "script:generators/canonical-sheetjs.mjs" --license MIT --tags kind:golden,generator:sheetjs \
 *     [--expected canonical/canonical.json] [--expected-error ENCRYPTED] [--notes "..."]
 *
 * Re-registering an existing id replaces its entry (bytes/sha256 are always recomputed from disk).
 */
const FIXTURES_DIR = new URL('.', import.meta.url).pathname.replace(/\/$/, '');
const MANIFEST_PATH = join(FIXTURES_DIR, 'manifest.json');

const { values, positionals } = parseArgs({
  allowPositionals: true,
  options: {
    id: { type: 'string' },
    generator: { type: 'string' },
    provenance: { type: 'string' },
    license: { type: 'string', default: 'MIT' },
    tags: { type: 'string', default: '' },
    expected: { type: 'string' },
    'expected-error': { type: 'string' },
    notes: { type: 'string' },
    generated: { type: 'boolean', default: false },
  },
});

const [relativePath] = positionals;
if (!relativePath || !values.id || !values.generator || !values.provenance) {
  console.error(
    'usage: register.mjs <path> --id <id> --generator <name+version> --provenance <text> [--license --tags --expected --expected-error --notes]',
  );
  process.exit(2);
}
const absolutePath = join(FIXTURES_DIR, relativePath);
if (!existsSync(absolutePath)) {
  console.error(`missing file: ${absolutePath}`);
  process.exit(1);
}

const [generatorName, ...versionParts] = values.generator.split(' ');
const entry = {
  id: values.id,
  path: relativePath,
  bytes: statSync(absolutePath).size,
  sha256: createHash('sha256').update(readFileSync(absolutePath)).digest('hex'),
  generator: {
    name: generatorName,
    version: versionParts.join(' ') || 'unknown',
    os: process.platform === 'darwin' ? 'macOS' : process.platform,
  },
  provenance: values.provenance,
  license: values.license,
  tags: values.tags.split(',').filter(Boolean),
  expected: values.expected ?? null,
  generated: values.generated,
};
if (values['expected-error']) {
  entry.expectedError = values['expected-error'];
}
if (values.notes) {
  entry.notes = values.notes;
}

const manifest = JSON.parse(readFileSync(MANIFEST_PATH, 'utf8'));
const index = manifest.fixtures.findIndex(fixture => fixture.id === entry.id);
if (index >= 0) {
  manifest.fixtures[index] = entry;
} else {
  manifest.fixtures.push(entry);
}
writeFileSync(MANIFEST_PATH, JSON.stringify(manifest, null, 2) + '\n');
console.log(`${index >= 0 ? 'updated' : 'registered'} ${entry.id} (${entry.bytes} bytes)`);
