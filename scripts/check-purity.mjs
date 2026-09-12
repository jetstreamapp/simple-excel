import { readdirSync, readFileSync, statSync } from 'node:fs';
import { join, relative } from 'node:path';

/**
 * Nothing in this repository may import Jetstream monorepo libraries (`@jetstream/*`): the library, fixtures,
 * oracle and bench are self-contained. A guardrail, not a proof.
 */
const ROOT = new URL('..', import.meta.url).pathname.replace(/\/$/, '');
const SKIP_DIRS = new Set(['node_modules', '.generated', 'results', 'dist', 'docs', '.git', 'coverage']);
const SOURCE_FILE = /\.(mjs|cjs|js|ts|mts|cts|py|sh)$/;
const FORBIDDEN = [{ pattern: /['"]@jetstream\//, reason: 'imports a monorepo library (@jetstream/*)' }];

function* walk(directory) {
  for (const entry of readdirSync(directory)) {
    if (SKIP_DIRS.has(entry)) {
      continue;
    }
    const fullPath = join(directory, entry);
    if (statSync(fullPath).isDirectory()) {
      yield* walk(fullPath);
    } else if (SOURCE_FILE.test(entry)) {
      yield fullPath;
    }
  }
}

const violations = [];
for (const file of walk(ROOT)) {
  const source = readFileSync(file, 'utf8');
  for (const { pattern, reason } of FORBIDDEN) {
    if (pattern.test(source)) {
      violations.push(`${relative(ROOT, file)}: ${reason}`);
    }
  }
}

if (violations.length > 0) {
  console.error('purity check failed:\n  ' + violations.join('\n  '));
  process.exit(1);
}
console.log('purity check passed');
