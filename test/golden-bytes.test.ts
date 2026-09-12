/**
 * The deterministic writer's output for the canonical workbook is pinned byte-for-byte as
 * `fixtures/golden/simple-excel/canonical.xlsx` (registered in the manifest, so the oracle and the corpus suite cover
 * it too). Any writer change that alters bytes must regenerate it deliberately:
 *
 *   UPDATE_GOLDENS=1 npx vitest run --project corpus test/golden-bytes.test.ts
 *   node fixtures/register.mjs golden/simple-excel/canonical.xlsx --id golden-canonical-simple-excel ...
 *
 * and then pass the validator and the Excel oracle before committing.
 */
import { createHash } from 'node:crypto';
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { writeCanonicalWorkbook } from './helpers/canonical-writer';
import { FIXTURES_ROOT } from './helpers/fixtures';
import { WRITER_READY } from './helpers/ready';

const GOLDEN = join(FIXTURES_ROOT, 'golden', 'simple-excel', 'canonical.xlsx');

describe.skipIf(!WRITER_READY)('golden bytes: deterministic canonical workbook', () => {
  it('matches the committed golden byte for byte', async () => {
    const bytes = await writeCanonicalWorkbook();
    if (process.env.UPDATE_GOLDENS === '1' || !existsSync(GOLDEN)) {
      mkdirSync(dirname(GOLDEN), { recursive: true });
      writeFileSync(GOLDEN, bytes);
      console.log(
        `wrote ${GOLDEN} (${bytes.byteLength} bytes, sha256 ${createHash('sha256').update(bytes).digest('hex')}); re-register it in the manifest`,
      );
      return;
    }
    const committed = readFileSync(GOLDEN);
    expect(bytes.byteLength).toBe(committed.byteLength);
    expect(
      Buffer.from(bytes).equals(committed),
      'writer output drifted from the committed golden (see file header for how to update)',
    ).toBe(true);
  });

  it('is reproducible across two runs', async () => {
    const first = await writeCanonicalWorkbook();
    const second = await writeCanonicalWorkbook();
    expect(Buffer.from(first).equals(Buffer.from(second))).toBe(true);
  });
});
