/**
 * The deterministic writer's output for the canonical workbook (with the bounded shared-string table, `strings: 'auto'`;
 * `canonical.inline.xlsx` covers the inline default) is pinned byte-for-byte as
 * `fixtures/golden/simple-excel/canonical.xlsx` (registered in the manifest, so the oracle and the corpus suite cover
 * it too). Any writer change that alters bytes must regenerate it deliberately:
 *
 *   UPDATE_GOLDENS=1 npx vitest run --project corpus test/golden-bytes.test.ts
 *   node fixtures/register.mjs golden/simple-excel/canonical.xlsx --id golden-canonical-simple-excel ...
 *
 * and then pass the validator and the Excel oracle before committing.
 *
 * The workbook is written with `dates: 'local'`, and the canonical DST-gap cell (local 2024-03-10 02:30) does not exist
 * in US zones that observe DST: there JavaScript moves it to 03:30, while in UTC it stays 02:30. The bytes therefore
 * depend on the host time zone, so the suite pins a US zone, as the golden was generated in one (CI runs in UTC).
 */
import { createHash } from 'node:crypto';
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { writeCanonicalWorkbook } from './helpers/canonical-writer';
import { FIXTURES_ROOT } from './helpers/fixtures';
import { WRITER_READY } from './helpers/ready';

const GOLDEN = join(FIXTURES_ROOT, 'golden', 'simple-excel', 'canonical.xlsx');
const GOLDEN_TIME_ZONE = 'America/Los_Angeles';

describe.skipIf(!WRITER_READY)('golden bytes: deterministic canonical workbook', () => {
  const hostTimeZone = process.env.TZ;
  beforeAll(() => {
    process.env.TZ = GOLDEN_TIME_ZONE;
  });
  afterAll(() => {
    if (hostTimeZone === undefined) {
      delete process.env.TZ;
    } else {
      process.env.TZ = hostTimeZone;
    }
  });

  it('matches the committed golden byte for byte', async () => {
    const bytes = await writeCanonicalWorkbook({ strings: 'auto' });
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
    const first = await writeCanonicalWorkbook({ strings: 'auto' });
    const second = await writeCanonicalWorkbook({ strings: 'auto' });
    expect(Buffer.from(first).equals(Buffer.from(second))).toBe(true);
  });
});
