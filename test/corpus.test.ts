/**
 * Every fixture in the manifest with an expected dump must read back to it through the public API, up to the
 * mismatch categories `test/corpus-policies.json` allows for that fixture. Fixtures without an expected dump
 * (Jetstream assets, Salesforce report exports) must at least open and stream every sheet without throwing.
 */
import { describe, expect, it } from 'vitest';
import { compareDumps, describeReport } from './helpers/diff';
import { dumpWorkbook } from './helpers/dump';
import { allFixtures, corpusPolicies, fixturePolicies, readExpected, readFixture } from './helpers/fixtures';
import { READER_READY } from './helpers/ready';

const policies = corpusPolicies();
const readable = allFixtures().filter(fixture => !fixture.expectedError && !fixture.tags.includes('kind:hostile'));

describe.skipIf(!READER_READY)('corpus: reads every fixture to its expected dump', () => {
  for (const fixture of readable) {
    const policy = policies[fixture.id];
    if (policy?.skip) {
      it.skip(`${fixture.id} (${policy.skip})`, () => {});
      continue;
    }
    it(fixture.id, async () => {
      const dump = await dumpWorkbook(readFixture(fixture));
      expect(dump.sheets.length).toBeGreaterThan(0);
      if (!fixture.expected) {
        return;
      }
      const report = compareDumps(readExpected(fixture), dump, fixturePolicies(fixture));
      const allowed = new Set(policy?.allowedCategories ?? []);
      const unexpected = Object.keys(report.byCategory).filter(category => !allowed.has(category));
      expect(unexpected, describeReport(report)).toEqual([]);
    });
  }
});
