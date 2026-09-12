import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import type { ExpectedWorkbook } from './diff';

export const REPO_ROOT: string = join(import.meta.dirname, '..', '..');
export const FIXTURES_ROOT: string = join(REPO_ROOT, 'fixtures');

export interface Fixture {
  id: string;
  path: string;
  bytes: number;
  sha256: string;
  generator: { name: string; version: string; os: string };
  provenance: string;
  license: string;
  tags: string[];
  expected: string | null;
  expectedError?: string;
  generated: boolean;
  notes?: string;
}

interface Manifest {
  fixtures: Fixture[];
}

let cached: Fixture[] | undefined;

export function allFixtures(): Fixture[] {
  cached ??= (JSON.parse(readFileSync(join(FIXTURES_ROOT, 'manifest.json'), 'utf8')) as Manifest).fixtures;
  return cached;
}

export function fixturesWithTag(tag: string): Fixture[] {
  return allFixtures().filter(fixture => fixture.tags.includes(tag));
}

export function fixtureById(id: string): Fixture {
  const fixture = allFixtures().find(candidate => candidate.id === id);
  if (!fixture) {
    throw new Error(`unknown fixture ${id}`);
  }
  return fixture;
}

export function fixtureFile(fixture: Fixture): string {
  return join(fixture.generated ? join(REPO_ROOT, '.generated') : FIXTURES_ROOT, fixture.path);
}

export function readFixture(fixture: Fixture): Uint8Array {
  const buffer = readFileSync(fixtureFile(fixture));
  return new Uint8Array(buffer.buffer, buffer.byteOffset, buffer.byteLength);
}

export function readExpected(fixture: Fixture): ExpectedWorkbook {
  if (!fixture.expected) {
    throw new Error(`fixture ${fixture.id} has no expected dump`);
  }
  return JSON.parse(readFileSync(join(FIXTURES_ROOT, fixture.expected), 'utf8')) as ExpectedWorkbook;
}

/** `policy:<name>` tags on the fixture (e.g. `truncate-32767`). */
export function fixturePolicies(fixture: Fixture): Set<string> {
  return new Set(fixture.tags.filter(tag => tag.startsWith('policy:')).map(tag => tag.slice('policy:'.length)));
}

export interface CorpusPolicy {
  /** Mismatch categories this fixture is allowed to produce against its expected dump, with the reason. */
  allowedCategories?: string[];
  /**
   * Cells excluded from the SheetJS parity comparison (documented SheetJS quirks), each `Sheet!A1` or, when a whole
   * column of the canonical dataset probes the same quirk, `Sheet!A`.
   */
  parityExemptions?: string[];
  /** Skip the fixture entirely (reason required). */
  skip?: string;
  /** Skip only the parity comparison, for a file SheetJS cannot open (reason required). */
  paritySkip?: string;
  reason?: string;
}

export function corpusPolicies(): Record<string, CorpusPolicy> {
  return JSON.parse(readFileSync(join(REPO_ROOT, 'test', 'corpus-policies.json'), 'utf8')) as Record<string, CorpusPolicy>;
}
