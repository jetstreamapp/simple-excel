import { describe, expect, it } from 'vitest';
import type { SharedStringBudget } from '../../types';
import type { ZipEntrySummary, ZipEntryWriter } from '../../zip/zip-writer';
import { SharedStringWriter } from '../shared-strings';

const DEFAULT_BUDGET: Required<SharedStringBudget> = { maxUnique: 65_536, maxChars: 16 * 1024 * 1024, maxLength: 256 };
const UNBOUNDED_BUDGET: Required<SharedStringBudget> = {
  maxUnique: Number.POSITIVE_INFINITY,
  maxChars: Number.POSITIVE_INFINITY,
  maxLength: Number.POSITIVE_INFINITY,
};

/** Collects what the writer pushes, one entry per `write` call, so the chunking itself can be asserted. */
class CollectingEntry implements ZipEntryWriter {
  readonly chunks: string[] = [];

  async write(chunk: Uint8Array): Promise<void> {
    this.chunks.push(new TextDecoder().decode(chunk));
  }

  async close(): Promise<ZipEntrySummary> {
    return { name: 'xl/sharedStrings.xml', crc32: 0, compressedSize: 0, uncompressedSize: 0, offset: 0 };
  }

  get xml(): string {
    return this.chunks.join('');
  }
}

function budget(overrides: Partial<Required<SharedStringBudget>>): Required<SharedStringBudget> {
  return { ...DEFAULT_BUDGET, ...overrides };
}

async function serialize(writer: SharedStringWriter): Promise<CollectingEntry> {
  const entry = new CollectingEntry();
  await writer.writeTo(entry);
  return entry;
}

describe('SharedStringWriter interning', () => {
  it('hands out indexes in insertion order and counts every reference', () => {
    const writer = new SharedStringWriter(DEFAULT_BUDGET);
    expect(writer.intern('Account')).toBe(0);
    expect(writer.intern('Contact')).toBe(1);
    expect(writer.intern('Account')).toBe(0);
    expect(writer.intern('Account')).toBe(0);
    expect(writer.stats).toEqual({ count: 4, uniqueCount: 2, frozen: false });
  });

  it('writes strings over maxLength inline without touching the table', () => {
    const writer = new SharedStringWriter(budget({ maxLength: 8 }));
    expect(writer.intern('short')).toBe(0);
    expect(writer.intern('a string that is far too long to be worth interning')).toBe(-1);
    expect(writer.stats).toEqual({ count: 1, uniqueCount: 1, frozen: false });
  });

  it('freezes on maxUnique but keeps resolving strings it already knows (ADR-001)', () => {
    const writer = new SharedStringWriter(budget({ maxUnique: 2 }));
    expect(writer.intern('first')).toBe(0);
    expect(writer.intern('second')).toBe(1);
    expect(writer.stats.frozen).toBe(false);

    expect(writer.intern('third'), 'the budget is spent, so a new string goes inline').toBe(-1);
    expect(writer.stats.frozen).toBe(true);
    expect(writer.intern('first'), 'a known string still resolves after the freeze').toBe(0);
    expect(writer.intern('second')).toBe(1);
    expect(writer.intern('fourth')).toBe(-1);
    expect(writer.stats).toEqual({ count: 4, uniqueCount: 2, frozen: true });
  });

  it('freezes on maxChars before the next string would cross it', () => {
    const writer = new SharedStringWriter(budget({ maxChars: 10 }));
    expect(writer.intern('12345')).toBe(0);
    expect(writer.intern('67890')).toBe(1);
    expect(writer.intern('!')).toBe(-1);
    expect(writer.stats).toEqual({ count: 2, uniqueCount: 2, frozen: true });
  });

  it('keeps a million unique strings inside the default budget', () => {
    const writer = new SharedStringWriter(DEFAULT_BUDGET);
    let inlined = 0;
    for (let i = 0; i < 1_000_000; i++) {
      if (writer.intern(`003Xx0000000${String(i).padStart(6, '0')}AAA`) === -1) {
        inlined++;
      }
    }
    const { uniqueCount, count, frozen } = writer.stats;
    expect(frozen).toBe(true);
    expect(uniqueCount).toBe(DEFAULT_BUDGET.maxUnique);
    expect(count).toBe(uniqueCount);
    expect(inlined).toBe(1_000_000 - uniqueCount);
    // 65,536 x 21 characters is ~1.3 Mi, comfortably inside the 16 Mi ceiling the budget promises.
    expect(uniqueCount * 21).toBeLessThan(DEFAULT_BUDGET.maxChars);
  });

  it('interns everything when the budget is switched off (strings: shared)', () => {
    const writer = new SharedStringWriter(UNBOUNDED_BUDGET);
    for (let i = 0; i < 100_000; i++) {
      writer.intern(`value ${i}`);
    }
    expect(writer.intern('x'.repeat(100_000))).toBe(100_000);
    expect(writer.stats).toEqual({ count: 100_001, uniqueCount: 100_001, frozen: false });
  });
});

describe('SharedStringWriter serialization', () => {
  it('writes an sst with exact counts in insertion order', async () => {
    const writer = new SharedStringWriter(DEFAULT_BUDGET);
    writer.intern('Account');
    writer.intern('Contact');
    writer.intern('Account');
    const entry = await serialize(writer);

    expect(entry.xml).toBe(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="2">' +
        '<si><t>Account</t></si><si><t>Contact</t></si></sst>',
    );
  });

  it('writes an empty table when nothing was interned', async () => {
    const entry = await serialize(new SharedStringWriter(DEFAULT_BUDGET));
    expect(entry.xml).toContain('count="0" uniqueCount="0"');
    expect(entry.xml.endsWith('></sst>')).toBe(true);
  });

  it('preserves whitespace and escapes cell text (EC-XML-SPACE-PRESERVE, EC-XML-ESCAPE-LITERAL)', async () => {
    const writer = new SharedStringWriter(DEFAULT_BUDGET);
    writer.intern(' leading');
    writer.intern('trailing ');
    writer.intern('line1\nline2');
    writer.intern('a & b < c');
    writer.intern('_x0041_');
    writer.intern('ctlx');
    writer.intern('crlf\r\nx');
    const entry = await serialize(writer);

    expect(entry.xml).toContain('<si><t xml:space="preserve"> leading</t></si>');
    expect(entry.xml).toContain('<si><t xml:space="preserve">trailing </t></si>');
    expect(entry.xml).toContain('<si><t xml:space="preserve">line1\nline2</t></si>');
    expect(entry.xml).toContain('<si><t>a &amp; b &lt; c</t></si>');
    expect(entry.xml).toContain('<si><t>_x005F_x0041_</t></si>');
    expect(entry.xml).toContain('<si><t>ctl_x0001_x</t></si>');
    expect(entry.xml).toContain('<si><t xml:space="preserve">crlf_x000D_\nx</t></si>');
  });

  it('flushes in bounded chunks rather than building one string', async () => {
    const writer = new SharedStringWriter(DEFAULT_BUDGET);
    for (let i = 0; i < 20_000; i++) {
      writer.intern(`row ${i} value with enough text to add up`);
    }
    const entry = await serialize(writer);

    expect(entry.chunks.length).toBeGreaterThan(10);
    // Every flush happens at the first entry past 64 Ki characters, so no chunk is materially larger than that.
    for (const chunk of entry.chunks) {
      expect(chunk.length).toBeLessThan(64 * 1024 + 1024);
    }
    expect(entry.xml).toContain('count="20000" uniqueCount="20000"');
    expect(entry.xml.match(/<si>/g)?.length).toBe(20_000);
  });
});
