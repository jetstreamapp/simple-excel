import { describe, expect, it } from 'vitest';
import { SHEET_LONG_NAME } from '../../../fixtures/canonical/canonical.mjs';
import { isXlsxError } from '../../errors';
import { isValidSheetName, MAX_SHEET_NAME_LENGTH, quoteSheetName, sanitizeSheetName } from '../sheet-name';

const sanitize = (name: string, taken: Set<string> = new Set()): string => sanitizeSheetName(name, taken);

describe('sanitizeSheetName', () => {
  it('keeps a name Excel already accepts', () => {
    expect(sanitize('Data')).toBe('Data');
    expect(sanitize('Sheet 1')).toBe('Sheet 1');
    expect(sanitize('Ünïcödé 🚀')).toBe('Ünïcödé 🚀');
  });

  it('EC-SHEET-NAME-31-APOSTROPHE: the canonical 31-character name with an inner apostrophe is untouched', () => {
    expect(SHEET_LONG_NAME).toHaveLength(MAX_SHEET_NAME_LENGTH);
    expect(sanitize(SHEET_LONG_NAME)).toBe(SHEET_LONG_NAME);
    expect(isValidSheetName(SHEET_LONG_NAME)).toBe(true);
  });

  it('replaces the characters Excel forbids with underscores', () => {
    expect(sanitize('a:b\\c/d?e*f[g]h')).toBe('a_b_c_d_e_f_g_h');
    expect(sanitize('////')).toBe('____');
    expect(sanitize('Q1/Q2 Sales')).toBe('Q1_Q2 Sales');
  });

  it('strips leading and trailing apostrophes and whitespace, but not inner ones', () => {
    expect(sanitize("'Quoted'")).toBe('Quoted');
    expect(sanitize("''Quoted''")).toBe('Quoted');
    expect(sanitize("  '  Spaced  '  ")).toBe('Spaced');
    expect(sanitize("It's fine")).toBe("It's fine");
  });

  it('renames the reserved History sheet in any casing', () => {
    expect(sanitize('History')).toBe('History_');
    expect(sanitize('HISTORY')).toBe('HISTORY_');
    expect(sanitize('history')).toBe('history_');
    expect(sanitize('Histories')).toBe('Histories');
    expect(isValidSheetName('History')).toBe(false);
  });

  it('falls back to Sheet<n> for an empty name', () => {
    expect(sanitize('')).toBe('Sheet1');
    expect(sanitize("'''")).toBe('Sheet1');
    expect(sanitize('   ')).toBe('Sheet1');
    const taken = new Set(['sheet1']);
    expect(sanitize('', taken)).toBe('Sheet2');
    expect(sanitize('', taken)).toBe('Sheet3');
  });

  it('truncates to 31 characters', () => {
    const long = 'A'.repeat(40);
    expect(sanitize(long)).toBe('A'.repeat(MAX_SHEET_NAME_LENGTH));
    // Cutting at 31 leaves a trailing apostrophe, which Excel rejects, so the edges are cleaned again after the cut.
    expect(sanitize(`${'B'.repeat(30)}'C`)).toBe('B'.repeat(30));
  });

  it('never splits a surrogate pair when truncating', () => {
    expect(sanitize(`${'x'.repeat(29)}🚀`)).toBe(`${'x'.repeat(29)}🚀`);
    const dropped = sanitize(`${'x'.repeat(30)}🚀`);
    expect(dropped).toBe('x'.repeat(30));
    expect([...dropped].length).toBe(30);
  });

  it('dedupes case-insensitively with a numbered suffix', () => {
    const taken = new Set<string>();
    expect(sanitize('Data', taken)).toBe('Data');
    expect(sanitize('data', taken)).toBe('data (2)');
    expect(sanitize('DATA', taken)).toBe('DATA (3)');
    expect([...taken]).toEqual(['data', 'data (2)', 'data (3)']);
  });

  it('fits the suffix inside the 31-character budget', () => {
    const taken = new Set<string>();
    const first = sanitize('A'.repeat(31), taken);
    const second = sanitize('A'.repeat(31), taken);
    expect(first).toBe('A'.repeat(31));
    expect(second).toBe(`${'A'.repeat(27)} (2)`);
    expect(second).toHaveLength(MAX_SHEET_NAME_LENGTH);
  });

  it('does not leave a double space before the suffix', () => {
    const taken = new Set<string>();
    sanitize(`${'A'.repeat(26)} name`, taken);
    expect(sanitize(`${'A'.repeat(26)} name`, taken)).toBe(`${'A'.repeat(26)} (2)`);
  });

  it('records the name it returned, preserving the original casing', () => {
    const taken = new Set<string>();
    expect(sanitize('MixedCase', taken)).toBe('MixedCase');
    expect(taken.has('mixedcase')).toBe(true);
    expect(taken.has('MixedCase')).toBe(false);
  });

  it('accepts a name that sanitizes to underscores only', () => {
    expect(sanitize('[]')).toBe('__');
  });

  it('EC-XML-CONTROL-CHARS-METADATA: replaces the characters XML forbids and unpaired surrogates with underscores', () => {
    expect(sanitize('Q1\u0001Report')).toBe('Q1_Report');
    expect(sanitize('a\u0000b\u0008c\u000Bd\u000Ce\u001Ff')).toBe('a_b_c_d_e_f');
    expect(sanitize('non\uFFFEchar\uFFFF')).toBe('non_char_');
    expect(sanitize('lone \uD83D high')).toBe('lone _ high');
    expect(sanitize('lone \uDE00 low')).toBe('lone _ low');
    expect(sanitize('🚀 Launch'), 'a complete pair is a real character').toBe('🚀 Launch');
  });

  it("EC-XML-CONTROL-CHARS-METADATA: replaces tab, LF and CR too, which Excel's rename box refuses", () => {
    expect(sanitize('Line\nBreak')).toBe('Line_Break');
    expect(sanitize('Tab\tSeparated\r')).toBe('Tab_Separated_');
    expect(isValidSheetName('Tab\tName')).toBe(false);
  });

  it('throws INVALID_SHEET_NAME when the input is not a string', () => {
    try {
      sanitizeSheetName(undefined as unknown as string, new Set());
      expect.unreachable('expected an XlsxError');
    } catch (error) {
      expect(isXlsxError(error) && error.code).toBe('INVALID_SHEET_NAME');
    }
  });
});

describe('isValidSheetName', () => {
  it('is true only for names sanitizing would leave alone', () => {
    expect(isValidSheetName('Data')).toBe(true);
    expect(isValidSheetName("It's fine")).toBe(true);
    expect(isValidSheetName('')).toBe(false);
    expect(isValidSheetName('a/b')).toBe(false);
    expect(isValidSheetName("'quoted'")).toBe(false);
    expect(isValidSheetName('A'.repeat(32))).toBe(false);
    expect(isValidSheetName(' padded ')).toBe(false);
  });

  it('EC-XML-CONTROL-CHARS-METADATA: is false for control characters and unpaired surrogates', () => {
    expect(isValidSheetName('Q1\u0001Report')).toBe(false);
    expect(isValidSheetName('x\uFFFF')).toBe(false);
    expect(isValidSheetName('x\uD800')).toBe(false);
    expect(isValidSheetName('🚀')).toBe(true);
  });
});

describe('quoteSheetName', () => {
  it('leaves a plain identifier alone', () => {
    expect(quoteSheetName('Data')).toBe('Data');
    expect(quoteSheetName('_Data_1')).toBe('_Data_1');
  });

  it('quotes anything else and doubles inner apostrophes', () => {
    expect(quoteSheetName('My Sheet')).toBe("'My Sheet'");
    expect(quoteSheetName('1Sheet')).toBe("'1Sheet'");
    expect(quoteSheetName("It's a name")).toBe("'It''s a name'");
    expect(quoteSheetName(SHEET_LONG_NAME)).toBe("'It''s a very long sheet name 001'");
    expect(quoteSheetName('Ünïcödé')).toBe("'Ünïcödé'");
  });

  it('builds the reference a defined name needs', () => {
    expect(`${quoteSheetName(SHEET_LONG_NAME)}!$A$1:$C$3`).toBe("'It''s a very long sheet name 001'!$A$1:$C$3");
  });
});
