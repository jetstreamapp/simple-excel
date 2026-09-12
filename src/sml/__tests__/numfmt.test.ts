import { describe, expect, it } from 'vitest';
import { builtinIdForCode, BUILTIN_NUMFMTS, DEFAULT_DATE_FORMAT, isBuiltinDateId, isDateFormatCode } from '../numfmt';

/** ECMA-376 §18.8.30, "All Languages"; ids 5-8, 23-26 and 41-44 are locale-defined and have no fixed code. */
const EXPECTED_BUILTINS: readonly (readonly [number, string])[] = [
  [0, 'General'],
  [1, '0'],
  [2, '0.00'],
  [3, '#,##0'],
  [4, '#,##0.00'],
  [9, '0%'],
  [10, '0.00%'],
  [11, '0.00E+00'],
  [12, '# ?/?'],
  [13, '# ??/??'],
  [14, 'mm-dd-yy'],
  [15, 'd-mmm-yy'],
  [16, 'd-mmm'],
  [17, 'mmm-yy'],
  [18, 'h:mm AM/PM'],
  [19, 'h:mm:ss AM/PM'],
  [20, 'h:mm'],
  [21, 'h:mm:ss'],
  [22, 'm/d/yy h:mm'],
  [37, '#,##0 ;(#,##0)'],
  [38, '#,##0 ;[Red](#,##0)'],
  [39, '#,##0.00;(#,##0.00)'],
  [40, '#,##0.00;[Red](#,##0.00)'],
  [45, 'mm:ss'],
  [46, '[h]:mm:ss'],
  [47, 'mmss.0'],
  [48, '##0.0E+0'],
  [49, '@'],
];

/** Built-in ids every reader renders as a date or time, and the ids around them that it must not. */
const DATE_IDS: readonly number[] = [14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 30, 36, 45, 46, 47, 50, 55, 58];
const NON_DATE_IDS: readonly number[] = [0, 1, 2, 3, 4, 9, 10, 11, 12, 13, 13, 23, 26, 37, 38, 39, 40, 44, 48, 49, 59, 164];

/** Custom codes that appear in the corpus fixtures' styles.xml, plus the traps the heuristic has to survive. */
const DATE_CODES: readonly string[] = [
  'yyyy-mm-dd',
  'hh:mm:ss',
  'yyyy-mm-dd hh:mm:ss',
  'yyyy-mm-dd h:mm:ss',
  'yyyy"-"mm"-"dd"T"hh:mm:ss',
  String.raw`yyyy\-mm\-dd`,
  String.raw`yyyy\-mm\-dd\ hh:mm:ss`,
  String.raw`yyyy\-mm\-dd\Thh:mm:ss.000`,
  String.raw`[$-409]d\-mmm\-yy;@`,
  'mmm-yy',
  'hh:mm:ss.000',
  '[h]:mm:ss',
  '[mm]:ss',
  '[SS]',
  'h:mm AM/PM',
  'h:mm A/P',
  '"上午/下午 "hh"時"mm"分"ss"秒 "',
  DEFAULT_DATE_FORMAT,
];

const NON_DATE_CODES: readonly string[] = [
  'General',
  'general',
  '@',
  '0',
  '0.00',
  '0.0%',
  '#,##0',
  '#,##0%',
  '0.00E+00',
  '##0.0E+0',
  '$#,##0.00',
  String.raw`\$#,##0.00`,
  '"$"#,##0.00',
  '"USD" #,##0.00;"USD" -#,##0.00',
  '"months" 0',
  '#,##0 ;[Red](#,##0)',
  '[Red]0.00',
  '[$€-2]#,##0.00',
  '[>100]0.0;[<=100]0.000',
  String.raw`_(* #,##0.00_);_(* \(#,##0.00\);_(* "-"??_);_(@_)`,
  '',
];

describe('BUILTIN_NUMFMTS', () => {
  it('holds exactly the fixed-code built-ins', () => {
    expect([...BUILTIN_NUMFMTS.entries()]).toEqual(EXPECTED_BUILTINS.map(entry => [...entry]));
  });

  it.each(EXPECTED_BUILTINS)('resolves %i back from its code', (id, code) => {
    expect(builtinIdForCode(code)).toBe(id);
  });

  it('returns -1 for codes that are not built-ins', () => {
    expect(builtinIdForCode('yyyy-mm-dd')).toBe(-1);
    expect(builtinIdForCode(DEFAULT_DATE_FORMAT)).toBe(-1);
    expect(builtinIdForCode('0.000')).toBe(-1);
    expect(builtinIdForCode('')).toBe(-1);
    // The table is exact-match only: no trimming, no case folding.
    expect(builtinIdForCode('general')).toBe(-1);
    expect(builtinIdForCode(' General')).toBe(-1);
  });
});

describe('isBuiltinDateId (EC-DATE-DETECTION-VIA-NUMFMT)', () => {
  it.each(DATE_IDS)('treats %i as a date', id => {
    expect(isBuiltinDateId(id)).toBe(true);
  });

  it.each(NON_DATE_IDS)('does not treat %i as a date', id => {
    expect(isBuiltinDateId(id)).toBe(false);
  });
});

describe('isDateFormatCode (EC-DATE-DETECTION-VIA-NUMFMT)', () => {
  it.each(DATE_CODES)('reads %s as a date format', code => {
    expect(isDateFormatCode(code)).toBe(true);
  });

  it.each(NON_DATE_CODES)('reads %s as a plain number format', code => {
    expect(isDateFormatCode(code)).toBe(false);
  });

  it('agrees with isBuiltinDateId on every built-in code', () => {
    for (const [id, code] of BUILTIN_NUMFMTS) {
      expect(isDateFormatCode(code), `built-in ${id} (${code})`).toBe(isBuiltinDateId(id));
    }
  });
});
