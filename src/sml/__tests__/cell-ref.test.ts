import { describe, expect, it } from 'vitest';
import { FEATURES } from '../../../fixtures/canonical/canonical.mjs';
import { isXlsxError } from '../../errors';
import {
  columnIndexOf,
  columnIndexOfRef,
  columnLetters,
  formatRange,
  formatRef,
  MAX_COLUMNS,
  MAX_ROWS,
  parseRange,
  parseRef,
} from '../cell-ref';

describe('columnLetters', () => {
  it('maps the boundaries of every letter width', () => {
    expect(columnLetters(0)).toBe('A');
    expect(columnLetters(25)).toBe('Z');
    expect(columnLetters(26)).toBe('AA');
    expect(columnLetters(51)).toBe('AZ');
    expect(columnLetters(52)).toBe('BA');
    expect(columnLetters(701)).toBe('ZZ');
    expect(columnLetters(702)).toBe('AAA');
    expect(columnLetters(MAX_COLUMNS - 1)).toBe('XFD');
  });

  it('returns the same value from the cache on repeated calls', () => {
    expect(columnLetters(16_000)).toBe('WQK');
    expect(columnLetters(16_000)).toBe('WQK');
    expect(columnLetters(16_001)).toBe('WQL');
  });

  it('throws ROW_OUT_OF_RANGE outside 0..16383', () => {
    for (const invalid of [-1, MAX_COLUMNS, MAX_COLUMNS + 1, 1.5, Number.NaN, Number.POSITIVE_INFINITY]) {
      expect(() => columnLetters(invalid)).toThrow(/columns a sheet can hold/);
    }
    try {
      columnLetters(MAX_COLUMNS);
      expect.unreachable('expected an XlsxError');
    } catch (error) {
      expect(isXlsxError(error) && error.code).toBe('ROW_OUT_OF_RANGE');
    }
  });
});

describe('columnIndexOf', () => {
  it('inverts columnLetters over every column', () => {
    for (let index = 0; index < MAX_COLUMNS; index++) {
      expect(columnIndexOf(columnLetters(index))).toBe(index);
    }
  });

  it('rejects anything that is not a column', () => {
    for (const invalid of ['', 'a', 'aa', 'A1', '1', '$A', 'XFE', 'ZZZ', 'AAAA', 'A B', 'Ä']) {
      expect(columnIndexOf(invalid), invalid).toBe(-1);
    }
  });
});

describe('formatRef', () => {
  it('formats the corners of the sheet', () => {
    expect(formatRef(0, 0)).toBe('A1');
    expect(formatRef(29, 1)).toBe('B30');
    expect(formatRef(MAX_ROWS - 1, MAX_COLUMNS - 1)).toBe('XFD1048576');
  });

  it('throws ROW_OUT_OF_RANGE past the last row', () => {
    expect(() => formatRef(MAX_ROWS, 0)).toThrow(/rows a sheet can hold/);
    expect(() => formatRef(-1, 0)).toThrow(/rows a sheet can hold/);
    expect(() => formatRef(0.5, 0)).toThrow(/rows a sheet can hold/);
    expect(() => formatRef(0, MAX_COLUMNS)).toThrow(/columns a sheet can hold/);
  });
});

describe('parseRef', () => {
  it('parses plain and anchored references', () => {
    expect(parseRef('A1')).toEqual({ row: 0, col: 0 });
    expect(parseRef('B30')).toEqual({ row: 29, col: 1 });
    expect(parseRef('$AB$12')).toEqual({ row: 11, col: 27 });
    expect(parseRef('$AB12')).toEqual({ row: 11, col: 27 });
    expect(parseRef('AB$12')).toEqual({ row: 11, col: 27 });
    expect(parseRef('XFD1048576')).toEqual({ row: MAX_ROWS - 1, col: MAX_COLUMNS - 1 });
  });

  it('round trips every formatted reference', () => {
    for (const [row, col] of [
      [0, 0],
      [9, 25],
      [99, 26],
      [1000, 701],
      [MAX_ROWS - 1, MAX_COLUMNS - 1],
    ] as const) {
      expect(parseRef(formatRef(row, col))).toEqual({ row, col });
    }
  });

  it('returns null for malformed or out-of-range references', () => {
    for (const invalid of ['', 'A', '1', 'A0', 'a1', 'A1B', 'A 1', 'AB', '$', '$$A1', 'XFE1', 'XFD1048577', 'A1.5', 'A-1']) {
      expect(parseRef(invalid), invalid).toBeNull();
    }
  });
});

describe('columnIndexOfRef', () => {
  it('reads the column without allocating a parse result', () => {
    expect(columnIndexOfRef('A1')).toBe(0);
    expect(columnIndexOfRef('AB12')).toBe(27);
    expect(columnIndexOfRef('$AB$12')).toBe(27);
    expect(columnIndexOfRef('XFD1048576')).toBe(MAX_COLUMNS - 1);
  });

  it('agrees with parseRef on every valid reference', () => {
    for (let index = 0; index < MAX_COLUMNS; index += 7) {
      const ref = formatRef(index % MAX_ROWS, index);
      expect(columnIndexOfRef(ref)).toBe(parseRef(ref)?.col);
    }
  });

  it('returns -1 when there is no column or no row part', () => {
    for (const invalid of ['', 'A', 'AB', '1', '12A', 'a1', 'XFE1', 'AAAA1', '$']) {
      expect(columnIndexOfRef(invalid), invalid).toBe(-1);
    }
  });
});

describe('parseRange', () => {
  it('parses a range and a single cell', () => {
    expect(parseRange('A1:C3')).toEqual({ start: { row: 0, col: 0 }, end: { row: 2, col: 2 } });
    expect(parseRange('B2')).toEqual({ start: { row: 1, col: 1 }, end: { row: 1, col: 1 } });
  });

  it('normalizes reversed corners', () => {
    expect(parseRange('C3:A1')).toEqual(parseRange('A1:C3'));
    expect(parseRange('C1:A3')).toEqual(parseRange('A1:C3'));
    expect(parseRange('$C$3:$A$1')).toEqual(parseRange('A1:C3'));
  });

  it('parses the canonical Features sheet ranges', () => {
    expect(parseRange(FEATURES.autoFilter)).toEqual({ start: { row: 1, col: 0 }, end: { row: 11, col: 2 } });
    const [merge] = FEATURES.merges;
    expect(merge).toBeDefined();
    const mergeRange = `${formatRef(merge!.s.r, merge!.s.c)}:${formatRef(merge!.e.r, merge!.e.c)}`;
    expect(parseRange(mergeRange)).toEqual({ start: { row: 0, col: 0 }, end: { row: 0, col: 2 } });
  });

  it('returns null when either end is malformed', () => {
    for (const invalid of ['', ':', 'A1:', ':C3', 'A1:C3:D4', 'A:C', 'A1:XFE3', 'A1 : C3']) {
      expect(parseRange(invalid), invalid).toBeNull();
    }
  });
});

describe('formatRange', () => {
  it('formats and normalizes 0-based corners', () => {
    expect(formatRange(0, 0, 0, 2)).toBe('A1:C1');
    expect(formatRange(4, 3, 1, 1)).toBe('B2:D5');
    expect(formatRange(0, 0, 0, 0)).toBe('A1:A1');
  });

  it('rejects corners outside the sheet', () => {
    expect(() => formatRange(0, 0, 1_048_576, 0)).toThrow();
  });
});
