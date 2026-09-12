import { describe, expect, it } from 'vitest';
import {
  COLUMNS,
  components as canonicalComponents,
  isTemporal,
  ROW_COUNT,
  toSerial,
  type DateParts,
  type TypedValue,
} from '../../../fixtures/canonical/canonical.mjs';
import {
  componentsFromDate,
  componentsFromIso,
  componentsFromSerial,
  dateFromComponents,
  serialFromComponents,
  SERIAL_1900_03_01,
  SERIAL_1904_OFFSET,
  type DateComponents,
} from '../date';

const parts = (year: number, month: number, day: number, hour = 0, minute = 0, second = 0, millisecond = 0): DateComponents => ({
  year,
  month,
  day,
  hour,
  minute,
  second,
  millisecond,
});

const fromCanonical = ({ y, m, d, hh, mm, ss, ms }: DateParts): DateComponents => parts(y, m, d, hh, mm, ss, ms);

/** The fixture helpers take the temporal shape of the typed encoding; `isTemporal` has already vouched for it. */
const asTemporal = (typed: TypedValue): { $date?: string; $datetime?: string; $time?: string } =>
  typed as { $date?: string; $datetime?: string; $time?: string };

/** Every distinct Date/Time/DateTime value in the canonical dataset, in the order the columns declare them. */
function canonicalTemporalValues(): { label: string; typed: TypedValue }[] {
  const seen = new Set<string>();
  const values: { label: string; typed: TypedValue }[] = [];
  for (const column of COLUMNS) {
    for (let row = 0; row < ROW_COUNT; row++) {
      const typed = column.value(row);
      if (!isTemporal(typed)) {
        continue;
      }
      const label = JSON.stringify(typed);
      if (seen.has(label)) {
        continue;
      }
      seen.add(label);
      values.push({ label, typed });
    }
  }
  return values;
}

const TEMPORAL_VALUES = canonicalTemporalValues();

/**
 * The one place we deliberately differ from the fixture helper: `toSerial` reproduces what every other writer does
 * with a pre-1900 date (serial 0), while EC-DATE-PRE-1900 says such a date is not an Excel date at all, so the
 * writer must fall back to text. `serialFromComponents` reports that with null.
 */
const PRE_EPOCH_LABEL = JSON.stringify({ $date: '1899-12-31' });

describe('canonical dataset: serials', () => {
  it('covers the Date, Time and DateTime columns', () => {
    expect(TEMPORAL_VALUES.length).toBeGreaterThanOrEqual(22);
  });

  for (const { label, typed } of TEMPORAL_VALUES) {
    it(`${label} matches the fixture serial and round trips`, () => {
      const expected = toSerial(asTemporal(typed));
      const components = fromCanonical(canonicalComponents(asTemporal(typed)));
      const serial = serialFromComponents(components, false);

      if (label === PRE_EPOCH_LABEL) {
        expect(expected).toBe(0);
        expect(serial).toBeNull();
        return;
      }

      expect(serial).toBe(expected);
      expect(componentsFromSerial(serial as number, false)).toEqual(components);
    });
  }

  it('shifts every in-range value by exactly 1,462 days in the 1904 system', () => {
    for (const { label, typed } of TEMPORAL_VALUES) {
      const components = fromCanonical(canonicalComponents(asTemporal(typed)));
      const serial1900 = serialFromComponents(components, false);
      const serial1904 = serialFromComponents(components, true);
      const isTimeOnly = components.year === 1899;
      if (isTimeOnly) {
        expect(serial1904, label).toBe(serial1900);
      } else if (serial1900 !== null && serial1900 >= SERIAL_1904_OFFSET) {
        expect(serial1904, label).toBe(serial1900 - SERIAL_1904_OFFSET);
      } else {
        expect(serial1904, label).toBeNull();
      }
    }
  });
});

describe('serialFromComponents (1900 system)', () => {
  it('EC-DATE-1900-LEAP-BUG: 1900-02-28 is 59, the fake 1900-02-29 is rejected, 1900-03-01 is 61', () => {
    expect(serialFromComponents(parts(1900, 1, 1), false)).toBe(1);
    expect(serialFromComponents(parts(1900, 2, 28), false)).toBe(59);
    expect(serialFromComponents(parts(1900, 2, 29), false)).toBeNull();
    expect(serialFromComponents(parts(1900, 3, 1), false)).toBe(61);
    expect(serialFromComponents(parts(1900, 1, 1, 12), false)).toBe(1.5);
  });

  it('EC-DATE-PRE-1900: 1899-12-31 and earlier are not Excel dates', () => {
    expect(serialFromComponents(parts(1899, 12, 31), false)).toBeNull();
    expect(serialFromComponents(parts(1899, 12, 31, 23, 59, 59, 999), false)).toBeNull();
    expect(serialFromComponents(parts(1899, 12, 29), false)).toBeNull();
    expect(serialFromComponents(parts(1700, 1, 1), false)).toBeNull();
  });

  it('EC-DATE-TIME-ONLY-NEGATIVE-SERIAL: the 1899-12-30 sentinel day is the fraction alone', () => {
    expect(serialFromComponents(parts(1899, 12, 30), false)).toBe(0);
    expect(serialFromComponents(parts(1899, 12, 30, 12), false)).toBe(0.5);
    expect(serialFromComponents(parts(1899, 12, 30, 12, 34, 56, 789), false)).toBe(45_296_789 / 86_400_000);
    expect(serialFromComponents(parts(1899, 12, 30, 23, 59, 59, 999), false)).toBeLessThan(1);
    expect(serialFromComponents(parts(1899, 12, 30), true)).toBe(0);
  });

  it('matches the documented serials across the range', () => {
    expect(serialFromComponents(parts(1904, 1, 1), false)).toBe(SERIAL_1904_OFFSET);
    expect(serialFromComponents(parts(1970, 1, 1), false)).toBe(25_569);
    expect(serialFromComponents(parts(2024, 2, 29, 12), false)).toBe(45_351.5);
    expect(serialFromComponents(parts(9999, 12, 31), false)).toBe(2_958_465);
  });

  it('rejects invalid components', () => {
    expect(serialFromComponents(parts(2024, 13, 1), false)).toBeNull();
    expect(serialFromComponents(parts(2024, 0, 1), false)).toBeNull();
    expect(serialFromComponents(parts(2024, 1, 0), false)).toBeNull();
    expect(serialFromComponents(parts(2023, 2, 29), false)).toBeNull();
    expect(serialFromComponents(parts(2024, 4, 31), false)).toBeNull();
    expect(serialFromComponents(parts(2024, 1, 1, 24), false)).toBeNull();
    expect(serialFromComponents(parts(2024, 1, 1, 0, 60), false)).toBeNull();
    expect(serialFromComponents(parts(2024, 1, 1, 0, 0, 60), false)).toBeNull();
    expect(serialFromComponents(parts(2024, 1, 1, 0, 0, 0, 1000), false)).toBeNull();
    expect(serialFromComponents(parts(2024.5, 1, 1), false)).toBeNull();
    expect(serialFromComponents(parts(Number.NaN, 1, 1), false)).toBeNull();
  });
});

describe('serialFromComponents (1904 system)', () => {
  it('EC-DATE-1904: serial 0 is 1904-01-01 and there is no leap bug', () => {
    expect(serialFromComponents(parts(1904, 1, 1), true)).toBe(0);
    expect(serialFromComponents(parts(1904, 1, 2), true)).toBe(1);
    expect(serialFromComponents(parts(1904, 1, 1, 6), true)).toBe(0.25);
    expect(serialFromComponents(parts(2024, 2, 29, 12), true)).toBe(45_351.5 - SERIAL_1904_OFFSET);
    expect(serialFromComponents(parts(1903, 12, 31), true)).toBeNull();
    expect(serialFromComponents(parts(1900, 3, 1), true)).toBeNull();
  });
});

describe('componentsFromSerial', () => {
  it('EC-DATE-1900-LEAP-BUG: serial 60 is documented as 1900-03-01, the same as serial 61', () => {
    expect(componentsFromSerial(59, false)).toEqual(parts(1900, 2, 28));
    expect(componentsFromSerial(60, false)).toEqual(parts(1900, 3, 1));
    expect(componentsFromSerial(61, false)).toEqual(parts(1900, 3, 1));
    expect(componentsFromSerial(1, false)).toEqual(parts(1900, 1, 1));
    expect(componentsFromSerial(1.5, false)).toEqual(parts(1900, 1, 1, 12));
  });

  it('cannot round trip the fake leap day: serial 60 comes back as serial 61', () => {
    const fakeLeapDay = componentsFromSerial(60.5, false);
    expect(fakeLeapDay).toEqual(parts(1900, 3, 1, 12));
    expect(serialFromComponents(fakeLeapDay as DateComponents, false)).toBe(SERIAL_1900_03_01 + 0.5);
  });

  it('EC-TIME-ONLY-HAS-DATE-PART: serials below 1 land on the 1899-12-30 marker day', () => {
    expect(componentsFromSerial(0, false)).toEqual(parts(1899, 12, 30));
    expect(componentsFromSerial(0.5, false)).toEqual(parts(1899, 12, 30, 12));
    expect(componentsFromSerial(45_296_789 / 86_400_000, false)).toEqual(parts(1899, 12, 30, 12, 34, 56, 789));
  });

  it('decodes the documented serials', () => {
    expect(componentsFromSerial(SERIAL_1904_OFFSET, false)).toEqual(parts(1904, 1, 1));
    expect(componentsFromSerial(25_569, false)).toEqual(parts(1970, 1, 1));
    expect(componentsFromSerial(45_351.5, false)).toEqual(parts(2024, 2, 29, 12));
    expect(componentsFromSerial(2_958_465, false)).toEqual(parts(9999, 12, 31));
    expect(componentsFromSerial(0, true)).toEqual(parts(1904, 1, 1));
    expect(componentsFromSerial(45_351.5 - SERIAL_1904_OFFSET, true)).toEqual(parts(2024, 2, 29, 12));
  });

  it('rounds the fraction to the nearest millisecond without drifting', () => {
    expect(componentsFromSerial(45_000.123456, false)).toEqual(parts(2023, 3, 15, 2, 57, 46, 598));
    expect(componentsFromSerial(45_000 + 1 / 86_400_000, false)).toEqual(parts(2023, 3, 15, 0, 0, 0, 1));
    // A fraction that rounds up to a whole day carries into the next date instead of producing 24:00.
    expect(componentsFromSerial(45_000 + 86_399_999.7 / 86_400_000, false)).toEqual(parts(2023, 3, 16));
  });

  it('round trips every millisecond of a day', () => {
    const day = 45_351;
    for (let msOfDay = 0; msOfDay < 86_400_000; msOfDay += 997) {
      const serial = (day * 86_400_000 + msOfDay) / 86_400_000;
      const components = componentsFromSerial(serial, false);
      expect(components, String(msOfDay)).not.toBeNull();
      expect(serialFromComponents(components as DateComponents, false), String(msOfDay)).toBe(serial);
    }
  });

  it('returns null for negative and non-finite serials', () => {
    expect(componentsFromSerial(-1, false)).toBeNull();
    expect(componentsFromSerial(-0.5, false)).toBeNull();
    expect(componentsFromSerial(Number.NaN, false)).toBeNull();
    expect(componentsFromSerial(Number.POSITIVE_INFINITY, false)).toBeNull();
    expect(componentsFromSerial(-1, true)).toBeNull();
  });
});

describe('dateFromComponents and componentsFromDate', () => {
  it("'utc' never shifts the wall clock, whatever the host timezone is", () => {
    for (const { label, typed } of TEMPORAL_VALUES) {
      const components = fromCanonical(canonicalComponents(asTemporal(typed)));
      const date = dateFromComponents(components, 'utc');
      expect(componentsFromDate(date, 'utc'), label).toEqual(components);
    }
  });

  it('EC-DATE-DST-GAP: a wall clock inside the US spring-forward gap survives in utc mode', () => {
    const gap = parts(2024, 3, 10, 2, 30);
    const date = dateFromComponents(gap, 'utc');
    expect(date.getUTCHours()).toBe(2);
    expect(componentsFromDate(date, 'utc')).toEqual(gap);
    expect(serialFromComponents(gap, false)).toBe(45_361 + 2.5 / 24);
  });

  it("'local' reads back through the local getters", () => {
    const components = parts(2024, 1, 2, 15, 4, 5, 678);
    expect(componentsFromDate(dateFromComponents(components, 'local'), 'local')).toEqual(components);
  });

  it('keeps years below 100 out of the 1900s', () => {
    expect(dateFromComponents(parts(50, 1, 1), 'utc').getUTCFullYear()).toBe(50);
    expect(dateFromComponents(parts(50, 1, 1), 'local').getFullYear()).toBe(50);
    expect(componentsFromDate(dateFromComponents(parts(7, 6, 5), 'utc'), 'utc')).toEqual(parts(7, 6, 5));
  });

  it('returns null for an invalid Date', () => {
    expect(componentsFromDate(new Date(Number.NaN), 'utc')).toBeNull();
    expect(componentsFromDate(new Date('nope'), 'local')).toBeNull();
  });
});

describe('componentsFromIso', () => {
  it('parses the forms a Strict t="d" cell can hold', () => {
    expect(componentsFromIso('2024-02-29')).toEqual(parts(2024, 2, 29));
    expect(componentsFromIso('2024-02-29T12:00:00')).toEqual(parts(2024, 2, 29, 12));
    expect(componentsFromIso('2024-01-02T15:04')).toEqual(parts(2024, 1, 2, 15, 4));
    expect(componentsFromIso('2024-01-02T15:04:05.678')).toEqual(parts(2024, 1, 2, 15, 4, 5, 678));
    expect(componentsFromIso('  2024-01-02T15:04:05.678Z  ')).toEqual(parts(2024, 1, 2, 15, 4, 5, 678));
  });

  it('EC-STRICT-ISO-DATE-PRECISION: 17 fractional digits round to the nearest millisecond', () => {
    expect(componentsFromIso('2024-03-10T09:29:59.9999995809048412')).toEqual(parts(2024, 3, 10, 9, 30));
    expect(componentsFromIso('2024-03-10T09:29:59.99949')).toEqual(parts(2024, 3, 10, 9, 29, 59, 999));
    expect(componentsFromIso('2024-12-31T23:59:59.9999999')).toEqual(parts(2025, 1, 1));
  });

  it('EC-DATE-TIME-ONLY: a bare time lands on the 1899-12-30 marker day', () => {
    expect(componentsFromIso('19:34:56.789')).toEqual(parts(1899, 12, 30, 19, 34, 56, 789));
    expect(componentsFromIso('00:00:00')).toEqual(parts(1899, 12, 30));
    expect(componentsFromIso('23:59')).toEqual(parts(1899, 12, 30, 23, 59));
  });

  it('feeds serialFromComponents directly', () => {
    const iso = componentsFromIso('2024-02-29T12:00:00.000');
    expect(iso).not.toBeNull();
    expect(serialFromComponents(iso as DateComponents, false)).toBe(45_351.5);
  });

  it('returns null when the text is not a date', () => {
    for (const invalid of [
      '',
      'nope',
      '2024',
      '2024-1-2',
      '2024-13-01',
      '2024-02-30',
      '1900-02-29',
      '24:00:00',
      '12:60:00',
      '2024-01-02T15:04:05.678+05:00x',
      '45351.5',
    ]) {
      expect(componentsFromIso(invalid), invalid).toBeNull();
    }
  });
});
