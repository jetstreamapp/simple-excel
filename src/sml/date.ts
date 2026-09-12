/**
 * Excel serials are integer day counts with the time of day as a fraction, so every conversion here is integer
 * arithmetic on the proleptic Gregorian calendar (Howard Hinnant's days-from-civil / civil-from-days). JS `Date` is
 * used only at the two boundary helpers: a timezone or a DST gap must never reach the serial math (EC-DATE-DST-GAP,
 * EC-DATE-HISTORICAL-TZ-OFFSET).
 */

/** Wall-clock components; `month` is 1-12. */
export interface DateComponents {
  readonly year: number;
  readonly month: number;
  readonly day: number;
  readonly hour: number;
  readonly minute: number;
  readonly second: number;
  readonly millisecond: number;
}

/** Serial for 1900-03-01 in the 1900 system: the first serial after Excel's fake 1900-02-29 (serial 60). */
export const SERIAL_1900_03_01: number = 61;
/** Difference between the two date systems: 1904-01-01 is serial 0 in the 1904 system and 1462 in the 1900 system. */
export const SERIAL_1904_OFFSET: number = 1462;

const MS_PER_DAY = 86_400_000;
const MONTH_LENGTHS = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];

/** Day number of the 1900 system's epoch. Serials >= 61 are exactly this many days plus the serial. */
const EPOCH_1900_DAY = daysFromCivil(1899, 12, 30);
const EPOCH_1904_DAY = daysFromCivil(1904, 1, 1);

/** Days from 1970-01-01 to a civil date, valid for any proleptic Gregorian year. */
function daysFromCivil(year: number, month: number, day: number): number {
  const shiftedYear = month <= 2 ? year - 1 : year;
  const era = Math.floor(shiftedYear / 400);
  const yearOfEra = shiftedYear - era * 400;
  const dayOfYear = Math.floor((153 * (month + (month > 2 ? -3 : 9)) + 2) / 5) + day - 1;
  const dayOfEra = yearOfEra * 365 + Math.floor(yearOfEra / 4) - Math.floor(yearOfEra / 100) + dayOfYear;
  return era * 146_097 + dayOfEra - 719_468;
}

/** Inverse of `daysFromCivil`. */
function civilFromDays(days: number): { year: number; month: number; day: number } {
  const shifted = days + 719_468;
  const era = Math.floor(shifted / 146_097);
  const dayOfEra = shifted - era * 146_097;
  const yearOfEra = Math.floor(
    (dayOfEra - Math.floor(dayOfEra / 1460) + Math.floor(dayOfEra / 36_524) - Math.floor(dayOfEra / 146_096)) / 365,
  );
  const dayOfYear = dayOfEra - (365 * yearOfEra + Math.floor(yearOfEra / 4) - Math.floor(yearOfEra / 100));
  const monthIndex = Math.floor((5 * dayOfYear + 2) / 153);
  const day = dayOfYear - Math.floor((153 * monthIndex + 2) / 5) + 1;
  const month = monthIndex + (monthIndex < 10 ? 3 : -9);
  const year = yearOfEra + era * 400 + (month <= 2 ? 1 : 0);
  return { year, month, day };
}

function isLeapYear(year: number): boolean {
  return (year % 4 === 0 && year % 100 !== 0) || year % 400 === 0;
}

function daysInMonth(year: number, month: number): number {
  if (month === 2) {
    return isLeapYear(year) ? 29 : 28;
  }
  return MONTH_LENGTHS[month - 1] ?? 0;
}

/** Every field an integer and in range. 1900 is not a leap year here, so 1900-02-29 is rejected on the way in. */
function isValidComponents(components: DateComponents): boolean {
  const { year, month, day, hour, minute, second, millisecond } = components;
  if (!Number.isInteger(year) || !Number.isInteger(month) || !Number.isInteger(day)) {
    return false;
  }
  if (!Number.isInteger(hour) || !Number.isInteger(minute) || !Number.isInteger(second) || !Number.isInteger(millisecond)) {
    return false;
  }
  if (month < 1 || month > 12 || day < 1 || day > daysInMonth(year, month)) {
    return false;
  }
  return hour >= 0 && hour < 24 && minute >= 0 && minute < 60 && second >= 0 && second < 60 && millisecond >= 0 && millisecond < 1000;
}

function millisecondsOfDay(components: DateComponents): number {
  return ((components.hour * 60 + components.minute) * 60 + components.second) * 1000 + components.millisecond;
}

/**
 * One division of two exact integers, so the serial is the correctly rounded double for the instant - the same value
 * the fixture generators compute, and close enough that `componentsFromSerial` rounds back to the same millisecond.
 */
function serialOf(days: number, msOfDay: number): number {
  return (days * MS_PER_DAY + msOfDay) / MS_PER_DAY;
}

function componentsOf(dayNumber: number, msOfDay: number): DateComponents {
  const { year, month, day } = civilFromDays(dayNumber);
  return {
    year,
    month,
    day,
    hour: Math.floor(msOfDay / 3_600_000),
    minute: Math.floor(msOfDay / 60_000) % 60,
    second: Math.floor(msOfDay / 1000) % 60,
    millisecond: msOfDay % 1000,
  };
}

/**
 * Excel serial (days since the epoch, fraction = time of day) for a wall-clock instant. Rounded to the nearest
 * millisecond. Uses the 1900 system unless `date1904`; dates before the epoch (1900-01-01 or 1904-01-01) and the
 * non-existent 1900-02-29 return null. Time-only inputs (year 1899, month 12, day 30) yield a fraction in [0, 1).
 */
export function serialFromComponents(components: DateComponents, date1904: boolean): number | null {
  if (!isValidComponents(components)) {
    return null;
  }
  const msOfDay = millisecondsOfDay(components);
  // Time-only values ride on Excel's "Jan 0 1900" sentinel day in both date systems: the serial is the fraction alone
  // (EC-DATE-TIME-ONLY-NEGATIVE-SERIAL - writing them as a date on 1899-12-30 makes Excel show ########).
  if (components.year === 1899 && components.month === 12 && components.day === 30) {
    return msOfDay / MS_PER_DAY;
  }
  const dayNumber = daysFromCivil(components.year, components.month, components.day);
  if (date1904) {
    const days = dayNumber - EPOCH_1904_DAY;
    return days < 0 ? null : serialOf(days, msOfDay);
  }
  const daysSinceEpoch = dayNumber - EPOCH_1900_DAY;
  // Serial 0 is Excel's "1/0/1900" placeholder, not 1899-12-31: pre-1900 dates are not Excel dates (EC-DATE-PRE-1900).
  if (daysSinceEpoch < 2) {
    return null;
  }
  // Serial 60 is the fake 1900-02-29, so 1900-01-01..1900-02-28 are serials 1..59, one below the real day count,
  // and only 1900-03-01 onwards lines up with the 1899-12-30 epoch (EC-DATE-1900-LEAP-BUG).
  return serialOf(daysSinceEpoch < SERIAL_1900_03_01 ? daysSinceEpoch - 1 : daysSinceEpoch, msOfDay);
}

/**
 * Inverse of `serialFromComponents`. Serial 60 in the 1900 system (Excel's fake leap day) maps to 1900-03-01 and
 * serials below 61 shift by one day, matching what every other reader does. Returns null for negative serials
 * (Excel displays those as `#####`) and non-finite input. Fractions are rounded to the nearest millisecond.
 */
export function componentsFromSerial(serial: number, date1904: boolean): DateComponents | null {
  if (!Number.isFinite(serial) || serial < 0) {
    return null;
  }
  let wholeDays = Math.floor(serial);
  let msOfDay = Math.round((serial - wholeDays) * MS_PER_DAY);
  if (msOfDay >= MS_PER_DAY) {
    msOfDay -= MS_PER_DAY;
    wholeDays++;
  }
  if (date1904) {
    return componentsOf(EPOCH_1904_DAY + wholeDays, msOfDay);
  }
  // Serials below 1 are time-only values and stay on the 1899-12-30 sentinel day (EC-TIME-ONLY-HAS-DATE-PART);
  // 1..60 shift forward past the fake 1900-02-29, which makes serial 60 and serial 61 both 1900-03-01.
  let daysSinceEpoch = wholeDays;
  if (wholeDays >= 1 && wholeDays < SERIAL_1900_03_01) {
    daysSinceEpoch = wholeDays + 1;
  }
  return componentsOf(EPOCH_1900_DAY + daysSinceEpoch, msOfDay);
}

/** Build a JS Date whose local (`'local'`) or UTC (`'utc'`) fields equal the components. */
export function dateFromComponents(components: DateComponents, fields: 'local' | 'utc'): Date {
  const { year, month, day, hour, minute, second, millisecond } = components;
  // Years 0-99 are mapped into the 1900s by both constructors, so set them again explicitly.
  if (fields === 'utc') {
    const date = new Date(Date.UTC(year, month - 1, day, hour, minute, second, millisecond));
    if (year >= 0 && year < 100) {
      date.setUTCFullYear(year);
    }
    return date;
  }
  const date = new Date(year, month - 1, day, hour, minute, second, millisecond);
  if (year >= 0 && year < 100) {
    date.setFullYear(year);
  }
  return date;
}

/** Read the local or UTC fields of a Date as components. Invalid dates return null. */
export function componentsFromDate(date: Date, fields: 'local' | 'utc'): DateComponents | null {
  if (Number.isNaN(date.getTime())) {
    return null;
  }
  if (fields === 'utc') {
    return {
      year: date.getUTCFullYear(),
      month: date.getUTCMonth() + 1,
      day: date.getUTCDate(),
      hour: date.getUTCHours(),
      minute: date.getUTCMinutes(),
      second: date.getUTCSeconds(),
      millisecond: date.getUTCMilliseconds(),
    };
  }
  return {
    year: date.getFullYear(),
    month: date.getMonth() + 1,
    day: date.getDate(),
    hour: date.getHours(),
    minute: date.getMinutes(),
    second: date.getSeconds(),
    millisecond: date.getMilliseconds(),
  };
}

// Strict `t="d"` cells, which Excel writes with up to 17 fractional digits (EC-STRICT-ISO-DATE-PRECISION). A trailing
// zone designator is accepted and ignored: the library's contract is wall clock in, wall clock out (ADR-003).
const ISO_DATE_TIME = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2})(?:\.(\d+))?)?)?(?:Z|[+-]\d{2}:?\d{2})?$/;
const ISO_TIME = /^(\d{2}):(\d{2})(?::(\d{2})(?:\.(\d+))?)?(?:Z|[+-]\d{2}:?\d{2})?$/;

/**
 * Parse an ISO 8601 date/time from a Strict `t="d"` cell: `YYYY-MM-DD`, `YYYY-MM-DDThh:mm:ss` with any number of
 * fractional digits, or a bare `hh:mm:ss` time (day 1899-12-30). Returns null when malformed.
 */
export function componentsFromIso(text: string): DateComponents | null {
  const trimmed = text.trim();
  const dateTime = ISO_DATE_TIME.exec(trimmed);
  if (dateTime !== null) {
    return fromIsoParts(Number(dateTime[1]), Number(dateTime[2]), Number(dateTime[3]), dateTime[4], dateTime[5], dateTime[6], dateTime[7]);
  }
  const time = ISO_TIME.exec(trimmed);
  if (time !== null) {
    return fromIsoParts(1899, 12, 30, time[1], time[2], time[3], time[4]);
  }
  return null;
}

function fromIsoParts(
  year: number,
  month: number,
  day: number,
  hourText: string | undefined,
  minuteText: string | undefined,
  secondText: string | undefined,
  fractionDigits: string | undefined,
): DateComponents | null {
  const hour = hourText === undefined ? 0 : Number(hourText);
  const minute = minuteText === undefined ? 0 : Number(minuteText);
  const second = secondText === undefined ? 0 : Number(secondText);
  const candidate: DateComponents = { year, month, day, hour, minute, second, millisecond: 0 };
  if (!isValidComponents(candidate)) {
    return null;
  }
  const fraction = fractionDigits === undefined ? 0 : Number(`0.${fractionDigits}`);
  // 17 fractional digits round up to a whole second often enough that the carry has to walk the calendar:
  // 2024-03-10T09:29:59.9999995809048412 is 09:30:00.000.
  const msOfDay = millisecondsOfDay(candidate) + Math.round(fraction * 1000);
  const dayNumber = daysFromCivil(year, month, day);
  if (msOfDay >= MS_PER_DAY) {
    return componentsOf(dayNumber + 1, msOfDay - MS_PER_DAY);
  }
  return componentsOf(dayNumber, msOfDay);
}
