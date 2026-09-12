/**
 * The canonical dataset: one small workbook whose every cell is a known trap, used as the ground truth
 * for every generator golden and every reader parity check. `canonical.json` and `canonical.csv` are
 * rendered from this module by `build.mjs`; the module is the source of truth.
 *
 * Typed JSON encoding used by canonical.json, expected/*.json and every oracle dump:
 *   null                       blank cell (no <c> element, or an empty one)
 *   ""                         empty string cell
 *   number | boolean | string  as-is
 *   { "$date": "YYYY-MM-DD" }                  date-only (wall clock, no timezone)
 *   { "$datetime": "YYYY-MM-DDTHH:mm:ss.SSS" } date+time (wall clock)
 *   { "$time": "HH:mm:ss.SSS" }                time-only (serial < 1)
 *   { "$error": "#N/A" }                       error cell
 *   { "$formula": "A1+1", "cached": <value> }  formula with cached result (cached uses this same encoding)
 */

export const SHEET_DATA = 'Data';
export const SHEET_FEATURES = 'Features';
export const SHEET_LONG_NAME = "It's a very long sheet name 001"; // exactly 31 chars, apostrophe inside
export const SHEET_HIDDEN = 'Hidden';

export const ROW_COUNT = 30;

function pick(list, i) {
  return list[i % list.length];
}

/** Column definitions: name, optional number format, and the value for data row `i` (0-based). */
export const COLUMNS = [
  {
    name: 'Id',
    numFmt: null,
    value: i => `001Xx${String(100000 + i * 7919).padStart(9, '0')}AA${String.fromCharCode(65 + (i % 26))}`.slice(0, 18),
  },
  {
    name: 'Name',
    numFmt: null,
    value: i =>
      pick(
        [
          'Zoë Ångström',
          '田中 太郎',
          '🚀 Launch Co',
          'Ñandú & Cía',
          'Émilie Brontë',
          'Ærø Ø ApS',
          'Straße Müller GmbH',
          `O'Brien "Quoted"`,
          '<script>alert(1)</script>',
          'Ünïcödé Tëst',
        ],
        i,
      ),
  },
  { name: 'RTL', numFmt: null, value: i => (i % 2 === 0 ? 'שלום עולם' : 'مرحبا بالعالم') },
  {
    name: 'Spaces',
    numFmt: null,
    value: i => pick([' leading', 'trailing ', '  both  ', 'tab\there', 'line1\nline2', 'crlf\r\nline2', '   ', 'a  b', 'none'], i),
  },
  {
    name: 'Control',
    numFmt: null,
    // U+0001 / U+001F / U+0007 / U+0000 / U+001B are not valid XML 1.0 characters and must be written as _xHHHH_.
    value: i => pick(['a\u0001b', 'x\u001fy', 'bell\u0007', 'del\u007fx', 'plain', 'nul\u0000?', 'esc\u001b[0m'], i),
  },
  {
    name: 'EscapeLiteral',
    numFmt: null,
    // Literal text that LOOKS like an escape. A correct writer emits _x005F_x0041_ for the literal '_x0041_'.
    value: i => pick(['_x0041_', '_x005F_x0041_', 'under_score', '_x', '_xZZZZ_', '_X0041_', 'end_x0041_'], i),
  },
  {
    name: 'FormulaLikeText',
    numFmt: null,
    value: i => pick(['=SUM(A1)', '+1', '-1', '@x', "'quoted", '=1+1', '=cmd|calc', '-2+3', 'plain'], i),
  },
  { name: 'LeadingZeros', numFmt: null, value: i => pick(['00123', '007', '0', '0042', '000', '01', '1e3', '0x1F'], i) },
  {
    name: 'BigInt17',
    numFmt: null,
    value: i =>
      pick([9007199254740991, 9007199254740992, 12345678901234568, 1e15, 123456789012345, -9007199254740991, 1e17, 4294967296], i),
  },
  {
    name: 'Float',
    numFmt: null,
    value: i => pick([-0, 0.1 + 0.2, 1e21, 1e-7, -123.456, 3.14159265358979, 1 / 3, 2.5, 1e-300, 1.7976931348623157e308], i),
  },
  { name: 'Percent', numFmt: '0.0%', value: i => pick([0.5, 0.125, 1, 0, -0.25, 1.5, 0.333333333], i) },
  { name: 'Currency', numFmt: '$#,##0.00', value: i => pick([1234.5, -99.99, 0, 1000000, 0.005, 19.999, -0.5], i) },
  {
    name: 'Date',
    numFmt: 'yyyy-mm-dd',
    value: i =>
      date(
        pick(
          [
            '1899-12-31',
            '1900-02-28',
            '1900-03-01',
            '1904-01-01',
            '2024-02-29',
            '1999-12-31',
            '2000-01-01',
            '2038-01-19',
            '9999-12-31',
            '1970-01-01',
            '1900-01-01',
            '2100-02-28',
          ],
          i,
        ),
      ),
  },
  {
    name: 'Time',
    numFmt: 'hh:mm:ss',
    value: i => time(pick(['12:34:56.789', '00:00:00.000', '23:59:59.000', '06:00:00.000', '00:00:01.500'], i)),
  },
  {
    name: 'DateTime',
    numFmt: 'yyyy-mm-dd hh:mm:ss',
    value: i =>
      datetime(
        pick(
          [
            '2024-03-10T02:30:00.000', // US DST gap (this wall clock does not exist in America/Los_Angeles)
            '2024-11-03T01:30:00.000', // US DST overlap
            '2024-01-02T15:04:05.678',
            '2000-01-01T00:00:00.000',
            '1900-01-01T12:00:00.000',
            '2024-12-31T23:59:59.999',
          ],
          i,
        ),
      ),
  },
  { name: 'Bool', numFmt: null, value: i => i % 3 === 0 },
  { name: 'BlankOrEmpty', numFmt: null, value: i => (i % 2 === 0 ? null : '') },
  { name: 'Long32767', numFmt: null, value: i => (i === ROW_COUNT - 1 ? repeatTo('Lorem ipsum dolor sit amet, ', 32767) : 'short') },
  // Exceeds Excel's per-cell limit. Ground truth keeps the full value; writers apply the Jetstream policy
  // (slice to 32,767 including the '...(truncated)' suffix). See tag policy:truncate-32767 in the manifest.
  { name: 'Over32767', numFmt: null, value: i => (i === ROW_COUNT - 2 ? repeatTo('0123456789', 40000) : 'short') },
  {
    name: 'Json',
    numFmt: null,
    value: i =>
      JSON.stringify({
        records: [
          {
            Id: `003Xx0000000${String(i).padStart(3, '0')}AAA`,
            Name: `O'Brien "Q${i}"`,
            Note: 'line1\nline2\ttab',
            Amount: i * 1.5,
            Active: i % 2 === 0,
          },
        ],
        totalSize: 1,
        done: true,
      }),
  },
];

export const EXCEL_MAX_CELL_CHARS = 32767;
export const TRUNCATION_SUFFIX = '...(truncated)';

/** Jetstream's write-side policy for over-limit cells (prepareExcelFile). */
export function truncateForExcel(value) {
  if (typeof value === 'string' && value.length > EXCEL_MAX_CELL_CHARS) {
    return value.slice(0, EXCEL_MAX_CELL_CHARS - TRUNCATION_SUFFIX.length) + TRUNCATION_SUFFIX;
  }
  return value;
}

function repeatTo(chunk, length) {
  return chunk.repeat(Math.ceil(length / chunk.length)).slice(0, length);
}

export const date = iso => ({ $date: iso });
export const datetime = iso => ({ $datetime: iso });
export const time = iso => ({ $time: iso });
export const error = code => ({ $error: code });
export const formula = (text, cached) => ({ $formula: text, cached });

export function isTyped(value, key) {
  return value !== null && typeof value === 'object' && key in value;
}

export function isTemporal(value) {
  return isTyped(value, '$date') || isTyped(value, '$datetime') || isTyped(value, '$time');
}

/** Data sheet rows (header + ROW_COUNT data rows) in the typed encoding. */
export function dataRows() {
  const header = COLUMNS.map(({ name }) => name);
  const rows = [];
  for (let i = 0; i < ROW_COUNT; i++) {
    rows.push(COLUMNS.map(({ value }) => value(i)));
  }
  return [header, ...rows];
}

/**
 * Features sheet: exercises structural features rather than values. Layout (1-based):
 *   A1:C1  merged header "Merged header"
 *   A2     "Feature"   B2 "Value"  C2 "Note"   (header row; freeze pane below row 2, autofilter A2:C12)
 *   A3 formula   B3 =Data!A2&"x"   (cached: first Id + "x")
 *   A4 div0      B4 =1/0           (cached #DIV/0!)
 *   A5 na        B5 =NA()          (cached #N/A)
 *   A6 error     B6 #N/A literal error cell
 *   A7 hyperlink B7 "Jetstream" -> https://getjetstream.app
 *   A8 richtext  B8 "bold and plain" (bold run + plain run; readers flatten to plain text)
 *   A9 note      B9 "has a note"  (legacy comment/note on B9: "This is a note")
 *   A10 validation B10 "Red"  (list validation Red,Green,Blue on B10:B12)
 *   A11 condfmt  B11 42  (conditional format: >40 -> fill)
 *   Row 12 hidden; column D hidden.
 */
export function featureRows() {
  const firstId = COLUMNS[0].value(0);
  return [
    ['Merged header', null, null],
    ['Feature', 'Value', 'Note'],
    ['formula', formula('Data!A2&"x"', `${firstId}x`), 'string concat with cached value'],
    ['div0', formula('1/0', error('#DIV/0!')), 'error cached value'],
    ['na', formula('NA()', error('#N/A')), 'error cached value'],
    ['error', error('#N/A'), 'literal error cell (t="e")'],
    ['hyperlink', 'Jetstream', 'https://getjetstream.app'],
    ['richtext', 'bold and plain', 'rich text runs, flattened on read'],
    ['note', 'has a note', 'legacy comment on B9'],
    ['validation', 'Red', 'list validation Red,Green,Blue'],
    ['condfmt', 42, 'conditional format >40'],
    ['hidden row', 'row 12 is hidden', null],
  ];
}

export const FEATURES = {
  merges: [{ s: { r: 0, c: 0 }, e: { r: 0, c: 2 } }],
  freeze: { rows: 2, cols: 0 },
  autoFilter: 'A2:C12',
  hyperlink: { cell: 'B7', url: 'https://getjetstream.app' },
  richText: { cell: 'B8', runs: [{ text: 'bold', bold: true }, { text: ' and plain' }] },
  note: { cell: 'B9', text: 'This is a note' },
  validation: { range: 'B10:B12', list: ['Red', 'Green', 'Blue'] },
  conditionalFormat: { range: 'B11', operator: 'greaterThan', value: 40, fillColor: 'FFFFC7CE' },
  hiddenRows: [12],
  hiddenColumns: ['D'],
  columnWidths: { A: 16, B: 24, C: 40 },
};

export function longNameRows() {
  return [['Sheet name is 31 chars with an apostrophe'], [SHEET_LONG_NAME.length]];
}

export function hiddenSheetRows() {
  return [['This sheet is hidden'], ['secret', 42]];
}

/** Full workbook description in the typed encoding. */
export function canonicalWorkbook() {
  return {
    version: 1,
    generatedBy: 'fixtures/canonical/canonical.mjs',
    sheets: [
      {
        name: SHEET_DATA,
        hidden: false,
        rows: dataRows(),
        numFmts: Object.fromEntries(COLUMNS.filter(column => column.numFmt).map(column => [column.name, column.numFmt])),
      },
      { name: SHEET_FEATURES, hidden: false, rows: featureRows(), features: FEATURES },
      { name: SHEET_LONG_NAME, hidden: false, rows: longNameRows() },
      { name: SHEET_HIDDEN, hidden: true, rows: hiddenSheetRows() },
    ],
  };
}

// ---------- helpers shared by generators and oracle dumps ----------

const EXCEL_EPOCH_UTC = Date.UTC(1899, 11, 30);

/** Wall-clock components of a typed date/datetime/time value (no timezone involved). */
export function components(typed) {
  const iso = typed.$datetime ?? (typed.$date ? `${typed.$date}T00:00:00.000` : `1899-12-30T${typed.$time}`);
  const [datePart, timePart] = iso.split('T');
  const [y, m, d] = datePart.split('-').map(Number);
  const [hh, mm, ssms] = timePart.split(':');
  const [ss, ms = '0'] = ssms.split('.');
  return { y, m, d, hh: Number(hh), mm: Number(mm), ss: Number(ss), ms: Number(ms.padEnd(3, '0')) };
}

/** Excel serial (1900 system, with the Lotus leap-year bug) for a typed temporal value. */
export function toSerial(typed) {
  const { y, m, d, hh, mm, ss, ms } = components(typed);
  const utc = Date.UTC(y, m - 1, d, hh, mm, ss, ms);
  let serial = (utc - EXCEL_EPOCH_UTC) / 86400000;
  // Dates before 1900-03-01 are one day lower in Excel's 1900 system (there is no real 1900-02-29).
  if (serial >= 1 && serial < 61) {
    serial -= 1;
  }
  return serial;
}

/** Local-wall-clock JS Date (what SheetJS cellDates:true produces) for a typed temporal value. */
export function toLocalDate(typed) {
  const { y, m, d, hh, mm, ss, ms } = components(typed);
  return new Date(y, m - 1, d, hh, mm, ss, ms);
}

/** UTC JS Date with the same wall-clock digits (what libraries using Date.UTC produce). */
export function toUtcDate(typed) {
  const { y, m, d, hh, mm, ss, ms } = components(typed);
  return new Date(Date.UTC(y, m - 1, d, hh, mm, ss, ms));
}

const pad = (n, width = 2) => String(n).padStart(width, '0');

/** Inverse of toLocalDate: typed encoding from a JS Date read back by a library (local getters). */
export function fromLocalDate(value) {
  const y = value.getFullYear();
  const m = value.getMonth() + 1;
  const d = value.getDate();
  const hh = value.getHours();
  const mm = value.getMinutes();
  const ss = value.getSeconds();
  const ms = value.getMilliseconds();
  if (y === 1899 && m === 12 && d === 30) {
    return time(`${pad(hh)}:${pad(mm)}:${pad(ss)}.${pad(ms, 3)}`);
  }
  if (hh === 0 && mm === 0 && ss === 0 && ms === 0) {
    return date(`${pad(y, 4)}-${pad(m)}-${pad(d)}`);
  }
  return datetime(`${pad(y, 4)}-${pad(m)}-${pad(d)}T${pad(hh)}:${pad(mm)}:${pad(ss)}.${pad(ms, 3)}`);
}

/** Same as fromLocalDate but reading UTC getters (for libraries that construct Dates with Date.UTC). */
export function fromUtcDate(value) {
  return fromLocalDate(
    new Date(
      value.getUTCFullYear(),
      value.getUTCMonth(),
      value.getUTCDate(),
      value.getUTCHours(),
      value.getUTCMinutes(),
      value.getUTCSeconds(),
      value.getUTCMilliseconds(),
    ),
  );
}
