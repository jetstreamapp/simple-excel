/**
 * The smoke workbook: the dataset, the writer and the read-back verification, shared by `smoke.js` (main
 * thread) and `smoke.worker.js` so both threads write exactly the same bytes and are checked the same way.
 *
 * Served from the repo root, so the library is imported by absolute URL rather than through a bundler; every
 * module in this directory imports that same URL and therefore the same module instance.
 *
 * The dataset is shaped like a Jetstream export — ids, names, numbers, booleans, dates, blanks — with every
 * trap value the catalog cares about parked on the first data row so a reader check is one row lookup.
 */
import { createWorkbookWriter, openWorkbook } from '/dist/esm/index.mjs';

export const DATA_SHEET = 'Data';
export const HIDDEN_SHEET = 'Meta';

export const MAX_CELL_CHARS = 32_767;
export const TRUNCATION_SUFFIX = '...(truncated)';
export const LONG_CELL_CHARS = 40_000;
export const DATE_NUMBER_FORMAT = 'yyyy-mm-dd hh:mm:ss';

export const HEADER = [
  'Id',
  'Name',
  'Multiline',
  'Crlf',
  'Control',
  'EscapeLiteral',
  'NegativeZero',
  'Big',
  'Float',
  'MaxSafeInteger',
  'Flag',
  'When',
  'TimeOnly',
  'Blank',
  'Long',
];

/** Columns that carry a `Date` and therefore need the registered date number format to read back as one. */
const DATE_COLUMNS = new Set(['When', 'TimeOnly']);

/**
 * Control characters are built at runtime instead of written as `\u0001` escapes: the formatter rewrites such an
 * escape into the literal character, which makes the file unreadable and easy to corrupt in an editor.
 */
function control(codePoint) {
  return String.fromCodePoint(codePoint);
}

/**
 * Row 1 of the data: every value the reader checks by name. `-0`, `1e21`, `0.1 + 0.2` and `Number.MAX_SAFE_INTEGER`
 * are the number traps; the unicode, CRLF, control character and `_x0041_` literal are the string traps; the
 * time-only Date is a serial below 1; the long cell is over Excel's 32,767-character cell limit.
 */
export const SENTINEL_ROW = [
  'SMOKE-000000',
  'Zoë 田中 🚀',
  'line1\nline2',
  'crlf\r\nline2',
  `a${control(0x01)}b`,
  '_x0041_',
  -0,
  1e21,
  0.1 + 0.2,
  Number.MAX_SAFE_INTEGER,
  true,
  new Date(2024, 5, 15, 13, 45, 30),
  new Date(1899, 11, 30, 12, 34, 56, 789),
  null,
  'L'.repeat(LONG_CELL_CHARS),
];

const NAMES = ['Zoë Ångström', '田中 太郎', '🚀 Launch Co', 'O\'Brien "Quoted"', '<script>alert(1)</script>'];
const MULTILINE = ['line1\nline2', ' leading', 'trailing ', 'tab\there', 'plain'];
const CONTROL = [`a${control(0x01)}b`, `x${control(0x1f)}y`, `bell${control(0x07)}`, `esc${control(0x1b)}[0m`, 'plain'];
const ESCAPES = ['_x0041_', '_x005F_x0041_', 'under_score', '_xZZZZ_', 'end_x0041_'];
const NUMBERS = [0, -1.5, 1234.5678, -98765, -0.000001];

function pick(list, index) {
  return list[index % list.length];
}

/**
 * Data row `index` (0-based). Row 0 is the sentinel row; the rest cycle through the same families of trap values
 * so the whole sheet is mixed rather than 20,000 copies of one shape.
 */
export function smokeRow(index) {
  if (index === 0) {
    return SENTINEL_ROW;
  }
  return [
    `SMOKE-${String(index).padStart(6, '0')}`,
    pick(NAMES, index),
    pick(MULTILINE, index),
    index % 4 === 0 ? 'crlf\r\nline2' : `note ${index}`,
    pick(CONTROL, index),
    pick(ESCAPES, index),
    pick(NUMBERS, index),
    index * 1e15,
    index / 7,
    Number.MAX_SAFE_INTEGER - index,
    index % 2 === 0,
    new Date(2024, 0, 1 + (index % 365), 12, 0, 0),
    new Date(1899, 11, 30, 8 + (index % 12), index % 60, (index * 7) % 60),
    index % 3 === 0 ? null : `filled ${index}`,
    `row ${index}`,
  ];
}

function* smokeRows(rowCount) {
  for (let index = 0; index < rowCount; index++) {
    yield smokeRow(index);
  }
}

/**
 * Write the smoke workbook to `sink`: a `Data` sheet with a bold header, a frozen header row, an autofilter and a
 * date-formatted column, plus a hidden `Meta` sheet. `deterministic: true` so two runs produce identical bytes.
 */
export async function writeSmokeWorkbook(sink, rowCount) {
  const startedAt = performance.now();
  const workbook = createWorkbookWriter(sink, { deterministic: true });
  const dateStyle = workbook.registerStyle({ numFmt: DATE_NUMBER_FORMAT });
  const columnStyles = HEADER.map(name => (DATE_COLUMNS.has(name) ? dateStyle : undefined));

  const data = workbook.addSheet(DATA_SHEET, {
    header: HEADER,
    freeze: { rows: 1 },
    autoFilter: true,
    rowCount,
  });
  for (const row of smokeRows(rowCount)) {
    await data.writeRow(row, columnStyles);
  }
  await data.close();

  const meta = workbook.addSheet(HIDDEN_SHEET, { hidden: true, header: ['Key', 'Value'], rowCount: 2 });
  await meta.writeRow(['generator', 'simple-excel browser smoke']);
  await meta.writeRow(['rows', rowCount]);
  await meta.close();

  const result = await workbook.close();
  return { result, ms: performance.now() - startedAt };
}

function check(list, id, ok, detail) {
  list.push({ id, ok, detail });
}

function sameDate(value, expected) {
  return (
    value instanceof Date &&
    value.getFullYear() === expected.getFullYear() &&
    value.getMonth() === expected.getMonth() &&
    value.getDate() === expected.getDate() &&
    value.getHours() === expected.getHours() &&
    value.getMinutes() === expected.getMinutes() &&
    value.getSeconds() === expected.getSeconds() &&
    value.getMilliseconds() === expected.getMilliseconds()
  );
}

function describeDate(value) {
  if (!(value instanceof Date)) {
    return `${typeof value} ${JSON.stringify(value)}`;
  }
  const pad = (part, width = 2) => String(part).padStart(width, '0');
  return `${value.getFullYear()}-${pad(value.getMonth() + 1)}-${pad(value.getDate())} ${pad(value.getHours())}:${pad(value.getMinutes())}:${pad(value.getSeconds())}.${pad(value.getMilliseconds(), 3)}`;
}

function describe(value) {
  if (value instanceof Date) {
    return describeDate(value);
  }
  if (typeof value === 'string' && value.length > 48) {
    return `string(${value.length}) ${JSON.stringify(`${value.slice(0, 16)}…${value.slice(-16)}`)}`;
  }
  return `${typeof value} ${JSON.stringify(value)}`;
}

/**
 * Read the workbook back from `source` (a Blob, an ArrayBuffer or a Uint8Array) and compare every sentinel cell
 * with the value `SENTINEL_ROW` actually wrote. Returns one pass/fail record per check so the runner never has to
 * know what a correct value looks like.
 */
export async function verifySmokeWorkbook(source, rowCount) {
  const startedAt = performance.now();
  const checks = [];
  const workbook = await openWorkbook(source);
  try {
    const names = workbook.sheets.map(sheet => sheet.name);
    check(checks, 'sheet-names', names.join(',') === `${DATA_SHEET},${HIDDEN_SHEET}`, names.join(', '));
    const hidden = workbook.sheets.map(sheet => `${sheet.name}=${sheet.hidden}`);
    check(checks, 'sheet-hidden-flags', !workbook.sheets[0].hidden && workbook.sheets[1].hidden === true, hidden.join(', '));

    const { rows, headers } = await workbook.sheet(DATA_SHEET).toObjects({ defval: null });
    check(checks, 'headers', headers.join(',') === HEADER.join(','), headers.join(', '));
    check(checks, 'row-count', rows.length === rowCount, `${rows.length} of ${rowCount}`);

    const first = rows[0] ?? {};
    const expected = Object.fromEntries(HEADER.map((name, index) => [name, SENTINEL_ROW[index]]));
    const cell = name => first[name];

    check(checks, 'unicode', cell('Name') === expected.Name, describe(cell('Name')));
    check(checks, 'newline-lf', cell('Multiline') === expected.Multiline, describe(cell('Multiline')));
    check(checks, 'newline-crlf', cell('Crlf') === expected.Crlf, describe(cell('Crlf')));
    check(checks, 'control-character', cell('Control') === expected.Control, describe(cell('Control')));
    check(checks, 'escape-literal', cell('EscapeLiteral') === expected.EscapeLiteral, describe(cell('EscapeLiteral')));
    // Excel has no negative zero: `-0` is written as `0` and must read back as a plain zero, not as NaN or a string.
    check(
      checks,
      'negative-zero-as-zero',
      cell('NegativeZero') === 0 && !Object.is(cell('NegativeZero'), -0),
      describe(cell('NegativeZero')),
    );
    check(checks, 'number-1e21', cell('Big') === expected.Big, describe(cell('Big')));
    check(checks, 'float-0.1+0.2', cell('Float') === expected.Float, describe(cell('Float')));
    check(checks, 'max-safe-integer', cell('MaxSafeInteger') === expected.MaxSafeInteger, describe(cell('MaxSafeInteger')));
    check(checks, 'boolean', cell('Flag') === true, describe(cell('Flag')));
    check(checks, 'date-local-wall-clock', sameDate(cell('When'), expected.When), describe(cell('When')));
    check(checks, 'date-time-only', sameDate(cell('TimeOnly'), expected.TimeOnly), describe(cell('TimeOnly')));
    check(checks, 'blank-cell', cell('Blank') === null, describe(cell('Blank')));

    const long = cell('Long');
    const truncatedCorrectly = typeof long === 'string' && long.length === MAX_CELL_CHARS && long.endsWith(TRUNCATION_SUFFIX);
    check(checks, 'cell-truncated-at-32767', truncatedCorrectly, describe(long));

    // A row from the middle proves the whole sheet streamed, not just the sentinel row.
    const middleIndex = Math.floor(rowCount / 2);
    const middle = rows[middleIndex] ?? {};
    const middleExpected = smokeRow(middleIndex);
    const middleOk = middle.Id === middleExpected[0] && middle.Name === middleExpected[1] && sameDate(middle.When, middleExpected[11]);
    check(checks, 'middle-row', middleOk, `row ${middleIndex}: ${describe(middle.Id)}, ${describe(middle.Name)}, ${describe(middle.When)}`);

    const last = rows.at(-1) ?? {};
    check(checks, 'last-row', last.Id === smokeRow(rowCount - 1)[0], describe(last.Id));
  } finally {
    await workbook.close();
  }
  return { checks, ms: performance.now() - startedAt };
}
