#!/usr/bin/env node
// Produces the Excel-generated goldens by driving Microsoft Excel (full license required) through
// oracle/excel/save-as.applescript, then registers each file in the manifest.
//
//   node fixtures/generators/excel-goldens.mjs [--only from-csv,from-exceljs,...] [--no-register]
//
// Each golden is "what Excel writes after opening X": the CSV import shape, re-saves of the library goldens
// (Excel rewrites styles/SST/theme/calcChain), a 1904-date-system save, and BIFF8 / xlsb saves used as renamed-format
// hostile inputs. The Strict Open XML save is manual (no AppleScript file format for it).
import { spawnSync } from 'node:child_process';
import { existsSync, mkdirSync, renameSync, statSync } from 'node:fs';
import { join } from 'node:path';
import { parseArgs } from 'node:util';

const HERE = new URL('.', import.meta.url).pathname.replace(/\/$/, '');
const FIXTURES = join(HERE, '..');
const SCRIPT = join(FIXTURES, '../oracle/excel/save-as.applescript');
const OUT_DIR = join(FIXTURES, 'golden/excel-365');
const EXCEL_VERSION = (() => {
  const result = spawnSync('osascript', ['-e', 'tell application "Microsoft Excel" to version'], { encoding: 'utf8', timeout: 60000 });
  return (result.stdout || '').trim() || 'unknown';
})();
const GENERATOR = `Microsoft Excel ${EXCEL_VERSION} macOS (desktop)`;
const PROVENANCE = 'applescript:oracle/excel/save-as.applescript (fixtures/generators/excel-goldens.mjs)';
const EXP = 'canonical/canonical.json';
const EXP_CSV = 'expected/canonical-csv-import.json';

const GOLDENS = [
  {
    key: 'from-csv',
    input: 'canonical/canonical.csv',
    output: 'canonical.from-csv.xlsx',
    format: 'xlsx',
    id: 'golden-canonical-excel-365-from-csv',
    expected: EXP_CSV,
    tags: ['source:csv', 'EC-INPUT-CSV-IMPORT', 'EC-EXCEL-CSV-IMPORT-TYPING', 'EC-CSV-NO-BOM-MOJIBAKE', 'EC-EXCEL-CSV-FORMULA-EVAL'],
    notes:
      "Excel opened canonical.csv (UTF-8, no BOM) and saved it as .xlsx: every column typed by Excel's CSV heuristics, non-ASCII text mojibaked, '=text' evaluated.",
  },
  {
    key: 'from-sheetjs',
    input: 'golden/sheetjs/canonical.xlsx',
    output: 'canonical.from-sheetjs.xlsx',
    format: 'xlsx',
    id: 'golden-canonical-excel-365-resave-sheetjs',
    expected: EXP,
    tags: ['policy:truncate-32767'],
    notes: 'Excel re-save of the SheetJS golden - what a user gets after opening a Jetstream export in Excel and saving.',
  },
  {
    key: 'from-exceljs',
    input: 'golden/exceljs/canonical.xlsx',
    output: 'canonical.from-exceljs.xlsx',
    format: 'xlsx',
    id: 'golden-canonical-excel-365-resave',
    expected: EXP,
    tags: ['policy:truncate-32767'],
    notes: 'Excel re-save of the ExcelJS golden (styles, SST, theme, calcChain rewritten by Excel).',
  },
  {
    key: 'from-office-kit',
    input: 'golden/office-kit/canonical.xlsx',
    output: 'canonical.from-office-kit.xlsx',
    format: 'xlsx',
    id: 'golden-canonical-excel-365-resave-office-kit',
    expected: EXP,
    tags: ['policy:truncate-32767', 'EC-EXCEL-RESAVE-PRESERVES-SHIFTED-SERIALS'],
    notes: 'Excel re-save of the office-kit golden (its UTC-offset date serials and double-escaped _x005F literal survive the round trip).',
  },
  {
    key: 'strict',
    input: 'golden/exceljs/canonical.xlsx',
    output: 'canonical.from-exceljs.strict.xlsx',
    format: 'strict',
    id: 'golden-canonical-excel-365-strict',
    expected: EXP,
    tags: ['policy:truncate-32767', 'EC-STRICT-NAMESPACES', 'EC-STRICT-REAL-EXCEL-UNREADABLE', 'EC-STRICT-ISO-DATE-PRECISION'],
    notes:
      'Excel \'Strict Open XML Spreadsheet\' save (VBA FileFormat 61): purl.oclc.org namespaces, Strict relationship types, t="d" ISO cells with 17 fractional digits.',
  },
  {
    key: 'xls',
    input: 'golden/exceljs/canonical.xlsx',
    output: '../../hostile/biff8-xls-renamed.xlsx',
    format: 'xls',
    id: 'hostile-biff8-xls-renamed',
    expectedError: 'NOT_XLSX',
    tags: ['kind:hostile', 'EC-INPUT-XLS-RENAMED'],
    notes:
      'Excel 97-2004 (BIFF8, Excel98to2004 file format) save of the ExcelJS golden under an .xlsx name: the renamed-legacy-format hostile input.',
  },
  {
    key: 'xlsb',
    input: 'golden/exceljs/canonical.xlsx',
    output: '../../hostile/xlsb-renamed.xlsx',
    format: 'xlsb',
    id: 'hostile-xlsb-renamed',
    expectedError: 'NOT_XLSX',
    tags: ['kind:hostile', 'EC-INPUT-XLSB-RENAMED'],
    notes:
      'Excel Binary Workbook (.xlsb) save of the ExcelJS golden under an .xlsx name: a zip container whose parts are binary records, not XML.',
  },
  {
    key: '1904',
    input: 'golden/exceljs/canonical.xlsx',
    output: 'canonical.from-exceljs.1904.xlsx',
    format: 'xlsx',
    date1904: true,
    id: 'golden-canonical-excel-365-1904',
    expected: EXP,
    tags: ['policy:truncate-32767', 'EC-DATE-1904'],
    notes:
      'Excel re-save with the 1904 date system switched on (workbookPr date1904="1"); serials shift by 1,462 days while displayed dates stay the same.',
  },
];

const { values: options } = parseArgs({ options: { only: { type: 'string' }, 'no-register': { type: 'boolean', default: false } } });
const wanted = options.only ? new Set(options.only.split(',')) : null;
mkdirSync(OUT_DIR, { recursive: true });

let failures = 0;
for (const golden of GOLDENS) {
  if (wanted && !wanted.has(golden.key)) {
    continue;
  }
  const input = join(FIXTURES, golden.input);
  const output = join(OUT_DIR, golden.output);
  if (!existsSync(input)) {
    console.log(`${golden.key.padEnd(16)} SKIP missing input ${golden.input}`);
    continue;
  }
  const started = Date.now();
  if (golden.format === 'strict') {
    console.log(
      `${golden.key.padEnd(16)} MANUAL Excel 16 for Mac cannot save Strict Open XML from AppleScript; keep the hand-saved file (see golden/excel-365/STEPS.md)`,
    );
    continue;
  }
  // Excel refuses to write the binary format under an .xlsx name ('Parameter error'), so save with the native
  // extension and rename afterwards - the fixture's point is exactly that the bytes and the name disagree.
  const saveTarget = golden.format === 'xlsb' ? `${output}.tmp.xlsb` : output;
  const result = spawnSync('osascript', [SCRIPT, input, saveTarget, golden.format, String(Boolean(golden.date1904))], {
    encoding: 'utf8',
    timeout: 180000,
  });
  if (saveTarget !== output && existsSync(saveTarget)) {
    renameSync(saveTarget, output);
  }
  const stdout = (result.stdout || '').trim();
  const unsupported = /STRICT-UNSUPPORTED/.test(stdout);
  if (result.status !== 0 || !existsSync(output) || unsupported) {
    failures++;
    console.log(`${golden.key.padEnd(16)} FAIL ${(result.stderr || stdout).trim().split('\n').slice(-2).join(' | ').slice(0, 220)}`);
    continue;
  }
  const bytes = statSync(output).size;
  console.log(
    `${golden.key.padEnd(16)} ok   ${String(bytes).padStart(7)} bytes  ${Date.now() - started} ms  ${stdout.split('\n').slice(2).join('; ')}`,
  );
  if (options['no-register']) {
    continue;
  }
  const relative = golden.output.startsWith('../../') ? golden.output.replace('../../', '') : `golden/excel-365/${golden.output}`;
  const args = [
    join(FIXTURES, 'register.mjs'),
    relative,
    '--id',
    golden.id,
    '--generator',
    GENERATOR,
    '--provenance',
    PROVENANCE,
    '--license',
    'Microsoft Excel output',
    '--tags',
    [golden.expectedError ? '' : 'kind:golden', 'generator:excel-365', ...golden.tags].filter(Boolean).join(','),
    '--notes',
    golden.notes,
  ];
  if (golden.expected) {
    args.push('--expected', golden.expected);
  }
  if (golden.expectedError) {
    args.push('--expected-error', golden.expectedError);
  }
  const registered = spawnSync('node', args, { encoding: 'utf8' });
  process.stdout.write(`${''.padEnd(16)}      ${(registered.stdout || registered.stderr).trim()}\n`);
}
process.exitCode = failures > 0 ? 1 : 0;
