#!/usr/bin/env node
// Compatibility oracle runner: for each fixture (or the ones selected), run every reader, diff against the
// fixture's expected dump, and write per-(fixture, reader) result JSON plus a summary table under
// oracle/results/<YYYY-MM-DD>-<label>/.
//
//   node run.mjs [--fixtures id1,id2 | --tag kind:golden]
//                [--readers simple-excel,sheetjs,office-kit,office-kit-stream,openpyxl,calamine,validator,libreoffice,excel]
//                [--label baseline] [--python .generated/venv/bin/python]
//
// Readers without an expected dump (validator, libreoffice) record open/convert success only.
// `simple-excel` is this package itself, read through `dist/esm` - run `npm run build` before the oracle.
import { spawnSync } from 'node:child_process';
import { existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { hostname } from 'node:os';
import { join } from 'node:path';
import { parseArgs } from 'node:util';

const HERE = new URL('.', import.meta.url).pathname.replace(/\/$/, '');
const ROOT = join(HERE, '..');
const FIXTURES = join(ROOT, 'fixtures');
const ALL_READERS = [
  'simple-excel',
  'sheetjs',
  'office-kit',
  'office-kit-stream',
  'openpyxl',
  'calamine',
  'validator',
  'libreoffice',
  'excel',
];

/**
 * The manifest records reader-agnostic error names for the hostile fixtures. A reader that classifies its own
 * rejections (ours reports `dump.error.code`) PASSes one when the code it threw is listed here; readers that only
 * crash still record REJECTED. Same table as `test/hostile.test.ts`.
 */
const EXPECTED_ERROR_CODES = {
  XML_DOCTYPE: ['XML_DOCTYPE'],
  TRUNCATED: ['ZIP_TRUNCATED'],
  CRC_MISMATCH: ['ZIP_CRC_MISMATCH'],
  DUPLICATE_ENTRY: ['ZIP_DUPLICATE_ENTRY'],
  LIMIT_EXCEEDED: ['LIMIT_EXCEEDED'],
  NOT_XLSX: ['NOT_XLSX', 'ODS', 'LEGACY_XLS', 'XLSB'],
  ENCRYPTED: ['ENCRYPTED'],
  ZIP_BOMB: ['ZIP_BOMB'],
};

const { values: options } = parseArgs({
  options: {
    fixtures: { type: 'string' },
    tag: { type: 'string' },
    readers: { type: 'string', default: ALL_READERS.join(',') },
    label: { type: 'string', default: 'run' },
    python: { type: 'string', default: join(ROOT, '.generated/venv/bin/python') },
    soffice: { type: 'string', default: '/Applications/LibreOffice.app/Contents/MacOS/soffice' },
  },
});

const manifest = JSON.parse(readFileSync(join(FIXTURES, 'manifest.json'), 'utf8'));
let fixtures = manifest.fixtures.filter(fixture => !fixture.generated || existsSync(join(FIXTURES, fixture.path)));
if (options.fixtures) {
  const wanted = new Set(options.fixtures.split(','));
  fixtures = fixtures.filter(fixture => wanted.has(fixture.id));
}
if (options.tag) {
  fixtures = fixtures.filter(fixture => fixture.tags.includes(options.tag));
}
const readers = options.readers.split(',').filter(reader => ALL_READERS.includes(reader));

const date = new Date().toISOString().slice(0, 10);
const outDir = join(HERE, 'results', `${date}-${options.label}`);
const scratch = join(ROOT, '.generated', 'oracle', `${date}-${options.label}`);
mkdirSync(outDir, { recursive: true });
mkdirSync(scratch, { recursive: true });

function run(command, args, timeoutMs = 180000) {
  const result = spawnSync(command, args, { encoding: 'utf8', timeout: timeoutMs, maxBuffer: 64 * 1024 * 1024 });
  return { status: result.status, stdout: result.stdout ?? '', stderr: result.stderr ?? '', signal: result.signal };
}

/**
 * Read limits for the hostile fixtures, matching `test/hostile.test.ts`. Our library's defaults (1 GiB per part)
 * are deliberately generous - a 30 MB sheet is a legitimate export - so the zip bomb is only a bomb against the
 * kind of budget a host application sets.
 */
const HOSTILE_LIMITS = ['--max-inflated-bytes', String(8 * 1024 * 1024), '--max-shared-string-chars', String(4 * 1024 * 1024)];

function dumpCommand(reader, file, dumpPath, fixture) {
  switch (reader) {
    case 'simple-excel':
      return [
        'node',
        [
          join(HERE, 'simple-excel/read-dump.mjs'),
          file,
          '--out',
          dumpPath,
          ...(fixture.tags.includes('kind:hostile') ? HOSTILE_LIMITS : []),
        ],
      ];
    case 'sheetjs':
      return ['node', [join(HERE, 'sheetjs/read-dump.mjs'), file, '--out', dumpPath]];
    case 'office-kit':
      return ['node', [join(HERE, 'office-kit/read-dump.mjs'), file, '--clock', 'utc', '--out', dumpPath]];
    case 'office-kit-stream':
      return ['node', [join(HERE, 'office-kit/read-dump.mjs'), file, '--mode', 'stream', '--clock', 'utc', '--out', dumpPath]];
    case 'openpyxl':
    case 'calamine':
      return [options.python, [join(HERE, 'python/read_dump.py'), file, '--reader', reader, '--out', dumpPath]];
    default:
      throw new Error(`no dump command for ${reader}`);
  }
}

/** A hostile fixture is only PASSed by a reader that names the classified code the manifest asks for. */
function rejectionMatches(expectedError, code) {
  if (!code) {
    return false;
  }
  return (EXPECTED_ERROR_CODES[expectedError] ?? [expectedError]).includes(code);
}

/** The dump a reader wrote, or null when it wrote nothing usable (a crash rather than a reported failure). */
function readDump(dumpPath) {
  if (!existsSync(dumpPath)) {
    return null;
  }
  try {
    return JSON.parse(readFileSync(dumpPath, 'utf8'));
  } catch {
    return null;
  }
}

function summarizeError(text) {
  const lines = text
    .split('\n')
    .map(line => line.trim())
    .filter(line => line && !line.startsWith('at '));
  return lines.slice(-3).join(' | ').slice(0, 300);
}

const results = [];
for (const fixture of fixtures) {
  const file = join(FIXTURES, fixture.path);
  const expectedPath = fixture.expected ? join(FIXTURES, fixture.expected) : null;
  const policies = fixture.tags.filter(tag => tag.startsWith('policy:')).map(tag => tag.slice('policy:'.length));
  for (const reader of readers) {
    const started = Date.now();
    const record = { fixture: fixture.id, reader, file: fixture.path, status: 'UNKNOWN', ms: 0 };
    try {
      if (reader === 'validator') {
        const validator = run('npx', ['--no-install', 'ooxml-validator', file], 120000);
        record.exit = validator.status;
        record.stdout = validator.stdout.slice(0, 4000);
        record.stderr = summarizeError(validator.stderr);
        let parsed = null;
        try {
          parsed = JSON.parse(validator.stdout);
        } catch {
          parsed = null;
        }
        const errors = parsed?.errors ?? null;
        record.schemaErrors = errors ? errors.length : null;
        record.errorsSample = errors ? errors.slice(0, 5) : null;
        record.status = errors === null ? (validator.status === 0 ? 'PASS' : 'ERROR') : errors.length === 0 ? 'PASS' : 'FAIL';
      } else if (reader === 'excel') {
        // read-only Excel: open, detect repair (timeout/[Repaired]), read sentinel cells, close without saving
        const excel = run('node', [join(HERE, 'excel/run.mjs'), file, '--timeout', '150000'], 200000);
        let verdict = null;
        try {
          verdict = JSON.parse(excel.stdout);
        } catch {
          verdict = null;
        }
        record.status = verdict?.status ?? 'ERROR';
        record.workbookName = verdict?.workbookName ?? null;
        record.sheets = verdict?.sheets ?? null;
        record.cells = (verdict?.cells ?? null)?.map(cell => ({
          ...cell,
          value: String(cell.value ?? '').slice(0, 120),
          text: String(cell.text ?? '').slice(0, 120),
        }));
        record.repairLogs = verdict?.repairLogs;
        record.note = verdict?.note;
        record.error = verdict?.error ?? (excel.status !== 0 && !verdict ? summarizeError(excel.stderr) : undefined);
      } else if (reader === 'libreoffice') {
        const loOut = join(scratch, `${fixture.id}.libreoffice`);
        mkdirSync(loOut, { recursive: true });
        const convert = run(
          options.soffice,
          [
            '--headless',
            `-env:UserInstallation=file://${join(ROOT, '.generated/lo-profile')}`,
            '--convert-to',
            'xlsx',
            '--outdir',
            loOut,
            file,
          ],
          240000,
        );
        const produced = existsSync(
          join(
            loOut,
            fixture.path
              .split('/')
              .pop()
              .replace(/\.[^.]+$/, '.xlsx'),
          ),
        );
        record.exit = convert.status;
        record.stderr = summarizeError(convert.stderr + convert.stdout);
        record.status = convert.status === 0 && produced ? 'PASS' : 'FAIL';
        if (produced && expectedPath) {
          // read LibreOffice's re-save with SheetJS to see what survived the round trip
          const resaved = join(
            loOut,
            fixture.path
              .split('/')
              .pop()
              .replace(/\.[^.]+$/, '.xlsx'),
          );
          const dumpPath = join(scratch, `${fixture.id}.libreoffice-resave.sheetjs.json`);
          const dumped = run('node', [join(HERE, 'sheetjs/read-dump.mjs'), resaved, '--out', dumpPath]);
          if (dumped.status === 0) {
            const reportPath = join(scratch, `${fixture.id}.libreoffice-resave.diff.json`);
            run('node', [
              join(HERE, 'diff.mjs'),
              expectedPath,
              dumpPath,
              '--out',
              reportPath,
              '--quiet',
              ...policies.flatMap(policy => ['--policy', policy]),
            ]);
            const report = JSON.parse(readFileSync(reportPath, 'utf8'));
            record.resaveDiff = { matches: report.totals.matches, cells: report.totals.cells, byCategory: report.byCategory };
          }
        }
      } else {
        const dumpPath = join(scratch, `${fixture.id}.${reader}.json`);
        const [command, args] = dumpCommand(reader, file, dumpPath, fixture);
        rmSync(dumpPath, { force: true }); // a crash this time must never read the previous run's dump
        const dumped = run(command, args);
        // A written dump counts even when the reader exited non-zero: that is how a reader reports a classified
        // rejection (ours does, with `error.code`) as opposed to crashing without saying anything.
        const dump = readDump(dumpPath);
        if (!dump) {
          record.status = fixture.expectedError ? 'REJECTED' : 'ERROR';
          record.expectedError = fixture.expectedError;
          record.error =
            summarizeError(dumped.stderr || dumped.stdout) ||
            (dumped.signal ? `killed by ${dumped.signal} (timeout?)` : `exit ${dumped.status}`);
        } else {
          record.readerVersion = dump.reader;
          // `error` is a message string from the python readers and { code, message } from ours
          const readerError = dump.error ?? null;
          record.errorCode = readerError?.code;
          const errorText = readerError
            ? readerError.code
              ? `${readerError.code}: ${readerError.message}`
              : String(readerError)
            : undefined;
          if (fixture.expectedError) {
            // hostile fixture: a classified rejection is the desired outcome, the right code is a pass
            record.expectedError = fixture.expectedError;
            record.status = !readerError ? 'ACCEPTED' : rejectionMatches(fixture.expectedError, readerError.code) ? 'PASS' : 'REJECTED';
            record.error = errorText;
            record.sheets = dump.sheets?.map(sheet => ({ name: sheet.name, rows: sheet.rows.length }));
          } else if (readerError) {
            record.status = 'ERROR';
            record.error = errorText;
          } else if (!expectedPath) {
            record.status = 'OPENED';
            record.sheets = dump.sheets.map(sheet => ({ name: sheet.name, rows: sheet.rows.length }));
          } else {
            const reportPath = join(scratch, `${fixture.id}.${reader}.diff.json`);
            const diff = run('node', [
              join(HERE, 'diff.mjs'),
              expectedPath,
              dumpPath,
              '--out',
              reportPath,
              '--quiet',
              ...policies.flatMap(policy => ['--policy', policy]),
            ]);
            const report = JSON.parse(readFileSync(reportPath, 'utf8'));
            record.status = report.pass ? 'PASS' : 'DIFF';
            record.cells = report.totals.cells;
            record.matches = report.totals.matches;
            record.byCategory = report.byCategory;
            record.structural = Object.fromEntries(
              report.sheets.filter(sheet => sheet.structural).map(sheet => [sheet.name, sheet.structural]),
            );
            record.readerErrors = dump.errors?.length ? dump.errors : undefined;
            record.diffExit = diff.status;
          }
        }
      }
    } catch (error) {
      record.status = 'ERROR';
      record.error = error.message;
    }
    record.ms = Date.now() - started;
    results.push(record);
    const detail =
      record.status === 'DIFF' || (record.status === 'PASS' && record.byCategory)
        ? `${record.matches}/${record.cells} ${JSON.stringify(record.byCategory)}`
        : reader === 'excel'
          ? `${record.sheets?.length ?? '?'} sheets, ${record.cells?.length ?? 0} sentinel cells${record.repairLogs ? ' REPAIR: ' + record.repairLogs.flatMap(log => log.removed).join('; ') : ''}`
          : record.status === 'ACCEPTED'
            ? `opened ${JSON.stringify(record.sheets)} (expected ${record.expectedError})`
            : (record.error ?? record.stderr ?? '');
    console.log(`${record.status.padEnd(6)} ${fixture.id.padEnd(44)} ${reader.padEnd(18)} ${String(detail).slice(0, 140)}`);
    // per-record files are handy while debugging but redundant with results.json, so they go to scratch
    writeFileSync(join(scratch, `${fixture.id}.${reader}.record.json`), JSON.stringify(record, null, 2) + '\n');
  }
}

const meta = {
  date: new Date().toISOString(),
  host: hostname(),
  label: options.label,
  node: process.version,
  readers,
  fixtures: fixtures.map(fixture => fixture.id),
};
writeFileSync(join(outDir, 'results.json'), JSON.stringify({ meta, results }, null, 2) + '\n');

// summary.md: one row per fixture, one column per reader
const header = `| fixture | ${readers.join(' | ')} |`;
const separator = `|---|${readers.map(() => '---').join('|')}|`;
const rows = fixtures.map(fixture => {
  const cells = readers.map(reader => {
    const record = results.find(entry => entry.fixture === fixture.id && entry.reader === reader);
    if (!record) {
      return '';
    }
    if (record.expectedError) {
      return record.status === 'PASS' ? `PASS (${record.errorCode})` : record.status;
    }
    if (record.status === 'PASS' || record.status === 'DIFF') {
      const pct = record.cells ? Math.round((record.matches / record.cells) * 100) : 100;
      return `${record.status === 'PASS' ? 'PASS' : `DIFF ${pct}%`}`;
    }
    if (record.status === 'FAIL' && record.schemaErrors !== undefined) {
      return `FAIL (${record.schemaErrors} schema errors)`;
    }
    if (reader === 'excel' && record.status === 'PASS') {
      return `PASS (${record.sheets?.length ?? '?'} sheets)`;
    }
    return record.status;
  });
  return `| ${fixture.id} | ${cells.join(' | ')} |`;
});
const summary = [
  `# Oracle run ${meta.date} (${meta.label})`,
  '',
  `Host ${meta.host}, Node ${meta.node}. Cell = reader verdict against the fixture's expected dump (PASS = every cell equal after policies; DIFF n% = share of matching cells; OPENED = no ground truth; ERROR = reader threw).`,
  '',
  header,
  separator,
  ...rows,
  '',
].join('\n');
writeFileSync(join(outDir, 'summary.md'), summary);
console.log(`\nwrote ${outDir}/summary.md`);
