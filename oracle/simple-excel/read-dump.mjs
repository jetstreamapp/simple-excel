#!/usr/bin/env node
// Typed dump of an xlsx as @jetstreamapp/simple-excel sees it, through the built bundle (dist/esm/index.mjs) and
// nothing but the public API. This is `test/helpers/dump.ts` in plain Node: same options, same encoding, so the
// oracle's verdict and the corpus suite's verdict cannot drift apart. Run `npm run build` first.
//
//   node read-dump.mjs <file.xlsx> [--out dump.json] [--jetstream]
//                      [--max-inflated-bytes N] [--max-shared-string-chars N]
//     --jetstream  additionally include toObjects() of the first worksheet, the counterpart of the sheetjs
//                  reader's sheet_to_json section (that pair is what the parity suite compares)
//     --max-*      read limits, for the hostile fixtures that are only hostile against a budget: the library's
//                  defaults (1 GiB per part, 256 Mi shared-string chars) are deliberately generous, so a host
//                  application sets its own. test/hostile.test.ts uses 8 MiB / 4 Mi and so does the runner.
//
// A classified XlsxError is a result, not a crash: the dump becomes { reader, file, error: { code, message } },
// written to --out as usual, and the process exits non-zero. run.mjs checks that code against the fixture's
// expectedError. Anything else surfaces as code UNCLASSIFIED, which never matches and so never passes.
import { readFileSync, writeFileSync } from 'node:fs';
import { parseArgs } from 'node:util';
import { isXlsxError, openWorkbook } from '../../dist/esm/index.mjs';
import { error as typedError, fromJsDate, normalizeRows, stableStringify } from '../lib/typed.mjs';

const { values: options, positionals } = parseArgs({
  allowPositionals: true,
  options: {
    out: { type: 'string' },
    jetstream: { type: 'boolean', default: false },
    'max-inflated-bytes': { type: 'string' },
    'max-shared-string-chars': { type: 'string' },
  },
});
const [file] = positionals;
if (!file) {
  console.error('usage: read-dump.mjs <file.xlsx> [--out dump.json] [--jetstream] [--max-inflated-bytes N]');
  process.exit(2);
}

const limits = {};
if (options['max-inflated-bytes']) {
  limits.maxInflatedBytes = Number(options['max-inflated-bytes']);
}
if (options['max-shared-string-chars']) {
  limits.maxSharedStringChars = Number(options['max-shared-string-chars']);
}

const version = JSON.parse(readFileSync(new URL('../../package.json', import.meta.url), 'utf8')).version;
const reader = `simple-excel ${version}`;

/**
 * `dates: 'utc'` puts the file's wall clock in a Date's UTC fields, so reading it back with the UTC getters makes
 * the dump independent of the host timezone; `errors: 'object'` turns error cells into `{ error: '#REF!' }`.
 */
function typedValue(value) {
  if (value instanceof Date) {
    return fromJsDate(value, 'utc');
  }
  if (typeof value === 'object' && value !== null) {
    return typedError(value.error);
  }
  return value ?? null;
}

let workbook = null;
let dump;
try {
  const bytes = new Uint8Array(readFileSync(file));
  workbook = await openWorkbook(bytes, { dates: 'utc', errors: 'object', limits });
  const sheets = [];
  for (const info of workbook.sheets) {
    const rows = [];
    // chartsheets carry no cells; they stay in the list so sheet order and names still compare
    if (info.kind === 'worksheet') {
      for await (const row of workbook.sheet(info.index).rows({ blankRows: true })) {
        rows.push(row.map(typedValue));
      }
    }
    sheets.push({ name: info.name, hidden: info.hidden, rows: normalizeRows(rows) });
  }
  dump = { reader, file, sheets };
  if (Object.keys(limits).length > 0) {
    dump.limits = limits;
  }

  if (options.jetstream) {
    // Object mode with its defaults is Jetstream's parseWorkbook (defval '', blank rows skipped, SheetJS header
    // naming); the sheetjs reader's --jetstream section is the same call through sheet_to_json.
    const worksheet = workbook.sheets.find(info => info.kind === 'worksheet');
    const { rows, headers } = await workbook.sheet(worksheet.index).toObjects();
    dump.jetstream = {
      headers: headers.filter(field => !field.startsWith('__empty')),
      data: rows.map(row => Object.fromEntries(Object.entries(row).map(([key, value]) => [key, typedValue(value)]))),
    };
  }
} catch (caught) {
  dump = {
    reader,
    file,
    sheets: [],
    error: { code: isXlsxError(caught) ? caught.code : 'UNCLASSIFIED', message: String(caught?.message ?? caught) },
  };
} finally {
  try {
    await workbook?.close();
  } catch {
    // a workbook that failed mid-stream may also fail to close; the original error is the one that matters
  }
}

const json = stableStringify(dump);
if (options.out) {
  writeFileSync(options.out, json);
  console.log(`wrote ${options.out} (${dump.sheets.length} sheets)${dump.error ? ` ERROR ${dump.error.code}: ${dump.error.message}` : ''}`);
} else {
  process.stdout.write(json);
}
process.exit(dump.error ? 1 : 0);
