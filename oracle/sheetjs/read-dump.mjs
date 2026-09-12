#!/usr/bin/env node
// Typed dump of an xlsx as SheetJS CE sees it, using Jetstream's exact read options. Also the reference
// reader for parity checks (a replacement engine must produce the same values Jetstream sees today).
//
//   node read-dump.mjs <file.xlsx> [--out dump.json] [--jetstream]
//     --jetstream  additionally include Jetstream's parseWorkbook result for the first sheet
//                  (sheet_to_json with dateNF/defval:''/blankrows:false/rawNumbers:true)
import { readFileSync, writeFileSync } from 'node:fs';
import { parseArgs } from 'node:util';
import * as XLSX from 'xlsx';
import { a1, error, formula, fromJsDate, normalizeRows, stableStringify } from '../lib/typed.mjs';

const { values: options, positionals } = parseArgs({
  allowPositionals: true,
  options: { out: { type: 'string' }, jetstream: { type: 'boolean', default: false } },
});
const [file] = positionals;
if (!file) {
  console.error('usage: read-dump.mjs <file.xlsx> [--out dump.json] [--jetstream]');
  process.exit(2);
}

const bytes = readFileSync(file);
// Jetstream: XLSX.read(content, { cellText: false, cellDates: true, type: 'array' })
const workbook = XLSX.read(bytes, { cellText: false, cellDates: true, type: 'array' });

// cellText:false (Jetstream's option) suppresses `w`, so error cells only carry the BIFF error code in `v`.
const ERROR_NAMES = { 0: '#NULL!', 7: '#DIV/0!', 15: '#VALUE!', 23: '#REF!', 29: '#NAME?', 36: '#NUM!', 42: '#N/A', 43: '#GETTING_DATA' };

function typedCell(cell) {
  if (!cell || cell.t === 'z') {
    return null;
  }
  let value;
  switch (cell.t) {
    case 'e':
      value = error(cell.w ?? ERROR_NAMES[cell.v] ?? String(cell.v));
      break;
    case 'd':
      // SheetJS numdate() builds Dates from a Date.UTC epoch, so the spreadsheet wall clock lives in the UTC
      // components; sheet_to_json() (what Jetstream calls) converts them with utc_to_local. Reading the UTC
      // components here yields exactly what sheet_to_json hands Jetstream.
      value = fromJsDate(cell.v instanceof Date ? cell.v : new Date(cell.v), 'utc');
      break;
    case 'n':
    case 'b':
    case 's':
    case 'str':
    default:
      value = cell.v instanceof Date ? fromJsDate(cell.v, 'utc') : (cell.v ?? null);
  }
  return cell.f !== undefined ? formula(cell.f, value) : value;
}

const sheetStates = workbook.Workbook?.Sheets ?? [];
const sheets = workbook.SheetNames.map((name, index) => {
  const worksheet = workbook.Sheets[name];
  const rows = [];
  if (worksheet['!ref']) {
    const range = XLSX.utils.decode_range(worksheet['!ref']);
    for (let r = 0; r <= range.e.r; r++) {
      const row = [];
      for (let c = 0; c <= range.e.c; c++) {
        row.push(typedCell(worksheet[a1(r, c)]));
      }
      rows.push(row);
    }
  }
  const hyperlinks = [];
  const comments = [];
  for (const [address, cell] of Object.entries(worksheet)) {
    if (address.startsWith('!') || typeof cell !== 'object' || cell === null) {
      continue;
    }
    if (cell.l?.Target) {
      hyperlinks.push({ cell: address, url: cell.l.Target });
    }
    if (Array.isArray(cell.c) && cell.c.length > 0) {
      comments.push({ cell: address, text: cell.c.map(comment => comment.t).join('') });
    }
  }
  return {
    name,
    hidden: Boolean(sheetStates[index]?.Hidden),
    rows: normalizeRows(rows),
    merges: (worksheet['!merges'] ?? []).map(range => XLSX.utils.encode_range(range)),
    hiddenRows: (worksheet['!rows'] ?? []).map((info, i) => (info?.hidden ? i + 1 : null)).filter(v => v !== null),
    hiddenColumns: (worksheet['!cols'] ?? []).map((info, i) => (info?.hidden ? XLSX.utils.encode_col(i) : null)).filter(v => v !== null),
    autoFilter: worksheet['!autofilter']?.ref ?? null,
    hyperlinks,
    comments,
  };
});

const dump = { reader: `sheetjs ${XLSX.version}`, file, sheets };

if (options.jetstream) {
  // Exactly libs/shared/ui-utils parseWorkbook (first sheet; the modal-driven sheet choice is out of scope)
  const selectedSheet = workbook.Sheets[workbook.SheetNames[0]];
  const data = XLSX.utils.sheet_to_json(selectedSheet, {
    dateNF: 'yyyy"-"mm"-"dd"T"hh:mm:ss',
    defval: '',
    blankrows: false,
    rawNumbers: true,
  });
  const headers = data.length > 0 ? Object.keys(data[0]) : [];
  dump.jetstream = {
    headers: headers.filter(field => !field.startsWith('__empty')),
    data: data.map(row =>
      Object.fromEntries(Object.entries(row).map(([key, value]) => [key, value instanceof Date ? fromJsDate(value, 'local') : value])),
    ),
  };
}

const json = stableStringify(dump);
if (options.out) {
  writeFileSync(options.out, json);
  console.log(`wrote ${options.out} (${sheets.length} sheets)`);
} else {
  process.stdout.write(json);
}
