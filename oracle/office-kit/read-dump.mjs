#!/usr/bin/env node
// Typed dump of an xlsx as @office-kit/xlsx sees it. Two modes:
//   --mode document   loadWorkbook (full model; structural features available)      [default]
//   --mode stream     loadWorkbookStream + iterValues (values only; what a Jetstream reader would use)
//
//   node read-dump.mjs <file.xlsx> [--out dump.json] [--mode document|stream] [--clock local|utc]
import { readFileSync, writeFileSync } from 'node:fs';
import { parseArgs } from 'node:util';
import { isErrorValue, isFormulaValue, isRichTextValue, richTextToString } from '@office-kit/xlsx/cell';
import { loadWorkbook } from '@office-kit/xlsx/io';
import { fromFile } from '@office-kit/xlsx/node';
import { loadWorkbookStream } from '@office-kit/xlsx/streaming';
import { getSheetState, iterWorksheets } from '@office-kit/xlsx/workbook';
import { isDateFormat } from '@office-kit/xlsx/styles';
import { excelToDate } from '@office-kit/xlsx/utils';
import { getAutoFilter, getFreezePanes, getMergedCells, listComments, listHyperlinks } from '@office-kit/xlsx/worksheet';
import { columnLetter, error, formula, fromJsDate, normalizeRows, stableStringify } from '../lib/typed.mjs';

const { values: options, positionals } = parseArgs({
  allowPositionals: true,
  options: { out: { type: 'string' }, mode: { type: 'string', default: 'document' }, clock: { type: 'string', default: 'local' } },
});
const [file] = positionals;
if (!file) {
  console.error('usage: read-dump.mjs <file.xlsx> [--out dump.json] [--mode document|stream] [--clock local|utc]');
  process.exit(2);
}
const version = JSON.parse(readFileSync(new URL('../../node_modules/@office-kit/xlsx/package.json', import.meta.url), 'utf8')).version;

function typedValue(value) {
  if (value === null || value === undefined) {
    return null;
  }
  if (value instanceof Date) {
    return fromJsDate(value, options.clock);
  }
  if (isFormulaValue(value)) {
    return formula(value.formula, typedValue(value.cachedValue ?? null));
  }
  if (isErrorValue(value)) {
    return error(value.code);
  }
  if (isRichTextValue(value) || value?.kind === 'rich-text') {
    // loaded rich text is { kind: 'rich-text', runs: TextRun[] }; richTextToString wants the runs array
    return richTextToString(value.runs ?? value);
  }
  if (typeof value === 'object') {
    return { $unknown: JSON.stringify(value) };
  }
  return value;
}

const BUILTIN_DATE_NUMFMT_IDS = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47]);

/** Stream mode yields raw serials for date cells; resolve date-ness through the stylesheet like a consumer must. */
function makeDateResolver(styles) {
  const cache = new Map();
  return styleId => {
    if (cache.has(styleId)) {
      return cache.get(styleId);
    }
    const xf = styles?.cellXfs?.[styleId];
    const numFmtId = xf?.numFmtId ?? 0;
    const isDate = BUILTIN_DATE_NUMFMT_IDS.has(numFmtId) || isDateFormat(styles?.numFmts?.get(numFmtId));
    cache.set(styleId, isDate);
    return isDate;
  };
}

const errors = [];
function safe(label, fn, fallback = null) {
  try {
    return fn();
  } catch (caught) {
    errors.push(`${label}: ${caught?.message ?? caught}`);
    return fallback;
  }
}

function setCellAt(rows, row1, col1, value) {
  while (rows.length < row1) {
    rows.push([]);
  }
  const row = rows[row1 - 1];
  while (row.length < col1) {
    row.push(null);
  }
  row[col1 - 1] = value;
}

let sheets;
if (options.mode === 'stream') {
  const workbook = await loadWorkbookStream(fromFile(file));
  sheets = [];
  for (const name of workbook.sheetNames) {
    const worksheet = workbook.openWorksheet(name);
    const isDateStyle = makeDateResolver(workbook.styles);
    const rows = [];
    let rowNumber = 0;
    for await (const cells of worksheet.iterRows()) {
      rowNumber++;
      for (const cell of cells) {
        let value = cell.value;
        if (typeof value === 'number' && isDateStyle(cell.styleId)) {
          value = excelToDate(value, workbook.date1904);
        }
        setCellAt(rows, cell.row, cell.col, typedValue(value));
      }
    }
    sheets.push({
      name,
      hidden: null,
      rows: normalizeRows(rows),
      note: `iterRows yielded ${rowNumber} rows; dates resolved via styles + excelToDate`,
    });
  }
  await workbook.close();
} else {
  const workbook = await loadWorkbook(fromFile(file));
  // workbook.sheets holds SheetRef entries; the Worksheet objects come from iterWorksheets()
  sheets = [...iterWorksheets(workbook)].map(worksheet => {
    const rows = [];
    // iterCells(ws) threw "reading 'length'" in 0.11.0; Worksheet.rows is a public Map<row, Map<col, Cell>>
    // The document loader also leaves date cells as raw serials; resolve through the stylesheet like stream mode.
    const isDateStyle = makeDateResolver(workbook.styles);
    safe(`rows ${worksheet.title}`, () => {
      for (const [rowNumber, cols] of worksheet.rows) {
        for (const [colNumber, cell] of cols) {
          let value = cell.value;
          if (typeof value === 'number' && isDateStyle(cell.styleId)) {
            value = excelToDate(value, workbook.date1904 ?? false);
          }
          setCellAt(rows, rowNumber, colNumber, typedValue(value));
        }
      }
    });
    let state = 'visible';
    try {
      state = getSheetState(workbook, worksheet.title);
    } catch {
      // ignore
    }
    return {
      name: worksheet.title,
      hidden: state !== 'visible',
      rows: normalizeRows(rows),
      merges: safe(
        'getMergedCells',
        () =>
          [...getMergedCells(worksheet)].map(range =>
            typeof range === 'string'
              ? range
              : `${columnLetter(range.minCol - 1)}${range.minRow}:${columnLetter(range.maxCol - 1)}${range.maxRow}`,
          ),
        [],
      ),
      autoFilter: safe('getAutoFilter', () => getAutoFilter(worksheet)?.ref ?? null),
      freeze: safe('getFreezePanes', () => getFreezePanes(worksheet) ?? null),
      hyperlinks: safe(
        'listHyperlinks',
        () => listHyperlinks(worksheet).map(link => ({ cell: link.ref, url: link.target ?? link.location ?? null })),
        [],
      ),
      comments: safe('listComments', () => listComments(worksheet).map(comment => ({ cell: comment.ref, text: comment.text })), []),
    };
  });
}

const json = stableStringify({ reader: `office-kit ${version} (${options.mode}, ${options.clock} clock)`, file, sheets, errors });
if (options.out) {
  writeFileSync(options.out, json);
  console.log(`wrote ${options.out} (${sheets.length} sheets)`);
} else {
  process.stdout.write(json);
}
