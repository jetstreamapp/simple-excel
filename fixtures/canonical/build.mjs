#!/usr/bin/env node
import { writeFileSync } from 'node:fs';
import { canonicalWorkbook, COLUMNS, dataRows, isTyped } from './canonical.mjs';

const here = new URL('.', import.meta.url).pathname;
const workbook = canonicalWorkbook();
writeFileSync(`${here}canonical.json`, JSON.stringify(workbook, null, 2) + '\n');

/** CSV rendering of the Data sheet only (what a user would upload / what Excel for the web imports). */
function csvCell(value) {
  if (value === null) {
    return '';
  }
  if (isTyped(value, '$date')) {
    return value.$date;
  }
  if (isTyped(value, '$datetime')) {
    return value.$datetime;
  }
  if (isTyped(value, '$time')) {
    return value.$time;
  }
  const text = typeof value === 'string' ? value : String(value);
  return /[",\r\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

const csv = dataRows()
  .map(row => row.map(csvCell).join(','))
  .join('\r\n');
writeFileSync(`${here}canonical.csv`, csv + '\r\n');
console.log(
  `canonical.json: ${workbook.sheets.length} sheets, Data ${dataRows().length - 1} rows x ${COLUMNS.length} cols; canonical.csv written`,
);
