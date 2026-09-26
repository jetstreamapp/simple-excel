// Golden: canonical workbook written by write-excel-file (values + formats only; no structural features).
import { createRequire } from 'node:module';
import writeXlsxFile from 'write-excel-file/node';
import { toLocalDate } from '../canonical/canonical.mjs';
import {
  canonicalWorkbook,
  ensureDir,
  isTemporal,
  isTyped,
  newTracker,
  runStandalone,
  truncateRows,
  writeFeatureSidecar,
} from './_shared.mjs';

const require = createRequire(import.meta.url);
const { version } = require('write-excel-file/package.json');

function toCell(value, format) {
  if (value === null || value === '') {
    return { value: null };
  }
  if (isTemporal(value)) {
    return { value: toLocalDate(value), type: Date, format: format ?? 'yyyy-mm-dd hh:mm:ss' };
  }
  if (isTyped(value, '$error')) {
    return { value: value.$error, type: String };
  }
  if (isTyped(value, '$formula')) {
    return { value: `=${value.$formula}`, type: String };
  }
  if (typeof value === 'number') {
    return { value, type: Number, ...(format ? { format } : {}) };
  }
  if (typeof value === 'boolean') {
    return { value, type: Boolean };
  }
  return { value: String(value), type: String };
}

export default async function generate(outPath) {
  const { sheets } = canonicalWorkbook();
  const tracker = newTracker();
  const visibleSheets = sheets.filter(sheet => !sheet.hidden);
  tracker.skipped.push({ feature: 'hiddenSheet', reason: 'write-excel-file has no sheet state option' });
  tracker.skipped.push({ feature: 'features', reason: 'write-excel-file writes values and cell formats only' });

  const workbookSheets = visibleSheets.map(sheet => {
    const rows = truncateRows(sheet.rows);
    const header = sheet.rows[0];
    const data = rows.map((row, r) =>
      row.map((value, c) => {
        const format = r > 0 && sheet.numFmts ? sheet.numFmts[header[c]] : undefined;
        const cell = toCell(value, format);
        return r === 0 ? { ...cell, fontWeight: 'bold' } : cell;
      }),
    );
    return { data, sheet: sheet.name };
  });

  ensureDir(outPath);
  await writeXlsxFile(workbookSheets).toFile(outPath);
  writeFeatureSidecar(outPath, `write-excel-file ${version}`, tracker.applied, tracker.skipped);
  return { version };
}

await runStandalone(import.meta.url, generate, new URL('../golden/write-excel-file/canonical.xlsx', import.meta.url).pathname);
