// Golden: canonical workbook written by SheetJS CE 0.20.3 with Jetstream's production options
// (aoa_to_sheet dense, bookSST:false, cells > 32,767 chars truncated).
import { writeFileSync } from 'node:fs';
import * as XLSX from 'xlsx';
import { toLocalDate } from '../canonical/canonical.mjs';
import {
  attempt,
  canonicalWorkbook,
  ensureDir,
  isTemporal,
  isTyped,
  newTracker,
  runStandalone,
  truncateRows,
  writeFeatureSidecar,
} from './_shared.mjs';

// SheetJS error cells carry the BIFF error code in `v` and the text in `w`.
const ERROR_CODES = { '#NULL!': 0x00, '#DIV/0!': 0x07, '#VALUE!': 0x0f, '#REF!': 0x17, '#NAME?': 0x1d, '#NUM!': 0x24, '#N/A': 0x2a };

function toSheetJsValue(value) {
  if (isTemporal(value)) {
    return toLocalDate(value);
  }
  if (isTyped(value, '$error') || isTyped(value, '$formula')) {
    return null; // placed as cell objects afterwards
  }
  return value;
}

function errorCell(code) {
  return { t: 'e', v: ERROR_CODES[code] ?? 0x2a, w: code };
}

export default async function generate(outPath) {
  const { sheets } = canonicalWorkbook();
  const workbook = XLSX.utils.book_new();
  const tracker = newTracker();

  for (const sheet of sheets) {
    const rows = truncateRows(sheet.rows).map(row => row.map(toSheetJsValue));
    const worksheet = XLSX.utils.aoa_to_sheet(rows, { cellDates: true });

    if (sheet.numFmts) {
      // Column number formats: SheetJS has no column-level format, so stamp every data cell's `z`.
      const header = sheet.rows[0];
      for (let r = 1; r < sheet.rows.length; r++) {
        header.forEach((name, c) => {
          const format = sheet.numFmts[name];
          const address = XLSX.utils.encode_cell({ r, c });
          if (format && worksheet[address]) {
            worksheet[address].z = format;
          }
        });
      }
    }

    // Typed cells (errors / formulas) that aoa_to_sheet cannot express
    sheet.rows.forEach((row, r) => {
      row.forEach((value, c) => {
        const address = XLSX.utils.encode_cell({ r, c });
        if (isTyped(value, '$error')) {
          worksheet[address] = errorCell(value.$error);
        } else if (isTyped(value, '$formula')) {
          const cached = value.cached;
          worksheet[address] = isTyped(cached, '$error')
            ? { ...errorCell(cached.$error), f: value.$formula }
            : { t: typeof cached === 'number' ? 'n' : typeof cached === 'boolean' ? 'b' : 's', v: cached, f: value.$formula };
        }
      });
    });

    if (sheet.features) {
      const features = sheet.features;
      await attempt(tracker, 'merges', () => {
        worksheet['!merges'] = features.merges;
      });
      await attempt(tracker, 'autoFilter', () => {
        worksheet['!autofilter'] = { ref: features.autoFilter };
      });
      await attempt(tracker, 'hyperlink', () => {
        worksheet[features.hyperlink.cell].l = { Target: features.hyperlink.url };
      });
      await attempt(tracker, 'note', () => {
        worksheet[features.note.cell].c = [{ a: 'Jetstream', t: features.note.text }];
      });
      await attempt(tracker, 'columnWidths+hiddenColumns', () => {
        worksheet['!cols'] = [];
        for (const [letter, width] of Object.entries(features.columnWidths)) {
          worksheet['!cols'][XLSX.utils.decode_col(letter)] = { wch: width };
        }
        for (const letter of features.hiddenColumns) {
          worksheet['!cols'][XLSX.utils.decode_col(letter)] = { hidden: true };
        }
      });
      await attempt(tracker, 'hiddenRows', () => {
        worksheet['!rows'] = [];
        for (const rowNumber of features.hiddenRows) {
          worksheet['!rows'][rowNumber - 1] = { hidden: true };
        }
      });
      tracker.skipped.push(
        { feature: 'freeze', reason: 'SheetJS CE does not write sheetViews/pane' },
        { feature: 'richText', reason: 'SheetJS CE writes plain text only' },
        { feature: 'validation', reason: 'not supported in CE' },
        { feature: 'conditionalFormat', reason: 'not supported in CE' },
      );
    }

    XLSX.utils.book_append_sheet(workbook, worksheet, sheet.name);
    if (sheet.hidden) {
      workbook.Workbook ??= { Sheets: [] };
      workbook.Workbook.Sheets ??= [];
      workbook.Workbook.Sheets[workbook.SheetNames.length - 1] = { Hidden: 1 };
      tracker.applied.push('hiddenSheet');
    }
  }

  ensureDir(outPath);
  const bytes = XLSX.write(workbook, { bookType: 'xlsx', bookSST: false, type: 'array', compression: true });
  writeFileSync(outPath, Buffer.from(bytes));
  writeFeatureSidecar(outPath, `SheetJS CE ${XLSX.version}`, tracker.applied, tracker.skipped);
  return { bytes: bytes.byteLength, version: XLSX.version };
}

await runStandalone(import.meta.url, generate, new URL('../golden/sheetjs/canonical.xlsx', import.meta.url).pathname);
