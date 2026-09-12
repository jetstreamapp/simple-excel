// Golden: canonical workbook written by ExcelJS (document API, richest feature coverage of the JS libraries).
import { createRequire } from 'node:module';
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

const require = createRequire(import.meta.url);
const ExcelJS = require('exceljs');
const { version } = require('exceljs/package.json');

function toExcelJsValue(value) {
  if (value === null) {
    return null;
  }
  if (isTemporal(value)) {
    return toLocalDate(value);
  }
  if (isTyped(value, '$error')) {
    return { error: value.$error };
  }
  if (isTyped(value, '$formula')) {
    const cached = value.cached;
    return { formula: value.$formula, result: isTyped(cached, '$error') ? { error: cached.$error } : cached };
  }
  return value;
}

export default async function generate(outPath) {
  const { sheets } = canonicalWorkbook();
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'xlsx-engine research';
  workbook.created = new Date(Date.UTC(2026, 8, 11));
  workbook.modified = workbook.created;
  const tracker = newTracker();

  for (const sheet of sheets) {
    const worksheet = workbook.addWorksheet(sheet.name, sheet.hidden ? { state: 'hidden' } : undefined);
    if (sheet.hidden) {
      tracker.applied.push('hiddenSheet');
    }
    const rows = truncateRows(sheet.rows);
    for (const row of rows) {
      worksheet.addRow(row.map(toExcelJsValue));
    }

    if (sheet.numFmts) {
      sheet.rows[0].forEach((name, index) => {
        const format = sheet.numFmts[name];
        if (format) {
          worksheet.getColumn(index + 1).numFmt = format;
        }
      });
      worksheet.getRow(1).font = { bold: true };
    }

    if (sheet.features) {
      const features = sheet.features;
      await attempt(tracker, 'merges', () => {
        for (const { s, e } of features.merges) {
          worksheet.mergeCells(s.r + 1, s.c + 1, e.r + 1, e.c + 1);
        }
      });
      await attempt(tracker, 'freeze', () => {
        worksheet.views = [{ state: 'frozen', ySplit: features.freeze.rows, xSplit: features.freeze.cols }];
      });
      await attempt(tracker, 'autoFilter', () => {
        worksheet.autoFilter = features.autoFilter;
      });
      await attempt(tracker, 'hyperlink', () => {
        worksheet.getCell(features.hyperlink.cell).value = { text: 'Jetstream', hyperlink: features.hyperlink.url };
      });
      await attempt(tracker, 'richText', () => {
        worksheet.getCell(features.richText.cell).value = {
          richText: features.richText.runs.map(({ text, bold }) => (bold ? { text, font: { bold: true } } : { text })),
        };
      });
      await attempt(tracker, 'note', () => {
        worksheet.getCell(features.note.cell).note = features.note.text;
      });
      await attempt(tracker, 'validation', () => {
        worksheet.dataValidations.add(features.validation.range, {
          type: 'list',
          allowBlank: true,
          formulae: [`"${features.validation.list.join(',')}"`],
        });
      });
      await attempt(tracker, 'conditionalFormat', () => {
        worksheet.addConditionalFormatting({
          ref: features.conditionalFormat.range,
          rules: [
            {
              type: 'cellIs',
              operator: features.conditionalFormat.operator,
              formulae: [features.conditionalFormat.value],
              style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: features.conditionalFormat.fillColor } } },
            },
          ],
        });
      });
      await attempt(tracker, 'hiddenRows', () => {
        for (const rowNumber of features.hiddenRows) {
          worksheet.getRow(rowNumber).hidden = true;
        }
      });
      await attempt(tracker, 'hiddenColumns', () => {
        for (const letter of features.hiddenColumns) {
          worksheet.getColumn(letter).hidden = true;
        }
      });
      await attempt(tracker, 'columnWidths', () => {
        for (const [letter, width] of Object.entries(features.columnWidths)) {
          worksheet.getColumn(letter).width = width;
        }
      });
    }
  }

  ensureDir(outPath);
  await workbook.xlsx.writeFile(outPath);
  writeFeatureSidecar(outPath, `ExcelJS ${version}`, tracker.applied, tracker.skipped);
  return { version };
}

await runStandalone(import.meta.url, generate, new URL('../golden/exceljs/canonical.xlsx', import.meta.url).pathname);
