// Golden: canonical workbook written by @office-kit/xlsx (document API, so structural features can be exercised;
// the streaming write-only API is benchmarked separately). Every feature is attempted independently and the
// outcome recorded in the .features.json sidecar - API friction here feeds the evaluation (doc 10).
import { readFileSync } from 'node:fs';
import { makeErrorValue, makeRichText, setFormula } from '@office-kit/xlsx/cell';
import { saveWorkbook } from '@office-kit/xlsx/io';
import { toFile } from '@office-kit/xlsx/node';
import { addDxf, makeDifferentialStyle, makeFill, rgbColor, setBold, setCellNumberFormat } from '@office-kit/xlsx/styles';
import { addWorksheet, createWorkbook } from '@office-kit/xlsx/workbook';
import {
  addConditionalFormatting,
  addDataValidation,
  freezePanes,
  getCell,
  hideColumn,
  hideRow,
  makeAutoFilter,
  makeCfRule,
  makeConditionalFormatting,
  makeDataValidation,
  mergeCells,
  setAutoFilter,
  setCell,
  setColumnWidth,
  setComment,
  setHyperlink,
} from '@office-kit/xlsx/worksheet';
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

// package.json is not in the package's exports map, so read it relative to the workspace's node_modules
const { version } = JSON.parse(readFileSync(new URL('../../node_modules/@office-kit/xlsx/package.json', import.meta.url), 'utf8'));

const COLUMN_LETTERS = { A: 1, B: 2, C: 3, D: 4 };

export default async function generate(outPath) {
  const { sheets } = canonicalWorkbook();
  const workbook = createWorkbook();
  const tracker = newTracker();

  for (const sheet of sheets) {
    const worksheet = addWorksheet(workbook, sheet.name, sheet.hidden ? { state: 'hidden' } : undefined);
    if (sheet.hidden) {
      tracker.applied.push('hiddenSheet');
    }
    const rows = truncateRows(sheet.rows);
    const header = sheet.rows[0];

    rows.forEach((row, r) => {
      row.forEach((value, c) => {
        const rowNumber = r + 1;
        const colNumber = c + 1;
        if (value === null) {
          return;
        }
        if (isTyped(value, '$formula')) {
          const cached = value.cached;
          if (isTyped(cached, '$error')) {
            // setCellFormula/setCellRichText are declared in worksheet.d.ts but not exported from the subpath (0.11.0);
            // the cell-level setFormula works. cachedValue only accepts number | string | boolean.
            tracker.skipped.push({
              feature: `formula cached error at r${rowNumber}c${colNumber}`,
              reason: 'setFormula cachedValue cannot be an error value',
            });
            setFormula(setCell(worksheet, rowNumber, colNumber), value.$formula);
          } else {
            setFormula(setCell(worksheet, rowNumber, colNumber), value.$formula, { cachedValue: cached });
          }
          return;
        }
        if (isTyped(value, '$error')) {
          setCell(worksheet, rowNumber, colNumber, makeErrorValue(value.$error));
          return;
        }
        if (isTemporal(value)) {
          const cell = setCell(worksheet, rowNumber, colNumber, toLocalDate(value));
          const format = sheet.numFmts?.[header[c]] ?? 'yyyy-mm-dd hh:mm:ss';
          setCellNumberFormat(workbook, cell, format);
          return;
        }
        const cell = setCell(worksheet, rowNumber, colNumber, value);
        if (r === 0 && sheet.numFmts) {
          setBold(workbook, cell, true);
        } else if (r > 0 && typeof value === 'number' && sheet.numFmts?.[header[c]]) {
          setCellNumberFormat(workbook, cell, sheet.numFmts[header[c]]);
        }
      });
    });

    if (sheet.features) {
      const features = sheet.features;
      await attempt(tracker, 'merges', () => {
        for (const { s, e } of features.merges) {
          // CellRange's object form is { minRow, minCol, maxRow, maxCol }; the A1 string form is simplest here
          mergeCells(worksheet, `${String.fromCharCode(65 + s.c)}${s.r + 1}:${String.fromCharCode(65 + e.c)}${e.r + 1}`);
        }
      });
      await attempt(tracker, 'freeze', () => freezePanes(worksheet, features.freeze.rows, features.freeze.cols));
      await attempt(tracker, 'autoFilter', () => setAutoFilter(worksheet, makeAutoFilter({ ref: features.autoFilter })));
      await attempt(tracker, 'hyperlink', () =>
        setHyperlink(worksheet, features.hyperlink.cell, { target: features.hyperlink.url, display: 'Jetstream' }),
      );
      // setCellRichText is declared but not exported in 0.11.0. The loaded value shape is
      // { kind: 'rich-text', runs }, and setCell accepts that object (a bare runs array is rejected at save).
      await attempt(tracker, 'richText', () =>
        setCell(worksheet, 8, 2, {
          kind: 'rich-text',
          runs: makeRichText(features.richText.runs.map(({ text, bold }) => (bold ? { text, font: { b: true } } : { text }))),
        }),
      );
      await attempt(tracker, 'note', () =>
        setComment(worksheet, { ref: features.note.cell, author: 'Jetstream', text: features.note.text }),
      );
      await attempt(tracker, 'validation', () =>
        addDataValidation(
          worksheet,
          makeDataValidation({
            type: 'list',
            sqref: features.validation.range,
            formula1: `"${features.validation.list.join(',')}"`,
            allowBlank: true,
          }),
        ),
      );
      await attempt(tracker, 'conditionalFormat', () => {
        const dxfId = addDxf(
          workbook.styles,
          makeDifferentialStyle({
            fill: makeFill({
              patternType: 'solid',
              fgColor: rgbColor(features.conditionalFormat.fillColor),
              bgColor: rgbColor(features.conditionalFormat.fillColor),
            }),
          }),
        );
        addConditionalFormatting(
          worksheet,
          makeConditionalFormatting({
            sqref: features.conditionalFormat.range,
            rules: [
              makeCfRule({
                type: 'cellIs',
                priority: 1,
                operator: features.conditionalFormat.operator,
                formulas: [String(features.conditionalFormat.value)],
                dxfId,
              }),
            ],
          }),
        );
      });
      await attempt(tracker, 'hiddenRows', () => features.hiddenRows.forEach(rowNumber => hideRow(worksheet, rowNumber)));
      await attempt(tracker, 'hiddenColumns', () =>
        features.hiddenColumns.forEach(letter => hideColumn(worksheet, COLUMN_LETTERS[letter])),
      );
      await attempt(tracker, 'columnWidths', () => {
        for (const [letter, width] of Object.entries(features.columnWidths)) {
          setColumnWidth(worksheet, COLUMN_LETTERS[letter], width);
        }
      });
      // sanity: the header cell survived the merge
      if (!getCell(worksheet, 1, 1)) {
        tracker.skipped.push({ feature: 'merge-keeps-top-left', reason: 'A1 missing after mergeCells' });
      }
    }
  }

  ensureDir(outPath);
  await saveWorkbook(workbook, toFile(outPath));
  writeFeatureSidecar(outPath, `@office-kit/xlsx ${version}`, tracker.applied, tracker.skipped);
  return { version };
}

await runStandalone(import.meta.url, generate, new URL('../golden/office-kit/canonical.xlsx', import.meta.url).pathname);
