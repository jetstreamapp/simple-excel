import {
  canonicalWorkbook,
  FEATURES,
  isTyped,
  toLocalDate,
  truncateForExcel,
  type TypedValue,
} from '../../fixtures/canonical/canonical.mjs';
import { collectToBytes, createWorkbookWriter } from '../../src/index';
import type { CellInput, StyleId, WorkbookWriterOptions } from '../../src/types';

function toCellInput(value: TypedValue): CellInput {
  if (value === null || typeof value !== 'object') {
    return value;
  }
  if (isTyped(value, '$formula')) {
    // The writer does not write formulas; the cached value is what a data export would contain.
    return toCellInput((value as { cached: TypedValue }).cached);
  }
  if (isTyped(value, '$error')) {
    return { error: (value as { $error: string }).$error as never };
  }
  return toLocalDate(value as { $date?: string; $datetime?: string; $time?: string });
}

/**
 * The canonical workbook written through our own API: every trap value, column number formats, the merged header,
 * freeze pane, autofilter, hidden sheet and 31-character sheet name. Hyperlinks, notes, rich text, validation and
 * conditional formats are out of scope for the writer and are recorded as skipped in the golden's sidecar.
 */
export async function writeCanonicalWorkbook(options: WorkbookWriterOptions = {}): Promise<Uint8Array> {
  const sink = collectToBytes();
  const workbook = createWorkbookWriter(sink, { deterministic: true, dates: 'local', ...options });
  for (const sheetSpec of canonicalWorkbook().sheets) {
    const rows = sheetSpec.rows.map(row => row.map(value => truncateForExcel(toCellInput(value))));
    const header = rows[0] ?? [];
    const columnStyles: (StyleId | undefined)[] = header.map(name =>
      typeof name === 'string' && sheetSpec.numFmts?.[name] ? workbook.registerStyle({ numFmt: sheetSpec.numFmts[name] }) : undefined,
    );
    const isFeatures = sheetSpec.features !== undefined;
    // `<dimension>` comes from the declared width and readers clip to it, so declare the widest row rather than the
    // header: the Hidden sheet's data row ('secret', 42) is one column wider than its header.
    const widestRow = Math.max(...rows.map(row => row.length));
    const sheet = workbook.addSheet(sheetSpec.name, {
      hidden: sheetSpec.hidden,
      header: isFeatures ? undefined : header,
      freeze: isFeatures ? FEATURES.freeze : undefined,
      autoFilter: isFeatures,
      columns: isFeatures
        ? Object.entries(FEATURES.columnWidths).map(([, width]) => ({ width }))
        : Array.from({ length: widestRow }, () => ({})),
      rowCount: rows.length - (isFeatures ? 0 : 1),
    });
    if (isFeatures) {
      for (const [index, row] of rows.entries()) {
        await sheet.writeRow(row, index === 1 ? workbook.registerStyle({ font: { bold: true } }) : undefined);
      }
      for (const merge of FEATURES.merges) {
        await Promise.resolve(sheet.merge(`A${merge.s.r + 1}:${String.fromCharCode(65 + merge.e.c)}${merge.e.r + 1}`));
      }
    } else {
      for (const row of rows.slice(1)) {
        await sheet.writeRow(row, columnStyles);
      }
    }
    await sheet.close();
  }
  await workbook.close();
  return sink.result();
}
