import { openWorkbook } from '../../src/index';
import type { SourceInput } from '../../src/zip/source';
import { normalizeRows, toTyped, type TypedDump, type TypedValue } from './typed';

/**
 * Typed dump of a workbook through the public API, in the same shape the oracle readers produce. UTC date fields are
 * used so the dump does not depend on the machine's timezone (`dates: 'utc'` + UTC getters = the file's wall clock).
 */
export async function dumpWorkbook(input: SourceInput): Promise<TypedDump> {
  const workbook = await openWorkbook(input, { dates: 'utc', errors: 'object' });
  const sheets: TypedDump['sheets'] = [];
  try {
    for (const info of workbook.sheets) {
      const rows: TypedValue[][] = [];
      if (info.kind === 'worksheet') {
        for await (const row of workbook.sheet(info.index).rows({ blankRows: true })) {
          rows.push(row.map(value => toTyped(value, 'utc')));
        }
      }
      sheets.push({ name: info.name, hidden: info.hidden, rows: normalizeRows(rows) });
    }
  } finally {
    await workbook.close();
  }
  return { reader: 'simple-excel', sheets };
}
