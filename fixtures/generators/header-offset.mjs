// Goldens: sheets whose header row is not row 1 (EC-HEADER-ROW-FIRST-NONEMPTY) and whose first column is empty
// (EC-HEADER-DIMENSION-START-COLUMN), written by three producers so the parity suite compares object mode with
// SheetJS `sheet_to_json` on them. Every earlier golden starts at A1, which is how the row-1 header default slipped
// past parity.
//
//   node fixtures/generators/header-offset.mjs [--python .generated/venv/bin/python] \
//     [--soffice /Applications/LibreOffice.app/Contents/MacOS/soffice]
//
// Writes fixtures/golden/header-offset/{sheetjs,openpyxl,libreoffice}.xlsx. The openpyxl file needs openpyxl in the
// given Python; the LibreOffice file is converted from a CSV whose first two lines are blank. Run once and register
// the outputs with register.mjs; the committed bytes are the fixtures, not this script's output on another machine.
import { execFileSync } from 'node:child_process';
import { mkdirSync, mkdtempSync, renameSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { parseArgs } from 'node:util';
import * as XLSX from 'xlsx';

const OUT_DIR = new URL('../golden/header-offset/', import.meta.url).pathname;

const HEADER = ['Id', 'Name', 'Amount', 'Active'];
const RECORDS = [
  ['001Xx000003DHP0AAO', 'Zoë Ångström', 1234.5, true],
  ['001Xx000003DHP1AAO', '田中 太郎', -0.25, false],
  ['001Xx000003DHP2AAO', 'Acme, Inc.', 42, true],
];

/** Row 1 blank, headers on row 2 (Excel writes `<dimension ref="A2:D5"/>`). */
const ROW2_SHEET = 'Row2Headers';
/** Rows 1-2 blank and column A empty, headers at B3 (`<dimension ref="B3:E6"/>`). */
const B3_SHEET = 'OffsetB3';

/**
 * `aoa_to_sheet` with an `origin` still starts `!ref` at A1, which would make SheetJS itself read row 1 as the
 * header. Pin the used range the way Excel, openpyxl and LibreOffice write it.
 */
function offsetSheet(origin, range) {
  const sheet = XLSX.utils.aoa_to_sheet([HEADER, ...RECORDS], { origin });
  sheet['!ref'] = range;
  return sheet;
}

function writeSheetJs(path) {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, offsetSheet('A2', 'A2:D5'), ROW2_SHEET);
  XLSX.utils.book_append_sheet(workbook, offsetSheet('B3', 'B3:E6'), B3_SHEET);
  writeFileSync(path, XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer', bookSST: false }));
}

function writeOpenpyxl(path, python) {
  const script = `
import json, sys
from openpyxl import Workbook
header, records, out = json.loads(sys.argv[1]), json.loads(sys.argv[2]), sys.argv[3]
workbook = Workbook()
first = workbook.active
first.title = ${JSON.stringify(ROW2_SHEET)}
second = workbook.create_sheet(${JSON.stringify(B3_SHEET)})
for row_offset, values in enumerate([header] + records):
    for col_offset, value in enumerate(values):
        first.cell(row=2 + row_offset, column=1 + col_offset, value=value)
        second.cell(row=3 + row_offset, column=2 + col_offset, value=value)
workbook.save(out)
`;
  execFileSync(python, ['-c', script, JSON.stringify(HEADER), JSON.stringify(RECORDS), path]);
}

function writeLibreOffice(path, soffice) {
  const scratch = mkdtempSync(join(tmpdir(), 'header-offset-'));
  try {
    const csv = ['', '', ...[HEADER, ...RECORDS].map(values => ',' + values.map(value => JSON.stringify(value)).join(','))].join('\n');
    const csvPath = join(scratch, 'libreoffice.csv');
    writeFileSync(csvPath, csv + '\n');
    execFileSync(soffice, [
      `-env:UserInstallation=file://${join(scratch, 'profile')}`,
      '--headless',
      '--infilter=CSV:44,34,76,1',
      '--convert-to',
      'xlsx',
      '--outdir',
      scratch,
      csvPath,
    ]);
    renameSync(join(scratch, 'libreoffice.xlsx'), path);
  } finally {
    rmSync(scratch, { recursive: true, force: true });
  }
}

const { values } = parseArgs({
  options: {
    python: { type: 'string', default: '.generated/venv/bin/python' },
    soffice: { type: 'string', default: '/Applications/LibreOffice.app/Contents/MacOS/soffice' },
  },
});

mkdirSync(OUT_DIR, { recursive: true });
writeSheetJs(join(OUT_DIR, 'sheetjs.xlsx'));
writeOpenpyxl(join(OUT_DIR, 'openpyxl.xlsx'), values.python);
writeLibreOffice(join(OUT_DIR, 'libreoffice.xlsx'), values.soffice);
console.log(`wrote ${OUT_DIR}{sheetjs,openpyxl,libreoffice}.xlsx`);
