// Golden: canonical workbook written by @jetstreamapp/simple-excel itself, in three container/string variants.
// Run with the TypeScript loader so the writer is exercised through its real source:
//
//   node --import tsx fixtures/generators/canonical-simple-excel.mjs [outDir]
//
// The workbook is built by test/helpers/canonical-writer.ts, the same implementation the golden-bytes suite pins,
// so the fixtures and the byte snapshot can never drift apart.
import { writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { writeCanonicalWorkbook } from '../../test/helpers/canonical-writer.ts';
import { ensureDir, newTracker, writeFeatureSidecar } from './_shared.mjs';

const DEFAULT_OUT_DIR = new URL('../golden/simple-excel/', import.meta.url).pathname;

/** Features the writer applies to the canonical workbook, and the ones it deliberately does not write yet. */
const APPLIED = ['freeze', 'autoFilter', 'merges', 'columnWidths', 'hiddenSheet', 'numFmts'];
const SKIPPED = [
  { feature: 'hyperlink', reason: 'the writer has no hyperlink API yet (worksheet <hyperlinks> + rel)' },
  { feature: 'note', reason: 'legacy comments need comments1.xml + vmlDrawing, not implemented' },
  { feature: 'richText', reason: 'cells are written as plain text; <r> runs are read-only for now' },
  { feature: 'validation', reason: 'no dataValidations API yet' },
  { feature: 'conditionalFormat', reason: 'no conditionalFormatting API yet' },
  { feature: 'hiddenRows', reason: 'row options are not exposed by writeRow' },
  { feature: 'hiddenColumns', reason: 'ColumnOptions.hidden exists but the canonical hidden column D has no width entry' },
];

const VARIANTS = [
  { file: 'canonical.xlsx', options: { strings: 'auto' }, note: "strings: 'auto' (bounded shared-string table), deflate, zip64 auto" },
  { file: 'canonical.inline.xlsx', options: { strings: 'inline' }, note: 'the default: every string inline, no sharedStrings part' },
  { file: 'canonical.zip64.xlsx', options: { zip64: true }, note: 'zip64 declared in every local header' },
];

export default async function generate(outDirOrFile = DEFAULT_OUT_DIR) {
  // `fixtures/generate.mjs` hands a single file path; a bare directory writes all three variants.
  const outDir = outDirOrFile.endsWith('.xlsx') ? dirname(outDirOrFile) : outDirOrFile;
  const written = [];
  for (const { file, options, note } of VARIANTS) {
    const outPath = join(outDir, file);
    ensureDir(outPath);
    const bytes = await writeCanonicalWorkbook(options);
    writeFileSync(outPath, bytes);

    const tracker = newTracker();
    tracker.applied.push(...APPLIED);
    tracker.skipped.push(...SKIPPED);
    writeFeatureSidecar(outPath, `@jetstreamapp/simple-excel (${note})`, tracker.applied, tracker.skipped);
    written.push({ file, bytes: bytes.byteLength });
  }
  return written;
}

if (process.argv[1] && new URL(import.meta.url).pathname === process.argv[1]) {
  const results = await generate(process.argv[2] ?? DEFAULT_OUT_DIR);
  for (const { file, bytes } of results) {
    console.log(`wrote ${file} (${bytes} bytes)`);
  }
}
