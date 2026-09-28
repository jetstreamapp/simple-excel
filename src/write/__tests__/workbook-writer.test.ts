import { execFileSync } from 'node:child_process';
import { existsSync, mkdtempSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { runInNewContext } from 'node:vm';
import { describe, expect, it } from 'vitest';
import * as XLSX from 'xlsx';
import { isXlsxError, XlsxError, type XlsxErrorCode } from '../../errors';
import { collectToBlob, collectToBytes } from '../../sinks';
import type {
  ByteSink,
  CellInput,
  Deflater,
  DeflaterFactory,
  StyleId,
  WorkbookWriteResult,
  WorkbookWriterOptions,
  WriteProgress,
} from '../../types';
import { XmlTokenizer } from '../../xml/tokenizer';
import { sourceFrom } from '../../zip/source';
import { ZipReader } from '../../zip/zip-reader';
import { createWorkbookWriter } from '../workbook-writer';

const ZIP_LIMITS = { maxEntries: 1000, maxInflatedBytes: 1024 * 1024 * 1024 };
const STATIC_PART_ORDER = ['_rels/.rels', 'docProps/core.xml'];
const CLOSING_PART_ORDER = ['xl/styles.xml', 'xl/workbook.xml', 'xl/_rels/workbook.xml.rels', 'docProps/app.xml', '[Content_Types].xml'];

interface Package {
  readonly names: string[];
  readonly bytes: Uint8Array;
  text(name: string): Promise<string>;
  localHeaderVersion(name: string): Promise<number>;
}

async function openPackage(bytes: Uint8Array): Promise<Package> {
  const reader = await ZipReader.open(sourceFrom(bytes), ZIP_LIMITS);
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  return {
    bytes,
    names: [...reader.entries.keys()],
    text: (name: string) => reader.readText(name),
    async localHeaderVersion(name: string): Promise<number> {
      const entry = reader.entries.get(name);
      if (entry === undefined) {
        throw new Error(`no entry ${name}`);
      }
      return view.getUint16(entry.localHeaderOffset + 4, true);
    },
  };
}

/**
 * Elements in one part, which is only countable when the part parses: the tokenizer throws on malformed XML and
 * this throws on a tag that closes something else or is never closed.
 */
function countElements(xml: string, what: string): number {
  const open: string[] = [];
  let elements = 0;
  const tokenizer = new XmlTokenizer({
    start: (name: string) => {
      elements++;
      open.push(name);
    },
    text: () => undefined,
    end: (name: string) => {
      const expected = open.pop();
      if (expected !== name) {
        throw new Error(`${what}: </${name}> closes ${expected ?? 'nothing'}`);
      }
    },
  });
  tokenizer.push(xml);
  tokenizer.end();
  if (open.length > 0) {
    throw new Error(`${what}: unclosed ${open.join(', ')}`);
  }
  return elements;
}

interface Written {
  readonly result: WorkbookWriteResult;
  readonly pkg: Package;
}

/** Write a two-sheet workbook and read the archive back. */
async function writeSimpleWorkbook(options: WorkbookWriterOptions = {}): Promise<Written> {
  const sink = collectToBytes();
  const workbook = createWorkbookWriter(sink, { deterministic: true, ...options });
  const sheet = workbook.addSheet('Accounts', { header: ['Id', 'Name', 'Amount'], rowCount: 2, autoFilter: true });
  await sheet.writeRow(['001', 'Acme', 10]);
  await sheet.writeRow(['002', 'Globex', 20.5]);
  await sheet.close();
  const notes = workbook.addSheet('Notes', { hidden: true, rowCount: 1 });
  await notes.writeRow(['hidden sheet']);
  await notes.close();
  const result = await workbook.close();
  return { result, pkg: await openPackage(sink.result()) };
}

/** The `XlsxErrorCode` a call rejects (or throws) with, so the assertion stays in the test that cares about it. */
async function errorCode(run: () => unknown): Promise<XlsxErrorCode | string> {
  try {
    await run();
  } catch (error) {
    return isXlsxError(error) ? error.code : `not an XlsxError: ${String(error)}`;
  }
  return 'no error was thrown';
}

const delay = (milliseconds: number): Promise<void> => new Promise(resolve => setTimeout(resolve, milliseconds));

describe('package structure', () => {
  it('writes the parts in streaming order with content types last (EC-ZIP-CONTENT-TYPES-LAST)', async () => {
    const { pkg } = await writeSimpleWorkbook({ strings: 'auto' });
    expect(pkg.names).toEqual([
      ...STATIC_PART_ORDER,
      'xl/worksheets/sheet1.xml',
      'xl/worksheets/sheet2.xml',
      'xl/sharedStrings.xml',
      ...CLOSING_PART_ORDER,
    ]);
  });

  it('produces well-formed XML in every part', async () => {
    const { pkg } = await writeSimpleWorkbook();
    for (const name of pkg.names) {
      expect(countElements(await pkg.text(name), name), `${name} parses and is balanced`).toBeGreaterThan(0);
    }
  });

  it('wires content types, relationships and the sheet list together', async () => {
    const { pkg } = await writeSimpleWorkbook({ strings: 'auto' });
    const contentTypes = await pkg.text('[Content_Types].xml');
    for (const name of pkg.names) {
      if (name.endsWith('.xml') && name !== '[Content_Types].xml') {
        expect(contentTypes, `${name} needs an Override`).toContain(`PartName="/${name}"`);
      }
    }

    const rels = await pkg.text('xl/_rels/workbook.xml.rels');
    expect(rels).toContain('Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"');
    expect(rels).toContain('Target="worksheets/sheet1.xml"');
    expect(rels).toContain('Target="worksheets/sheet2.xml"');
    expect(rels, 'styles take the id after the last sheet').toContain('Id="rId3"');
    expect(rels, 'shared strings take the one after that').toContain('Id="rId4"');

    const workbook = await pkg.text('xl/workbook.xml');
    expect(workbook).toContain('<sheet name="Accounts" sheetId="1" r:id="rId1"/>');
    expect(workbook).toContain('<sheet name="Notes" sheetId="2" state="hidden" r:id="rId2"/>');
    expect(workbook, 'an autofilter also needs the hidden defined name Excel writes').toContain(
      '<definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">\'Accounts\'!$A$1:$C$3</definedName>',
    );

    const sheet = await pkg.text('xl/worksheets/sheet1.xml');
    expect(sheet).toContain('<autoFilter ref="A1:C3"/>');
    expect(sheet, 'the first visible sheet is the selected tab').toContain('tabSelected="1"');
    expect(await pkg.text('xl/worksheets/sheet2.xml')).not.toContain('tabSelected');
  });

  it('reports what it wrote', async () => {
    const { result } = await writeSimpleWorkbook({ strings: 'auto' });
    expect(result.sheets).toEqual([
      { name: 'Accounts', rows: 3, columns: 3 },
      { name: 'Notes', rows: 1, columns: 1 },
    ]);
    expect(result.truncatedCells).toBe(0);
    expect(result.sharedStrings).toEqual({ count: 8, uniqueCount: 8, frozen: false });
    expect(result.bytes).toBeGreaterThan(0);
  });

  it('writes byte-identical output across two deterministic runs', async () => {
    const first = await writeSimpleWorkbook();
    const second = await writeSimpleWorkbook();
    expect(Buffer.from(first.pkg.bytes).equals(Buffer.from(second.pkg.bytes))).toBe(true);
  });

  it('accepts a WritableStream as the sink', async () => {
    const chunks: Uint8Array[] = [];
    const stream = new WritableStream<Uint8Array>({
      write(chunk: Uint8Array) {
        chunks.push(chunk.slice());
      },
    });
    const workbook = createWorkbookWriter(stream, { deterministic: true });
    const sheet = workbook.addSheet('Data', { header: ['a'] });
    await sheet.writeRow([1]);
    await sheet.close();
    const result = await workbook.close();

    const total = chunks.reduce((sum, chunk) => sum + chunk.byteLength, 0);
    expect(total).toBe(result.bytes);
  });
});

describe('sheets', () => {
  it('sanitizes and de-duplicates sheet names', async () => {
    const sink = collectToBytes();
    const workbook = createWorkbookWriter(sink, { deterministic: true });
    const names: string[] = [];
    for (const name of ['Report: 2026/Q1', 'History', 'Report_ 2026_Q1', '']) {
      const sheet = workbook.addSheet(name);
      names.push(sheet.name);
      await sheet.writeRow(['x']);
      await sheet.close();
    }
    await workbook.close();

    expect(names).toEqual(['Report_ 2026_Q1', 'History_', 'Report_ 2026_Q1 (2)', 'Sheet4']);
    const workbookXml = await (await openPackage(sink.result())).text('xl/workbook.xml');
    for (const name of names) {
      expect(workbookXml).toContain(`name="${name}"`);
    }
  });

  it('writes the header row with the bold preset unless it is switched off', async () => {
    const sink = collectToBytes();
    const workbook = createWorkbookWriter(sink, { deterministic: true, strings: 'inline' });
    const styled = workbook.addSheet('Styled', { header: ['Id'] });
    await styled.close();
    const plain = workbook.addSheet('Plain', { header: ['Id'], headerStyle: false });
    await plain.close();
    await workbook.close();

    const pkg = await openPackage(sink.result());
    expect(await pkg.text('xl/worksheets/sheet1.xml')).toContain('<c r="A1" s="1" t="inlineStr">');
    expect(await pkg.text('xl/worksheets/sheet2.xml')).toContain('<c r="A1" t="inlineStr">');
    expect(await pkg.text('xl/styles.xml')).toContain('<font><b/>');
  });

  it('points activeTab at the first visible sheet', async () => {
    const sink = collectToBytes();
    const workbook = createWorkbookWriter(sink, { deterministic: true });
    for (const [name, hidden] of [
      ['Secret', true],
      ['Visible', false],
    ] as const) {
      const sheet = workbook.addSheet(name, { hidden });
      await sheet.writeRow(['x']);
      await sheet.close();
    }
    await workbook.close();

    const pkg = await openPackage(sink.result());
    expect(await pkg.text('xl/workbook.xml')).toContain('activeTab="1"');
    expect(await pkg.text('xl/worksheets/sheet1.xml')).not.toContain('tabSelected');
    expect(await pkg.text('xl/worksheets/sheet2.xml')).toContain('tabSelected="1"');
  });

  it('tracks the next row across the header and the rows written', async () => {
    const workbook = createWorkbookWriter(collectToBytes(), { deterministic: true });
    const sheet = workbook.addSheet('Data', { header: ['Id'] });
    expect(sheet.nextRow).toBe(2);
    await sheet.writeRow([1]);
    expect(sheet.nextRow).toBe(3);
    await sheet.close();
    await workbook.close();
  });

  it('applies a merge registered before the zip entry finished opening', async () => {
    const sink = collectToBytes();
    const workbook = createWorkbookWriter(sink, { deterministic: true });
    const sheet = workbook.addSheet('Data');
    sheet.merge('A1:C1');
    await sheet.writeRow(['Merged header']);
    await sheet.close();
    await workbook.close();

    const sheetXml = await (await openPackage(sink.result())).text('xl/worksheets/sheet1.xml');
    expect(sheetXml).toContain('<mergeCells count="1"><mergeCell ref="A1:C1"/></mergeCells>');
  });

  it('EC-MERGE-OVERLAP: validates a merge synchronously, even before the zip entry has opened', async () => {
    const sink = collectToBytes();
    const workbook = createWorkbookWriter(sink, { deterministic: true });
    const sheet = workbook.addSheet('Data');
    expect(() => sheet.merge('A1:A1')).toThrowError(expect.objectContaining({ code: 'WRITER_STATE' }));
    sheet.merge('a1:b1');
    expect(() => sheet.merge('B1:C2')).toThrowError(/overlaps the merged range "A1:B1"/);
    expect(() => sheet.merge('A1:B1')).toThrowError(/already merged/);
    expect(() => sheet.merge('A1:XFE2')).toThrowError(expect.objectContaining({ code: 'WRITER_STATE' }));
    // A refused merge does not poison the sheet: nothing was written for it.
    await sheet.writeRow(['Merged header']);
    sheet.merge('A2:C3');
    await sheet.close();
    await workbook.close();

    const sheetXml = await (await openPackage(sink.result())).text('xl/worksheets/sheet1.xml');
    expect(sheetXml).toContain('<mergeCells count="2"><mergeCell ref="A1:B1"/><mergeCell ref="A2:C3"/></mergeCells>');
  });

  it('lets the next sheet be added without awaiting the previous close', async () => {
    const sink = collectToBytes();
    const workbook = createWorkbookWriter(sink, { deterministic: true });
    const first = workbook.addSheet('One');
    await first.writeRow(['a']);
    const closing = first.close();
    const second = workbook.addSheet('Two');
    await second.writeRow(['b']);
    await Promise.all([closing, second.close()]);
    await workbook.close();

    const pkg = await openPackage(sink.result());
    expect(pkg.names.filter(name => name.startsWith('xl/worksheets/'))).toEqual(['xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml']);
    expect(await pkg.text('xl/worksheets/sheet2.xml')).toContain('<row r="1">');
  });

  it('writes rows from a sync iterable and from an async generator', async () => {
    async function* generate(): AsyncGenerator<readonly CellInput[]> {
      for (let i = 0; i < 3; i++) {
        yield [i, `row ${i}`];
      }
    }
    const sink = collectToBytes();
    const workbook = createWorkbookWriter(sink, { deterministic: true });
    const sheet = workbook.addSheet('Data');
    await sheet.writeRows([
      [1, 'one'],
      [2, 'two'],
    ]);
    await sheet.writeRows(generate());
    const summary = await sheet.close();
    await workbook.close();

    expect(summary.rows).toBe(5);
    const sheetXml = await (await openPackage(sink.result())).text('xl/worksheets/sheet1.xml');
    expect(sheetXml).toContain('<row r="5">');
  });
});

describe('strings, styles and limits', () => {
  it('writes inline strings by default: no shared-string part, no t="s" cells (ADR-001)', async () => {
    const { pkg, result } = await writeSimpleWorkbook();
    expect(pkg.names).not.toContain('xl/sharedStrings.xml');
    expect(await pkg.text('xl/worksheets/sheet1.xml')).not.toContain('t="s"');
    expect(await pkg.text('xl/worksheets/sheet1.xml')).toContain('t="inlineStr"');
    expect(result.sharedStrings).toEqual({ count: 0, uniqueCount: 0, frozen: false });
  });

  it('omits the shared-string part entirely with strings: inline (EC-SST-ABSENT-INLINE-ONLY)', async () => {
    const { pkg, result } = await writeSimpleWorkbook({ strings: 'inline' });
    expect(pkg.names).not.toContain('xl/sharedStrings.xml');
    expect(await pkg.text('[Content_Types].xml')).not.toContain('sharedStrings');
    expect(await pkg.text('xl/_rels/workbook.xml.rels')).not.toContain('sharedStrings');
    expect(await pkg.text('xl/worksheets/sheet1.xml')).toContain('t="inlineStr"');
    expect(result.sharedStrings).toEqual({ count: 0, uniqueCount: 0, frozen: false });
  });

  it('interns past the default budget with strings: shared', async () => {
    const sink = collectToBytes();
    const workbook = createWorkbookWriter(sink, { deterministic: true, strings: 'shared' });
    const sheet = workbook.addSheet('Data');
    await sheet.writeRow(['x'.repeat(500)]);
    await sheet.close();
    const result = await workbook.close();

    expect(result.sharedStrings).toEqual({ count: 1, uniqueCount: 1, frozen: false });
    expect(await (await openPackage(sink.result())).text('xl/worksheets/sheet1.xml')).toContain('t="s"');
  });

  it('freezes the table mid-sheet with a small budget and keeps the file valid', async () => {
    const sink = collectToBytes();
    const workbook = createWorkbookWriter(sink, { deterministic: true, strings: 'auto', sstBudget: { maxUnique: 2 } });
    const sheet = workbook.addSheet('Data');
    await sheet.writeRow(['a', 'b', 'c', 'a']);
    await sheet.close();
    const result = await workbook.close();

    expect(result.sharedStrings).toEqual({ count: 3, uniqueCount: 2, frozen: true });
    const sheetXml = await (await openPackage(sink.result())).text('xl/worksheets/sheet1.xml');
    expect(sheetXml).toContain('<c r="C1" t="inlineStr"><is><t>c</t></is></c>');
    expect(sheetXml).toContain('<c r="D1" t="s"><v>0</v></c>');
  });

  it('skips the shared-string part when nothing was interned', async () => {
    const sink = collectToBytes();
    const workbook = createWorkbookWriter(sink, { deterministic: true });
    const sheet = workbook.addSheet('Numbers');
    await sheet.writeRow([1, 2, 3]);
    await sheet.close();
    await workbook.close();
    expect((await openPackage(sink.result())).names).not.toContain('xl/sharedStrings.xml');
  });

  it('EC-CELL-32767-LIMIT: reports the workbook truncation total once, at close', async () => {
    const counts: number[] = [];
    const sink = collectToBytes();
    const workbook = createWorkbookWriter(sink, { deterministic: true, onCellTruncated: count => counts.push(count) });
    const first = workbook.addSheet('One');
    await first.writeRow(['x'.repeat(40_000), 'short']);
    await first.close();
    expect(counts, 'nothing is reported as a sheet closes').toEqual([]);
    const second = workbook.addSheet('Two');
    await second.writeRow(['y'.repeat(40_000), 'z'.repeat(33_000)]);
    await second.close();
    const result = await workbook.close();

    expect(counts).toEqual([3]);
    expect(result.truncatedCells).toBe(3);
  });

  it('EC-CELL-32767-LIMIT: does not call onCellTruncated when nothing was truncated or the file failed', async () => {
    const counts: number[] = [];
    const clean = createWorkbookWriter(collectToBytes(), { deterministic: true, onCellTruncated: count => counts.push(count) });
    const sheet = clean.addSheet('Clean');
    await sheet.writeRow(['short']);
    await sheet.close();
    await clean.close();

    const failing = createWorkbookWriter(collectToBytes(), { deterministic: true, onCellTruncated: count => counts.push(count) });
    const broken = failing.addSheet('Broken');
    await broken.writeRow(['x'.repeat(40_000)]);
    expect(await errorCode(() => broken.writeRow([{ not: 'a cell' } as unknown as CellInput]))).toBe('WRITER_STATE');
    expect(await errorCode(() => failing.close())).toBe('WRITER_STATE');
    expect(counts).toEqual([]);
  });

  it('reports progress every 5,000 rows and at sheet close', async () => {
    const progress: WriteProgress[] = [];
    const workbook = createWorkbookWriter(collectToBytes(), { deterministic: true, onProgress: entry => progress.push(entry) });
    const sheet = workbook.addSheet('Data');
    for (let i = 0; i < 12_000; i++) {
      await sheet.writeRow([i]);
    }
    await sheet.close();
    await workbook.close();

    expect(progress.map(entry => entry.rows)).toEqual([5000, 10_000, 12_000]);
    expect(progress.every(entry => entry.sheet === 'Data')).toBe(true);
    expect(progress.at(-1)?.bytesOut).toBeGreaterThan(0);
  });

  it('registers styles that the sheets can use for the whole workbook', async () => {
    const sink = collectToBytes();
    const workbook = createWorkbookWriter(sink, { deterministic: true, strings: 'inline' });
    const currency = workbook.registerStyle({ numFmt: '$#,##0.00' });
    const sheet = workbook.addSheet('Data');
    await sheet.writeRow([1234.5], currency);
    await sheet.close();
    await workbook.close();

    const pkg = await openPackage(sink.result());
    expect(await pkg.text('xl/worksheets/sheet1.xml')).toContain(`<c r="A1" s="${currency}"><v>1234.5</v></c>`);
    expect(await pkg.text('xl/styles.xml')).toContain('formatCode="$#,##0.00"');
  });
});

describe('container decisions', () => {
  it('stores entries uncompressed when asked (EC-ZIP-STORED-ENTRIES)', async () => {
    const { pkg } = await writeSimpleWorkbook({ compression: 'store' });
    const reader = await ZipReader.open(sourceFrom(pkg.bytes), ZIP_LIMITS);
    for (const entry of reader.entries.values()) {
      expect(entry.method).toBe('store');
      expect(entry.compressedSize).toBe(entry.uncompressedSize);
    }
    // Stored bytes are still a readable package, which is what the no-CompressionStream fallback relies on.
    expect(await pkg.text('xl/workbook.xml')).toContain('<sheet name="Accounts"');
  });

  it('keeps small known sheets 32-bit under zip64: auto', async () => {
    const { pkg } = await writeSimpleWorkbook({ strings: 'auto' });
    expect(await pkg.localHeaderVersion('xl/worksheets/sheet1.xml')).toBe(20);
    expect(await pkg.localHeaderVersion('xl/worksheets/sheet2.xml')).toBe(20);
    expect(await pkg.localHeaderVersion('xl/sharedStrings.xml')).toBe(20);
  });

  it('keeps a sheet with an unknown row count 32-bit under zip64: auto (ADR-002: SheetJS cannot read zip64)', async () => {
    const sink = collectToBytes();
    const workbook = createWorkbookWriter(sink, { deterministic: true, strings: 'auto' });
    const streaming = workbook.addSheet('Streaming', { header: ['Id'] });
    await streaming.writeRow(['a']);
    await streaming.close();
    await workbook.close();

    const pkg = await openPackage(sink.result());
    expect(await pkg.localHeaderVersion('xl/worksheets/sheet1.xml')).toBe(20);
    expect(await pkg.localHeaderVersion('xl/sharedStrings.xml')).toBe(20);
  });

  it('gives the shared-string table zip64 when any sheet had it', async () => {
    const sink = collectToBytes();
    const workbook = createWorkbookWriter(sink, { deterministic: true, strings: 'auto' });
    const huge = workbook.addSheet('Huge', { columns: Array.from({ length: 40 }, () => ({})), rowCount: 1_500_000 });
    await huge.writeRow(['a']);
    await huge.close();
    const known = workbook.addSheet('Known', { header: ['Id'], rowCount: 10 });
    await known.writeRow(['b']);
    await known.close();
    await workbook.close();

    const pkg = await openPackage(sink.result());
    expect(await pkg.localHeaderVersion('xl/worksheets/sheet1.xml')).toBe(45);
    expect(await pkg.localHeaderVersion('xl/worksheets/sheet2.xml')).toBe(20);
    expect(await pkg.localHeaderVersion('xl/sharedStrings.xml'), 'the table follows the sheets it serves').toBe(45);
    expect(await pkg.localHeaderVersion('[Content_Types].xml'), 'static parts stay 32-bit').toBe(20);
  });

  it('declares zip64 for a sheet with more than 40M declared cells', async () => {
    const sink = collectToBytes();
    const workbook = createWorkbookWriter(sink, { deterministic: true });
    const huge = workbook.addSheet('Huge', { columns: Array.from({ length: 40 }, () => ({})), rowCount: 1_500_000 });
    await huge.writeRow(['a']);
    await huge.close();
    await workbook.close();

    expect(await (await openPackage(sink.result())).localHeaderVersion('xl/worksheets/sheet1.xml')).toBe(45);
  });

  it('forces zip64 everywhere with zip64: true and nowhere with false', async () => {
    const forced = await writeSimpleWorkbook({ zip64: true });
    expect(await forced.pkg.localHeaderVersion('xl/worksheets/sheet1.xml')).toBe(45);

    const never = await writeSimpleWorkbook({ zip64: false });
    expect(await never.pkg.localHeaderVersion('xl/worksheets/sheet1.xml')).toBe(20);
  });
});

describe('writer state', () => {
  it('allows one open sheet at a time', async () => {
    const workbook = createWorkbookWriter(collectToBytes(), { deterministic: true });
    const sheet = workbook.addSheet('One');
    expect(await errorCode(() => workbook.addSheet('Two'))).toBe('WRITER_STATE');
    expect(await errorCode(() => workbook.close())).toBe('WRITER_STATE');
    await sheet.close();
    const second = workbook.addSheet('Two');
    await second.close();
    await workbook.close();
  });

  it('refuses everything after close', async () => {
    const workbook = createWorkbookWriter(collectToBytes(), { deterministic: true });
    const sheet = workbook.addSheet('One');
    await sheet.close();
    expect(await errorCode(() => sheet.close())).toBe('WRITER_STATE');
    expect(await errorCode(() => sheet.writeRow(['late']))).toBe('WRITER_STATE');
    await workbook.close();
    expect(await errorCode(() => workbook.addSheet('Two'))).toBe('WRITER_STATE');
    expect(await errorCode(() => workbook.registerStyle({ font: { bold: true } }))).toBe('WRITER_STATE');
    expect(await errorCode(() => workbook.close())).toBe('WRITER_STATE');
  });

  it('rejects with ABORTED after abort()', async () => {
    const sink = collectToBytes();
    const workbook = createWorkbookWriter(sink, { deterministic: true });
    const sheet = workbook.addSheet('One');
    await sheet.writeRow(['a']);
    await workbook.abort(new Error('user cancelled'));

    expect(await errorCode(() => sheet.writeRow(['b']))).toBe('ABORTED');
    expect(await errorCode(() => workbook.addSheet('Two'))).toBe('ABORTED');
    expect(await errorCode(() => workbook.close())).toBe('ABORTED');
    await workbook.abort();
  });

  it('aborts when the caller signal fires', async () => {
    const controller = new AbortController();
    const workbook = createWorkbookWriter(collectToBytes(), { deterministic: true, signal: controller.signal });
    const sheet = workbook.addSheet('One');
    await sheet.writeRow(['a']);
    controller.abort(new Error('navigated away'));

    expect(await errorCode(() => sheet.writeRow(['b']))).toBe('ABORTED');
    expect(await errorCode(() => workbook.close())).toBe('ABORTED');
  });
});

describe('back-pressure and throughput', () => {
  it('never runs ahead of a slow sink', async () => {
    let concurrent = 0;
    let maxConcurrent = 0;
    let writes = 0;
    const slowSink: ByteSink = {
      async write(): Promise<void> {
        concurrent++;
        maxConcurrent = Math.max(maxConcurrent, concurrent);
        writes++;
        await delay(2);
        concurrent--;
      },
      async close(): Promise<void> {},
      async abort(): Promise<void> {},
    };

    const started = Date.now();
    const workbook = createWorkbookWriter(slowSink, { deterministic: true, compression: 'store' });
    const sheet = workbook.addSheet('Data');
    for (let i = 0; i < 4000; i++) {
      // Each awaited row that crosses a flush boundary waits for the sink, so nothing queues up behind it.
      await sheet.writeRow([i, `row ${i} with enough text to fill the 64 Ki buffer a few times over`]);
      expect(concurrent).toBe(0);
    }
    await sheet.close();
    await workbook.close();

    expect(maxConcurrent).toBe(1);
    expect(writes).toBeGreaterThan(4);
    expect(Date.now() - started).toBeGreaterThanOrEqual(writes * 2 - 10);
  });

  it('streams 100k x 20 from an async generator', async () => {
    const rowCount = 100_000;
    const columnCount = 20;
    const header = Array.from({ length: columnCount }, (_, index) => `Column ${index}`);
    async function* generate(): AsyncGenerator<readonly CellInput[]> {
      for (let row = 0; row < rowCount; row++) {
        const values: CellInput[] = [];
        for (let column = 0; column < columnCount; column++) {
          values.push(
            column % 4 === 0 ? row * column : column % 4 === 1 ? `value ${row}-${column}` : column % 4 === 2 ? row % 2 === 0 : null,
          );
        }
        yield values;
      }
    }

    let bytesOut = 0;
    const countingSink: ByteSink = {
      async write(chunk: Uint8Array): Promise<void> {
        bytesOut += chunk.byteLength;
      },
      async close(): Promise<void> {},
      async abort(): Promise<void> {},
    };

    const started = Date.now();
    const workbook = createWorkbookWriter(countingSink, { deterministic: true });
    const sheet = workbook.addSheet('Big', { header, rowCount });
    await sheet.writeRows(generate());
    const summary = await sheet.close();
    const result = await workbook.close();
    const elapsed = Date.now() - started;

    expect(summary).toEqual({ name: 'Big', rows: rowCount + 1, columns: columnCount });
    expect(result.bytes).toBe(bytesOut);
    // Not asserted: the 06 baseline gate is 1.2 s for this shape on the reference machine.
    console.log(`write ${rowCount} x ${columnCount}: ${elapsed} ms, ${result.bytes} bytes out`);
  }, 120_000);
});

// ---- audit fixes ----------------------------------------------------------------------------------------------------

const VALIDATOR = join(process.cwd(), 'node_modules', '.bin', 'ooxml-validator');

/** The Open XML SDK validator's verdict on a package, or null when the validator is not installed. */
function validate(bytes: Uint8Array): { ok: boolean; errors: unknown[] } | null {
  if (!existsSync(VALIDATOR)) {
    return null;
  }
  const file = join(mkdtempSync(join(tmpdir(), 'simple-excel-writer-')), 'book.xlsx');
  writeFileSync(file, bytes);
  return JSON.parse(execFileSync(VALIDATOR, [file], { encoding: 'utf8' })) as { ok: boolean; errors: unknown[] };
}

/** Every sheet as SheetJS reads it: rows of values (header: 1), plus the range it computed. */
function readWithSheetJS(bytes: Uint8Array): { names: string[]; sheets: Record<string, { ref: string | undefined; rows: unknown[][] }> } {
  const workbook = XLSX.read(bytes, { type: 'array', cellDates: true });
  const sheets: Record<string, { ref: string | undefined; rows: unknown[][] }> = {};
  for (const name of workbook.SheetNames) {
    const sheet = workbook.Sheets[name];
    if (sheet === undefined) {
      continue;
    }
    sheets[name] = { ref: sheet['!ref'], rows: XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: null }) };
  }
  return { names: workbook.SheetNames, sheets };
}

/** What a call rejects (or throws) with. */
async function rejection(run: () => unknown): Promise<unknown> {
  try {
    await run();
  } catch (error) {
    return error;
  }
  return new Error('no error was thrown');
}

/** The first character XML 1.0 forbids in a decoded part, or -1. CR, LF and tab are legal. */
function findIllegalXmlCharacter(text: string): number {
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    if ((code < 0x20 && code !== 9 && code !== 10 && code !== 13) || code === 0xfffe || code === 0xffff) {
      return code;
    }
    if (code >= 0xd800 && code <= 0xdbff) {
      const nextCode = text.charCodeAt(i + 1);
      if (!(nextCode >= 0xdc00 && nextCode <= 0xdfff)) {
        return code;
      }
      i++;
    } else if (code >= 0xdc00 && code <= 0xdfff) {
      return code;
    }
  }
  return -1;
}

interface SpySink extends ByteSink {
  readonly writes: number;
  readonly aborts: unknown[];
  readonly closes: number;
}

/** A sink that records aborts and closes, and whose write number `failOnWrite` (1-based) rejects with `failure`. */
function createSpySink(failOnWrite = Number.POSITIVE_INFINITY, failure: unknown = new Error('disk full')): SpySink {
  let writes = 0;
  let closes = 0;
  const aborts: unknown[] = [];
  return {
    get writes(): number {
      return writes;
    },
    get closes(): number {
      return closes;
    },
    aborts,
    async write(): Promise<void> {
      writes++;
      if (writes >= failOnWrite) {
        throw failure;
      }
    },
    async close(): Promise<void> {
      closes++;
    },
    async abort(reason?: unknown): Promise<void> {
      aborts.push(reason);
    },
  };
}

/** Run `body` and collect every unhandled rejection it causes, including ones that surface a macrotask later. */
async function collectUnhandledRejections(body: () => Promise<void>): Promise<unknown[]> {
  const unhandled: unknown[] = [];
  const listener = (reason: unknown): void => {
    unhandled.push(reason);
  };
  process.on('unhandledRejection', listener);
  try {
    await body();
    await delay(50);
  } finally {
    process.off('unhandledRejection', listener);
  }
  return unhandled;
}

describe('workbook-level rules', () => {
  it('EC-WORKBOOK-NO-SHEETS: close() without a sheet rejects with WRITER_STATE and aborts the sink', async () => {
    const spy = createSpySink();
    const workbook = createWorkbookWriter(spy, { deterministic: true });
    const failure = await rejection(() => workbook.close());
    expect(failure).toMatchObject({ code: 'WRITER_STATE' });
    expect((failure as Error).message).toContain('at least one sheet');
    expect(spy.aborts, 'the sink is aborted exactly once, with the error').toEqual([failure]);
    expect(spy.closes).toBe(0);
    expect(await rejection(() => workbook.close()), 'later calls report the same error').toBe(failure);
    expect(await rejection(() => workbook.addSheet('Late'))).toBe(failure);

    const bytes = collectToBytes();
    await rejection(() => createWorkbookWriter(bytes).close());
    expect(() => bytes.result()).toThrowError(expect.objectContaining({ code: 'WRITER_STATE', message: (failure as Error).message }));
  });

  it('EC-ALL-SHEETS-HIDDEN: refuses a workbook whose sheets are all hidden', async () => {
    const spy = createSpySink();
    const workbook = createWorkbookWriter(spy, { deterministic: true });
    for (const name of ['One', 'Two']) {
      const sheet = workbook.addSheet(name, { hidden: true });
      await sheet.writeRow(['x']);
      await sheet.close();
    }
    const failure = await rejection(() => workbook.close());
    expect(failure).toMatchObject({ code: 'WRITER_STATE' });
    expect((failure as Error).message).toContain('hidden');
    expect(spy.aborts).toEqual([failure]);
    expect(spy.closes).toBe(0);
  });

  it('EC-ALL-SHEETS-HIDDEN: a hidden first sheet leaves the active and selected tab on the first visible one', async () => {
    const sink = collectToBytes();
    const workbook = createWorkbookWriter(sink, { deterministic: true });
    for (const [name, hidden] of [
      ['Lookup', true],
      ['Accounts', false],
      ['Contacts', false],
      ['Secret', true],
    ] as const) {
      const sheet = workbook.addSheet(name, { hidden, header: ['Id', 'Name'] });
      await sheet.writeRow([`${name}-1`, 'value']);
      await sheet.close();
    }
    await workbook.close();
    const bytes = sink.result();

    const pkg = await openPackage(bytes);
    expect(await pkg.text('xl/workbook.xml')).toContain('activeTab="1"');
    const selected = [];
    for (let index = 1; index <= 4; index++) {
      selected.push((await pkg.text(`xl/worksheets/sheet${index}.xml`)).includes('tabSelected="1"'));
    }
    expect(selected).toEqual([false, true, false, false]);

    const verdict = validate(bytes);
    if (verdict !== null) {
      expect(verdict.errors).toEqual([]);
      expect(verdict.ok).toBe(true);
    }
    const read = readWithSheetJS(bytes);
    expect(read.names).toEqual(['Lookup', 'Accounts', 'Contacts', 'Secret']);
    expect(read.sheets.Accounts?.rows).toEqual([
      ['Id', 'Name'],
      ['Accounts-1', 'value'],
    ]);
    expect(XLSX.read(bytes, { type: 'array' }).Workbook?.Sheets?.map(sheet => sheet.Hidden)).toEqual([1, 0, 0, 1]);
  });

  it('EC-DIMENSION-FROM-HINT: a rowCount that is too low or too high still reads back fully and exactly', async () => {
    const sink = collectToBytes();
    const workbook = createWorkbookWriter(sink, { deterministic: true });
    const expected: Record<string, unknown[][]> = {};

    // Too low: the hint says 2 rows, 6 arrive, and two of them are wider than the header.
    const low = workbook.addSheet('Low', { header: ['Id', 'Name'], rowCount: 2 });
    expected.Low = [['Id', 'Name', null, null]];
    for (let row = 1; row <= 6; row++) {
      const values = row % 3 === 0 ? [row, `name ${row}`, 'extra', row * 10] : [row, `name ${row}`];
      await low.writeRow(values);
      expected.Low.push([...values, ...Array.from({ length: 4 - values.length }, () => null)]);
    }
    await low.close();

    // Too high: the hint says 1,000 rows, 3 arrive.
    const high = workbook.addSheet('High', { header: ['Id'], rowCount: 1000, columns: [{ width: 12 }, { width: 20 }] });
    expected.High = [['Id']];
    for (let row = 1; row <= 3; row++) {
      await high.writeRow([row]);
      expected.High.push([row]);
    }
    await high.close();
    await workbook.close();

    const bytes = sink.result();
    const pkg = await openPackage(bytes);
    expect(await pkg.text('xl/worksheets/sheet1.xml')).not.toContain('<dimension');
    expect(await pkg.text('xl/worksheets/sheet2.xml')).not.toContain('<dimension');
    const read = readWithSheetJS(bytes);
    expect(read.sheets.Low).toEqual({ ref: 'A1:D7', rows: expected.Low });
    expect(read.sheets.High).toEqual({ ref: 'A1:A4', rows: expected.High });
  });

  it('EC-DIMENSION-FROM-HINT: refuses a rowCount that is not a non-negative safe integer, at addSheet', async () => {
    const sink = collectToBytes();
    const workbook = createWorkbookWriter(sink, { deterministic: true });
    for (const rowCount of [-1, 2.5, Number.NaN, Number.POSITIVE_INFINITY]) {
      expect(await errorCode(() => workbook.addSheet('Data', { rowCount })), String(rowCount)).toBe('WRITER_STATE');
    }
    // Nothing was registered for the refused sheets: the name is still free and the workbook still writes.
    const sheet = workbook.addSheet('Data', { rowCount: 0 });
    expect(sheet.name).toBe('Data');
    await sheet.writeRow(['ok']);
    await sheet.close();
    await workbook.close();
    expect((await openPackage(sink.result())).names).toContain('xl/worksheets/sheet1.xml');
    expect((await openPackage(sink.result())).names).not.toContain('xl/worksheets/sheet2.xml');
  });

  it('EC-DOCPROPS-CREATED-RANGE: refuses an invalid created date before touching the sink', () => {
    const tenThousand = new Date(Date.UTC(10_000, 0, 1));
    const yearZero = new Date(Date.UTC(2000, 0, 1));
    yearZero.setUTCFullYear(0);
    for (const created of [new Date(Number.NaN), tenThousand, yearZero, '2026-01-01' as unknown as Date]) {
      const stream = new WritableStream<Uint8Array>();
      expect(() => createWorkbookWriter(stream, { properties: { created } }), String(created)).toThrowError(
        expect.objectContaining({ code: 'WRITER_STATE' }),
      );
      expect(stream.locked, 'the caller keeps an unlocked stream').toBe(false);
    }
    expect(() => createWorkbookWriter(collectToBytes(), { properties: { created: new Date(Date.UTC(9999, 11, 31)) } })).not.toThrow();
  });

  it('accepts a Date from another realm for properties.created and as a cell value', async () => {
    const foreign = runInNewContext('new Date(Date.UTC(2024, 5, 15, 13, 45, 30))') as Date;
    expect(foreign instanceof Date, 'the vm realm has its own Date').toBe(false);
    const sink = collectToBytes();
    const workbook = createWorkbookWriter(sink, { deterministic: true, dates: 'utc', properties: { created: foreign } });
    const sheet = workbook.addSheet('Data');
    await sheet.writeRow([foreign]);
    await sheet.close();
    await workbook.close();
    const pkg = await openPackage(sink.result());
    expect(await pkg.text('docProps/core.xml')).toContain('2024-06-15T13:45:30Z');
    expect(await pkg.text('xl/worksheets/sheet1.xml')).toMatch(/<c r="A1" s="\d+"><v>45458\.573263888/);
  });

  it('EC-DOCPROPS-CREATED-RANGE: a part that fails to write makes close() reject instead of leaving it out', async () => {
    // Stored parts take three writes each (header, data, descriptor): write 5 is docProps/core.xml's data, inside
    // the step nobody awaits.
    const spy = createSpySink(5);
    const unhandled = await collectUnhandledRejections(async () => {
      const workbook = createWorkbookWriter(spy, { deterministic: true, compression: 'store' });
      const failure = await rejection(async () => {
        const sheet = workbook.addSheet('Data');
        await sheet.writeRow(['x']);
        await sheet.close();
        await workbook.close();
      });
      expect(failure).toMatchObject({ code: 'ABORTED', detail: { cause: new Error('disk full') } });
      expect(await rejection(() => workbook.close())).toBe(failure);
    });
    expect(spy.closes).toBe(0);
    expect(spy.aborts).toHaveLength(1);
    expect(unhandled).toEqual([]);
  });
});

describe('rows and cells', () => {
  it('EC-ROW-NOT-ARRAY: refuses an object row instead of writing an empty one, and fails the workbook', async () => {
    const spy = createSpySink();
    const workbook = createWorkbookWriter(spy, { deterministic: true });
    const sheet = workbook.addSheet('Accounts', { header: ['Id', 'Name'] });
    await sheet.writeRow(['001', 'Acme']);
    const failure = await rejection(() => sheet.writeRow({ Id: '002', Name: 'Globex' } as unknown as CellInput[]));
    expect(failure).toMatchObject({ code: 'WRITER_STATE' });
    expect((failure as Error).message).toContain('Sheet "Accounts" row 3: a row must be an array');
    expect(await rejection(() => sheet.close())).toBe(failure);
    expect(await rejection(() => workbook.close())).toBe(failure);
    expect(spy.aborts).toEqual([failure]);

    const second = createWorkbookWriter(collectToBytes(), { deterministic: true });
    const rows = second.addSheet('Rows');
    expect(await errorCode(() => rows.writeRows([['a'], { b: 1 } as unknown as CellInput[]]))).toBe('WRITER_STATE');
  });

  it('EC-CELL-UNSUPPORTED-TYPE: names the sheet, the cell and the type', async () => {
    const workbook = createWorkbookWriter(collectToBytes(), { deterministic: true });
    const sheet = workbook.addSheet('Accounts', { header: ['Id', 'Name', 'Owner'] });
    for (let row = 0; row < 5; row++) {
      await sheet.writeRow([row, `name ${row}`, 'owner']);
    }
    const failure = await rejection(() => sheet.writeRow([6, 'name 6', { Id: '005' } as unknown as CellInput]));
    expect((failure as Error).message).toBe(
      'Sheet "Accounts" cell C7: a value of type object cannot be written. Convert it to a string, number, boolean, Date or null first.',
    );
  });

  it('EC-ERROR-CODE-UNKNOWN: refuses an error cell whose code is not an Excel error', async () => {
    const workbook = createWorkbookWriter(collectToBytes(), { deterministic: true });
    const sheet = workbook.addSheet('Data');
    await sheet.writeRow([{ error: '#N/A' }]);
    expect(await errorCode(() => sheet.writeRow([{ error: 'Something went wrong' } as unknown as CellInput]))).toBe('WRITER_STATE');
  });

  it('EC-ERROR-CODE-UNKNOWN: writes a newer Excel error the reader can return (#SPILL!) as its text', async () => {
    const sink = collectToBytes();
    const workbook = createWorkbookWriter(sink, { deterministic: true });
    const sheet = workbook.addSheet('Data');
    await sheet.writeRow([{ error: '#N/A' }, { error: '#SPILL!' } as unknown as CellInput, { error: '#CALC!' } as unknown as CellInput]);
    await sheet.close();
    await workbook.close();
    const xml = await (await openPackage(sink.result())).text('xl/worksheets/sheet1.xml');
    expect(xml).toContain('<c r="A1" t="e"><v>#N/A</v></c>');
    expect(xml).toContain('<c r="B1" t="inlineStr"><is><t>#SPILL!</t></is></c>');
    expect(xml).not.toContain('t="e"><v>#SPILL!');
    expect(xml).toContain('<c r="C1" t="inlineStr"><is><t>#CALC!</t></is></c>');
  });

  it('EC-STYLE-ID-RANGE: refuses unregistered style ids for rows, headers and columns', async () => {
    const workbook = createWorkbookWriter(collectToBytes(), { deterministic: true });
    const bold = workbook.registerStyle({ font: { bold: true } });
    expect(await errorCode(() => workbook.addSheet('Header', { header: ['Id'], headerStyle: bold + 1 }))).toBe('WRITER_STATE');
    expect(await errorCode(() => workbook.addSheet('Columns', { columns: [{ style: 99 }] }))).toBe('WRITER_STATE');
    expect(await errorCode(() => workbook.addSheet('Columns', { columns: [{ style: -1 }] }))).toBe('WRITER_STATE');

    const sheet = workbook.addSheet('Rows', { header: ['Id'], headerStyle: bold, columns: [{ style: bold }] });
    await sheet.writeRow(['a'], bold);
    await sheet.writeRow(['a', 'b'], [bold, 0]);
    expect(await errorCode(() => sheet.writeRow(['a'], bold + 5))).toBe('WRITER_STATE');
    expect(await errorCode(() => sheet.close()), 'the refused row failed the workbook').toBe('WRITER_STATE');
  });

  it('EC-COLS-WIDTH-OVER-255: clamps a wide column and refuses a NaN, negative or infinite width', async () => {
    const sink = collectToBytes();
    const workbook = createWorkbookWriter(sink, { deterministic: true });
    for (const width of [Number.NaN, -5, Number.POSITIVE_INFINITY]) {
      expect(await errorCode(() => workbook.addSheet('Data', { columns: [{ width }] })), String(width)).toBe('WRITER_STATE');
    }
    const sheet = workbook.addSheet('Data', { columns: [{ width: 300 }, { width: 40 }] });
    await sheet.writeRow(['a', 'b']);
    await sheet.close();
    await workbook.close();
    const bytes = sink.result();
    expect(await (await openPackage(bytes)).text('xl/worksheets/sheet1.xml')).toContain(
      '<cols><col min="1" max="1" width="255" customWidth="1"/><col min="2" max="2" width="40" customWidth="1"/></cols>',
    );
    const verdict = validate(bytes);
    if (verdict !== null) {
      expect(verdict.errors).toEqual([]);
    }
  });

  it('refuses frozen panes that leave no scrolling cell inside the grid, at addSheet', async () => {
    const sink = collectToBytes();
    const workbook = createWorkbookWriter(sink, { deterministic: true });
    const invalid = [{ rows: -1 }, { rows: 1.5 }, { rows: 1_048_576 }, { cols: 16_384 }, { rows: '1' as unknown as number }];
    for (const freeze of invalid) {
      expect(await errorCode(() => workbook.addSheet('Data', { freeze })), JSON.stringify(freeze)).toBe('WRITER_STATE');
    }
    expect(await errorCode(() => workbook.addSheet('Data', { freeze: 'A2' as unknown as { rows: number } }))).toBe('WRITER_STATE');
    const sheet = workbook.addSheet('Data', { freeze: { rows: 1_048_575, cols: 16_383 } });
    await sheet.writeRow(['a']);
    await sheet.close();
    await workbook.close();
    expect(await (await openPackage(sink.result())).text('xl/worksheets/sheet1.xml')).toContain('topLeftCell="XFD1048576"');
  });

  it('EC-STYLE-FIELD-RANGE: registerStyle refuses a bad field without failing the workbook', async () => {
    const sink = collectToBytes();
    const workbook = createWorkbookWriter(sink, { deterministic: true });
    expect(await errorCode(() => workbook.registerStyle({ fill: {} as { color: string } }))).toBe('WRITER_STATE');
    expect(await errorCode(() => workbook.registerStyle({ font: { size: 1000 } }))).toBe('WRITER_STATE');
    const sheet = workbook.addSheet('Data');
    await sheet.writeRow(['still fine'], workbook.registerStyle({ font: { size: 12 } }));
    await sheet.close();
    await workbook.close();
  });

  it('EC-DATE-STYLE-MERGE: a date column styled bold with a fill keeps both, and reads back as dates', async () => {
    const sink = collectToBytes();
    const workbook = createWorkbookWriter(sink, { deterministic: true, dates: 'utc' });
    const highlighted = workbook.registerStyle({ font: { bold: true }, fill: { color: '#FFF2CC' }, border: 'thin' });
    const sheet = workbook.addSheet('Dates', { header: ['When', 'What'] });
    await sheet.writeRow([new Date(Date.UTC(2024, 1, 29, 12)), 'leap day'], [highlighted, highlighted]);
    await sheet.close();
    await workbook.close();
    const bytes = sink.result();

    const cell = XLSX.read(bytes, { type: 'array', cellStyles: true, cellDates: true }).Sheets.Dates?.A2;
    expect(cell?.t).toBe('d');
    const pkg = await openPackage(bytes);
    const styleId = /<c r="A2" s="(\d+)">/.exec(await pkg.text('xl/worksheets/sheet1.xml'))?.[1];
    const xfs = [
      ...(await pkg.text('xl/styles.xml')).matchAll(/<xf numFmtId="(\d+)" fontId="(\d+)" fillId="(\d+)" borderId="(\d+)"/g),
    ].slice(1);
    const derived = xfs[Number(styleId)];
    expect(derived?.slice(1), 'date format, bold font, fill and border all present').toEqual(['164', '1', '2', '1']);
    const verdict = validate(bytes);
    if (verdict !== null) {
      expect(verdict.errors).toEqual([]);
    }
  });

  it('EC-ROW-SNAPSHOT: rows queued before the entry opens keep the values they had when written', async () => {
    const sink = collectToBytes();
    const workbook = createWorkbookWriter(sink, { deterministic: true });
    const header: CellInput[] = ['Id', 'Name'];
    const columns = [{ width: 10 }];
    const sheet = workbook.addSheet('Data', { header, columns });
    header[0] = 'Changed';
    header.push('Extra');
    columns[0] = { width: 99 };
    const row: CellInput[] = [0, 'row 0'];
    const styles: (StyleId | undefined)[] = [undefined, undefined];
    const pending: Promise<void>[] = [];
    for (let index = 0; index < 3; index++) {
      row[0] = index;
      row[1] = `row ${index}`;
      pending.push(sheet.writeRow(row, styles));
    }
    await Promise.all(pending);
    await sheet.close();
    await workbook.close();

    const read = readWithSheetJS(sink.result());
    expect(read.sheets.Data?.rows).toEqual([
      ['Id', 'Name'],
      [0, 'row 0'],
      [1, 'row 1'],
      [2, 'row 2'],
    ]);
    expect(await (await openPackage(sink.result())).text('xl/worksheets/sheet1.xml')).toContain('width="10"');
  });

  it('EC-XML-CONTROL-CHARS-METADATA: every part stays well-formed with control characters in the metadata', async () => {
    const sink = collectToBytes();
    const workbook = createWorkbookWriter(sink, {
      deterministic: true,
      properties: { title: 'Q1\u0001 export\uFFFF', creator: 'Jet\u0000stream\uD800' },
    });
    // A font name with a control character is refused outright (EC-STYLE-FIELD-RANGE); an unpaired surrogate is dropped.
    expect(await errorCode(() => workbook.registerStyle({ font: { name: 'Cal\u0002ibri' } }))).toBe('WRITER_STATE');
    expect(await errorCode(() => workbook.registerStyle({ numFmt: '0.00\u0003' }))).toBe('WRITER_STATE');
    const style = workbook.registerStyle({ font: { name: 'Cal\uD800ibri' }, numFmt: '0.00\uD800' });
    const sheet = workbook.addSheet('Acc\u0001ounts\u001F', { header: ['Name\u0001'] });
    await sheet.writeRow(['text with \u0001 control', 1.5], [undefined, style]);
    await sheet.close();
    await workbook.close();

    expect(sheet.name).toBe('Acc_ounts_');
    const pkg = await openPackage(sink.result());
    for (const name of pkg.names) {
      const text = await pkg.text(name);
      expect(findIllegalXmlCharacter(text), `${name} holds a character XML forbids`).toBe(-1);
      expect(countElements(text, name), `${name} parses`).toBeGreaterThan(0);
    }
    expect(await pkg.text('docProps/core.xml')).toContain('<dc:title>Q1 export</dc:title><dc:creator>Jetstream</dc:creator>');
    expect(await pkg.text('xl/worksheets/sheet1.xml'), 'cell text still keeps them as escapes').toContain('_x0001_');
    const verdict = validate(sink.result());
    if (verdict !== null) {
      expect(verdict.errors).toEqual([]);
    }
  });
});

describe('failure propagation (EC-WRITER-FAILURE-STICKY)', () => {
  const longRow = (index: number): CellInput[] => [index, `row ${index} with enough text to cross the 64 Ki flush boundary soon`];

  it('rejects a refused row at once even while the sink is stalled on back-pressure', async () => {
    let stalled = false;
    const sink: ByteSink = {
      write: () => (stalled ? new Promise<void>(() => {}) : Promise.resolve()),
      close: () => Promise.resolve(),
      abort: () => Promise.resolve(),
    };
    const workbook = createWorkbookWriter(sink, { deterministic: true, compression: 'store' });
    const sheet = workbook.addSheet('Data');
    await sheet.writeRow(longRow(0));
    stalled = true;
    const settled = (promise: Promise<unknown>): Promise<string> =>
      Promise.race([
        promise.then(
          () => 'resolved',
          (error: unknown) => `rejected ${(error as XlsxError).code}`,
        ),
        new Promise<string>(resolve => setTimeout(() => resolve('pending'), 200)),
      ]);
    let stalledFlush: Promise<void> | undefined;
    for (let index = 1; index < 5000 && stalledFlush === undefined; index++) {
      const write = sheet.writeRow(longRow(index));
      if ((await settled(write)) === 'pending') {
        stalledFlush = write;
      }
    }
    expect(stalledFlush, 'a flush should be waiting on the stalled sink').toBeDefined();
    expect(await settled(sheet.writeRow([{ not: 'a cell' } as unknown as CellInput]))).toBe('rejected WRITER_STATE');
  });

  it('a sink that fails mid-sheet aborts once and every later call reports the original error', async () => {
    // Writes 1-6 are the two leading parts and 7 the worksheet's local header, so write 9 is sheet data.
    const spy = createSpySink(9);
    const unhandled = await collectUnhandledRejections(async () => {
      const workbook = createWorkbookWriter(spy, { deterministic: true, compression: 'store' });
      const sheet = workbook.addSheet('Data');
      let failure: unknown;
      let failedAt = -1;
      for (let index = 0; index < 10_000 && failure === undefined; index++) {
        failure = await rejection(() => sheet.writeRow(longRow(index))).then(error =>
          error instanceof Error && error.message === 'no error was thrown' ? undefined : error,
        );
        failedAt = index;
      }
      expect(failedAt, 'the failure surfaced from a flush well into the sheet').toBeGreaterThan(100);
      expect(failure).toMatchObject({ code: 'ABORTED', detail: { cause: new Error('disk full') } });
      expect(await rejection(() => sheet.writeRow(['late']))).toBe(failure);
      expect(await rejection(() => sheet.close())).toBe(failure);
      expect(await rejection(() => workbook.close())).toBe(failure);
      expect(await rejection(() => workbook.addSheet('Next'))).toBe(failure);
      await workbook.abort();
    });
    expect(spy.aborts).toHaveLength(1);
    expect(spy.closes).toBe(0);
    expect(unhandled).toEqual([]);
  });

  it('a refused row makes sheet.close() and workbook.close() report it, not a derived zip error', async () => {
    const spy = createSpySink();
    const unhandled = await collectUnhandledRejections(async () => {
      const workbook = createWorkbookWriter(spy, { deterministic: true });
      const sheet = workbook.addSheet('Data');
      await sheet.writeRow(['fine']);
      const failure = await rejection(() => sheet.writeRow([Symbol('bad') as unknown as CellInput]));
      expect(failure).toMatchObject({ code: 'WRITER_STATE' });
      expect(await rejection(() => sheet.close())).toBe(failure);
      const closing = await rejection(() => workbook.close());
      expect(closing).toBe(failure);
      expect((closing as Error).message).not.toContain('Cannot start the zip entry');
      expect(spy.aborts).toEqual([failure]);
    });
    expect(unhandled).toEqual([]);
  });

  it('a refused row that nobody awaited still reaches close(), and aborts the sink once', async () => {
    const spy = createSpySink();
    const unhandled = await collectUnhandledRejections(async () => {
      const workbook = createWorkbookWriter(spy, { deterministic: true });
      const sheet = workbook.addSheet('Data');
      const ignored = sheet.writeRow([(() => 'not a value') as unknown as CellInput]);
      const sheetClose = rejection(() => sheet.close());
      const workbookClose = rejection(() => workbook.close());
      const failure = await rejection(() => ignored);
      expect(failure).toMatchObject({ code: 'WRITER_STATE' });
      expect(await sheetClose).toBe(failure);
      expect(await workbookClose).toBe(failure);
    });
    expect(spy.aborts).toHaveLength(1);
    expect(unhandled).toEqual([]);
  });

  it('a failing deflater aborts the sink and is reported by every later call', async () => {
    const deflateFailure = new XlsxError('WRITER_STATE', 'compressor broke');
    const brokenDeflater: DeflaterFactory = (): Deflater => ({
      push: async (): Promise<void> => {
        throw deflateFailure;
      },
      finish: async (): Promise<void> => undefined,
      abort: async (): Promise<void> => undefined,
      bytesIn: 0,
      bytesOut: 0,
    });
    const spy = createSpySink();
    const unhandled = await collectUnhandledRejections(async () => {
      const workbook = createWorkbookWriter(spy, { deterministic: true, deflater: brokenDeflater });
      const sheet = workbook.addSheet('Data');
      const failure = await rejection(() => sheet.close());
      expect(failure).toBe(deflateFailure);
      expect(await rejection(() => workbook.close())).toBe(deflateFailure);
      expect(await rejection(() => workbook.addSheet('Next'))).toBe(deflateFailure);
    });
    expect(spy.aborts).toHaveLength(1);
    expect(unhandled).toEqual([]);
  });

  it('an abort after a failure keeps the original error and does not abort the sink twice', async () => {
    const spy = createSpySink();
    const workbook = createWorkbookWriter(spy, { deterministic: true });
    const sheet = workbook.addSheet('Data');
    const failure = await rejection(() => sheet.writeRow([{} as CellInput]));
    await workbook.abort(new Error('user gave up'));
    expect(await rejection(() => workbook.close())).toBe(failure);
    expect(spy.aborts).toHaveLength(1);
  });

  it('collectToBlob rejects with the original error when the writer failed', async () => {
    const sink = collectToBlob();
    const workbook = createWorkbookWriter(sink, { deterministic: true });
    const sheet = workbook.addSheet('Data');
    const failure = await rejection(() => sheet.writeRow([{ error: 'nope' } as unknown as CellInput]));
    await expect(sink.result()).rejects.toBe(failure);
  });
});
