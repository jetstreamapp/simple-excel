import { describe, expect, it } from 'vitest';
import { isXlsxError, type XlsxErrorCode } from '../../errors';
import { collectToBytes } from '../../sinks';
import type { ByteSink, CellInput, WorkbookWriteResult, WorkbookWriterOptions, WriteProgress } from '../../types';
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
    const { pkg } = await writeSimpleWorkbook();
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
    const { pkg } = await writeSimpleWorkbook();
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
    const { result } = await writeSimpleWorkbook();
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

  it('surfaces an invalid early merge instead of swallowing it', async () => {
    const workbook = createWorkbookWriter(collectToBytes(), { deterministic: true });
    const sheet = workbook.addSheet('Data');
    sheet.merge('A1:A1');
    expect(await errorCode(() => sheet.writeRow(['x']))).toBe('WRITER_STATE');
    await workbook.abort();
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
    const workbook = createWorkbookWriter(sink, { deterministic: true, sstBudget: { maxUnique: 2 } });
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

  it('reports the running truncation count once per sheet and at close', async () => {
    const counts: number[] = [];
    const sink = collectToBytes();
    const workbook = createWorkbookWriter(sink, { deterministic: true, onCellTruncated: count => counts.push(count) });
    const first = workbook.addSheet('One');
    await first.writeRow(['x'.repeat(40_000), 'short']);
    await first.close();
    const second = workbook.addSheet('Two');
    await second.writeRow(['y'.repeat(40_000)]);
    await second.close();
    const result = await workbook.close();

    expect(counts).toEqual([1, 2]);
    expect(result.truncatedCells).toBe(2);
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
    const { pkg } = await writeSimpleWorkbook();
    expect(await pkg.localHeaderVersion('xl/worksheets/sheet1.xml')).toBe(20);
    expect(await pkg.localHeaderVersion('xl/worksheets/sheet2.xml')).toBe(20);
    expect(await pkg.localHeaderVersion('xl/sharedStrings.xml')).toBe(20);
  });

  it('keeps a sheet with an unknown row count 32-bit under zip64: auto (ADR-002: SheetJS cannot read zip64)', async () => {
    const sink = collectToBytes();
    const workbook = createWorkbookWriter(sink, { deterministic: true });
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
    const workbook = createWorkbookWriter(sink, { deterministic: true });
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
