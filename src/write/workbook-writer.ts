import { notImplemented } from '../internal/not-implemented';
import type { ByteSink, WorkbookWriter, WorkbookWriterOptions } from '../types';

/**
 * Create a streaming workbook writer over a sink (or a `WritableStream<Uint8Array>`). Parts are written in this
 * order: `[Content_Types].xml`, `_rels/.rels`, `docProps/app.xml`, `docProps/core.xml`, worksheets (one open at a
 * time), `xl/sharedStrings.xml`, `xl/styles.xml`, `xl/workbook.xml`, `xl/_rels/workbook.xml.rels`, central directory.
 */
export function createWorkbookWriter(sink: ByteSink | WritableStream<Uint8Array>, options?: WorkbookWriterOptions): WorkbookWriter {
  void sink;
  void options;
  throw notImplemented('write/workbook-writer');
}
