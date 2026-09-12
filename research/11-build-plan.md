# 11 - Build plan

The contract every module is built and tested against. Numbers, rules and edge cases referenced here are settled in
02 (format), 04 (catalog), 05 (matrix), 06 (baseline) and 07 (architecture); this document turns them into files,
signatures, work packages and gates.

_Accurate as of 2026-09-12. Package: `@jetstreamapp/simple-excel`. Target: browsers with Web Streams and
`CompressionStream` (evergreen), Node 20+. Zero runtime dependencies._

## 1. Goals and their gates

| Goal                                                                                | Gate that proves it                                                                                                                                                                                                                                                                                                                                                                                                                                                          |
| ----------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Files we write open cleanly in Excel (Win/Mac), Google Sheets, LibreOffice, Numbers | Every writer fixture: Open XML SDK validator clean; Excel oracle PASS with no recovery log; LibreOffice re-save PASS; manual Google Sheets/Numbers checklist before 1.0                                                                                                                                                                                                                                                                                                      |
| Files from those apps read correctly                                                | Every corpus fixture reads at the verdict recorded in `fixtures/manifest.json` / `test/corpus`; parity with SheetJS-with-Jetstream-options on the goldens after documented policies                                                                                                                                                                                                                                                                                          |
| Flat memory, any size                                                               | `bench`: write 1M×20 in Node under 250 MB heap growth and in a Chrome module worker; 18M cells under 3 GB; read 1M×20 streaming under 200 MB + SST                                                                                                                                                                                                                                                                                                                           |
| Fast                                                                                | On the `mixed` bench dataset (1.7 KB of text per row, 170 MB of XML at 100k×20; Apple M4, Node 24): write ≤ 0.5× and read-typed ≤ 0.5× SheetJS, flat across sizes. Measured 2026-09-12 (`bench/results/*phase-e*`): write 2.77 s platform deflate / 1.71 s zlib vs 9.75 s (0.28× / 0.18×), read-typed 1.62 s vs 6.17 s (0.26×); the platform write is ~1.9 s of deflate, which is the floor. The original 1.2 s / 1.5 s absolutes were set without a dataset and are retired |
| Hostile inputs never crash the host                                                 | every `kind:hostile` fixture → classified `XlsxError`, bounded memory                                                                                                                                                                                                                                                                                                                                                                                                        |
| Browser-first, no bloat                                                             | core bundle ≤ 40 KB min+brotli; no `node:` imports outside `src/node`; no DOM globals                                                                                                                                                                                                                                                                                                                                                                                        |

## 2. Package layout

```
src/
  index.ts                 public API (browser-safe)
  node/index.ts            node entry: fromFile, toFile, toWritable, zlib deflater
  types.ts                 CellValue, CellError, RawCell, Row, options, XlsxError codes
  errors.ts                XlsxError class + codes
  zip/crc32.ts             table-driven CRC32 (slicing-by-8 later if measured)
  zip/zip-writer.ts        streaming zip writer: local headers, data descriptors, zip64 up-front, central dir
  zip/zip-reader.ts        EOCD/zip64 scan, central directory, per-entry inflate stream
  zip/source.ts            RandomAccessSource over ArrayBuffer/Uint8Array/Blob
  compress/deflater.ts     CompressionStream('deflate-raw') adapter + StoredDeflater fallback
  compress/inflater.ts     DecompressionStream('deflate-raw') adapter with byte caps
  xml/escape.ts            escapeText/escapeAttr, encodeCellText/decodeCellText (_xHHHH_), needsSpacePreserve
  xml/tokenizer.ts         push tokenizer (start/text/end), prefix-agnostic, DOCTYPE-rejecting, entity-limited
  sml/cell-ref.ts          column letters cache, parseRef, formatRef
  sml/date.ts              serial <-> wall-clock components; 1900/1904; leap-bug rules
  sml/numfmt.ts            built-in id table, isDateFormat(code), codes we emit
  sml/sheet-name.ts        sanitize + case-insensitive dedupe (31 chars, forbidden chars, apostrophes, History)
  sml/styles.ts            registry (fonts/fills/borders/numFmts/cellXfs) + styles.xml writer + reader
  sml/shared-strings.ts    bounded SST writer (chunked serialization)
  sml/shared-strings-reader.ts  SST reader
  sml/package-parts.ts     [Content_Types].xml, rels, workbook.xml, docProps writers + readers
  sml/worksheet-writer.ts  row-at-a-time XML with chunked encoding, cols/freeze/autofilter/merges
  sml/worksheet-reader.ts  tokenizer-driven cell state machine -> typed rows (async iterator)
  sinks.ts                 ByteSink interface, collectToBlob, collectToBytes, fromWritableStream, toWritableStream
  write/workbook-writer.ts createWorkbookWriter facade
  read/sniff.ts            magic-byte sniffing incl. CFB EncryptedPackage probe, ods mimetype, xlsb content type
  read/open.ts             openWorkbook facade: sheets list, sheet(name).rows()/head()/toObjects()
```

Every module exports plain functions/classes with explicit types; no module imports from a sibling's internals except
through its index. `src/index.ts` re-exports only the public surface.

## 3. Public API (target for 1.0)

> **Shipped surface:** the authoritative API is `src/types.ts` + `src/index.ts` (and `src/node/index.ts`), documented
> under `docs/docs/`. The sketch below was the target the modules were built to; where it differs, the code won:
> `SheetWriter.close()` returns a `SheetWriteSummary`, `Workbook.close()` is async, `defval` lives on
> `ObjectRowsOptions`, the empty-header option is `dropEmptyHeaders`, `ReadLimits` carries `maxXmlDepth` /
> `maxTextLength` while `maxRows` / `maxColumns` are per-read `RowsOptions`, `openWorkbook` also accepts a
> `RandomAccessSource`, `sniff` returns `'zip' | 'html' | 'empty'` too, and the error codes include
> `ZIP_UNSUPPORTED`, `ZIP_CRC_MISMATCH`, `CELL_TOO_LONG`, `ABORTED` and `UNSUPPORTED_ENVIRONMENT`. Cell-ref and
> sheet-name helpers are exported for callers that key data by sheet name or build ranges (08 §6).

```ts
// values
type CellValue = string | number | boolean | Date | null;
interface CellError { readonly error: '#NULL!' | '#DIV/0!' | '#VALUE!' | '#REF!' | '#NAME?' | '#NUM!' | '#N/A' | '#GETTING_DATA' }
type CellInput = CellValue | undefined | bigint | CellError;

// writer
function createWorkbookWriter(sink: ByteSink | WritableStream<Uint8Array>, options?: WorkbookWriterOptions): WorkbookWriter;
interface WorkbookWriterOptions {
  strings?: 'auto' | 'inline' | 'shared';                 // ADR-001; auto = bounded hybrid
  sstBudget?: { maxUnique?: number; maxChars?: number; maxLength?: number };
  zip64?: 'auto' | boolean;                               // ADR-002
  compression?: 'deflate' | 'store';
  deterministic?: boolean;                                // fixed timestamps for golden bytes
  dates?: 'local' | 'utc';                                // ADR-003; which Date fields hold the wall clock
  cellOverflow?: 'truncate' | 'throw';                    // 32,767 policy
  truncationSuffix?: string;
  onCellTruncated?: (count: number) => void;
  onProgress?: (progress: { sheet: string; rows: number; bytesOut: number }) => void;
  signal?: AbortSignal;
  properties?: { creator?: string; title?: string; created?: Date };
}
interface WorkbookWriter {
  addSheet(name: string, options?: SheetOptions): SheetWriter;   // one open sheet at a time
  registerStyle(style: CellStyle): StyleId;
  close(): Promise<WorkbookWriteResult>;                            // finalize: SST, styles, workbook, rels, central dir
  abort(reason?: unknown): Promise<void>;
}
interface SheetOptions {
  header?: CellInput[]; headerStyle?: StyleId | false;
  columns?: Array<{ width?: number; style?: StyleId; hidden?: boolean }>;
  freeze?: { rows?: number; cols?: number };
  autoFilter?: boolean;
  hidden?: boolean;
  rowCount?: number;                                                // enables <dimension> and zip64 sizing
}
interface SheetWriter {
  readonly name: string;                                            // after sanitizing/dedupe
  writeRow(values: readonly CellInput[], styles?: StyleId | readonly (StyleId | undefined)[]): Promise<void>;
  writeRows(rows: Iterable<readonly CellInput[]> | AsyncIterable<readonly CellInput[]>): Promise<void>;
  merge(range: string): void;                                       // 'A1:C1'
  close(): Promise<void>;
}
interface CellStyle { font?: {...}; fill?: {...}; border?: ...; alignment?: {...}; numFmt?: string | number }  // 07 A6

// sinks
interface ByteSink { write(chunk: Uint8Array): Promise<void>; close(): Promise<void>; abort(reason?: unknown): Promise<void> }
function collectToBlob(type?: string): ByteSink & { result(): Promise<Blob> };
function collectToBytes(options?: { maxBytes?: number }): ByteSink & { result(): Uint8Array };
function fromWritableStream(stream: WritableStream<Uint8Array>): ByteSink;
function toWritableStream(sink: ByteSink): WritableStream<Uint8Array>;

// reader
function openWorkbook(input: ArrayBuffer | Uint8Array | Blob, options?: OpenOptions): Promise<Workbook>;
function sniff(bytes: Uint8Array): 'xlsx' | 'xlsb' | 'ods' | 'cfb-encrypted' | 'cfb-legacy' | 'xml' | 'text' | 'unknown';
interface OpenOptions {
  dates?: 'local' | 'utc' | 'serial';
  errors?: 'string' | 'object' | 'null';
  limits?: { maxInflatedBytes?: number; maxSharedStringChars?: number; maxEntries?: number; maxRows?: number; maxColumns?: number };
}
interface Workbook {
  readonly sheets: ReadonlyArray<{ name: string; index: number; kind: 'worksheet' | 'chartsheet'; hidden: boolean }>;
  readonly date1904: boolean;
  sheet(nameOrIndex: string | number): Sheet;
  close(): void;
}
interface Sheet {
  rows(options?: RowsOptions & { mode?: 'array' }): AsyncIterable<CellValue[]>;
  rows(options: RowsOptions & { mode: 'object' }): AsyncIterable<Record<string, CellValue>>;
  toObjects(options?: RowsOptions & ObjectOptions): Promise<{ rows: Record<string, CellValue>[]; headers: string[]; truncated: boolean }>;
  head(rowCount: number): Promise<Map<string, RawCell>>;           // A1-keyed prefix, cancels the inflate stream
}
interface RowsOptions { startRow?: number; maxRows?: number; blankRows?: boolean; defval?: CellValue; formulas?: 'value' | 'text' }
interface ObjectOptions { headerRow?: number; emptyHeaderColumns?: 'drop' | 'keep'; headerNaming?: 'sheetjs' | 'index' }

// errors
class XlsxError extends Error { readonly code: XlsxErrorCode; readonly detail?: unknown }
type XlsxErrorCode = 'NOT_XLSX' | 'ENCRYPTED' | 'LEGACY_XLS' | 'XLSB' | 'ODS' | 'ZIP_TRUNCATED' | 'ZIP_BOMB' | 'ZIP_DUPLICATE_ENTRY' | 'XML_DOCTYPE' | 'XML_MALFORMED' | 'LIMIT_EXCEEDED' | 'SHEET_NOT_FOUND' | 'INVALID_SHEET_NAME' | 'ROW_OUT_OF_RANGE' | 'ENTRY_TOO_LARGE' | 'WRITER_STATE';
```

`@jetstreamapp/simple-excel/node`: `fromFile(path)`, `toFile(path)`, `toWritable(writable)`, `nodeDeflater({ level })`.

## 4. Module contracts (what each agent builds)

Each row is one work package: the module, its exact exports, the rules it owns (catalog ids), and the tests that
prove it. All tests are vitest in `src/<module>/__tests__/`.

### A1 `zip/crc32.ts`

- `crc32(bytes: Uint8Array, seed?: number): number` (running CRC: `seed` is the previous value), `CRC32_INITIAL`.
- Tests: known vectors ("", "a", "123456789" → 0xCBF43926), incremental == whole, 1 MB random vs a reference implementation in the test.

### A2 `zip/zip-writer.ts`

- `class ZipWriter { constructor(sink: ByteSink, options: { zip64: boolean; deterministic: boolean }); async beginEntry(name: string, options?: { method?: 'deflate' | 'store' }): EntryWriter; async close(): Promise<{ bytes: number; entries: number }> }`
- `EntryWriter { write(chunk: Uint8Array): Promise<void>; close(): Promise<void> }` — streams: local header (bit 3 data descriptor, bit 11 UTF-8), deflater output straight to the sink, data descriptor after; central directory buffered in memory; zip64 extra field in every local header and central entry when `zip64: true`, plus zip64 EOCD/locator; otherwise a hard `XlsxError('ENTRY_TOO_LARGE')` before any field overflows 32 bits.
- Rules: 02 §9, EC-ZIP-DATA-DESCRIPTORS, ADR-002. Deterministic mode uses DOS date 2026-01-01 00:00.
- Tests: round trip through our own zip-reader and Node's `zlib.inflateRawSync` per entry; entry order; UTF-8 names (`Ünïcödé.xml`); zip64 mode with small content (structure asserted by parsing offsets); `store` method; abort mid-entry leaves the sink aborted.

### A3 `zip/zip-reader.ts` + `zip/source.ts`

- `interface RandomAccessSource { readonly size: number; read(offset: number, length: number): Promise<Uint8Array> }`; `sourceFrom(input: ArrayBuffer | Uint8Array | Blob): RandomAccessSource`.
- `class ZipReader { static open(source, limits): Promise<ZipReader>; readonly entries: ReadonlyMap<string, ZipEntry>; has(name); stream(name): AsyncIterable<Uint8Array>; read(name): Promise<Uint8Array>; readText(name): Promise<string>; close() }` — EOCD scan of the last 64 KiB + 22, zip64 EOCD/locator, central directory parse, names normalized (`\`→`/`, leading `/` stripped), duplicate names → `XlsxError('ZIP_DUPLICATE_ENTRY')`, local header parsed for its own extra-field length, method 0 and 8 only, inflated-byte caps and ratio guard → `XlsxError('ZIP_BOMB')`, truncated → `XlsxError('ZIP_TRUNCATED')`. CRC verified on both paths; a stream throws `ZIP_CRC_MISMATCH` after its last chunk.
- Tests: every file our writer produces; the `fixtures/edge` zip variants (data descriptors, stored, SST after sheets); `fixtures/hostile` truncated / duplicate / bomb / CRC.

### A4 `compress/deflater.ts`, `compress/inflater.ts`

- `interface Deflater { push(chunk: Uint8Array): Promise<void>; finish(): Promise<void>; readonly bytesIn: number; readonly bytesOut: number }` created by `createDeflater(onChunk: (chunk: Uint8Array) => Promise<void>, options?: { method: 'deflate' | 'store' })`; uses `CompressionStream('deflate-raw')` with the writer/reader pump pattern (never a foreign-realm stream); `StoredDeflater` when unavailable.
- `createInflater(onChunk, { maxBytes })` mirrors it with `DecompressionStream('deflate-raw')` and throws `XlsxError('ZIP_BOMB')` past the cap.
- Tests: round trip 0 B / 1 B / 1 MB random / 30 MB repetitive; back-pressure (a slow sink delays `push`); abort; cap enforcement.

### A5 `xml/escape.ts`

- `escapeText(s)`, `escapeAttr(s)`, `encodeCellText(s): string` (charCode scan; returns the input when clean; `_xHHHH_` for chars < 0x20 except 0x09/0x0A/0x0D and for 0xFFFE/0xFFFF; escapes CR as `_x000D_`; escapes every `_` that starts an `_xHHHH_`-shaped run as `_x005F_`, including overlapping runs; replaces lone surrogates with U+FFFD), `decodeCellText(s)` (lowercase `_x` + 4 hex only; left-to-right, resume after each match), `needsSpacePreserve(s)`.
- Rules: EC-XML-CONTROL-CHARS, -ESCAPE-LITERAL, -ESCAPE-OVERLAP, -ESCAPE-CASE, -ESCAPE-UNDECODED, EC-CRLF-NORMALIZED, EC-XML-SPACE-PRESERVE.
- Tests: the canonical Control/EscapeLiteral/Spaces columns (Excel's verdicts from 05 are the expected values: `_x0041_`→`_x005F_x0041_`, `_x005F_x0041_`→`_x005F_x005F_x005F_x0041_`, `_X0041_` untouched); property test: `decode(encode(s)) === s` for random strings including control chars and surrogates.

### A6 `xml/tokenizer.ts`

- `class XmlTokenizer { constructor(handler: { start(name: string, attrs: AttrReader): void; text(text: string): void; end(name: string): void }); push(chunk: string): void; end(): void }` — local names only (prefix stripped), attributes read lazily by local name, entities: the five predefined + `&#...;`/`&#x...;`, CDATA, comments and PIs skipped, `<!DOCTYPE` → `XlsxError('XML_DOCTYPE')`, malformed → `XlsxError('XML_MALFORMED')`, depth cap (default 256) → `LIMIT_EXCEEDED`, text-node cap. Chunk boundaries anywhere (tags, attributes, entities split across pushes).
- Tests: prefixed elements, CDATA, split-chunk fuzz (every split point of a sample document yields identical events), hostile DOCTYPE/deep nesting, `xml:space` attribute access, numeric char refs (`&#13;`, `&#x1F600;`).

### A7 `sml/cell-ref.ts`, `sml/date.ts`, `sml/numfmt.ts`, `sml/sheet-name.ts`

- `columnLetter(index0)`, `columnIndex(letters)`, `formatRef(row0, col0)`, `parseRef(ref)`, `parseRange(ref)`; limits 1,048,576 × 16,384.
- `serialFromComponents({y,m,d,hh,mm,ss,ms}, date1904)`, `componentsFromSerial(serial, date1904)`; rules: serial 60 fake day, serials < 61 shift, negative/pre-epoch rejected (`ROW_OUT_OF_RANGE` is not it — use a `DateOutOfRange` result), time-only = fraction of a day, ms rounding; `dateFromComponents(c, 'local'|'utc')`, `componentsFromDate(d, 'local'|'utc')`.
- `BUILTIN_NUMFMTS`, `isDateFormat(code | id)`, `isBuiltinDateId(id)`, `DEFAULT_DATE_FORMAT = 'yyyy-mm-dd hh:mm:ss'`.
- `sanitizeSheetName(name, taken: Set<string>)` → replaces `: \ / ? * [ ]` with `_`, strips leading/trailing apostrophes, trims to 31, dedupes case-insensitively with ` (2)` fitted inside 31, `History` → `History_`, empty → `Sheet1`.
- Rules: EC-DATE-1900-LEAP-BUG, EC-DATE-PRE-1900, EC-DATE-TIME-ONLY-NEGATIVE-SERIAL, EC-DATE-1904, EC-DATE-DETECTION-VIA-NUMFMT, EC-SHEET-NAME-31-APOSTROPHE.
- Tests: table-driven from the canonical Date/Time/DateTime columns and the 05 Excel verdicts; round trips both epochs; numFmt detection over the built-in table and the custom codes in the corpus.

### B1 `sml/styles.ts`

- `class StyleRegistry { register(style: CellStyle): StyleId; readonly count: number; toXml(): string }` seeded with Excel's defaults (font 0 Calibri 11, fills none + gray125, border 0 empty, cellXfs 0), custom numFmt ids from 164, built-in codes resolved to ids, `CT_Font` children emitted in schema order, `apply*` attributes, `cellStyleXfs`, `cellStyles` Normal, `dxfs count=0`, `tableStyles`. `parseStyles(xml): { isDateByXf: Uint8Array; numFmts: Map<number,string> }` for the reader.
- Rules: 02 §7, EC-STYLES-FONT-ELEMENT-ORDER (we pass the validator), EC-POI-RGB-6-HEX (accept 6/8 hex on read).
- Tests: validator-clean styles.xml for every style combination in the canonical Features sheet; reader over every corpus styles.xml (incl. POI's `u val="none"`, 6-hex colours).

### B2 `sml/shared-strings.ts`

- `class SharedStringWriter { constructor(budget); intern(text: string): number | -1 (inline); readonly count; readonly uniqueCount; async writeTo(entry: EntryWriter): Promise<void> }` — serializes `<si>` in 64 Ki-char chunks, never one string. `parseSharedStrings(stream, { maxChars }): Promise<string[]>` flattening `<r>` runs, skipping `<rPh>`, decoding escapes.
- Rules: ADR-001, EC-SST-ABSENT-INLINE-ONLY, EC-SST-INDEX-OUT-OF-RANGE (reader returns '' with a warning), EC-RICH-TEXT-RUNS.
- Tests: budget crossing mid-sheet; 1M unique strings stays under the char cap; parse every corpus sharedStrings.xml.

### B3 `sml/package-parts.ts`

- Writers: `contentTypesXml(sheets, hasSst)`, `rootRelsXml()`, `workbookXml(sheets, date1904)`, `workbookRelsXml(sheets, hasSst)`, `coreXml(props, deterministic)`, `appXml(sheets)`. Readers: `parseContentTypes`, `parseRels(base)` (absolute/relative/backslash targets, type matched by suffix so Strict works), `parseWorkbook` (sheets order/name/state/r:id, date1904, chartsheets).
- Rules: 02 §2/§8, EC-STRICT-NAMESPACES, EC-ABSOLUTE-REL-TARGETS, EC-BACKSLASH-REL-TARGETS, EC-HIDDEN-SHEETS, EC-PART-NONSTANDARD-NAMES.
- Tests: parse every corpus fixture's package (goldens + edge); writer output validator-clean.

### B4 `sml/worksheet-writer.ts`

- `class WorksheetWriter` used by the facade: prologue (`sheetPr` hidden state via workbook, `dimension` when known, `sheetViews` freeze, `sheetFormatPr`, `cols`), `writeRow(values, styles)` appending to a string builder flushed at 64 Ki chars through `TextEncoder` → `EntryWriter`, epilogue (`autoFilter`, `mergeCells`), monotonic row/col enforcement, cell serialization fast paths (number → `''+n` with uppercase exponent, boolean, Date via date.ts + style, string → SST or inline with `xml:space`, error `t="e"`, empty → no `<c>`), truncation policy hook.
- Rules: 02 §3 element order, EC-CELL-32767-LIMIT, EC-NUM-*, EC-DATE-SERIAL-EPOCH-UTC-FIELDS (dates option), EC-FORMULA-LIKE-TEXT-WRITTEN-AS-FORMULA (never infer formulas).
- Tests: XML snapshot per canonical row; element order; 32,767 truncation with suffix + count; `=text` stays a string.

### B5 `write/workbook-writer.ts` + `sinks.ts`

- `createWorkbookWriter` per §3; part order: `[Content_Types].xml` first (static with `<Default Extension="xml">` + Overrides) — verified in Phase A oracle, fallback last; `_rels/.rels`, docProps, sheets, SST, styles, workbook, workbook rels; progress/abort; sheet-name sanitizing; single open sheet.
- Sinks: `collectToBlob` folds chunks into sub-Blobs every 32 MiB; `collectToBytes` with `maxBytes`; stream adapters.
- Tests: golden bytes (deterministic) for the canonical workbook; validator on output; read back with our reader; back-pressure with a slow sink; abort; `writeRows` from an async generator of 100k rows under a heap ceiling (`--expose-gc` in a Node test).

### C1 `sml/worksheet-reader.ts`

- `readRows(stream: ReadableStream<Uint8Array>, context: { sst: string[]; isDateByXf: Uint8Array; date1904; options }): AsyncIterable<{ row: number; cells: CellValue[] }>` — tokenizer-driven state machine (row/col inference when `r` is missing, `t` handling for n/s/str/inlineStr/b/e/d, `<v/>` empty, `<f>` skipped or captured, `<is>` runs flattened, escapes decoded only when `_x` present, date-ness via `isDateByXf`), batched per decoded chunk, limits.
- Rules: 02 §3, EC-MISSING-R-ATTRIBUTES, EC-EMPTY-V-ELEMENT, EC-INLINE-STRINGS-CDATA, EC-PREFIXED-ELEMENTS, EC-STRICT-ISO-DATE-PRECISION (`t="d"` with any fraction digits and bare times), EC-NUM-MAX-DOUBLE, EC-TIME-ONLY-HAS-DATE-PART (document our marker: 1899-12-30).
- Tests: every `fixtures/edge` variant against its expected dump; corpus goldens against `canonical.json` with the diff categories allowed per fixture policy.

### C2 `read/sniff.ts` + `read/open.ts`

- `sniff` per §3 (`PK\x03\x04` → zip then content-types check for xlsb/ods; `D0CF11E0` → scan 64 KiB for `EncryptedPackage` → `cfb-encrypted` else `cfb-legacy`; `<?xml`/`<` → `xml`; else `text`). `openWorkbook` wires zip-reader → package-parts → styles → lazy SST → sheet handles; `head()`; object mode with SheetJS naming; errors with messages from a table (encrypted: "This workbook is password-protected. Remove the password in Excel and save it again as .xlsx.").
- Tests: every hostile fixture → its `expectedError`; Jetstream parity suite (`test/parity`): `sheet_to_json`-equivalent output equals the SheetJS dump for every golden except documented SheetJS quirks (serial 60, `_X0041_`, CRLF collapse).

### D `src/node/index.ts`

- `fromFile(path): RandomAccessSource` (FileHandle), `toFile(path): ByteSink`, `toWritable(stream): ByteSink`, `nodeDeflater(level)` swap-in. Tests run in Node only.

### E Integration and tooling

- `bench/engines/simple-excel.mjs` + `simple-excel-zlib.mjs` adapters (shared body in `_simple-excel.mjs`); `oracle` reader `ours` (typed dump through the public API); `test/corpus.test.ts` (manifest-driven), `test/golden-bytes.test.ts`, `test/roundtrip.test.ts`, `test/hostile.test.ts`, `test/parity.test.ts`, `test/memory.test.ts` (Node, `--expose-gc`).

## 5. Phases and fan-out

| Phase | Work packages                                                                  | Parallel?                              | Exit criterion                                                                                                      |
| ----- | ------------------------------------------------------------------------------ | -------------------------------------- | ------------------------------------------------------------------------------------------------------------------- |
| 0     | Scaffold: types, errors, module stubs with signatures, test harness wiring, CI | no (lead)                              | `npm test` green on stubs; corpus/golden suites skip with a reason                                                  |
| A     | A1–A7                                                                          | yes (up to 6 agents, disjoint files)   | each package's tests green; lint/typecheck green                                                                    |
| B     | B1–B5                                                                          | B1/B2/B3 in parallel, then B4, then B5 | canonical workbook written, validator clean, opens in Excel/LibreOffice (oracle run committed), golden bytes pinned |
| C     | C1–C2                                                                          | C1 then C2                             | corpus + hostile + parity suites green                                                                              |
| D     | node entry, bench adapter, oracle `ours` reader, memory tests                  | yes                                    | 06 gates measured and recorded                                                                                      |
| E     | README comparison, docs site pages, CHANGELOG, 1.0 checklist                   | yes                                    | manual Google Sheets/Numbers/Excel Windows pass recorded                                                            |

Agents receive: this document, the module row, the relevant catalog ids, the signatures from `src/types.ts`, and the
rule "touch only your module's files and tests". The lead integrates, runs the full suite, and owns `src/index.ts`,
`types.ts` and the facades.

## 6. Test harness

- **Unit** (`src/**/__tests__`): rules per module, table-driven from the catalog where possible.
- **Corpus** (`test/corpus.test.ts`): iterates `fixtures/manifest.json`; for each fixture with `expected`, reads with
  our reader, diffs with the typed-dump comparator (`oracle/diff.mjs` logic ported to TS), and asserts the mismatch
  categories are a subset of the fixture's allowed set (`policies` and a per-fixture `allowedCategories` list added to
  the manifest as we learn each generator's inherent deviations). Fixtures with `expectedError` assert the code.
- **Golden bytes** (`test/golden-bytes.test.ts`): deterministic writer output for the canonical workbook and the
  bench `mixed` 1k dataset, sha256-pinned; updating requires `UPDATE_GOLDENS=1` and a validator + oracle pass.
- **Round trip**: write → read for every canonical column and every style; property tests for escapes/dates.
- **Hostile**: every `kind:hostile` fixture → classified error, and a 30 MB bomb under the byte cap.
- **Parity**: our object-mode read vs the SheetJS dump for every golden (SheetJS installed as a devDependency).
- **Memory/perf** (`test/memory.test.ts`, Node with `--expose-gc`, opt-in via `RUN_PERF=1`): 200k-row write heap
  growth < 64 MB; 100k read heap < 150 MB.
- **Oracle** (manual, pre-release): `npm run oracle -- --tag kind:golden` including our writer's fixtures; results
  committed; Excel recovery-log detection is the pass/fail.
- **CI**: format, lint, typecheck, build, unit + corpus + golden + hostile + parity, `ooxml-validator` on the
  writer fixtures, bundle-size check. Python readers and LibreOffice stay local (documented) until a container job is
  added.

## 7. Definition of done for 1.0

Status on 2026-09-12:

1. ~~All gates in §1 measured and recorded in 06 (bench) and 05 (oracle) with our engine included.~~ Done
   (`bench/results/2026-09-12-*`, `oracle/results/2026-09-12-*`).
2. ~~Every corpus fixture at its recorded verdict; every hostile fixture rejected with a classified error.~~ Done
   (`test/corpus.test.ts`, `test/hostile.test.ts`, `test/parity.test.ts`, `test/corpus-policies.json`).
3. Golden bytes pinned, validator clean, Excel 365 (macOS) and LibreOffice oracle clean: done. Google Sheets manual
   check of `canonical.xlsx`: PASS (2026-09-12, application-side display quirks only; the zip64 variant is refused
   by Drive, as expected). **Numbers and Windows Excel checks are still open** (`fixtures/golden/simple-excel/STEPS.md`).
4. ~~README comparison table backed by the bench numbers; docs site has write/read/streaming/errors pages.~~ Done.
5. ~~Jetstream adapter notes (08) updated to the final API.~~ Done.
6. Open: create the GitHub repository and push; first release (`npm run release` derives 0.1.0 from the changelog);
   the Jetstream migration itself (08 §5) happens in the Jetstream repo behind a flag.
