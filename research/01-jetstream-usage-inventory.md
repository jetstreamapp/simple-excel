# 01 Jetstream usage inventory

Every place Jetstream produces or consumes spreadsheet files today, the exact SheetJS semantics each site depends on, and the platform constraints any replacement engine inherits.

_Accurate as of 2026-09-11 (branch chore/xlsx-hand-roll)._

Source of truth is the code; line numbers are from this branch and will drift. Three read-only inventories were taken (write paths, read paths, infrastructure) and cross-checked against the plan appendices.

---

## A. Write side

### A1. The choke point: `libs/shared/ui-utils/src/lib/shared-ui-utils.ts`

| Symbol                                                                  | Lines (approx.) | Behavior                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               |
| ----------------------------------------------------------------------- | --------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `prepareExcelFile(data, header?, defaultSheetName='Records', options?)` | 295–372         | Two overloads: `any[]` (array of objects, one sheet named `defaultSheetName`, header auto-detected from `Object.keys(data[0])`), or `Record<string, any[]>` (map of sheet name to rows, each sheet array-of-objects with optional per-sheet header, or array-of-arrays written verbatim). Every sheet is built with `XLSX.utils.aoa_to_sheet(rows, { dense: true })`. Compression is forced on when any sheet exceeds `COMPRESS_SHEET_ROW_COUNT = 10_000` rows. Returns an `ArrayBuffer`. Synchronous. |
| `PrepareExcelFileOptions`                                               | 295–301         | `XLSX.WritingOptions & { onCellsTruncated?: (count) => void }`. The callback fires once, only when at least one cell was truncated.                                                                                                                                                                                                                                                                                                                                                                    |
| `EXCEL_MAX_CELL_CHARS = 32_767`, `truncateCellsForExcel`                | 374–410         | Scans rows, re-allocates only rows containing an oversized string, slices to `32_767 - '...(truncated)'.length` and appends `'...(truncated)'`. Never mutates caller arrays. Only applied inside `prepareExcelFile`. Comment explains why: Salesforce long text is up to 131,072 chars and subquery records are JSON-stringified into one cell.                                                                                                                                                        |
| `excelWorkbookToArrayBuffer(workbook, options?)`                        | 420–431         | `XLSX.write(workbook, { bookType: 'xlsx', bookSST: false, type: 'array', compression: false, ...options })`. Comment: "Compression=true is slower, but helps avoid Invalid Array Length errors on large files."                                                                                                                                                                                                                                                                                        |
| `prepareCsvFile(data, header)`                                          | 433–441         | papaparse `unparse` with `{ header: true, quotes: true, delimiter: detectDelimiter() }` (locale-aware `,` vs `;`).                                                                                                                                                                                                                                                                                                                                                                                     |
| `getMaxWidthFromColumnContent(data, skipRows, defaultIfSkipped=15)`     | 443–460         | Width per column = longest stringified value + 2, ignoring `skipRows`; returns `XLSX.ColInfo[]`.                                                                                                                                                                                                                                                                                                                                                                                                       |
| `saveFile(content, filename, type)`                                     | 470–473         | `new Blob([content], { type })` then file-saver `saveAs`. The only local-download primitive in the repo.                                                                                                                                                                                                                                                                                                                                                                                               |
| `ensureXlsxCodepageTable()`                                             | 66–85           | Lazy `import('xlsx/dist/cpexcel.full.mjs')` then `XLSX.set_cptable`; resolves immediately in the browser extension. Only needed for legacy `.xls` and non-UTF-8 CSV. Not called by `parseFile`.                                                                                                                                                                                                                                                                                                        |

### A2. The bespoke builder: Permission Manager export

`libs/features/manage-permissions/src/utils/permission-manager-export-utils.ts`, `generateExcelWorkbookFromTable` (88–107). Calls `ensureXlsxCodepageTable()`, builds four worksheets, returns `excelWorkbookToArrayBuffer(workbook, { bookSST: true, compression: true })`. Sheet names are four hard-coded literals under 31 characters.

| Sheet              | `aoa_to_sheet` line | Options                                                              | `!cols`                                                                                    | `!merges`                                            | Values                                                  |
| ------------------ | ------------------- | -------------------------------------------------------------------- | ------------------------------------------------------------------------------------------ | ---------------------------------------------------- | ------------------------------------------------------- |
| Object Permissions | 188                 | none (**not dense**)                                                 | `getMaxWidthFromColumnContent(rows, new Set([0]))`                                         | 7-wide per profile/permission set on header row 0    | `'TRUE'`/`'FALSE'` strings                              |
| Field Permissions  | 234                 | `{ cellDates: true, dateNF: 'yyyy-mm-dd hh:mm:ss' }` (**not dense**) | widths measured against `formatDate(date, yyyy_MM_dd_HH_mm_ss)` for Date cells, skip row 0 | 2-wide (`Read`, `Edit`) offset past 7 prefix columns | strings plus real `Date` objects in the 4 audit columns |
| Tab Visibility     | 282                 | none (**not dense**)                                                 | skip row 0                                                                                 | 2-wide (`Available`, `Visible`)                      | `'TRUE'`/`'FALSE'` strings                              |
| System Permissions | 419                 | none (**not dense**)                                                 | skip rows 0 and 1                                                                          | none                                                 | strings                                                 |

Facts that matter for a replacement:

- This is the only writer that uses `!cols`, `!merges`, `cellDates`/`dateNF`, or `bookSST: true`.
- It bypasses `prepareExcelFile`, so nothing truncates cells to 32,767 characters; SheetJS will throw on an oversized cell and the caller catches and logs.
- Missing `dense: true` is why `RangeError: Too many properties to enumerate` occurs on large orgs (V8 caps own properties near 8.4 million).
- Worst case is `#objects x #fields` rows by `2 x (#profiles + #permission sets)` columns on Field Permissions; real orgs routinely exceed 50k rows and hundreds of columns.
- The `Date` cell contract is pinned by `__tests__/permission-manager-field-export.spec.ts:84-97` (`cell.t === 'd'`, `cell.z === 'yyyy-mm-dd hh:mm:ss'`) and the merge offsets at `:110-115`.
- CSV export path (`generateCsvFilesFromTable`, 109–135) writes four CSVs into a JSZip (`compression: 'STORE'`) and downloads a `.zip`.
- Caller: `libs/features/manage-permissions/src/ManagePermissionsEditor.tsx:492-570` (`saveFile` at 535, Google upload via `UploadToGoogleJob`, zip download at 558).

### A3. Every `prepareExcelFile` call site

**`libs/ui/src/lib/file-download-modal/RecordDownloadModal.tsx`** (lib `ui`)

- 482–512: the format branch. `gdrive` or `RADIO_ALL_SERVER` hands off to `onDownloadFromServer` (job runner). Otherwise `xlsx` builds `data['records'] = flattenRecords(...)` or `getMapOfBaseAndSubqueryRecords(...)` when subqueries are included, then `prepareExcelFile(data, undefined, undefined, { onCellsTruncated: notifyExcelCellsTruncated })` at 491. `RADIO_FORMAT_XLSX_LOAD_TEMPLATE` builds `prepareLoadMultiObjectTemplate(...)` then `prepareExcelFile` at 505. `csv` at 511, `json` inline. `saveFile` at 524.
- Thresholds: `ALLOW_BULK_API_COUNT = 5_000` (126) offers the Bulk API radio; `REQUIRE_BULK_API_COUNT = 500_000` (132) force-sets CSV plus Bulk API (281–288), with the comment that a standard download "fails with RangeError: Invalid string length once the generated file crosses the ~512MB max string length. XLSX reaches that ceiling first". `FILE_FORMAT_ALLOWED_BULK_API = new Set(['csv', 'gdrive'])` (125) downgrades xlsx to csv when Bulk API is chosen (291–294). `MAX_RECORDS_PER_GROUP = 500` (`libs/shared/constants/src/lib/shared-constants.ts:1162`) is a warning only (741–746).
- Error handler 539–553: logs and tracks, then maps a message containing `32767` to "One or more values exceed Excel's 32,767 character cell limit. Download as CSV or JSON to get full values."; any other xlsx failure to "There was a problem preparing your file download. Try downloading as CSV or JSON."
- Callers of the modal: `libs/features/query/src/QueryResults/QueryResultsDownloadButton.tsx:139`, `libs/ui/src/lib/data-table/grid/renderers/SubqueryRenderer.tsx:529`, `libs/shared/ui-core/src/record/ViewEditCloneRecord.tsx:579`.

**`libs/shared/ui-core/src/jobs/JobWorker.ts`** (lib `ui-core`)

- 287–334: `prepareExcelFile(getMapOfBaseAndSubqueryRecords(...))` (289), `prepareExcelFile(flattenRecords(...), fields)` (291), load-template (300–315), `prepareCsvFile` (316); the same pair again for `gdrive` (327/329) with `mimeType = MIME_TYPES.GSHEET`. **No `onCellsTruncated` is passed**, so background downloads truncate silently.
- `WorkerAdapter` (95–108) is documented as "mimics a web-worker" and calls `handleMessage` synchronously; `Jobs.tsx:55` instantiates it once. There is no `new Worker` for jobs anywhere. `JobWorker.ts:271-281` posts "Preparing file for download…" before building because "Building the file blocks the worker and can take a while".
- Bulk API branch (249–269): outside the extension and canvas the result is a server link `${serverUrl}/static/bulk/${jobId}/${batchResult.id}/file?...` with `mimeType: MIME_TYPES.CSV`, streamed by the browser through `DownloadFileStream.tsx`. Inside the extension or canvas (252) records are fetched into memory and built locally.
- `Jobs.tsx:64-68` documents the known main-thread limit ("throws `RangeError: Invalid string length` past ~512MB").

**`libs/ui/src/lib/file-download-modal/FileDownloadModal.tsx`** (lib `ui`)

- 197–200 pass-through when `data instanceof ArrayBuffer`; 203 single sheet `prepareExcelFile(_data, headerFields, undefined, { onCellsTruncated })`; 205–207 multi-sheet map; 216 CSV; 242 `saveFile`; 262–276 `handleUploadToGoogle` repeats the build for an `UploadToGoogleJob`. Errors (248–251) only call `onError`; no CSV fallback or 32,767 message here.
- 23 direct render sites: `WhereIsThisFieldUsed.tsx:83`, `SalesforceRecordDataTable.tsx:922`, `GridDownloadButton.tsx:61`, `MassUpdateRecordsDeploymentRow.tsx:168`, `MassUpdateRecordsObjectRow.tsx:208`, `AutomationControlEditor.tsx:244`, `DeployMetadataDeployment.tsx:279`, `DeployMetadataHistoryModal.tsx:223,241`, `DeployMetadataToOrg.tsx:106`, `AddToChangeset.tsx:120`, `DeployMetadataPackage.tsx:100`, `DeleteMetadataModal.tsx:80,98`, `DeployComparedMetadataModal.tsx:110`, `ViewOrCompareMetadataModal.tsx:254`, `LoadRecordsBatchApiResults.tsx:458`, `LoadRecordsBulkApiResults.tsx:718`, `PerformLoadCustomMetadata.tsx:211`, `LoadRecordsMultiObjectDownloadModal.tsx:29`, `DataHistoryFormatDownloadModal.tsx:111`, `BulkUpdateProposedChangesPreview.tsx:134`, `CreateFieldsImportExport.tsx:136`, `CreateFieldsDeployModal.tsx:113`, `PlatformEventMonitor.tsx:207`, `FormulaEvaluatorDeployModal.tsx:164`, `SObjectExport.tsx:221`.
- 8 indirect sites through `GridDownloadButton` (all Permission Analysis trees, grids and tabs). `GridDownloadButton` flattens via `buildGridExportData` (`libs/ui/src/lib/data-table/gridExport.ts:32-44`) into `{ header: string[], data: Record<string, string>[] }`, so everything is stringified and objects/arrays become `''`.

**Multi-sheet producers**

| Producer                                                                                                      | Shape                                                                                                                               | Note                                                                                                                         |
| ------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------- |
| `getMapOfBaseAndSubqueryRecords` (`libs/shared/utils/src/lib/utils.ts:914-964`)                               | base sheet `records` plus one sheet per subquery via `getSubquerySheetName`; injects `_ParentId` (`SUBQUERY_PARENT_ID_COLUMN`, 882) | used by RecordDownloadModal and JobWorker                                                                                    |
| `libs/features/sobject-export/src/sobject-export-utils.ts:120-212`                                            | `ERRORS`, `Object Metadata`, `Field Metadata`, or one sheet per sObject at 204                                                      | `output[sobject.substring(0, 31)]` with no dedupe: two API names that collide at 31 characters silently overwrite each other |
| `prepareLoadMultiObjectTemplate` (`libs/shared/ui-utils/src/lib/load-multi-object-template.utils.ts:278-328`) | `Record<string, any[][]>` (array-of-arrays, written verbatim): 4 preamble rows, `['Reference Id', ...fields]`, then data            | root sheet `getExcelSafeSheetName(sanitizeExcelSheetName(sobject))`; child sheets `getSubquerySheetName(path, existing)`     |
| `PlatformEventMonitor.tsx:217`, `CreateFieldsDeployModal.tsx:121`, `FormulaEvaluatorDeployModal.tsx:172`      | worksheet map plus header map                                                                                                       |                                                                                                                              |

Sheet-name rules live in `libs/shared/utils/src/lib/utils.ts:972-1021`: `EXCEL_FORBIDDEN_SHEET_NAME_CHARS = /[:\\/?*[\]]/g`, `EXCEL_MAX_SHEET_NAME_LENGTH = 31`, `sanitizeExcelSheetName` strips forbidden characters, `getExcelSafeSheetName` truncates to 31 and appends an incrementing numeric suffix on collision, `getSubquerySheetName` falls back to `${relationshipName} (L{depth})`. Leading/trailing apostrophes and the reserved name `History` are not handled.

### A4. SheetJS feature usage (write)

| Feature                                                                                                                                                                                                            | Used            | Where                                                |
| ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ | --------------- | ---------------------------------------------------- |
| `aoa_to_sheet` with `dense: true`                                                                                                                                                                                  | yes             | `prepareExcelFile` only                              |
| `compression`                                                                                                                                                                                                      | conditional     | auto above 10,000 rows; forced in Permission Manager |
| `bookSST`                                                                                                                                                                                                          | `false` default | `true` only in Permission Manager                    |
| `!cols` widths                                                                                                                                                                                                     | yes             | Permission Manager (4 sheets)                        |
| `!merges`                                                                                                                                                                                                          | yes             | Permission Manager (3 sheets)                        |
| `cellDates` + `dateNF` on write                                                                                                                                                                                    | yes             | Permission Manager Field Permissions                 |
| `bookType: 'xlsx'`, `type: 'array'`                                                                                                                                                                                | yes             | `excelWorkbookToArrayBuffer`                         |
| `set_cptable`                                                                                                                                                                                                      | yes             | lazy, skipped in extension                           |
| freeze panes, autofilter, `!rows`, styles/fonts/colors, `xlsx-js-style`, other number formats, formulas, hyperlinks, workbook props, `json_to_sheet`, `sheet_add_aoa`, `writeFile`, `XLSX.stream`, protection, VBA | **no**          | nowhere                                              |

### A5. Delivery mechanisms

| Mechanism                               | Where                                                                                                                                                                         | Notes                                                                                                                                                                                                                                                                                                                                                             |
| --------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Blob + file-saver                       | `saveFile` (`shared-ui-utils.ts:470-473`); xlsx callers `RecordDownloadModal.tsx:524`, `FileDownloadModal.tsx:242`, `ManagePermissionsEditor.tsx:535`, `Jobs.tsx:433/528/703` | the only local path; whole file in one `Blob`                                                                                                                                                                                                                                                                                                                     |
| Google Drive resumable upload           | `googleUploadFile` (`libs/shared/data/src/lib/client-data.ts:1447-1481`), called from `Jobs.tsx:192-208`                                                                      | `POST .../upload/drive/v3/files?uploadType=resumable` then `PUT`; bytes labelled `X-Upload-Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`; resource `mimeType = application/vnd.google-apps.spreadsheet`, so Drive converts to a native Google Sheet; on failure `handleGoogleUploadFailure` (`Jobs.tsx:684-704`) saves locally |
| Server-streamed link (CSV only)         | `JobWorker.ts:256-268`, `Jobs.tsx:386-393`, `libs/shared/ui-core/src/app/DownloadFileStream.tsx`                                                                              | Bulk API results only; hidden anchor with `download`; hidden on desktop (`Jobs.tsx:398`)                                                                                                                                                                                                                                                                          |
| Electron                                | `apps/jetstream-desktop/src/services/protocol.service.ts` `will-download` (~130–157)                                                                                          | renderer still calls `saveAs`; main process auto-saves to `downloadPreferences.downloadPath` via `toSafeDownloadFileName`; true streaming to disk exists only for zips and Bulk API files (`file-download.service.ts`, `pipeline(Readable.fromWeb(webStream), createWriteStream)`)                                                                                |
| Browser extension, Salesforce canvas    | same `saveAs`                                                                                                                                                                 | no `chrome.downloads`; no `showSaveFilePicker` anywhere in the repo                                                                                                                                                                                                                                                                                               |
| OPFS / File System Access / Electron fs | `libs/shared/ui-data-history/src/lib/file-store/*` (`HistoryFileStore.createWriteStream`, `gzip-utils.ts` `StreamEncoder`)                                                    | Data History only; stores CSV/JSON gzip; a ready-made platform-abstracted streaming sink with a contract suite (`file-store-contract.spec.ts`)                                                                                                                                                                                                                    |

### A6. Constants

`libs/shared/constants/src/lib/shared-constants.ts:126-166`: `MIME_TYPES.XLSX = 'application/octet-stream;charset=utf-8'` (the real type is only in `XLSX_OPEN_OFFICE`), `MIME_TYPES.GSHEET`, `fileExtToMimeType.xlsx = MIME_TYPES.XLSX`, `fileExtToGoogleDriveMimeType.xlsx = GSHEET`, `FILE_FORMAT_XLSX_LOAD_TEMPLATE = 'xlsx-load-template'`, `MAX_BINARY_DOWNLOAD_SIZE_BYTES = 1 GB` (1156), `MAX_RECORDS_PER_GROUP = 500` (1162).

---

## B. Read side

### B1. Call sites

| #   | `XLSX.read` site                                                              | Options                                                | Feeds                                                                                   |
| --- | ----------------------------------------------------------------------------- | ------------------------------------------------------ | --------------------------------------------------------------------------------------- |
| 1   | `libs/shared/ui-utils/src/lib/shared-ui-utils.ts:1299`                        | `{ cellText: false, cellDates: true, type: 'binary' }` | `parseWorkbook`; dead branch (no caller sets `isBinaryString`)                          |
| 2   | `libs/shared/ui-utils/src/lib/shared-ui-utils.ts:1300`                        | `{ cellText: false, cellDates: true, type: 'array' }`  | `parseWorkbook`: Load Records (single object), Create Fields import, Electron open-with |
| 3   | `libs/features/load-records-multi-object/src/LoadRecordsMultiObject.tsx:156`  | `{ cellText: false, cellDates: true, type: 'array' }`  | multi-object `parseWorkbook`                                                            |
| 4   | `libs/ui/src/lib/form/file-selector/GoogleFileSelector.tsx:101`               | `{ cellText: false, cellDates: true, type: 'binary' }` | Google Drive (web app, gapi) → both load features                                       |
| 5   | `libs/ui/src/lib/form/file-selector/GoogleFileSelectorExternalButton.tsx:128` | `{ cellText: false, cellDates: true, type: 'array' }`  | Google Drive (desktop, extension, canvas, REST) → both load features                    |

| #   | `sheet_to_json` site                                                                     | Options                                                                                   |
| --- | ---------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------- |
| A   | `libs/shared/ui-utils/src/lib/shared-ui-utils.ts:1323-1328`                              | `{ dateNF: 'yyyy"-"mm"-"dd"T"hh:mm:ss', defval: '', blankrows: false, rawNumbers: true }` |
| B   | `libs/features/load-records-multi-object/src/load-records-multi-object-utils.ts:343-350` | same four plus `range: 4, header: 1`                                                      |

Not used anywhere: `raw`, `dense`, `sheets`, `sheetRows`, `bookSheets`, `bookProps`, `skipHidden`, `cellNF`, `decode_range`, `encode_cell`, `!ref`. The only direct cell access is `worksheet['B1'|'B2'|'B3'|'A5'].v` in the multi-object loader. All parsing runs on the main thread. No Node or server code parses xlsx; the `xlsx` dependency is re-injected into the API's dist `package.json` by `scripts/replace-package-deps.mjs:20-24` only because of the type leak in B6.

### B2. `parseFile` / `parseWorkbook` contract (`shared-ui-utils.ts:1245-1336`)

`parseFile(content, { onParsedMultipleWorkbooks?, isBinaryString?, isPasteFromClipboard?, extension? })`: `.json` → `parseJson`; string content → papaparse with `header: true`, `skipEmptyLines: true`, delimiter from `detectDelimiter(extension)` (tab for `.tsv`, locale `;` when the decimal separator is `,`), with a re-parse on a likely-wrong delimiter; anything else → `XLSX.read` → `parseWorkbook`.

| Aspect           | Behavior                                                                                                                                                                                                                                        |
| ---------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Sheet selection  | first sheet by default; if more than one sheet and a callback is supplied, the user picks via `XlsxSheetSelectionModalPromise` (`libs/ui/src/lib/modal/`); an unknown or cancelled pick falls back to sheet 1; only one sheet is ever converted |
| Header row       | implicit row 1 of the used range (no `range`, no `header`)                                                                                                                                                                                      |
| Header naming    | SheetJS defaults: blank header → `__EMPTY`, `__EMPTY_1`, …; duplicate header → `Name`, `Name_1`, …; `headers` are `Object.keys(data[0])`, so a column that is blank on the first data row is absent from `headers` while present in `data`      |
| `__EMPTY` filter | `headers.filter((field) => !field.startsWith('__empty'))` uses lowercase and never matches SheetJS's `__EMPTY`: a no-op (the multi-object loader lowercases first and works)                                                                    |
| Empty cells      | `defval: ''` → empty string, never `undefined`                                                                                                                                                                                                  |
| Blank rows       | `blankrows: false`, then `removeEmptyRows` (1338–1361) drops rows whose values are all null, `''` or whitespace and toasts the count                                                                                                            |
| Numbers          | `rawNumbers: true` → JS numbers (a `$1,234.50` cell arrives as `1234.5`)                                                                                                                                                                        |
| Dates            | `cellDates: true` → JS `Date` for date-formatted cells; `cellText: false` suppresses `.w`; `dateNF` is effectively unused for typed values                                                                                                      |
| Booleans         | `b` cells → JS booleans                                                                                                                                                                                                                         |
| Trimming         | none, for headers or cells                                                                                                                                                                                                                      |
| Row limit        | none at parse time                                                                                                                                                                                                                              |
| Errors           | always `errors: []` for xlsx; papaparse errors for CSV                                                                                                                                                                                          |

### B3. Multi-object template (`load-records-multi-object-utils.ts`)

`WORKSHEET_LOCATIONS = { sobject: 'B1', operation: 'B2', externalId: 'B3', referenceId: 'A5', dataStartCell: 'A5', dataStartRow: 4 }` (22–29).

| Excel row | Column A                   | Column B                | Read as                                                                                     |
| --------- | -------------------------- | ----------------------- | ------------------------------------------------------------------------------------------- |
| 1         | `Object Api Name`          | `Account`               | `worksheet['B1'].v.toLowerCase()`                                                           |
| 2         | `Operation`                | `Insert`                | `worksheet['B2'].v.toUpperCase()` (`INSERT`, `UPDATE`, `UPSERT`)                            |
| 3         | `External Id (for upsert)` | `My_Ext_Id__c`          | `worksheet['B3']?.v` (required only for UPSERT)                                             |
| 4         | blank                      |                         | spacer                                                                                      |
| 5         | `Reference Id`             | `Name`, `{ParentId}`, … | `sheet_to_json({ range: 4, header: 1 })[0]`; `A5` also read raw for `referenceColumnHeader` |
| 6+        | data                       |                         | `getExcelRow(idx) = dataStartRow + 2 + idx` for error locations                             |

Behavior (268–431):

- Sheets whose lowercased name includes `instructions` are skipped (warning unless the name is exactly `instructions`); zero remaining sheets is a workbook-level error. All remaining sheets are parsed eagerly; there is no picker.
- Each raw cell read is wrapped in its own try/catch with a targeted message ("Cell B1 must contain the Object API Name …", "Cell B2 must contain the operation …", "Cell B3 must contain the API name of an External Id field …"). These depend on the sparse address-keyed worksheet shape; a dense worksheet makes `worksheet['B1']` undefined and produces a misleading "cell is blank" error.
- `dataHeaders = (data[0] || []).filter(Boolean).map((h) => h.trim())`: a blank middle header is removed and shifts every later index, misaligning the `row.reduce` mapping. A numeric header cell throws at `.trim()` and is caught by the outer handler ("Jetstream could not read the data on this worksheet. Check that headers are on row 5, data starts on row 6 …").
- Duplicate headers and duplicate Reference Ids produce row-located errors; an empty sheet produces "This worksheet has no data rows. Add records starting on row 6, or remove the worksheet."
- Header trimming happens here (unlike the single-object path); `normalizeHeader` strips `{}`; reference columns become `@{refId.id}`.
- Group limit `MAX_RECORDS_PER_GROUP = 500` enforced at 847–858. `validateObjectData` describes each sObject and annotates field-level errors.

### B4. Downstream typing: `transformRecordForDataLoad` (`libs/shared/utils/src/lib/utils.ts:663-773`)

| Incoming value | Field type    | Result                                                                                                                                                                                         |
| -------------- | ------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `Date`         | date          | `formatISO(value, { representation: 'date' })`; invalid Date → `null`                                                                                                                          |
| `Date`         | datetime      | `formatISO(value, { representation: 'complete' })` (local offset); invalid Date throws `DATE_ERR_MESSAGE`                                                                                      |
| string         | date/datetime | parsed with the user-selected format (`MM/dd/yyyy`, `dd/MM/yyyy`, `yyyy-MM-dd`), so whether the reader yields a `Date` or a string changes the result for ambiguous dates such as `03/04/2024` |
| boolean        | boolean       | returned untouched                                                                                                                                                                             |
| string/number  | boolean       | `REGEX.BOOLEAN_STR_TRUE` (`t…` or `1…`) → boolean                                                                                                                                              |
| number         | text/number   | returned as a number, stringified later by the CSV/SOAP builders (`rawNumbers` is load-bearing)                                                                                                |
| `''` (defval)  | any           | `null`, then the insert-nulls setting decides skip vs null                                                                                                                                     |

Preview and mapping also rely on types: `CellRenderers.tsx:38` renders `value instanceof Date` with the date formatter and booleans as checkboxes; `LoadRecordsFieldMappingRow.tsx:11-19` shows `Date.toJSON()`; `autoMapFields` (`load-records-utils.tsx:165-256`) matches headers by API name, lowercase, and alphanumeric-stripped variants with no trimming. Preview caps: `MAX_RECORD_FOR_PREVIEW = 100_000`, `MAX_COLUMNS_TO_KEEP_SET_FILTER = 2000` (`LoadRecordsDataPreview.tsx:29-30`).

### B5. Input paths

| Path                                    | Accept                                                                                                                                                      | Bytes                                                                                    | Notes                                                                                                                                                                                                                                            |
| --------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| Load Records (single)                   | `.csv .tsv .xlsx .json` (`SelectObjectAndFile.tsx:41-47`)                                                                                                   | `FileReader` (`readFile`, text for `.csv/.tsv/.xml/.json`, else ArrayBuffer)             | `handleFile` 123–158; toast for removed rows; **`password-protected` substring** in the error maps to "Your file is password protected, remove the password and try again." (146–150); everything else "There was an error reading your file. …" |
| Load Records (multi-object)             | `.xlsx` only (`LoadRecordsMultiObject.tsx:217`)                                                                                                             | ArrayBuffer                                                                              | no password special case; clipboard paste is a no-op here                                                                                                                                                                                        |
| Create Fields import                    | `.csv .xlsx` (`CreateFieldsImportExport.tsx:178`)                                                                                                           | text or ArrayBuffer                                                                      | `extension` not passed to `parseFile`, so TSV branch unreachable; headers normalized to lowercase alphanumerics                                                                                                                                  |
| Google Drive (web, gapi)                | picker `ViewId.SPREADSHEETS`, no MIME filter                                                                                                                | native Sheets exported as xlsx; anything else raw via `alt=media` as a **binary string** | a Drive-hosted CSV is parsed by SheetJS's CSV path, not papaparse                                                                                                                                                                                |
| Google Drive (desktop/extension/canvas) | external picker popup, same views                                                                                                                           | `fetch(...).arrayBuffer()`                                                               | `GoogleFileSelectorExternalButton.tsx:120-132`                                                                                                                                                                                                   |
| Electron open-with                      | `.csv .xlsx` (`protocol.service.ts:159-205`, `CFBundleDocumentTypes`)                                                                                       | `readFileSync` slice as ArrayBuffer over IPC                                             | a `.csv` arrives as an ArrayBuffer and takes the XLSX branch (SheetJS sniffs CSV from bytes)                                                                                                                                                     |
| Validation                              | `readFileForUpload` (`file-selector-utils.ts`): extension-only, case-insensitive suffix; no MIME or magic-byte check                                        |                                                                                          | `.xls`, `.xlsb`, `.ods` renamed to `.xlsx` reach SheetJS and parse today                                                                                                                                                                         |
| Size limits                             | none for spreadsheets; the attachments zip is 50 MB free / 1000 MB with Drive / unlimited on desktop, extension, canvas (`SelectObjectAndFile.tsx:120-121`) |                                                                                          | no cap on local, Drive or open-with spreadsheet bytes                                                                                                                                                                                            |

### B6. Type leak

`libs/types/src/lib/ui/types.ts:4` imports `type * as XLSX from 'xlsx'`; `:732` declares `InputReadGoogleSheet.workbook: XLSX.WorkBook`. Consumers: `GoogleFileSelector.tsx:6,33`, `GoogleFileSelectorExternalButton.tsx:6,64`, `SelectObjectAndFile.tsx:20,160`, `LoadRecordsMultiObject.tsx:15,170`, `useProcessLoadFile.ts:6,57`, `load-records-multi-object-utils.ts:266`, `shared-ui-utils.ts:1306` (read) and `:420` (write). `XLSX.WorkSheet` is never named. The pickers parse eagerly inside UI components and hand a materialized workbook to the feature.

### B7. Tests that pin read/write behavior

| File                                                                                            | Pins                                                                                                                                              |
| ----------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------- |
| `libs/shared/ui-utils/src/lib/__tests__/shared-ui-utils.spec.ts`                                | truncation to 32,767 + suffix, typed passthrough, multi-sheet counting, no mutation; reads back with `XLSX.read` + `sheet_to_json({ header: 1 })` |
| `libs/shared/ui-utils/src/lib/__tests__/prepare-load-multi-object-template.spec.ts`             | exact template AOA layout; round-trip through `prepareExcelFile` and `XLSX.read`                                                                  |
| `libs/features/load-records-multi-object/src/__tests__/load-records-multi-object-utils.spec.ts` | builds workbooks with `aoa_to_sheet`; instructions skipping, B1/B2/B3 errors, duplicate refs, `{}` references                                     |
| `libs/features/manage-permissions/src/utils/__tests__/permission-manager-field-export.spec.ts`  | `cell.t === 'd'`, `cell.z`, merge offsets                                                                                                         |
| `libs/shared/ui-utils/src/lib/__tests__/remove-empty-rows.spec.ts`, `parse-file-json.spec.ts`   | empty-row removal, JSON/CSV dispatch                                                                                                              |
| `libs/ui/src/lib/file-download-modal/__tests__/RecordDownloadModal.spec.tsx`                    | Bulk API / RangeError regression (mocks `saveFile`)                                                                                               |

No Playwright test parses a downloaded xlsx; the one upload test uses `records-Product2.csv` and asserts only on the suggested filename.

---

## C. Platform and infrastructure

### C1. Dependencies

- `xlsx`: `https://cdn.sheetjs.com/xlsx-0.20.3/xlsx-0.20.3.tgz` (`package.json:359`; npm stops at 0.18.5, which carries CVE-2023-30533 and CVE-2024-22363). `scripts/replace-package-deps.mjs:20-24` carries the tarball spec into the API dist package.json; `scripts/build-electron.mjs:59` excludes `xlsx` from the Electron package (bundled into the renderer); `electron-builder.config.js:372` registers `.xlsx` for open-with.
- Present: `papaparse ^5.7.0`, `file-saver ^2.0.5`, `jszip ^3.10.1` (13 call sites, DEFLATE level 1 for metadata packages), `@jetstreamapp/simple-xml ^1.1.1` (in-house XML; the fast-xml-parser swap is guarded by `libs/salesforce-api/src/lib/__tests__/xml-parsing-snapshot.spec.ts`), `xml2js` (SAML only), `dexie`.
- Absent: `exceljs`, `fflate`, `pako`, `fast-xml-parser`, `sax`, `write-excel-file`, `excelize-wasm`, `xlsx-js-style`.
- Compression primitives are platform APIs: `CompressionStream('gzip')` in `gzip-utils.ts` and `libs/shared/utils/src/lib/compression.ts`; Node `zlib` in the desktop data-history service.

### C2. Workers

The only real Web Worker is `libs/shared/ui-data-history/src/lib/file-store/history-storage.worker.ts`, instantiated at `opfs-file-store.ts:134` with the literal `new Worker(new URL('./history-storage.worker.ts', import.meta.url), { type: 'module', name: 'jetstream-data-history-storage' })` (the bundler pattern-matches this call; no `worker: {}` block exists in any vite config). Protocol: promise-map RPC (`worker-messages.ts`), selective transferables, `Blob` returned by reference, client-allocated stream ids, respawn on crash. Desktop uses the same op protocol over Electron IPC instead of a worker. Jobs use the fake `WorkerAdapter` (A3), so every spreadsheet build and parse runs on the main thread.

### C3. CSP per platform

| Platform  | Definition                                                                               | `script-src`                                                                                                         | Workers                                 | wasm                  |
| --------- | ---------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------- | --------------------------------------- | --------------------- |
| Web       | `apps/api/src/app/utils/security-headers.ts:100-135, 205-212`                            | `'self'` + nonce (+ `'strict-dynamic'` on `/app`), never `blob:` (test-enforced by `security-headers.spec.ts:77-90`) | `worker-src`/`child-src` `'self' blob:` | no `wasm-unsafe-eval` |
| Desktop   | `apps/jetstream-desktop/src/utils/utils.ts:11-34`                                        | `'self' 'unsafe-inline' https://*.google.com`                                                                        | `worker-src 'self' blob:`               | none                  |
| Extension | `apps/jetstream-web-extension/src/manifest.json` (MV3, no `content_security_policy` key) | MV3 default `script-src 'self'`                                                                                      | no `blob:` workers                      | none                  |

A pure-JS engine needs no CSP change on any surface; a wasm engine needs three.

### C4. Versions and targets

Node `>=24 <25` (`.nvmrc` 24), pnpm 11.20.0 (`devEngines`), TypeScript 6.0.3, Vite ^8.2.2 (rolldown), Vitest 5.0.0, Electron ^43, Nx 23.2.0, React ^19.2. `tsconfig.base.json`: target `es2022`, lib `["es2023", "dom"]`, `moduleResolution: "bundler"`. No browserslist anywhere; the only floors are the extension manifest (`minimum_chrome_version: 121`, Firefox `strict_min_version: 120.0`). `CompressionStream`/`DecompressionStream('deflate-raw')` is therefore available on every target (Chrome 103+, Firefox 113+, Safari 16.4+, Node 21.2+).

### C5. Fixture coverage

Four tracked binaries: `apps/jetstream-e2e/src/assets/records-Product2.csv` (1 KB), `records-Product2.xlsx` (17 KB, SheetJS-generated, inline strings only), `apps/api/src/assets/content/Jetstream - Load Records to Multiple Objects - Template.xlsx` (35 KB, a Google Sheets export with data-descriptor zip entries, `[Content_Types].xml` last, `xl/metadata` and `xl/commentsmeta*` parts, legacy VML comments, three sheets including `Instructions`), and a metadata deploy zip. All unit tests build workbooks in memory with `XLSX.utils.aoa_to_sheet`. jsdom's `Blob` does not interoperate with Node's `CompressionStream` (documented in at least six spec files; `libs/shared/ui-data-history/src/test-setup.ts` installs a Node `Blob` shim).

---

## D. Requirements a replacement engine must satisfy

Write

- W1. Accept array-of-objects and array-of-arrays per sheet, multiple sheets, raw JS values (string, number, boolean, `Date`, null/undefined), and produce a valid `.xlsx` (A1).
- W2. Truncate strings over 32,767 characters with the `...(truncated)` suffix and report the count through a callback (A1, A3).
- W3. Column widths, rectangular header merges, a `Date` cell number format (`yyyy-mm-dd hh:mm:ss`), and exact `TRUE`/`FALSE` strings for the Permission Manager workbook (A2).
- W4. Sanitize and de-duplicate sheet names to Excel's rules (31 characters, forbidden characters, case-insensitive uniqueness), including the collision that `sobject-export-utils.ts:204` currently ignores (A3).
- W5. Produce output as bytes that `saveFile` can wrap in a `Blob`, and as a stream for a sink such as `HistoryFileStore.createWriteStream` or Electron IPC (A5).
- W6. Bounded memory for 1M rows x 20 columns and for the 18M-cell case, with progress and cancellation, so the 500,000-row CSV-only rule can be lifted (A3).
- W7. Run unchanged in a module Web Worker, on the Electron renderer, in the MV3 extension (no `blob:` workers, no eval, no wasm) and in the canvas iframe (C2, C3).

Read

- R1. Sniff input bytes: zip → xlsx; CFB → a clear error whose message contains `password-protected` for encrypted files and a "legacy .xls, save as .xlsx" message otherwise; text → hand back to the CSV path (B5).
- R2. List sheet names and visibility without parsing sheet data, for the picker modal and `instructions` skipping (B2, B3).
- R3. Object mode that matches `sheet_to_json` semantics: header row 1, `__EMPTY`/`__EMPTY_1` and `Name_1` naming, `defval: ''`, `blankrows: false`, typed numbers and booleans, `Date` objects for date-formatted cells with local wall-clock semantics, no trimming (B2, B4).
- R4. Array mode with a start-row offset (`range: 4, header: 1`) and cheap raw access to `B1`, `B2`, `B3`, `A5` (B3).
- R5. Decode `_xHHHH_` escapes, rich-text runs, inline strings, shared strings, cached formula values, error cells, missing `<dimension>` and missing `r` attributes, prefixed and Strict namespaces (fixtures in `fixtures/`).
- R6. Enforce decompressed-size and entry caps (zip bombs), reject DOCTYPE, and cap rows for the 100,000-row preview without inflating the whole sheet (B4, B5).
- R7. Accept `ArrayBuffer`, `Uint8Array`, `Blob`/`File`, and a binary string (until the gapi picker is refactored) (B5).
- R8. Replace `XLSX.WorkBook` in `libs/types` with an engine-neutral handle or raw bytes (B6).

Program

- P1. No new CSP directives on any platform (C3).
- P2. Tests runnable under vitest in the `node` environment, with Jetstream specs keeping the Node `Blob` shim (C5).
- P3. Zero or reviewed runtime dependencies; no CDN tarballs (C1).
