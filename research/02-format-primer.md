# xlsx format primer — working reference for the streaming reader/writer

_Accurate as of 2026-09-11._

## 1. Purpose and scope

This is the working reference for engineers building or evaluating the streaming xlsx engine. It is not a spec summary; it records the facts that actually bite when you write bytes Excel has to open, or read bytes that Google Sheets, SheetJS, Excel, LibreOffice and Numbers produce. Every claim is grounded in one of three ways: inspection of the two committed fixtures (`fixtures/jetstream/multi-object-template.gsheets.xlsx`, a Google Sheets export, and `fixtures/jetstream/records-product2.sheetjs.xlsx`, a SheetJS export), probe files generated with the libraries installed under `node_modules` (output lands in the gitignored `.generated/primer-probes/`), or a cited spec/vendor page (section 13). Anything I could not ground is marked **(verify)**.

Scope:

- **Read + write:** ECMA-376 / ISO 29500 **Transitional** SpreadsheetML (`.xlsx`), which is what every real producer emits.
- **Read only:** ISO 29500 **Strict** (`http://purl.oclc.org/ooxml/spreadsheetml/main` namespace, `t="d"` ISO date cells). Excel only writes Strict when a user picks "Strict Open XML Spreadsheet" explicitly.
- **Out of scope:** `.xls` (BIFF8 in a CFB container), `.xlsb` (binary records in a zip), `.ods`, `.xlsm` macro preservation, encrypted packages (we detect and refuse, section 11).

## 2. Package anatomy

### 2.1 OPC in one paragraph

An xlsx is a ZIP archive that follows the Open Packaging Conventions (ECMA-376 Part 2). Each zip entry is a **part** whose **part name** is the entry name with a leading `/` (`xl/workbook.xml` is part `/xl/workbook.xml`). Part-name comparison is ASCII case-insensitive, so two entries differing only by case are a duplicate-part error. Parts are wired together by **relationships**: `_rels/.rels` holds the package-level relationships, and any part `dir/name.ext` may have relationships in `dir/_rels/name.ext.rels`. Every part's content type comes from `[Content_Types].xml`, which is not itself a part and has no relationships.

### 2.2 The minimal valid workbook

The smallest package I would ship has seven entries. The XML below is verbatim from the `@office-kit/xlsx` buffered-write probe (`officekit.xlsx`, 2,963 bytes), which is XSD-validated by that library's CI and is a good baseline. Whitespace was added for readability only.

`[Content_Types].xml`

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>
```

`_rels/.rels`

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
```

`xl/workbook.xml`

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Probe" sheetId="1" r:id="rId1"/></sheets>
</workbook>
```

`xl/_rels/workbook.xml.rels`

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
```

`xl/worksheets/sheet1.xml` (one row shown)

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1:B20"/>
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetData>
    <row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>0.30000000000000004</v></c></row>
  </sheetData>
</worksheet>
```

`xl/styles.xml` (the minimum Excel is happy with; see section 7 for the fuller skeleton)

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><name val="Calibri"/><family val="2"/><color theme="1"/><sz val="11"/><scheme val="minor"/></font></fonts>
  <fills count="2"><fill><patternFill/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
  <dxfs count="0"/>
</styleSheet>
```

`xl/sharedStrings.xml`

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="29" uniqueCount="29">
  <si><t>plain</t></si>
</sst>
```

Notes on what is and is not required:

- `docProps/core.xml` and `docProps/app.xml` are optional. Excel, SheetJS and exceljs write them; `@office-kit/xlsx` and `write-excel-file` do not. The Google Sheets fixture writes only `core.xml` (a single `dcterms:created`).
- `xl/theme/theme1.xml` is optional **unless** you reference theme colours (`<color theme="1"/>`). SheetJS ships a 7.6 KB theme even in its tiny fixture; Google Sheets ships a 3.7 KB "Sheets" theme; `@office-kit/xlsx` and `write-excel-file` ship none and still use `<color theme="1"/>` in the default font. Excel renders that as black, apparently by falling back to its built-in Office theme **(verify on current Excel; it works in the probes as far as tooling round-trips are concerned)**. Emitting `rgb` colours only sidesteps the question (section 7).
- `xl/sharedStrings.xml` is optional if every string is inline (SheetJS fixture has no SST at all).
- **`xl/styles.xml` is formally optional but effectively mandatory.** A package with no styles part cannot display a date (dates are numbers plus a number format that lives only in `cellXfs`), cannot carry any `s="…"` attribute, and several consumers assume the part exists. Always emit it.

### 2.3 Content types: Default vs Override

`Default` maps a file extension to a content type; `Override` maps one exact part name. The two `Default`s every package needs are `rels` and `xml`. Because `.xml` defaults to the generic `application/xml`, **every SpreadsheetML part must be `Override`n** (workbook, each worksheet, styles, sharedStrings, theme, comments, docProps). Omitting the worksheet override makes Excel treat the part as generic XML and it will not find the sheet.

Producers add noise here freely. The SheetJS fixture declares `Default`s for `bin`, `png`, `jpeg`, `pdf`, `emf`, `data`, and `vml` that nothing uses; harmless. Google Sheets declares `Override ContentType="application/binary"` for its proprietary `/xl/metadata` and `/xl/commentsmeta0..2` parts, which a reader must simply ignore.

### 2.4 Locating parts correctly

Do not hard-code `xl/worksheets/sheet1.xml`. Walk the relationships:

1. Read `_rels/.rels`, find the relationship whose `Type` ends in `/officeDocument`. Its `Target` (`xl/workbook.xml`) is resolved against the package root.
2. Read `xl/_rels/workbook.xml.rels`. Worksheet targets are relative to the **source part's directory** (`worksheets/sheet1.xml` relative to `xl/`).
3. For each `<sheet r:id="…">` in `xl/workbook.xml`, look the `Id` up in that rels part and check its `Type`: `…/relationships/worksheet`, `…/chartsheet`, `…/dialogsheet` and `…/xlMacrosheet` all appear in `<sheets>`.

Resolution rules that matter:

- A `Target` beginning with `/` is an absolute part name; otherwise resolve relative to the source part's folder and normalise `..` segments. The Google Sheets fixture's `xl/worksheets/_rels/sheet1.xml.rels` points at `../comments1.xml` and `../drawings/vmlDrawing1.vml`.
- `TargetMode="External"` targets are URIs, not parts (`Target="https://getjetstream.app/app/feedback"` in the same rels file). Never try to open them as parts, never fetch them.
- Zip entry names have no leading slash; `[Content_Types].xml` `PartName`s do. Normalise before comparing, and compare ASCII case-insensitively.
- Some broken writers emit backslashes (`xl\worksheets\sheet1.xml`). Normalise `\` to `/` on read; never emit them.
- The relationship `Id` is arbitrary text, and `rId` numbers do not encode order. In the exceljs probe the only worksheet is `rId4`; in the SheetJS fixture `_rels/.rels` lists `rId2, rId3, rId1` in that order. Sheet order is the document order of `<sheets>`.
- Unknown relationship types and parts must be tolerated. Google Sheets adds `<Relationship Id="rId7" Type="http://customschemas.google.com/relationships/workbookmetadata" Target="metadata"/>` to the workbook rels, and references it from `<extLst><ext uri="GoogleSheetsCustomDataVersion2"><go:sheetsCustomData … r:id="rId7"/></ext></extLst>` inside `workbook.xml`. Note the `ext uri` is not a GUID; do not validate it as one.

### 2.5 What real producers emit

| Producer                                     | Parts (order in zip)                                                                                                                                                                                                                                                                                                                                                                         | Notable traits                                                                                                                                                                                                                                                                                                                                                                    |
| -------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Google Sheets** (fixture, 30 entries)      | `xl/comments{1..3}.xml` + rels, `xl/drawings/vmlDrawing{1..3}.vml`, `xl/drawings/drawing{1..3}.xml` (empty `xdr:wsDr`), `xl/worksheets/sheet{1..3}.xml` + rels, `docProps/core.xml`, `xl/theme/theme1.xml`, `xl/sharedStrings.xml`, `xl/styles.xml`, `xl/workbook.xml`, `xl/_rels/workbook.xml.rels`, `_rels/.rels`, `xl/metadata`, `xl/commentsmeta{1,0,2}`, **`[Content_Types].xml` last** | Deflate with data descriptors (flag `0x0808`), no `<dimension>`, no `<numFmts>`, fill 1 is `lightGray` not `gray125`, numbers written as `1.0`, ~1,000 empty `<row>`s carrying only `customHeight`, empty styled cells `<c r="C2" s="8"/>`, one `<author></author>` and a "======\nID#…" comment convention, hyperlinks via rels, `<dataValidations>` with a literal list formula |
| **SheetJS CE** (fixture, 9 entries)          | `docProps/core.xml`, `docProps/app.xml`, `xl/worksheets/sheet1.xml`, `xl/workbook.xml`, `xl/theme/theme1.xml`, `xl/styles.xml`, `[Content_Types].xml`, `_rels/.rels`, `xl/_rels/workbook.xml.rels`                                                                                                                                                                                           | **Stored (method 0)**, version-needed 10, no SST, strings as `t="str"` (not `inlineStr`), booleans `t="b"`, `<ignoredErrors numberStoredAsText="1" sqref="A1:D18"/>`, always emits `<numFmt numFmtId="56" formatCode="&quot;上午/下午 &quot;hh&quot;時&quot;…"/>` (a zh-TW built-in override, harmless), `<workbookPr codeName="ThisWorkbook"/>`, empty `core.xml`                |
| **exceljs 4.4** (probe, 16 entries)          | `[Content_Types].xml` first, then **directory entries** (`_rels/`, `xl/`, `xl/_rels/`, …), `_rels/.rels`, workbook rels, sheet, sharedStrings, theme, styles, docProps, `xl/workbook.xml` last                                                                                                                                                                                               | No data descriptors, version-needed 10 even for deflate, `<fileVersion appName="xl" …/>`, `<calcPr calcId="171027"/>`, `x14ac:dyDescent` on rows with `mc:Ignorable="x14ac"`, `spans="1:2"`, `pageSetup horizontalDpi="4294967295"`, styles `extLst` with slicer/timeline defaults, **does not escape `_x0041_`**, strips CR and C0 controls                                      |
| **@office-kit/xlsx 0.11** (probe, 7 entries) | buffered: workbook, workbook rels, sheet, styles, sst, `_rels/.rels`, `[Content_Types].xml` last; streaming: sheet first                                                                                                                                                                                                                                                                     | Deflate + data descriptors (flag `0x0008`), version-needed 20, streaming variant omits `<dimension>`, escapes everything including tab/LF/CR and **surrogate pairs** (`_xD83D__xDE00_`), no `cellStyles`/`tableStyles`, dates written as bare serials with no number format                                                                                                       |
| **write-excel-file 2.3** (probe, 14 entries) | rels, `[Content_Types].xml`, directory entries incl. unused `xl/media/` and `xl/drawings/`, rels, sst, styles, workbook, sheet                                                                                                                                                                                                                                                               | Custom number format at **`numFmtId="100"`** (inside the reserved range), `<xf ></xf>` with no attributes, no `cellStyleXfs`/`cellStyles`, `<sst>` without counts, no `<dimension>`, empty-string cells omitted, `_x0041_` unescaped                                                                                                                                              |
| **Excel 365** (typical; not in repo)         | `[Content_Types].xml`, `_rels/.rels`, `xl/workbook.xml`, `xl/_rels/workbook.xml.rels`, `xl/worksheets/sheetN.xml`, `xl/theme/theme1.xml`, `xl/styles.xml`, `xl/sharedStrings.xml`, `docProps/core.xml`, `docProps/app.xml`; plus `xl/calcChain.xml` when formulas exist, `xl/printerSettings/`, `xl/worksheets/_rels/` for tables/comments                                                   | Deflate, no data descriptors, SST for all strings, `<dimension>`, `<sheetViews>`, `x14ac`/`xr` namespaces with `mc:Ignorable`, `<definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">` for autofilters. **(verify against a fresh Excel save before relying on any detail here)**                                                                                |
| **LibreOffice / Numbers**                    | Both emit SST strings, theme, docProps; LibreOffice adds `docProps/custom.xml` when custom properties exist.                                                                                                                                                                                                                                                                                 | **(verify; no fixture in repo)**                                                                                                                                                                                                                                                                                                                                                  |

## 3. Worksheet XML

### 3.1 CT_Worksheet child order

The schema is an `xsd:sequence`; Excel rejects out-of-order children with the repair dialog. Full order (ISO 29500-1 §18.3.1.99, via the OpenXML SDK `Worksheet` class), with the ones the engine touches in bold:

1. `sheetPr`
2. **`dimension`**
3. **`sheetViews`**
4. **`sheetFormatPr`**
5. **`cols`**
6. **`sheetData`**
7. `sheetCalcPr`
8. `sheetProtection`
9. `protectedRanges`
10. `scenarios`
11. **`autoFilter`**
12. `sortState`
13. `dataConsolidate`
14. `customSheetViews`
15. **`mergeCells`**
16. `phoneticPr`
17. `conditionalFormatting`
18. **`dataValidations`**
19. **`hyperlinks`**
20. `printOptions`
21. **`pageMargins`**
22. `pageSetup`
23. `headerFooter`
24. `rowBreaks`
25. `colBreaks`
26. `customProperties`
27. `cellWatches`
28. `ignoredErrors`
29. `smartTags`
30. `drawing`
31. `legacyDrawing`
32. `legacyDrawingHF`
33. `drawingHF`
34. `picture`
35. `oleObjects`
36. `controls`
37. `webPublishItems`
38. `tableParts`
39. `extLst`

The classic mistake is writing `autoFilter` after `mergeCells` (it goes before), or `cols` after `sheetData`. Both fixtures obey the order; the Google Sheets tail is `…</sheetData><mergeCells count="11">…</mergeCells><hyperlinks>…</hyperlinks><printOptions/><pageMargins …/><pageSetup orientation="landscape"/><drawing r:id="rId3"/><legacyDrawing r:id="rId4"/></worksheet>`.

### 3.2 Prologue elements

`<dimension ref="A1:D18"/>` (SheetJS fixture) is advisory. Google Sheets omits it; `@office-kit/xlsx` streaming omits it; some writers lie. Readers must derive extent from the cells they actually see. A streaming writer cannot know the extent until the end, so omit it or accept a second pass.

`<sheetViews>` needs at least `<sheetView workbookViewId="0"/>` if present at all (`workbookViewId` is required). A top-row freeze, verbatim from the exceljs probe:

```xml
<sheetViews>
  <sheetView workbookViewId="0">
    <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
    <selection pane="bottomLeft"/>
  </sheetView>
</sheetViews>
```

`xSplit`/`ySplit` are counts of frozen columns/rows when `state="frozen"`; `topLeftCell` is the first scrollable cell. `activePane` is `bottomRight` when both are frozen.

`<sheetFormatPr defaultRowHeight="15"/>`: `defaultRowHeight` is the one required attribute. Google Sheets adds `customHeight="1" defaultColWidth="12.63"`.

`<cols>` groups column ranges; `min`/`max` are 1-based and required, `width` is in character units (section 7.6), and Excel only honours `width` when `customWidth="1"`:

```xml
<cols>
  <col customWidth="1" min="1" max="1" width="2.88"/>
  <col customWidth="1" min="15" max="26" width="7.63"/>
</cols>
```

Ranges must be ascending and non-overlapping.

### 3.3 `sheetData`, `row`, `c`

`<row>` attributes: `r` (1-based index), `spans` (`"1:2"`, an optimisation hint only; Excel writes it, nobody needs it), `s` + `customFormat="1"` (row default style), `ht` (points) + `customHeight="1"`, `hidden`, `outlineLevel`, `collapsed`, `thickTop`, `thickBot`, `ph`. Google Sheets emits `<row r="21" ht="15.75" customHeight="1"/>` a thousand times with no cells; a reader must not treat cell-less rows as data.

`<c>` attributes: `r` (`"B12"`), `s` (0-based index into `cellXfs`, default 0), `t` (cell type, default `n`), plus `cm`/`vm` (metadata indices, ignore) and `ph`. Children in order: `<f>` (formula), `<v>` (value), `<is>` (inline string, only when `t="inlineStr"`). A `<c>` with no children is a blank cell that carries a style: `<c r="C2" s="8"/>`.

All `t` values, with what `<v>` holds:

| `t`                 | `<v>`                                                             | Example                                                |
| ------------------- | ----------------------------------------------------------------- | ------------------------------------------------------ |
| `n` (default, omit) | xsd:double                                                        | `<c r="B9"><v>0.30000000000000004</v></c>`             |
| `s`                 | 0-based SST index                                                 | `<c r="B1" s="2" t="s"><v>1</v></c>`                   |
| `str`               | literal string ("formula string"; Excel accepts it with no `<f>`) | `<c r="A1" t="str"><v>Id</v></c>` (SheetJS fixture)    |
| `inlineStr`         | none; text lives in `<is>`                                        | `<c r="A1" t="inlineStr"><is><t>Id</t></is></c>`       |
| `b`                 | `0` or `1`                                                        | `<c r="C2" t="b"><v>1</v></c>`                         |
| `e`                 | an error literal                                                  | `<c r="D2" t="e"><v>#DIV/0!</v></c>`                   |
| `d`                 | ISO 8601 date-time (ISO 29500; Excel writes it only in Strict)    | `<c r="E2" t="d" s="3"><v>2024-02-29T12:00:00</v></c>` |

A common reader bug is treating `t="str"` as an SST index; it is the literal text. Another is failing on `t="d"`; Excel 2010+ reads it in Transitional files too **(verify)**.

Formulas: `<c r="A3"><f>SUM(A1:A2)</f><v>3</v></c>`. `<f>` must not start with `=`. `<v>` is the cached result; omit it and set `<calcPr fullCalcOnLoad="1"/>` in the workbook so Excel recalculates. Shared formulas use `<f t="shared" ref="B2:B10" si="0">` on the master and `<f t="shared" si="0"/>` on followers.

**Missing `r` inference.** Both `row/@r` and `c/@r` are optional in the schema. If a row has no `r`, it is the previous row plus one (first row is 1). If a cell has no `r`, it is the next column after the previous cell in that row (first cell is column A). Excel and SheetJS write `r` on everything; readers must implement inference because small writers do omit it. Excel accepts `r`-less cells **(verify)**.

Rows must be in ascending `r` order and cells in ascending column order within a row; duplicates or reversals trigger the repair dialog.

### 3.4 Epilogue elements

```xml
<mergeCells count="11"><mergeCell ref="B9:B24"/><mergeCell ref="B1:H1"/>…</mergeCells>
```

Only the top-left cell of a merge carries a value; the others may be present as styled blanks. A 1×1 merge (`A1:A1`) and overlapping merges both trigger repair.

```xml
<autoFilter ref="A1:B1"/>
```

Excel also writes a hidden workbook-level defined name for each autofilter: `<definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">Probe!$A$1:$B$1</definedName>`. The filter arrows appear without it (exceljs probe has only the element), but emit both to match Excel.

```xml
<hyperlinks><hyperlink r:id="rId2" ref="B56"/></hyperlinks>
```

External links go through a rel with `TargetMode="External"`; in-workbook links use `location="Sheet2!A1"` and no `r:id`. Optional `display` and `tooltip` attributes.

```xml
<dataValidations>
  <dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1"
                  prompt="Invalid Operation - Only the predetermined list of operations are allowed." sqref="B2">
    <formula1>"Insert,Update,Upsert"</formula1>
  </dataValidation>
</dataValidations>
```

(Google Sheets fixture, sheet 2; the `"` inside `formula1` is stored as `&quot;`.) The engine only needs to preserve or drop these; it never evaluates them.

```xml
<pageMargins bottom="0.75" footer="0.0" header="0.0" left="0.7" right="0.7" top="0.75"/>
```

All six attributes are required by schema (inches). Excel's defaults are `left/right 0.7`, `top/bottom 0.75`, `header/footer 0.3`.

## 4. Strings

### 4.1 Shared string table

```xml
<sst xmlns="…/spreadsheetml/2006/main" count="215" uniqueCount="82">
  <si><t>Object Api Name</t></si>
  <si><t xml:space="preserve">  padded  </t></si>
  <si>
    <r><rPr><rFont val="Arial"/><b/><color theme="1"/></rPr><t xml:space="preserve">1. </t></r>
    <r><rPr><rFont val="Arial"/><color theme="1"/></rPr><t xml:space="preserve">Create one worksheet …
</t></r>
  </si>
</sst>
```

`count` is total references, `uniqueCount` is the number of `<si>`. Excel tolerates wrong counts and `write-excel-file` omits them entirely, but emit correct ones. An `<si>` is either one `<t>` or a sequence of rich-text runs `<r>` (each with optional `<rPr>` and one `<t>`), optionally followed by `<rPh>` phonetic runs and `<phoneticPr>` (Japanese furigana). A plain-text reader concatenates every `<r>/<t>` and ignores `<rPh>` entirely; including `<rPh>` text is the classic "furigana leaks into values" bug. The rich-text `<rPr>` uses `<rFont>` where `<font>` uses `<name>`.

Index `<v>` values must be `< uniqueCount`; the SST can appear anywhere in the zip (Google Sheets writes it after the sheets), so a one-pass streaming reader needs random access to the central directory anyway.

### 4.2 Inline strings and `t="str"`

`<c t="inlineStr"><is><t>text</t></is></c>` (`<is>` accepts the same rich-text children as `<si>`). This is the only string form a single-pass streaming writer can emit without buffering the whole SST. Excel reads inline strings and converts them to SST on save. Old consumers that ignored `inlineStr` (early Numbers, early POI) are essentially gone **(verify)**; SheetJS itself picks `t="str"` with `<v>`, which is also universally readable.

Cost trade-off: SST dedupes repeated values (215 refs → 82 entries in the Google fixture) and shrinks the sheet XML; inline strings make each sheet self-contained and let the writer stream. The pragmatic hybrid is to stream sheets with inline strings and never emit an SST, or to keep a bounded in-memory SST and flush it last (zip entry order is free).

### 4.3 Whitespace

XML parsers keep element text intact, but Excel normalises `<t>` content unless `xml:space="preserve"` is present: leading/trailing spaces are dropped and it may collapse runs **(verify exact behaviour)**. Every producer inspected adds `xml:space="preserve"` when the text starts or ends with whitespace or contains a newline; do the same.

### 4.4 ST_Xstring escaping (`_xHHHH_`)

SpreadsheetML strings are `ST_Xstring`: any character that cannot appear in XML 1.0 is written as `_xHHHH_` (four uppercase hex digits, UTF-16 code unit). The MS-OI29500 notes on §22.9.2.19 are the operative rules:

- Carriage return (U+000D) **shall** be escaped as `_x000D_` (this is why data extracted from Excel by naive tools shows literal `_x000D_`).
- Line feed and tab are escaped **only in attributes**; in element content Excel writes them raw and "does not allow" them to be escaped. Write `\n` and `\t` raw inside `<t>`.
- Underscore is escaped as `_x005F_` **only** to defuse a literal that would otherwise look like an escape: `_x0041_` in user data must be written `_x005F_x0041_`, otherwise Excel decodes it to `A`.
- Everything the XML 1.0 `Char` production forbids must be escaped: U+0000–U+0008, U+000B, U+000C, U+000E–U+001F, U+FFFE, U+FFFF, and lone surrogates (`_xD800_`), which cannot be UTF-8 encoded at all.
- Astral characters (emoji) are legal XML; write them as raw UTF-8. Do **not** escape surrogate pairs. `@office-kit/xlsx` 0.11 writes `_xD83D__xDE00_ emoji` for U+1F600; whether Excel reassembles that is unknown **(verify)**, and it is certainly not what Excel writes.
- On read, decode every `_xHHHH_` you see (including `_x0009_`/`_x000A_`, which strict Office would not have written), and decode `_x005F_` last so the un-defusing works.

What the probes did with `_x0041_ literal`, `tab\there`, `cr\rhere`, `ctlx`:

| Writer           | `_x0041_`                                   | tab       | CR           | U+0001       |
| ---------------- | ------------------------------------------- | --------- | ------------ | ------------ |
| exceljs          | **unescaped** (Excel will show `A literal`) | raw       | **stripped** | **stripped** |
| write-excel-file | **unescaped**                               | raw       | **stripped** | **stripped** |
| @office-kit/xlsx | `_x005F_x0041_`                             | `_x0009_` | `_x000D_`    | `_x0001_`    |

Only the last is round-trip safe; the over-escaping of tab is XSD-valid and Excel-readable but not what Excel writes.

### 4.5 Predefined entities, numeric references, limits

The five XML entities (`&amp; &lt; &gt; &quot; &apos;`) are the only named ones; `&nbsp;` is a repair-dialog error. Numeric references (`&#10;`, `&#x1F600;`) are legal for any XML `Char` but cannot smuggle a control character in (`&#1;` is a well-formedness error), so escaping is `_xHHHH_` or nothing. Emit entities for `& < >` in text and additionally `"` in attributes; the SheetJS fixture shows `formatCode="&quot;上午/下午 &quot;hh…"`.

A cell holds at most **32,767** characters. Excel repairs (and truncates) longer SST entries; the writer must truncate or reject before serialisation.

## 5. Numbers

`<v>` for a numeric cell is an `xsd:double` lexical value: optional sign, digits with optional fraction (`.5` is legal), optional exponent `[eE][+-]?digits`, and the specials `INF`, `-INF`, `NaN`. Excel rejects the specials with the repair dialog and has no representation for them; write an error cell (`t="e"`, `#NUM!`) or a blank instead. `Infinity`, `0x1p3`, thousands separators and locale decimal commas are all invalid.

JavaScript's `String(n)` (shortest round-trip, up to 17 significant digits) is a valid lexical form and every probed writer uses it verbatim:

| Input                  | exceljs / office-kit / write-excel-file wrote                                                                                              |
| ---------------------- | ------------------------------------------------------------------------------------------------------------------------------------------ |
| `0.1 + 0.2`            | `0.30000000000000004`                                                                                                                      |
| `1e21`                 | `1e+21` (lowercase e, explicit `+`: legal xsd:double; Excel itself writes `1E+21`)                                                         |
| `-0`                   | `0` (all three normalise; Excel has no negative zero)                                                                                      |
| `2 ** 53 + 1`          | `9007199254740992` (lost in the JS double before any writer saw it)                                                                        |
| `1 / 3`                | `0.3333333333333333`                                                                                                                       |
| `12345678901234567890` | `12345678901234567000` (exceljs, write-excel-file) vs `12345678901234567168` (office-kit, exact binary expansion). Same double, both fine. |

Excel calculates with 15 significant digits of display precision (it stores the full double), so 17-digit output is safe on read and merely noisy. Integers beyond 2^53 lose precision in JS before Excel is involved; if the source is a BigInt or a Salesforce 18-digit id, write it as a **string**. Excel's own range is ±1.7976931348623158e+308, smallest positive 2.2251e-308.

Error cells: `t="e"` with one of `#NULL!`, `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`, `#N/A`, `#GETTING_DATA` (the ISO 29500 set). Excel 365 also writes `#SPILL!`, `#CALC!`, `#FIELD!`, `#CONNECT!`, `#BLOCKED!`, `#UNKNOWN!` **(verify the full list)**; a reader should accept any `#…` literal for `t="e"`.

## 6. Dates and times

### 6.1 Serial numbers

A date is a number plus a date number format. In the default **1900 system**, serial 1 is 1900-01-01 00:00, and the fractional part is the time of day (`0.5` = 12:00). Excel deliberately reproduces the Lotus 1-2-3 bug that treats 1900 as a leap year: serial 60 displays as the nonexistent 1900-02-29, so serials 1–59 are one day off relative to the 1899-12-30 epoch that makes everything from 61 upward correct. Microsoft documents the bug as permanent.

| Serial  | Excel shows                 | Real date             |
| ------- | --------------------------- | --------------------- |
| 0       | 1/0/1900 (time-only values) | —                     |
| 1       | 1900-01-01                  | 1900-01-01            |
| 59      | 1900-02-28                  | 1900-02-28            |
| 60      | 1900-02-29                  | does not exist        |
| 61      | 1900-03-01                  | 1900-03-01            |
| 45351.5 | 2024-02-29 12:00            | 2024-02-29 12:00      |
| 2958465 | 9999-12-31                  | latest supported date |

Two of the three probed writers get this wrong: exceljs and write-excel-file both encode 1900-02-28 as **60** (should be 59); `@office-kit/xlsx` writes 59. All three encode 1900-03-01 as 61 and 2024-02-29 12:00 UTC as 45351.5.

The **1904 system** (`<workbookPr date1904="1"/>`, Mac Excel heritage) has serial 0 = 1904-01-01 and no leap-year bug; `serial1900 = serial1904 + 1462`. Honour the flag on read; never set it on write.

Precision: one millisecond is 1/86,400,000 ≈ 1.16e-8 days; at serial ≈ 45,000 a double still resolves ≈ 1e-11, so millisecond round-trips are exact if you round to the nearest millisecond when converting back (otherwise `.999` seconds appear).

### 6.2 JS conversions

```ts
const MS_PER_DAY = 86_400_000;
const EPOCH_1900_UTC = Date.UTC(1899, 11, 30); // serial 0 for serials >= 61
const EPOCH_1904_UTC = Date.UTC(1904, 0, 1);

/** Serial -> Date whose UTC fields equal the wall-clock fields in the sheet. */
export function serialToUtcDate(serial: number, date1904 = false): Date {
  if (date1904) {
    return new Date(Math.round(EPOCH_1904_UTC + serial * MS_PER_DAY));
  }
  // Serials 1..59 sit before the phantom 1900-02-29; shift them forward one day.
  const adjusted = serial < 60 ? serial + 1 : serial;
  return new Date(Math.round(EPOCH_1900_UTC + adjusted * MS_PER_DAY));
}

/** Date's UTC fields -> serial. Use for dates you built with Date.UTC(...) */
export function utcDateToSerial(date: Date, date1904 = false): number {
  if (date1904) {
    return (date.getTime() - EPOCH_1904_UTC) / MS_PER_DAY;
  }
  const serial = (date.getTime() - EPOCH_1900_UTC) / MS_PER_DAY;
  return serial < 61 ? serial - 1 : serial; // dates before 1900-03-01
}

/** Local-time variants: shift by the zone offset so local wall-clock fields are what lands in the sheet. */
export function localDateToSerial(date: Date, date1904 = false): number {
  const asUtc = new Date(date.getTime() - date.getTimezoneOffset() * 60_000);
  return utcDateToSerial(asUtc, date1904);
}
export function serialToLocalDate(serial: number, date1904 = false): Date {
  const utc = serialToUtcDate(serial, date1904);
  return new Date(utc.getTime() + utc.getTimezoneOffset() * 60_000);
}
```

Serial 60 itself has no real date; the code above maps it to 1900-03-01 (`60` is not `< 60`). Pick a policy and document it.

### 6.3 `t="d"` cells (Strict, read-only)

`<c t="d"><v>2024-02-29T12:00:00</v></c>` carries an ISO 8601 value (optionally with fraction and `Z`/offset). Parse it directly; do not run it through serial math. Only Strict-mode Excel writes it; still read it in Transitional files.

### 6.4 Deciding whether a number is a date

Nothing in `<c>` says "date". The only signal is the cell's style: `cellXfs[s].numFmtId`, resolved to either a built-in id or a custom `<numFmt formatCode>`.

Built-in ids 0–49 (ISO 29500-1 §18.8.30, "All Languages"; ids 5–8, 23–26, 41–44 are locale-defined and have no fixed code):

| id    | formatCode                                                     | date/time?              |
| ----- | -------------------------------------------------------------- | ----------------------- |
| 0     | `General`                                                      |                         |
| 1     | `0`                                                            |                         |
| 2     | `0.00`                                                         |                         |
| 3     | `#,##0`                                                        |                         |
| 4     | `#,##0.00`                                                     |                         |
| 5–8   | currency, locale-defined (`$#,##0_);($#,##0)` family in en-US) |                         |
| 9     | `0%`                                                           |                         |
| 10    | `0.00%`                                                        |                         |
| 11    | `0.00E+00`                                                     |                         |
| 12    | `# ?/?`                                                        |                         |
| 13    | `# ??/??`                                                      |                         |
| 14    | `mm-dd-yy`                                                     | **date**                |
| 15    | `d-mmm-yy`                                                     | **date**                |
| 16    | `d-mmm`                                                        | **date**                |
| 17    | `mmm-yy`                                                       | **date**                |
| 18    | `h:mm AM/PM`                                                   | **time**                |
| 19    | `h:mm:ss AM/PM`                                                | **time**                |
| 20    | `h:mm`                                                         | **time**                |
| 21    | `h:mm:ss`                                                      | **time**                |
| 22    | `m/d/yy h:mm`                                                  | **date-time**           |
| 23–26 | locale-defined                                                 |                         |
| 27–36 | locale-defined (zh/ja/ko date formats, e.g. `[$-404]e/m/d`)    | **date** in CJK locales |
| 37    | `#,##0 ;(#,##0)`                                               |                         |
| 38    | `#,##0 ;[Red](#,##0)`                                          |                         |
| 39    | `#,##0.00;(#,##0.00)`                                          |                         |
| 40    | `#,##0.00;[Red](#,##0.00)`                                     |                         |
| 41–44 | locale-defined accounting formats                              |                         |
| 45    | `mm:ss`                                                        | **time**                |
| 46    | `[h]:mm:ss`                                                    | **elapsed time**        |
| 47    | `mmss.0`                                                       | **time**                |
| 48    | `##0.0E+0`                                                     |                         |
| 49    | `@` (text)                                                     |                         |

Ids 50–58 are further CJK date formats and 59–81 are Thai variants (71–81 dates). Treat 14–22, 27–36, 45–47, 50–58, 71–81 as date/time. Excel writes `numFmtId="14"` for an unformatted date (exceljs probe `s="2"` → `numFmtId="14"`).

Custom codes (`numFmtId >= 164`, or in practice any id with an explicit `<numFmt>`): apply this heuristic to `formatCode`.

1. Take the first `;`-separated section (positive section) — or scan all sections if you want `[Red]`-style variants to count.
2. Remove quoted literals `"…"`, backslash-escaped characters `\x`, padding `_x` and fill `*x` tokens.
3. Remove bracketed tokens `[…]` **except** the elapsed-time tokens `[h]`, `[hh]`, `[m]`, `[mm]`, `[s]`, `[ss]` — those are time formats. `[Red]`, `[$-409]`, `[$€-2]`, `[>100]` are colour, locale, currency and condition tags.
4. If anything left matches `/[ymdhs]|AM\/PM|A\/P/i` (plus `e`/`g` era tokens in CJK codes and `b` Buddhist-year tokens in Thai), it is a date/time format. `General`, `@`, `0.00`, `#,##0` are not.

Formatting (not detection) is where `m` is ambiguous: `m` immediately after `h`/`hh` or immediately before `s`/`ss` means minutes, otherwise month. `AM/PM` anywhere switches `h` to 12-hour. `[h]` accumulates hours beyond 24 (durations).

## 7. `styles.xml`

### 7.1 Skeleton Excel accepts

Assembled from the exceljs and SheetJS fixtures; every collection is in the required order (`numFmts, fonts, fills, borders, cellStyleXfs, cellXfs, cellStyles, dxfs, tableStyles, colors, extLst`):

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="yyyy-mm-dd hh:mm"/>
  </numFmts>
  <fonts count="1">
    <font><sz val="11"/><color rgb="FF000000"/><name val="Calibri"/><family val="2"/></font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="3">
    <xf numFmtId="0"   fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="14"  fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
  <dxfs count="0"/>
  <tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>
```

Rules of thumb:

- **Fills 0 and 1 are reserved**: index 0 must be `none` and index 1 `gray125`; Excel's "No fill" UI maps to these. Google Sheets writes `lightGray` at index 1 and Excel tolerates it, but the two reserved slots are why you never see a 1-element `<fills>`.
- `cellStyleXfs[0]` plus `<cellStyle name="Normal" builtinId="0" xfId="0"/>` define the Normal style. `@office-kit/xlsx` omits `cellStyles` entirely and is XSD-valid; emitting it is safer.
- The `count` attributes are advisory; Excel ignores mismatches **(verify)**, but keep them right.
- `dxfs` (differential formats for conditional formatting/tables) can be `count="0"`; `tableStyles` can be omitted or `count="0"`.
- `<extLst>` in exceljs's output (slicer/timeline defaults) is optional noise.

### 7.2 How `s` resolves

`c/@s` is a 0-based index into `cellXfs`. Each `xf` there points at `numFmtId`, `fontId`, `fillId`, `borderId` (indices into the collections above) and `xfId` (index into `cellStyleXfs`, the named cell style it derives from). Index 0 is the default and is what a cell with no `s` gets. An out-of-range `s`, `fontId`, `fillId`, `borderId`, or a custom `numFmtId` with no `<numFmt>` element is a repair-dialog error.

The `apply*` attributes (`applyNumberFormat`, `applyFont`, `applyFill`, `applyBorder`, `applyAlignment`, `applyProtection`) on a `cellXfs` xf declare that this xf overrides the corresponding property of its parent cell style. Excel appears to honour the referenced ids regardless **(verify)**, but Excel itself sets them, and LibreOffice reads them, so set the flag for every property you assign. Alignment and protection are child elements (`<alignment wrapText="1" horizontal="center"/>`) rather than indices.

### 7.3 Custom number formats

`numFmtId` 0–163 are reserved for built-ins (some locale-specific, section 6.4). Custom formats start at **164** and each needs a `<numFmt numFmtId="…" formatCode="…"/>`. `write-excel-file` uses id 100, inside the reserved range; Excel tolerates it because 100 is undefined in every locale table, but the Thai table defines 59–81 and the CJK tables 27–36 and 50–58, so a writer that picks low ids will eventually collide. Dedupe by `formatCode`; the id is just a handle. `formatCode` is an ST_Xstring inside an attribute: escape `"` as `&quot;`, and escape LF/tab as `_x000A_`/`_x0009_` if they ever appear.

### 7.4 Colours

Three encodings appear on `<color>`, `<fgColor>`, `<bgColor>`:

- `rgb="FFF3F3F3"` — ARGB hex, alpha first. Excel ignores alpha; always write `FF`. Self-contained.
- `theme="1" tint="-0.25"` — index into the theme's `clrScheme`, with the historical swap: theme 0 = `lt1`, 1 = `dk1`, 2 = `lt2`, 3 = `dk2`, then `accent1..6`, `hlink`, `folHlink`. Resolving it requires parsing `xl/theme/theme1.xml` and applying the tint in HSL space.
- `indexed="64"` — legacy 56-colour palette (0–63; 64 is "system foreground"), optionally redefined by `<colors><indexedColors>` in the same part.

Writers should emit `rgb` only: no theme part, no palette, deterministic output, trivially readable. Readers must handle all three because every fixture uses `theme="1"` on the default font and Google Sheets uses `rgb` for fills and hyperlinks (`FF1155CC`).

### 7.5 Fonts, fills, borders in brief

`<font>` children in schema order: `b, i, strike, condense, extend, outline, shadow, u, vertAlign, sz, color, name, family, charset, scheme`. Google Sheets writes them in a different order (`sz, color, name, scheme`) and Excel accepts it, but keep the canonical order. `<patternFill patternType="solid"><fgColor rgb="…"/><bgColor rgb="…"/></patternFill>`: for solid fills the **foreground** colour is the visible one (Google writes both identical). `<border>` sides in order `left, right, top, bottom, diagonal`, each optionally `style="thin|medium|thick|…"` with a `<color>` child.

### 7.6 Column width units

From ISO 29500-1 §18.3.1.13, quoted: "Column width measured as the number of characters of the maximum digit width of the numbers 0, 1, 2, …, 9 as rendered in the normal style's font. There are 4 pixels of margin padding (two on each side), plus 1 pixel padding for the gridlines. width = Truncate([{Number of Characters} * {Maximum Digit Width} + {5 pixel padding}]/{Maximum Digit Width}*256)/256". With Calibri 11 at 96 dpi the maximum digit width is 7 px, so an 8-character column is `Truncate((8*7+5)/7*256)/256 = 8.7109375`, and pixels back out as `Truncate(((256*width + Truncate(128/7))/256)*7) = 61`. Practically: `width ≈ chars + 0.71` for Calibri 11, and a writer that wants "N characters" can emit `N + 5/7`. Row heights (`ht`, `defaultRowHeight`) are in points; 15 pt is the Calibri 11 default.

## 8. `workbook.xml`

Google Sheets fixture, trimmed:

```xml
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <workbookPr/>
  <sheets>
    <sheet state="visible" name="Instructions" sheetId="1" r:id="rId4"/>
    <sheet state="visible" name="Create Accounts" sheetId="2" r:id="rId5"/>
    <sheet state="visible" name="Create Contacts" sheetId="3" r:id="rId6"/>
  </sheets>
  <definedNames/>
  <calcPr/>
</workbook>
```

Child order (CT_Workbook): `fileVersion, fileSharing, workbookPr, workbookProtection, bookViews, sheets, functionGroups, externalReferences, definedNames, calcPr, oleSize, customWorkbookViews, pivotCaches, smartTagPr, smartTagTypes, webPublishing, fileRecoveryPr, webPublishObjects, extLst`.

- `<sheet>`: `name` (section 10 rules), `sheetId` (positive integer, unique, stable identity; not an index), `r:id` (required), `state` = `visible` (default) | `hidden` | `veryHidden` (only VBA can unhide). At least one sheet must be visible, and `bookViews/workbookView/@activeTab` (0-based position) must point at a visible one, or Excel repairs.
- `<workbookPr date1904="1"/>` switches the date system (section 6.1). SheetJS writes `codeName="ThisWorkbook"`; exceljs writes `defaultThemeVersion="164011"`; both are optional.
- `<definedNames>`: `<definedName name="MyRange">Sheet1!$A$1:$B$2</definedName>`; `localSheetId` scopes a name to the sheet at that 0-based position (this is the one place position, not `sheetId`, is used); `hidden="1"` for `_xlnm._FilterDatabase` and `_xlnm.Print_Area`. Sheet names with spaces or punctuation must be single-quoted in the formula (`'Create Accounts'!$A$1`) with embedded `'` doubled.
- `<calcPr calcId="171027"/>` (exceljs) records the calc engine version; add `fullCalcOnLoad="1"` when you write formulas without cached values.
- **Chartsheets**: a `<sheet>` whose relationship type is `…/relationships/chartsheet` points at `xl/chartsheets/sheetN.xml` (content type `…spreadsheetml.chartsheet+xml`) which has no `sheetData`. The engine should surface it as a sheet with zero rows rather than crash on a missing `<worksheet>` root.

## 9. Zip container

### 9.1 Layout

```
[local file header][name][extra][compressed data][data descriptor?]   ← one per entry, in stream order
…
[central directory header][name][extra][comment]                      ← one per entry
[zip64 end of central directory record][zip64 locator]                ← only when zip64
[end of central directory record][comment]
```

Local file header (30 bytes + name + extra), little-endian: signature `50 4B 03 04`, version needed (2), flags (2), method (2), DOS time (2), DOS date (2), CRC-32 (4), compressed size (4), uncompressed size (4), name length (2), extra length (2). Central directory header (46 bytes + name + extra + comment): signature `50 4B 01 02`, version made by (2), version needed (2), flags, method, time, date, CRC, sizes, name/extra/comment lengths, disk number, internal attrs (2), external attrs (4), **relative offset of the local header (4)**. End of central directory (22 bytes): signature `50 4B 05 06`, disk numbers, entry count (2, twice), CD size (4), CD offset (4), comment length. A reader finds the EOCD by scanning backwards from the end (comment can be up to 65,535 bytes), then reads the CD; the local headers are only consulted to find where data starts (`offset + 30 + nameLen + extraLen`, using the **local** name/extra lengths, which can differ from the CD's).

### 9.2 Data descriptors (flag bit 3) and streaming

A streaming writer does not know CRC or sizes when it emits the local header. Set general-purpose bit 3 (`0x0008`), write CRC and both sizes as `0` in the local header, and append a data descriptor after the compressed data: signature `50 4B 07 08` (optional per APPNOTE, but always write it; readers must accept both forms), CRC-32, compressed size, uncompressed size (4 bytes each; 8 bytes each under zip64). The central directory carries the real values.

Excel accepts descriptors: the Google Sheets fixture uses flag `0x0808` (bit 3 + bit 11) on all 30 entries, `@office-kit/xlsx` and `write-excel-file` use `0x0008`. Excel's own writer, exceljs and SheetJS write flag `0` with sizes in the local header.

### 9.3 Methods, CRC, timestamps, names

- **Method 8 (deflate, raw RFC 1951 stream)** is the norm. **Method 0 (stored)** is legal and the SheetJS fixture uses it for every entry (17,599 bytes of which 16,573 are payload). No other method is safe.
- Version needed: 10 for stored, 20 for deflate, 45 for zip64. exceljs writes 10 with method 8 and Excel does not care; write 20 for deflate anyway.
- **CRC-32** (IEEE, polynomial `0xEDB88320`) is over the **uncompressed** bytes. Excel verifies it; a wrong CRC is a repair-dialog error **(verify that it is a repair rather than a silent accept)**.
- DOS time/date: `time = (hour << 11) | (minute << 5) | (second >> 1)`, `date = ((year - 1980) << 9) | (month << 5) | day`, local time, 2-second resolution. The Google fixture's `0x5D0A`/`0x9AB8` decode to 2026-08-10 19:21:48. A constant such as `0x0021 0x0000` (1980-01-01 00:00) is acceptable for generated files.
- Bit 11 (`0x0800`) declares UTF-8 names. OPC part names are ASCII in practice, so it is moot; Google sets it, others do not.
- Entry order is unconstrained. Excel writes `[Content_Types].xml` first; Google Sheets writes it **last**; `@office-kit/xlsx` streaming writes the worksheet first. A reader must go through the central directory, never assume first-entry semantics.
- **Duplicate names** (byte-identical or differing only in case) violate OPC; Excel refuses or repairs. Never emit them; on read, treat as a hard error rather than "last wins".
- **Directory entries** (`xl/`, `_rels/`, method 0, size 0, name ending in `/`) are emitted by exceljs and write-excel-file and ignored by Excel. Do not emit them; skip them on read.
- External attributes are irrelevant (`0x81A40020` from write-excel-file, `0` elsewhere).

### 9.4 Zip64

Zip64 is required when any of these overflow: an entry's compressed or uncompressed size ≥ `0xFFFFFFFF`, a local-header offset ≥ `0xFFFFFFFF` (archive > 4 GiB), more than `0xFFFF` entries, or a central directory ≥ 4 GiB. For xlsx the realistic trigger is a single worksheet XML over 4 GiB; rzymek reproduced it with 1,000,000 rows × 100 columns (≈4.4 GiB of `sheet1.xml`). Entry count never matters for us.

What Excel demands (rzymek's byte-level testing, corroborated by apache/poi PR #154, COMPRESS-474, golang/go#77060 and minizip-ng#894):

- The **local file header** must say version needed **45** (`0x2D`) and carry a **zip64 extended information extra field** (`0x0001`) with both 8-byte size fields, **up front**. With bit 3 set the sizes in that extra field are `0` placeholders and the data descriptor then carries **8-byte** compressed and uncompressed sizes.
- Excel is "not that strict with central directory": the CD entry should still be version 45 with `0xFFFFFFFF` in the 32-bit size/offset fields and the real values in its `0x0001` extra field, followed by the zip64 EOCD record (`50 4B 06 06`), the zip64 locator (`50 4B 06 07`) and a normal EOCD with `0xFFFF`/`0xFFFFFFFF` sentinels.
- **Retroactive zip64** — writing a version-20 local header without the extra field and only fixing things up in the central directory when the entry turns out to be big — is exactly what `java.util.zip`, Commons Compress and Go's `archive/zip` do, and it produces "We found a problem with some content". The repaired file is fine, which is why the bug survives.

Consequence for a streaming writer: decide per entry, before emitting its local header, whether it _might_ exceed 4 GiB. Either always write the zip64 local header for worksheet entries (Excel opens small entries with version 45 fine) or bound the row count. `@office-kit/xlsx` 0.11 chooses to throw on a > 4 GiB entry and only handles the entry-count case.

### 9.5 `CompressionStream('deflate-raw')`

The Web `CompressionStream` accepts `gzip` (RFC 1952 header + trailer), `deflate` (zlib RFC 1950 wrapper) and `deflate-raw` (bare RFC 1951), and `deflate-raw` is precisely what zip method 8 stores. Baseline-available in browsers since 2023 and in Node 18+. What it does **not** give you: the number of input bytes consumed (you must count the uncompressed bytes yourself with a pass-through `TransformStream`), the CRC-32 (compute it on the uncompressed side, table-driven, before the bytes enter the compressor), the compressed size (sum the output chunk lengths), a compression level, a flush/sync point (output for a small entry may only arrive on `close`), or a dictionary. In Node, `zlib.createDeflateRaw({ level })` offers level control and `bytesWritten`; in the browser you live with the defaults.

## 10. Limits

| Limit                                 | Value                                                                                                        | Source                          |
| ------------------------------------- | ------------------------------------------------------------------------------------------------------------ | ------------------------------- |
| Rows per worksheet                    | 1,048,576 (`r` ≤ 1048576)                                                                                    | Excel specifications and limits |
| Columns per worksheet                 | 16,384 (`A`…`XFD`)                                                                                           | same                            |
| Characters per cell                   | 32,767                                                                                                       | same                            |
| Line feeds per cell                   | 253                                                                                                          | same                            |
| Column width                          | 255 characters                                                                                               | same                            |
| Row height                            | 409 points                                                                                                   | same                            |
| Sheet name                            | ≤ 31 chars; not blank; no `/ \ ? * : [ ]`; no leading/trailing `'`; not `History`; unique case-insensitively | Microsoft "Rename a worksheet"  |
| Sheets per workbook                   | limited by memory                                                                                            | Excel specifications and limits |
| Formula length                        | 8,192 characters (16,384 bytes internal)                                                                     | same                            |
| Nested function levels                | 64                                                                                                           | same                            |
| Hyperlinks per worksheet              | 65,530                                                                                                       | same                            |
| Unique cell formats (`cellXfs`)       | 65,490                                                                                                       | same                            |
| Fonts per workbook                    | 512                                                                                                          | same                            |
| Number formats                        | 200–250 depending on language                                                                                | same                            |
| Defined names                         | limited by memory                                                                                            | same                            |
| Numeric precision                     | 15 significant digits displayed; ±1.7976931348623158E+308                                                    | same                            |
| Date range                            | 1900-01-01 (1904-01-01 in the 1904 system) to 9999-12-31                                                     | same                            |
| Zip entry size / offset without zip64 | < 4,294,967,295 bytes                                                                                        | PKWARE APPNOTE                  |
| Zip entries without zip64             | < 65,535                                                                                                     | PKWARE APPNOTE                  |

## 11. Security

- **DOCTYPE / XXE / entity expansion.** OPC forbids DTDs in package XML (ECMA-376 Part 2, "XML usage" **(verify clause number)**), and Excel rejects parts containing them. Reject any part whose prologue contains `<!DOCTYPE` before it reaches the parser, and configure the parser with no external entity resolution and no DTD processing so that "billion laughs" and file-disclosure payloads are impossible even if the pre-check is bypassed.
- **Zip bombs.** Apache POI's `ZipSecureFile` defaults are the reference: fail if compressed/uncompressed ratio drops below `MIN_INFLATE_RATIO = 0.01` after a 100 KB grace, and cap any single entry at `MAX_ENTRY_SIZE = 0xFFFFFFFF`. A ratio check alone false-positives on legitimate sheets (the Google fixture's sheets already compress 89–92%, and a million styled empty rows compress far better), so prefer absolute caps: maximum inflated bytes per part, maximum total inflated bytes, maximum rows/cells actually materialised, all configurable. Stream the worksheet through the SAX parser instead of inflating it whole so the cap can be on rows rather than bytes. Ignore entries you never asked for; never inflate a nested `.zip`.
- **Central-directory lies.** Sizes in the CD can disagree with the actual stream; a decompressor must stop at the cap regardless of what the headers promise. Overlapping entries (two CD records pointing into the same bytes) are a known quine trick; harmless if you only inflate named parts with a byte budget.
- **Encrypted and legacy files.** If the first 8 bytes are `D0 CF 11 E0 A1 B1 1A E1` the file is a CFB (OLE2) compound document, which for spreadsheets means either a BIFF `.xls` (`Workbook` stream) or a **password-encrypted OOXML** file (MS-OFFCRYPTO: `EncryptionInfo` + `EncryptedPackage` streams wrapping the real zip). Refuse both with a clear message; the zip is not reachable without the password. A leading `PK` with no `[Content_Types].xml` is some other zip, not an xlsx.
- **External targets.** `TargetMode="External"` hyperlinks, `xl/externalLinks/`, and OLE links are URIs to be preserved or dropped, never dereferenced.
- **Formula injection is not an xlsx concern on write**: a string that starts with `=` is stored as a string cell (`t="s"`/`inlineStr`) and stays text in Excel. It becomes a concern only if the same value is later emitted to CSV.

## 12. Repair-dialog checklist

Things that make Excel say "We found a problem with some content in 'file.xlsx'. Do you want us to try to recover as much as we can?":

1. Worksheet, workbook or styleSheet children out of schema order (`cols` after `sheetData`, `autoFilter` after `mergeCells`, `hyperlinks` before `mergeCells`, `numFmts` after `fonts`).
2. Rows not ascending, cells within a row not ascending, or a duplicate `r`.
3. A cell reference outside `A1:XFD1048576`, or `row/@r="0"`.
4. A raw control character (U+0000–U+0008, U+000B, U+000C, U+000E–U+001F), an unpaired surrogate, invalid UTF-8, or a string over 32,767 characters.
5. `s` ≥ `cellXfs` count; `fontId`/`fillId`/`borderId` out of range; a custom `numFmtId` with no `<numFmt>`.
6. `<v>` that is not a valid double (`NaN`, `Infinity`, `1,5`, empty text with `t="n"`), `t="b"` with anything but `0`/`1`, `t="s"` index ≥ `uniqueCount`, `t="e"` with a non-error literal.
7. `<f>` starting with `=`, a shared formula whose `si` has no master, an array formula with a bad `ref`.
8. A 1×1 `mergeCell`, overlapping merges, or a `mergeCell` referencing cells beyond the limits.
9. `sheetView` without `workbookViewId`; `pageMargins` missing any of its six attributes; `dataValidation` with an empty `sqref`.
10. Sheet name rule violations (section 10), two sheets with the same name ignoring case, every sheet hidden, `activeTab` pointing at a hidden sheet, a `<sheet r:id>` with no matching relationship, or a relationship pointing at a part that does not exist.
11. A worksheet part with no `Override` in `[Content_Types].xml`, or a missing `Default` for `rels`/`xml`.
12. Elements or attributes in a non-standard namespace without `mc:Ignorable` on the root (`x14ac:dyDescent` without `mc:Ignorable="x14ac"`).
13. A `<!DOCTYPE>` or an undefined entity such as `&nbsp;` in any part.
14. Zip-level: wrong CRC-32, sizes in a non-descriptor local header that disagree with the data, a central-directory offset that does not land on `50 4B 03 04`, duplicate entry names, backslashes in names, or zip64 sizes without version 45 and the `0x0001` extra field in the **local** header.
15. Fill index 0 not `none` / index 1 not `gray125` is _not_ a repair, just wrong "No Fill" behaviour; likewise wrong `count` attributes and missing `dimension` are tolerated.

## 13. Sources

- Fixtures: `fixtures/jetstream/multi-object-template.gsheets.xlsx`, `fixtures/jetstream/records-product2.sheetjs.xlsx` (inspected with `unzip -lv`, `unzip -p`, and a Node zip-header parser).
- Probes: `.generated/primer-probes/{probe.mjs,zipinfo.mjs}` producing `exceljs.xlsx`, `officekit.xlsx`, `officekit-stream.xlsx`, `wef.xlsx` from exceljs 4.4.0, @office-kit/xlsx 0.11.0, write-excel-file 2.3.10.
- ISO/IEC 29500-1 text as reproduced in the OpenXML SDK reference: [Worksheet class (CT_Worksheet sequence)](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.worksheet), [NumberingFormat class (built-in numFmt table §18.8.30)](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.numberingformat), [Column class (width formula §18.3.1.13)](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.column).
- [MS-OI29500 §22.9.2.19 ST_Xstring implementation notes](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/d34ae755-c53f-4a44-a363-c6dd3ee018a4).
- [ECMA-376 standard landing page](https://ecma-international.org/publications-and-standards/standards/ecma-376/) (Parts 1 and 2; Part 2 is the OPC).
- [Excel specifications and limits](https://support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3).
- [Rename a worksheet (sheet-name rules)](https://support.microsoft.com/en-us/office/rename-a-worksheet-3f1f7148-ee83-404d-8ef0-9ff99fbad1f9).
- [Excel incorrectly assumes that the year 1900 is a leap year](https://learn.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year).
- [PKWARE APPNOTE.TXT (ZIP file format specification)](https://pkware.cachefly.net/webdocs/casestudies/APPNOTE.TXT).
- [rzymek, "Excel and ZIP64"](https://rzymek.github.io/post/excel-zip64/); [apache/poi PR #154 (bug 57342)](https://github.com/apache/poi/pull/154); [COMPRESS-474](https://issues.apache.org/jira/browse/COMPRESS-474); [golang/go#77060](https://github.com/golang/go/issues/77060); [minizip-ng#894](https://github.com/zlib-ng/minizip-ng/issues/894).
- [Apache POI ZipSecureFile](https://poi.apache.org/apidocs/dev/org/apache/poi/openxml4j/util/ZipSecureFile.html).
- [MDN CompressionStream()](https://developer.mozilla.org/en-US/docs/Web/API/CompressionStream/CompressionStream).
- [ClosedXML #608 (escaping `_xHHHH_` in user strings)](https://github.com/ClosedXML/ClosedXML/issues/608).
