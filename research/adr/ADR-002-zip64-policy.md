# ADR-002: Zip64 policy

_Status: Accepted (2026-09-12); revised the same day after the Excel oracle._

## Context

Excel repairs a workbook when an entry exceeds 4 GiB unless Zip64 was declared up-front in the _local file
header_ (version-needed 4.5 + zip64 extra field); declaring it only in the central directory is not enough
(Apache POI bug 57342, rzymek's write-up). The Phase B oracle run added two constraints Excel enforces on any
zip64 archive, small or large (EC-ZIP64-LOCAL-SIZES-BIT3, EC-ZIP64-EOCD-SENTINELS): with bit 3 set the local
header sizes must be literal zeros, and once the zip64 end records are written every field of the 32-bit end
record must be a sentinel. It also showed that SheetJS 0.20.3 cannot open zip64 archives at all, and a
manual Google Drive upload showed Google Sheets fails to convert them (EC-ZIP64-SMALL), while Excel, LibreOffice,
openpyxl and calamine can.

## Decision

- The zip writer declares zip64 per entry, up-front, in the shape Excel accepts; static parts stay 32-bit; the
  archive finishes with the zip64 end records whenever any entry used zip64 or a count/offset overflows.
- `zip64: 'auto'` (the default) is conservative: a worksheet gets zip64 only when its announced `rowCount`
  makes a >4 GiB part plausible (about 40M cells). A sheet with an unknown row count stays 32-bit, and the
  writer fails with `ENTRY_TOO_LARGE` (naming `zip64: true` as the fix) in the unlikely case a part passes
  4 GiB. Excel's own 1,048,576-row limit means such a sheet needs more than 4 KB of XML per row.
- `zip64: true` forces zip64 on every streamed part for callers who know their output is huge and whose
  readers support it; `false` never emits it.

## Consequences

Files written with the defaults open in every reader in the compatibility matrix, SheetJS and Google Sheets
included. Google Sheets caps a spreadsheet at 10M cells, so nothing that genuinely needs zip64 could be imported
there under any policy. The 4 GiB-per-part ceiling for unknown-size streams is an accepted v1 limit;
callers that can announce `rowCount` get zip64 automatically when it matters.
