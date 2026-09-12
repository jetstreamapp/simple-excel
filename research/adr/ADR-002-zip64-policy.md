# ADR-002: Zip64 policy

_Status: Accepted (2026-09-12)._

## Context

Excel repairs a workbook when an entry exceeds 4 GiB unless Zip64 was declared up-front in the _local file
header_ (version-needed 4.5 + zip64 extra field); declaring it only in the central directory is not enough
(Apache POI bug 57342, rzymek's write-up). `@office-kit/xlsx` 0.11 writes a zip64 EOCD for > 65,535 entries
but caps individual entries at 4 GiB.

## Decision

Declare Zip64 in every worksheet and shared-strings local header whenever the estimated uncompressed size is
unknown or above 3 GiB; keep static parts 32-bit. Fail with a clear error rather than emit a file Excel will
repair. For adoption of office-kit this is a requested feature (entry > 4 GiB), tracked in doc 10; Jetstream's
18M-cell case is ~0.8 GB of XML and fits without it.

## Consequences

The 4 GiB-per-entry ceiling is an accepted limit for v1 (≈ 100M cells); anything larger is split across
sheets or served as CSV.
