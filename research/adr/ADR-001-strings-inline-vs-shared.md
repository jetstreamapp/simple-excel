# ADR-001: Shared strings vs inline strings on write

_Status: Accepted (2026-09-12)._

## Context

Jetstream writes with `bookSST:false` today (every string inline) except the Permission Manager export
(`bookSST:true`). Inline strings are accepted by Excel, LibreOffice, Google Sheets, openpyxl, calamine and
office-kit (fixture `golden-canonical-sheetjs`, `edge-no-shared-strings-part`). An unbounded shared string
table is what makes SheetJS's memory explode on unique-heavy exports (Ids, names).

## Decision

Bounded hybrid. Intern strings of ≤ 256 chars while the table holds fewer than 65,536 unique entries and
fewer than 16 Mi chars; after that the map is frozen (repeats still resolve) and new strings go inline.
Always emit `xl/sharedStrings.xml` with exact `count`/`uniqueCount`. If the chosen engine (`@office-kit/xlsx`)
does not expose this policy, its streaming writer's default is acceptable as long as it stays bounded; verify
with the `strings-unique` benchmark dataset.

## Consequences

Files stay Excel-fast for low-cardinality Salesforce columns, memory stays flat for unique-heavy ones, and
mixed `s`/`inlineStr` cells in one sheet must be covered by the oracle (they are: Excel opens the goldens).
