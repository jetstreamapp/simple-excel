# ADR-001: Shared strings vs inline strings on write

_Status: Revised 2026-09-12 (default flipped to inline; see addendum). Originally accepted (2026-09-12)._

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

## Addendum (2026-09-12): inline is the default

Two measurements after the library was built changed the default from the bounded hybrid to `'inline'`:

- Apple Numbers 14.4 reads shared strings unfaithfully: it truncates them at the first control character
  (`a\u0001b` becomes `a`) and mis-decodes protected escapes (`_x005F_x0041_` becomes `_x005FA`), while it reads the
  same values from inline strings correctly (goldens `golden-canonical-numbers-resave-simple-excel` vs `-inline`).
  Excel, Google Sheets, LibreOffice, openpyxl and calamine read both shapes identically.
- On the `mixed` 100k×20 benchmark data the table saves 1.4% of compressed size and costs about 15% of write time
  with the platform deflater (44.49 MB in 4.86 s vs 45.10 MB in 4.09 s; equal with zlib). Unique-heavy data gains
  nothing from the table at all.

Inline also matches what Jetstream ships today (`bookSST: false` in SheetJS) and removes the only data structure
whose size depended on the data. The bounded hybrid stays available as `strings: 'auto'` for very low-cardinality
exports where size matters more than speed; the goldens pin both shapes.
