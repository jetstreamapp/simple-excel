# ADR-003: Date semantics at the engine boundary

_Status: Accepted (2026-09-12)._

## Context

The oracle showed three conventions colliding (04, EC-DATE-*): SheetJS writes serials from local Date fields
and its `sheet_to_json` hands Jetstream local-wall-clock Dates; `@office-kit/xlsx` uses UTC fields in both
directions, so a local-midnight Date lands in the file with the host offset as a time fraction (Excel shows
7:00 AM) and time-only values become negative serials (Excel shows ########). JS Dates in a DST gap or before
1901 add their own shifts.

## Decision

Jetstream's contract stays "local wall clock in, local wall clock out". The engine adapter converts at the
boundary: on write `new Date(Date.UTC(y, m, d, h, i, s, ms))` from local fields; on read, UTC fields become
local fields. Time-only values are written as fractions of a day. Dates before 1900-01-01 are written as ISO
text. The adapter, not the UI, owns this; parity tests run in `TZ=UTC` and in a US timezone.

## Consequences

No engine change is needed for correctness, but the conversion must be covered by the canonical fixture's
Date/Time/DateTime columns in both timezones, and doc 10's gap list carries the office-kit convention.
