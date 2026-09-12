# ADR-008: Engine selection

_Status: Accepted (2026-09-12) - dual track, time-boxed._

## Context

The user's rule: adopt `@office-kit/xlsx` if it meets 100% of Jetstream's needs; work with the maintainers on
gaps; fork only as a last resort. Doc 07 is the rubric; docs 04/05/06/10 hold the evidence.

## Findings that decide it

- office-kit's "fixed-memory streaming" writer buffers the worksheet and an unbounded shared string table
  until `finalize()` and fails 1M×20 and the 18M-cell shape with the same `RangeError` as SheetJS (10, W6/W7).
  Scale is the reason this work exists, so this is blocking, not a nice-to-have.
- Its reader needs four fixes to read other generators' files faithfully (escapes, missing `r`, Strict,
  backslash targets) and a documented Date convention; everything else is adapter glue.
- The missing writer pieces (streaming zip entries, chunked SST, inline/bounded strings) are exactly Phase B
  of the hand-rolled design and are small relative to the whole library.

## Decision

1. Open the upstream issues listed in 10 §8 now, leading with the streaming-writer and SST ones, each with a
   fixture from this repo and an offer to contribute the fix.
2. In parallel, build Phase A-B of the hand-rolled writer (07): zip + escapes + minimal styles + a truly
   streaming worksheet writer with the bounded-SST policy, validated by this oracle. It is the fallback and
   also the reference implementation for the upstream proposal.
3. Time-box: four weeks from the issue date. Then choose:
   - **ADOPT + CONTRIBUTE** if the maintainer accepts streaming/SST changes (ours or theirs) and R1 lands;
   - **MIXED** (our writer, office-kit reader behind the adapter) if only the reader items land;
   - **HAND-ROLL** (07 Phases C-E) otherwise.
4. Fork only if office-kit becomes unmaintained after adoption (ADR-006).

## Consequences

The corpus, oracle and benchmark in this repository are the acceptance suite for whichever path wins;
Jetstream's adapter contracts (dates, escapes, sniffing, `password-protected`, object mode) are engine-neutral.
