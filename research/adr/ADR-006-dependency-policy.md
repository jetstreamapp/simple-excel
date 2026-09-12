# ADR-006: Dependency policy

_Status: Accepted (2026-09-12)._

## Context

SheetJS CE is unmaintained on npm (0.18.5 with two CVEs; fixes only on the vendor CDN). ExcelJS is inactive.
`@office-kit/xlsx` is MIT with three runtime dependencies (fflate, fast-xml-parser, saxes), one primary
maintainer, and a two-month release history.

## Decision

A hand-rolled engine has zero runtime dependencies. Adopting office-kit is acceptable with: pinned versions,
Renovate-style upgrade review, a vendored fallback plan (fork on a Jetstream org) if maintenance stops, and
the parity/oracle suite run against every upgrade. Dependencies must be pure JS with no `eval`/wasm.

## Consequences

The oracle and benchmark harness in this repository become the upgrade gate for whichever engine ships.
