# xlsx engine replacement — research

Research, usage inventory, format reference, fixture catalog and engine evaluation behind replacing
SheetJS in Jetstream. Tooling lives at the repository root in `fixtures/`, `oracle/` and `bench/`; see
`tooling-notes.md`.

- **Audience:** Jetstream engineering; maintainers of `@jetstreamapp/simple-excel` (released 0.1.0 on 2026-09-12).
- **Accurate as of:** 2026-09-12 (see each document's own date line).
- **Start here:** ADR-008 (addendum) for the engine decision, 11 for the build contract, 05 for what each
  reader/writer actually does; 00 and 10 are the pre-decision evaluation.
- **Generated vs hand-written:** 04, 05 and 06 are rendered by `regenerate.sh` from `fixtures/edge-cases.json` and
  `fixtures/manifest.json` (04), `oracle/results/` (05) and `bench/results/` (06); everything else is hand-written.

## Document index

| #   | File                              | Status                                                                         |
| --- | --------------------------------- | ------------------------------------------------------------------------------ |
| 00  | `00-feasibility-verdict.md`       | written                                                                        |
| 01  | `01-jetstream-usage-inventory.md` | written                                                                        |
| 02  | `02-format-primer.md`             | written                                                                        |
| 03  | `03-library-landscape.md`         | written                                                                        |
| 04  | `04-edge-case-catalog.md`         | generated table + prose                                                        |
| 05  | `05-compatibility-matrix.md`      | generated tables + prose                                                       |
| 06  | `06-performance-baseline.md`      | generated results + prose                                                      |
| 07  | `07-reference-architecture.md`    | written                                                                        |
| 08  | `08-migration-plan.md`            | written (hardening candidates documented only)                                 |
| 09  | `09-risks-open-questions.md`      | written                                                                        |
| 10  | `10-office-kit-evaluation.md`     | written (decision in ADR-008)                                                  |
| 11  | `11-build-plan.md`                | written (module contracts, phases, gates)                                      |
| —   | `tooling-notes.md`                | written                                                                        |
| adr | `adr/ADR-001` … `ADR-008`         | accepted (ADR-001 revised: inline default; ADR-008 superseded by its addendum) |
