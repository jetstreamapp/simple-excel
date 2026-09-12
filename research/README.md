# xlsx engine replacement — research

Research, usage inventory, format reference, fixture catalog and engine evaluation behind replacing
SheetJS in Jetstream. Tooling (fixtures, oracle, benchmark) lives in ``.

- **Audience:** Jetstream engineering; future maintainers of the standalone engine repo (this folder
  and `` are written to lift out unchanged).
- **Accurate as of:** 2026-09-12 (see each document's own date line).
- **Start here:** 00 for the verdict, 10 for the engine decision, 05 for what each reader/writer actually does.
- **Generated vs hand-written:** 04, 05 and 06 are rendered by `regenerate.sh` from JSON under
  ``; everything else is hand-written.

## Document index

| #   | File                              | Status                                         |
| --- | --------------------------------- | ---------------------------------------------- |
| 00  | `00-feasibility-verdict.md`       | written                                        |
| 01  | `01-jetstream-usage-inventory.md` | written                                        |
| 02  | `02-format-primer.md`             | written                                        |
| 03  | `03-library-landscape.md`         | written                                        |
| 04  | `04-edge-case-catalog.md`         | generated table + prose                        |
| 05  | `05-compatibility-matrix.md`      | generated tables + prose                       |
| 06  | `06-performance-baseline.md`      | generated results + prose                      |
| 07  | `07-reference-architecture.md`    | written                                        |
| 08  | `08-migration-plan.md`            | written (hardening candidates documented only) |
| 09  | `09-risks-open-questions.md`      | written                                        |
| 10  | `10-office-kit-evaluation.md`     | written (decision in ADR-008)                  |
| 11  | `11-build-plan.md`                | written (module contracts, phases, gates)      |
| adr | `adr/ADR-001` … `ADR-008`         | accepted                                       |
