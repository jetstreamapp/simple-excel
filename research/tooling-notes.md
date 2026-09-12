# Research tooling notes

Fixtures, compatibility oracle and benchmark harness behind the xlsx engine replacement
(see `research/` for the research documents this tooling feeds).

This tooling was built inside the Jetstream monorepo and lifted here unchanged: **nothing in it imports
`@jetstream/*`** (`node scripts/check-purity.mjs` enforces this). The comparison libraries (`exceljs`,
`@office-kit/xlsx`, `write-excel-file`, `read-excel-file`, `xlsx`, `@xarsh/ooxml-validator`) are devDependencies.

| Folder      | What                                                                                                                 | Entry point                              |
| ----------- | -------------------------------------------------------------------------------------------------------------------- | ---------------------------------------- |
| `fixtures/` | Canonical dataset, goldens per generator, edge/hostile files, `manifest.json` (single source of truth)               | `npm run fixtures:check`                 |
| `oracle/`   | "Does it open / do the values match" across simple-excel, Excel, LibreOffice, validator, Python, SheetJS, office-kit | `npm run oracle -- <fixture-id or path>` |
| `bench/`    | Engine benchmark (sheetjs, exceljs, @office-kit/xlsx, write-excel-file, simple-excel)                                | `npm run bench -- --help`                |

## Rules

- Commit fixtures under 200 KB (15 MB total budget). Anything larger is `generated: true` in the
  manifest, built into `.generated/` (gitignored) by `fixtures/generators/<id>.mjs`, and pinned by sha256.
- Every fixture has provenance (where it came from, which generator/version, license) in the manifest.
- Benchmark and oracle results are committed under dated folders; the docs are regenerated from them.
- Generated, unversioned files go under `.generated/` (gitignored).
- `fixtures/sfdc/load-report-rows.mjs` is the one Jetstream-coupled helper (it shells out to `pnpm sf:api` in a
  Jetstream checkout named by `JETSTREAM_REPO` to load report-fixture rows into a dev org).

## Prerequisites (local oracle)

- `npm run build` first: the `simple-excel` reader (`oracle/simple-excel/read-dump.mjs`) reads through
  `dist/esm/index.mjs`, so an unbuilt or stale `dist/` silently grades the wrong code.
- macOS with Microsoft Excel (a read-only license is enough: the oracle opens and inspects, never saves),
  Numbers, and LibreOffice.app. The first run prompts for Automation/Accessibility permissions.
- Python 3: `python3 -m venv .generated/venv && .generated/venv/bin/pip install -r oracle/python/requirements.txt`
- `@xarsh/ooxml-validator` is installed with the workspace (bundled Open XML SDK binary, no .NET needed).
