# ADR-004: Where spreadsheet work runs

_Status: Accepted (2026-09-12)._

## Context

All SheetJS work runs on the main thread today (`JobWorker.ts` is a fake worker). The one real module worker
pattern in the repo (`opfs-file-store.ts`) bundles with a literal `new Worker(new URL(...), { type: 'module' })`.
The MV3 extension forbids `blob:` workers and `wasm-unsafe-eval`; the web CSP allows workers from `'self'`
and `blob:` only.

## Decision

The engine (own or office-kit) must be thread-agnostic: no DOM globals, `Uint8Array`/`Blob` in and out.
Jetstream hosts exports and large parses in a dedicated module worker using the proven literal-URL pattern;
Canvas and small files stay on the main thread with yielding. Pure JS only - no wasm on any surface.

## Consequences

office-kit qualifies (pure JS, Web Streams); its bundle must be verified not to spawn `blob:` workers
(fflate's async API does) - see doc 10.
