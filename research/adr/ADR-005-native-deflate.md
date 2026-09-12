# ADR-005: Compression primitive

_Status: Accepted (2026-09-12)._

## Context

`CompressionStream`/`DecompressionStream('deflate-raw')` exist on every Jetstream target (Chrome 121+,
Firefox 120+, Safari 16.4+, Node 24). A JS deflate implementation costs 300-500 lines and CPU.

## Decision

For a hand-rolled engine: native streams with a stored (method 0) fallback, no JS deflate in v1. For
office-kit: it bundles fflate (pure JS); accept it, but measure it against native streams in the benchmark
and treat "native when available" as an upstream proposal if the gap matters.

## Consequences

Node-side callers can pick a deflate level through zlib; browser callers cannot (platform API has no level).
