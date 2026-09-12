# ADR-007: Drop legacy formats and codepage tables

_Status: Accepted (2026-09-12)._

## Context

SheetJS silently parses BIFF8 `.xls`, `.ods`, CSV bytes and SpreadsheetML 2003 when they reach `XLSX.read`
(Drive `alt=media`, Electron open-with, renamed uploads). Jetstream's users are business users on modern
tools; the codepage table (`cpexcel.full.mjs`) exists only for those legacy paths and needs an extension CSP
workaround.

## Decision

Support `.xlsx`/`.xlsm` (Transitional; Strict read-only) only. Sniff inputs: zip → xlsx (or `.ods` by its
`mimetype` entry → clear error), CFB → distinguish legacy `.xls` from encrypted (`EncryptedPackage` stream)
and say which, text → route to papaparse. Remove the codepage table.

## Consequences

Error messages become a contract: the encrypted case must contain `password-protected` (the UI matches on
it). office-kit's current CFB message misclassifies `.xls` as encrypted and lacks that substring (doc 10).
