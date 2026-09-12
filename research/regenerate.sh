#!/usr/bin/env bash
#
# Re-render the generated sections of the xlsx-engine research docs from the tooling's JSON outputs.
#
#   04-edge-case-catalog.md   <- fixtures/edge-cases.json + manifest.json
#   05-compatibility-matrix.md <- newest oracle/results/<run>/results.json
#   06-performance-baseline.md <- newest bench/results/<run>/summary.md
#
# Idempotent: only the text between <!-- generated:start --> and <!-- generated:end --> changes.
# Usage: ./regenerate.sh            (all)      ./regenerate.sh 04|05|06   (one document)
set -euo pipefail
cd "$(dirname "${BASH_SOURCE[0]}")"
TOOLS=..
WHAT="${1:-all}"

if [[ "$WHAT" == "all" || "$WHAT" == "04" ]]; then
  node "$TOOLS/fixtures/render-catalog.mjs" --doc 04-edge-case-catalog.md
fi
if [[ "$WHAT" == "all" || "$WHAT" == "05" ]]; then
  node "$TOOLS/oracle/render-matrix.mjs" --doc 05-compatibility-matrix.md
fi
if [[ "$WHAT" == "all" || "$WHAT" == "06" ]]; then
  # prefer a merged run (name ends in -combined), else the newest folder
  latest=$(ls -d "$TOOLS"/bench/results/*-combined/ 2>/dev/null | sort | tail -1 || true)
  if [[ -z "$latest" ]]; then
    latest=$(ls -d "$TOOLS"/bench/results/*/ 2>/dev/null | sort | tail -1 || true)
  fi
  if [[ -n "$latest" && -f "$latest/summary.md" ]]; then
    node - "$latest/summary.md" 06-performance-baseline.md <<'NODE'
const fs = require('node:fs');
const [summaryPath, docPath] = process.argv.slice(2);
const summary = fs.readFileSync(summaryPath, 'utf8').replace(/^# .*\n/, '');
const doc = fs.readFileSync(docPath, 'utf8');
const start = doc.indexOf('<!-- generated:start');
const end = doc.indexOf('<!-- generated:end -->');
if (start < 0 || end < 0) { throw new Error('markers not found in ' + docPath); }
const block = `<!-- generated:start (bench/results) -->\n\nSource run: \`${summaryPath.split('/').slice(-2, -1)[0]}\`\n\n${summary.trim()}\n\n<!-- generated:end -->`;
fs.writeFileSync(docPath, doc.slice(0, start) + block + doc.slice(end + '<!-- generated:end -->'.length));
console.log('rendered ' + summaryPath + ' into ' + docPath);
NODE
  else
    echo "06: no bench results yet (bench/results/*/summary.md) - skipped"
  fi
fi
