#!/usr/bin/env node
// Runs the read-only Excel oracle on one file and prints/saves a JSON verdict.
//   node run.mjs <file.xlsx> [--sheet Data] [--refs A2,B2] [--out result.json] [--timeout 120000]
import { spawnSync } from 'node:child_process';
import { existsSync, readdirSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import { homedir } from 'node:os';
import { resolve } from 'node:path';
import { parseArgs } from 'node:util';

const { values: options, positionals } = parseArgs({
  allowPositionals: true,
  options: {
    sheet: { type: 'string', default: 'Data' },
    refs: { type: 'string' },
    out: { type: 'string' },
    timeout: { type: 'string', default: '120000' },
  },
});
const [file] = positionals;
if (!file) {
  console.error('usage: run.mjs <file.xlsx> [--sheet Data] [--refs A2,B2] [--out result.json]');
  process.exit(2);
}
const script = new URL('./open-check.applescript', import.meta.url).pathname;
const args = [script, resolve(file), options.sheet];
if (options.refs) {
  args.push(options.refs);
}
// Under automation Excel repairs silently (no prompt, no "[Repaired]" in the name); the only evidence is a new
// "Repair Result to <name>N.xml" recovery log in Excel's container tmp folder, so snapshot it around the open.
const REPAIR_LOG_DIR = `${homedir()}/Library/Containers/com.microsoft.Excel/Data/tmp`;
function repairLogs() {
  if (!existsSync(REPAIR_LOG_DIR)) {
    return new Map();
  }
  return new Map(
    readdirSync(REPAIR_LOG_DIR)
      .filter(name => /^Repair Result/i.test(name))
      .map(name => [name, statSync(`${REPAIR_LOG_DIR}/${name}`).mtimeMs]),
  );
}
const logsBefore = repairLogs();
const started = Date.now();
const result = spawnSync('osascript', args, { encoding: 'utf8', timeout: Number(options.timeout) });
const verdict = {
  file: resolve(file),
  ms: Date.now() - started,
  status: 'UNKNOWN',
  excel: 'Microsoft Excel (read-only license; open + inspect only)',
};
if (result.error?.code === 'ETIMEDOUT' || result.signal) {
  // `open` is blocked by a modal - almost always "We found a problem with some content… Do you want us to try
  // to recover?". Accept it (Return = Yes), then Excel shows the repair summary (Return = Close), and the
  // recovered workbook opens with "[Repaired]" in its name.
  verdict.status = 'HUNG';
  const accept = delayMs =>
    spawnSync(
      'osascript',
      ['-e', `delay ${delayMs / 1000}`, '-e', 'tell application "System Events" to tell process "Microsoft Excel" to keystroke return'],
      { timeout: 20000 },
    );
  accept(500);
  accept(4000);
  const name = spawnSync('osascript', ['-e', 'delay 3', '-e', 'tell application "Microsoft Excel" to name of active workbook'], {
    encoding: 'utf8',
    timeout: 30000,
  });
  verdict.workbookName = name.stdout?.trim() || null;
  if (/\[Repaired\]/i.test(verdict.workbookName ?? '')) {
    verdict.status = 'REPAIRED';
  }
  verdict.note =
    verdict.status === 'REPAIRED'
      ? 'Excel showed the repair prompt; the workbook opened as [Repaired].'
      : 'Excel did not return - a modal is blocking. Dismiss it manually; treat the file as REPAIRED/INVALID.';
  spawnSync('osascript', ['-e', 'tell application "Microsoft Excel" to close every workbook saving no'], { timeout: 20000 });
} else if (result.status !== 0) {
  verdict.status = 'ERROR';
  verdict.error = (result.stderr || result.stdout).trim().slice(0, 500);
} else {
  const lines = result.stdout.trim().split('\n');
  verdict.workbookName = lines.find(line => line.startsWith('workbook='))?.slice(9) ?? null;
  verdict.sheets = (lines.find(line => line.startsWith('sheets='))?.slice(7) ?? '').split('|').filter(Boolean);
  verdict.cells = lines
    .filter(line => !line.startsWith('workbook=') && !line.startsWith('sheets='))
    .map(line => {
      const [ref, value, text] = line.split('|||');
      const clip = input => (input !== undefined && input.length > 200 ? `${input.slice(0, 200)}… (${input.length} chars)` : input);
      return { ref, value: clip(value), text: clip(text) };
    });
  verdict.status = /\[Repaired\]/i.test(verdict.workbookName ?? '') ? 'REPAIRED' : 'PASS';
}
const newLogs = [...repairLogs()].filter(([name, mtime]) => !logsBefore.has(name) || logsBefore.get(name) !== mtime);
if (newLogs.length > 0) {
  verdict.status = 'REPAIRED';
  verdict.repairLogs = newLogs.map(([name]) => {
    const text = readFileSync(`${REPAIR_LOG_DIR}/${name}`, 'utf8');
    const removed = [...text.matchAll(/<removedRecord>([^<]*)<\/removedRecord>/g)].map(match => match[1]);
    const repaired = [...text.matchAll(/<repairedRecord>([^<]*)<\/repairedRecord>/g)].map(match => match[1]);
    return { log: name, removed, repaired };
  });
  verdict.note = 'Excel wrote a recovery log while opening this file (silent repair under automation).';
}
const json = JSON.stringify(verdict, null, 2) + '\n';
if (options.out) {
  writeFileSync(options.out, json);
}
process.stdout.write(json);
// process.exit() would truncate a large pending stdout write on a pipe (observed at 64 KB); let it drain
process.exitCode = verdict.status === 'PASS' ? 0 : 1;
