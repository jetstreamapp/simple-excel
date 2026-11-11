/**
 * Real-world corpus audit: every .xlsx under a directory through `worker.ts`, one child process per file, a pool of
 * them at a time. Results append to a JSON-lines file, so an interrupted run resumes where it stopped.
 *
 *   tsx scripts/corpus-audit/run.ts <dir> [--out results.jsonl] [--jobs 8] [--timeout 120] [--limit N]
 *
 * A child that outlives the timeout is killed and recorded as `timeout` with the phase it was in; one that dies
 * (out of memory, native crash) is recorded as `crash` with its exit code, signal and stderr tail.
 */
import { fork } from 'node:child_process';
import { appendFileSync, existsSync, readdirSync, readFileSync, statSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { parseArgs } from 'node:util';

const { values: options, positionals } = parseArgs({
  allowPositionals: true,
  options: {
    out: { type: 'string', default: 'tmp/corpus-audit/results.jsonl' },
    jobs: { type: 'string', default: '8' },
    timeout: { type: 'string', default: '120' },
    limit: { type: 'string' },
    heap: { type: 'string', default: '4096' },
  },
});
const [directory] = positionals;
if (!directory) {
  console.error('usage: run.ts <dir> [--out results.jsonl] [--jobs 8] [--timeout 120] [--limit N]');
  process.exit(2);
}

function listFiles(root: string): string[] {
  const files: string[] = [];
  const walk = (current: string): void => {
    for (const entry of readdirSync(current, { withFileTypes: true })) {
      const path = join(current, entry.name);
      if (entry.isDirectory()) {
        walk(path);
      } else if (/\.xls[xm]$/i.test(entry.name) && !entry.name.startsWith('._')) {
        files.push(resolve(path));
      }
    }
  };
  walk(root);
  return files.sort();
}

const outPath = resolve(options.out);
const done = new Set<string>();
if (existsSync(outPath)) {
  for (const line of readFileSync(outPath, 'utf8').split('\n')) {
    if (line.trim()) {
      done.add((JSON.parse(line) as { file: string }).file);
    }
  }
}
let queue = listFiles(directory).filter(file => !done.has(file));
if (options.limit) {
  queue = queue.slice(0, Number(options.limit));
}
const total = queue.length;
const jobs = Number(options.jobs);
const timeoutMs = Number(options.timeout) * 1000;
const workerPath = resolve(__dirname, 'worker.ts');
console.log(`${done.size} already done, ${total} to run, ${jobs} at a time, ${options.timeout}s timeout -> ${outPath}`);

let finished = 0;
const tally: Record<string, number> = {};
const startedAt = Date.now();

function record(entry: Record<string, unknown>): void {
  appendFileSync(outPath, `${JSON.stringify(entry)}\n`);
  finished++;
  const status = String(entry.status);
  tally[status] = (tally[status] ?? 0) + 1;
  if (finished % 25 === 0 || finished === total) {
    const rate = finished / ((Date.now() - startedAt) / 1000);
    console.log(`${finished}/${total} (${rate.toFixed(1)}/s) ${JSON.stringify(tally)}`);
  }
}

function statusOf(result: Record<string, any>): string {
  if (!result.ours?.ok) {
    return result.ours?.error?.classified ? 'rejected' : 'unclassified-error';
  }
  const problems: string[] = [];
  if (result.parity && !result.parity.pass) {
    problems.push('parity');
  }
  if (result.objects && !result.objects.pass) {
    problems.push('objects');
  }
  if (result.roundtrip && (result.roundtrip.error || !result.roundtrip.pass)) {
    problems.push('roundtrip');
  }
  return problems.length > 0 ? `diff:${problems.join('+')}` : 'pass';
}

function runOne(file: string): Promise<void> {
  return new Promise(resolvePromise => {
    let phase = 'start';
    let settled = false;
    let stderr = '';
    const child = fork(workerPath, [file], {
      execArgv: ['--import', 'tsx', `--max-old-space-size=${options.heap}`],
      stdio: ['ignore', 'ignore', 'pipe', 'ipc'],
    });
    child.stderr?.on('data', (chunk: Buffer) => {
      stderr = (stderr + chunk.toString()).slice(-2000);
    });
    const finish = (entry: Record<string, unknown>): void => {
      if (settled) {
        return;
      }
      settled = true;
      clearTimeout(timer);
      record({ file, ...entry });
      resolvePromise();
    };
    const timer = setTimeout(() => {
      child.kill('SIGKILL');
      finish({ status: 'timeout', phase, bytes: statSync(file).size });
    }, timeoutMs);
    child.on('message', (message: Record<string, any>) => {
      if (message.phase) {
        phase = message.phase;
      } else if (message.result) {
        finish({ status: statusOf(message.result), ...message.result });
      } else if (message.fatal) {
        finish({ status: 'harness-error', phase, error: message.fatal });
      }
    });
    child.on('exit', (code, signal) => {
      finish({ status: 'crash', phase, code, signal, stderr, bytes: statSync(file).size });
    });
  });
}

async function pool(): Promise<void> {
  let next = 0;
  const lane = async (): Promise<void> => {
    while (next < queue.length) {
      const file = queue[next++] as string;
      await runOne(file);
    }
  };
  await Promise.all(Array.from({ length: jobs }, lane));
  console.log(`done in ${((Date.now() - startedAt) / 1000).toFixed(0)}s ${JSON.stringify(tally)}`);
}

void pool();
