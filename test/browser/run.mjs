#!/usr/bin/env node
/**
 * Cross-browser smoke test runner. Serves the repository root with the COOP/COEP headers that unlock
 * `performance.measureUserAgentSpecificMemory` (same as `bench/lib/chrome.mjs`), opens
 * `test/browser/smoke.html` in each of Playwright's Chromium, Firefox and WebKit builds, waits for
 * `window.__smokeResults`, prints a table per browser and exits non-zero if any check failed.
 *
 *   node test/browser/run.mjs                              # all three browsers
 *   node test/browser/run.mjs --browsers firefox,webkit    # a subset
 *   node test/browser/run.mjs --serve                      # keep serving; open the URL in real Safari
 *
 * The page imports `/dist/esm/index.mjs`, so `npm run build` has to have run first (`npm run smoke:browsers`
 * does it for you). Firefox and WebKit need `npx playwright install firefox webkit`.
 */
import { createServer } from 'node:http';
import { existsSync, readFileSync, statSync } from 'node:fs';
import { extname, join, normalize, resolve, sep } from 'node:path';
import { fileURLToPath } from 'node:url';

/** `resolve` strips the trailing separator, so the containment check below can append one. */
const REPO_ROOT = resolve(fileURLToPath(new URL('../..', import.meta.url)));
const PAGE_PATH = '/test/browser/smoke.html';
const ALL_BROWSERS = ['chromium', 'firefox', 'webkit'];
const DEFAULT_TIMEOUT_SECONDS = 300;

const HELP = `
simple-excel cross-browser smoke test

Usage: node test/browser/run.mjs [options]

Options:
  --browsers <list>  chromium,firefox,webkit (default: all three)
  --serve            start the static server, print the URL and keep it running (for a manual Safari check)
  --port <n>         server port (default: an ephemeral port)
  --timeout <sec>    per-browser timeout (default: ${DEFAULT_TIMEOUT_SECONDS})
  --headed           show the browser window
  --help             this text
`;

const CONTENT_TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.mjs': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.map': 'application/json; charset=utf-8',
  '.csv': 'text/csv; charset=utf-8',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
};

/** A bad flag is a usage mistake, not a crash: it prints one line and the help text, never a stack. */
class UsageError extends Error {}

function parseArgs(argv) {
  const options = {
    browsers: [...ALL_BROWSERS],
    serve: false,
    port: 0,
    timeoutSeconds: DEFAULT_TIMEOUT_SECONDS,
    headed: false,
    help: false,
  };
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    const next = () => {
      const value = argv[++i];
      if (value === undefined) {
        throw new UsageError(`${arg} requires a value`);
      }
      return value;
    };
    switch (arg) {
      case '--browsers': {
        options.browsers = next()
          .split(',')
          .map(name => name.trim())
          .filter(Boolean);
        const unknown = options.browsers.filter(name => !ALL_BROWSERS.includes(name));
        if (unknown.length > 0) {
          throw new UsageError(`Unknown browser(s) ${unknown.join(', ')} (known: ${ALL_BROWSERS.join(', ')})`);
        }
        break;
      }
      case '--serve':
        options.serve = true;
        break;
      case '--port':
        options.port = Number(next());
        break;
      case '--timeout':
        options.timeoutSeconds = Number(next());
        break;
      case '--headed':
        options.headed = true;
        break;
      case '--help':
      case '-h':
        options.help = true;
        break;
      default:
        throw new UsageError(`Unknown option ${arg}`);
    }
  }
  return options;
}

/** Static file server over the repository root. Cross-origin isolated, so the memory step can run in Chromium. */
function startServer(port) {
  const server = createServer((request, response) => {
    const requestPath = decodeURIComponent(new URL(request.url, 'http://localhost').pathname);
    if (requestPath === '/favicon.ico') {
      response.writeHead(204).end();
      return;
    }
    const file = resolve(REPO_ROOT, `.${normalize(requestPath)}`);
    if (!file.startsWith(REPO_ROOT + sep) || !existsSync(file) || !statSync(file).isFile()) {
      response.writeHead(404).end('not found');
      return;
    }
    response.writeHead(200, {
      'Content-Type': CONTENT_TYPES[extname(file)] ?? 'application/octet-stream',
      // Cross-origin isolation; same-origin subresources satisfy require-corp on their own.
      'Cross-Origin-Opener-Policy': 'same-origin',
      'Cross-Origin-Embedder-Policy': 'require-corp',
      'Cache-Control': 'no-store',
    });
    response.end(readFileSync(file));
  });
  return new Promise((resolvePromise, rejectPromise) => {
    server.once('error', rejectPromise);
    server.listen(port, '127.0.0.1', () => resolvePromise({ server, origin: `http://127.0.0.1:${server.address().port}` }));
  });
}

class MissingBrowserError extends Error {}

/**
 * Chromium is launched from the full browser rather than Playwright's default headless shell: the shell rejects
 * `performance.measureUserAgentSpecificMemory()` with a SecurityError, which would silently skip the memory step.
 */
const LAUNCH_OPTIONS = { chromium: { channel: 'chromium' }, firefox: {}, webkit: {} };

async function launch(browserType, name, headed) {
  try {
    return await browserType.launch({ headless: !headed, ...LAUNCH_OPTIONS[name] });
  } catch (error) {
    if (/executable doesn't exist|playwright install/i.test(error.message)) {
      throw new MissingBrowserError(`Playwright's ${name} build is not installed`);
    }
    throw error;
  }
}

async function runBrowser(browserType, name, origin, options) {
  const browser = await launch(browserType, name, options.headed);
  const consoleErrors = [];
  try {
    const context = await browser.newContext();
    const page = await context.newPage();
    page.on('console', message => {
      if (message.type() === 'error') {
        consoleErrors.push(message.text());
      }
    });
    page.on('pageerror', error => consoleErrors.push(`pageerror: ${error.message}`));
    await page.goto(`${origin}${PAGE_PATH}`, { waitUntil: 'domcontentloaded' });
    let timedOut = null;
    try {
      await page.waitForFunction(() => window.__smokeResults?.done === true, undefined, { timeout: options.timeoutSeconds * 1000 });
    } catch (error) {
      timedOut = `timed out after ${options.timeoutSeconds}s: ${error.message}`;
    }
    // Read whatever the page got through even on a timeout: the last step it reported is the diagnosis.
    const results = await page.evaluate(() => window.__smokeResults ?? null);
    if (!results) {
      throw new Error(
        `${timedOut ?? 'the page never set window.__smokeResults'}${consoleErrors.length > 0 ? ` (${consoleErrors.join(' | ')})` : ''}`,
      );
    }
    if (timedOut) {
      results.fatal = timedOut;
      results.totals.failed += 1;
      results.totals.checks += 1;
    }
    return { name, version: browser.version(), results, consoleErrors };
  } finally {
    await browser.close().catch(() => {});
  }
}

function formatMs(ms) {
  return typeof ms === 'number' && Number.isFinite(ms) ? `${Math.round(ms)} ms` : '';
}

function printTable(rows) {
  if (rows.length === 0) {
    return;
  }
  const widths = rows[0].map((_, column) => Math.max(...rows.map(row => String(row[column] ?? '').length)));
  for (const row of rows) {
    const line = row.map((cell, column) => String(cell ?? '').padEnd(column === row.length - 1 ? 0 : widths[column])).join('  ');
    console.log(`  ${line.trimEnd()}`);
  }
}

function report({ name, version, results, consoleErrors }) {
  console.log(`\n${'='.repeat(110)}`);
  console.log(`${name} ${version}`);
  console.log(results.userAgent);
  console.log(`cross-origin isolated: ${results.crossOriginIsolated}`);
  console.log('='.repeat(110));

  const rows = [['STEP / CHECK', 'RESULT', 'TIME', 'DETAIL']];
  for (const step of results.steps) {
    const failed = step.checks.filter(check => !check.ok).length;
    rows.push([step.label, failed === 0 ? `${step.checks.length} ok` : `${failed} FAILED`, formatMs(step.ms), '']);
    for (const check of step.checks) {
      rows.push([`  ${check.id}`, check.ok ? 'pass' : 'FAIL', '', check.detail]);
    }
  }
  printTable(rows);

  const metrics = results.metrics;
  console.log('\n  metrics:');
  for (const [key, value] of Object.entries(metrics)) {
    console.log(`    ${key}: ${JSON.stringify(value)}`);
  }
  if (consoleErrors.length > 0) {
    console.log('\n  console errors:');
    for (const message of consoleErrors) {
      console.log(`    ${message}`);
    }
  }
  const { checks, failed } = results.totals;
  console.log(`\n  ${checks - failed}/${checks} checks passed${results.fatal ? ` — fatal: ${results.fatal}` : ''}`);
}

async function main() {
  const options = parseArgs(process.argv.slice(2));
  if (options.help) {
    console.log(HELP.trim());
    return;
  }
  if (!existsSync(join(REPO_ROOT, 'dist/esm/index.mjs'))) {
    console.error('[smoke] dist/esm/index.mjs is missing — run `npm run build` first.');
    process.exitCode = 1;
    return;
  }

  const { server, origin } = await startServer(options.port);
  const url = `${origin}${PAGE_PATH}`;

  if (options.serve) {
    console.log(`[smoke] serving ${REPO_ROOT}`);
    console.log(`[smoke] open this in Safari (or any browser) for the same table:\n\n    ${url}\n`);
    console.log('[smoke] Ctrl+C to stop.');
    return;
  }

  const playwright = await import('@playwright/test');
  const outcomes = [];
  const failures = [];
  try {
    for (const name of options.browsers) {
      console.log(`[smoke] ${name}…`);
      try {
        const outcome = await runBrowser(playwright[name], name, origin, options);
        outcomes.push(outcome);
        if (outcome.results.totals.failed > 0 || outcome.results.fatal) {
          failures.push(
            `${name}: ${outcome.results.totals.failed} failed check(s)${outcome.results.fatal ? `, fatal ${outcome.results.fatal}` : ''}`,
          );
        }
      } catch (error) {
        const hint =
          error instanceof MissingBrowserError
            ? `${error.message}\n            Run: npx playwright install firefox webkit`
            : `${error.name}: ${error.message}`;
        console.error(`[smoke] ${name} could not run: ${hint}`);
        failures.push(`${name}: ${error.message}`);
      }
    }
  } finally {
    server.close();
  }

  for (const outcome of outcomes) {
    report(outcome);
  }

  console.log(`\n${'='.repeat(110)}`);
  for (const outcome of outcomes) {
    const { checks, failed } = outcome.results.totals;
    console.log(`${outcome.name.padEnd(9)} ${outcome.version.padEnd(16)} ${checks - failed}/${checks} checks passed`);
  }
  if (failures.length > 0) {
    console.log('\nFAILED:');
    for (const failure of failures) {
      console.log(`  ${failure}`);
    }
    process.exitCode = 1;
    return;
  }
  console.log('\nAll browsers passed.');
}

main().catch(error => {
  console.error(error instanceof UsageError ? `[smoke] ${error.message}\n${HELP.trim()}` : `[smoke] ${error.stack ?? error.message}`);
  process.exitCode = 1;
});
