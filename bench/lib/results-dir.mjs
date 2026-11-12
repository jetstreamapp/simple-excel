import { existsSync, mkdirSync } from 'node:fs';
import os from 'node:os';
import { basename, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { machineLabel } from '../../scripts/lib/local-paths.mjs';

export const RESULTS_DIR = fileURLToPath(new URL('../results/', import.meta.url)).replace(/\/$/, '');

/** `results/<YYYY-MM-DD>-<platform>-<arch>-<label>` (local date), suffixed `-2`, `-3`… when it already exists. */
export function createResultsDir(label) {
  const now = new Date();
  const date = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
  const host = machineLabel();
  const safeLabel = label.toLowerCase().replace(/[^a-z0-9-]+/g, '-');
  let dir = join(RESULTS_DIR, `${date}-${host}-${safeLabel}`);
  for (let suffix = 2; existsSync(dir); suffix++) {
    dir = join(RESULTS_DIR, `${date}-${host}-${safeLabel}-${suffix}`);
  }
  mkdirSync(dir, { recursive: true });
  return { absolute: dir, relative: `bench/results/${basename(dir)}` };
}

export function machineInfo() {
  return {
    platform: os.platform(),
    release: os.release(),
    arch: os.arch(),
    cpuModel: os.cpus()[0]?.model ?? 'unknown',
    cpus: os.cpus().length,
    totalMemGB: Math.round(os.totalmem() / (1024 * 1024 * 1024)),
    nodeVersion: process.version,
    v8: process.versions.v8,
    loadAvgStart: os.loadavg(),
  };
}
