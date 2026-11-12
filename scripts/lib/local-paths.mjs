/**
 * Committed oracle and benchmark results must not record who ran them or where: reader errors, stack traces and
 * fixture paths all carry absolute paths, and the repo is public.
 */
import { homedir } from 'node:os';
import { fileURLToPath } from 'node:url';

const REPO_ROOT = fileURLToPath(new URL('../..', import.meta.url)).replace(/[\\/]$/, '');
const HOME = homedir();

/** Replaces this checkout's root with `<repo>` and the home directory with `~`. */
export function scrubLocalPaths(text) {
  return text.replaceAll(REPO_ROOT, '<repo>').replaceAll(HOME, '~');
}

/** `JSON.stringify` with every string value passed through `scrubLocalPaths` (so escaped Windows paths match too). */
export function stringifyScrubbed(value) {
  return JSON.stringify(value, (_key, entry) => (typeof entry === 'string' ? scrubLocalPaths(entry) : entry), 2);
}

/** A machine description that names no one: platform and architecture, never the hostname. */
export function machineLabel() {
  return `${process.platform}-${process.arch}`;
}
