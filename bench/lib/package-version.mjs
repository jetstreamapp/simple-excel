import { existsSync, readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

/**
 * Version of the package that owns `entrySpecifier`, found by walking up from the resolved entry file
 * to the nearest package.json with the expected name. Needed because packages with an `exports` map
 * (xlsx, @office-kit/xlsx) refuse `require('<pkg>/package.json')`.
 */
export function packageVersion(entrySpecifier, packageName) {
  let directory = dirname(fileURLToPath(import.meta.resolve(entrySpecifier)));
  while (true) {
    const candidate = join(directory, 'package.json');
    if (existsSync(candidate)) {
      const manifest = JSON.parse(readFileSync(candidate, 'utf8'));
      if (manifest.name === packageName) {
        return manifest.version;
      }
    }
    const parent = dirname(directory);
    if (parent === directory) {
      return 'unknown';
    }
    directory = parent;
  }
}
