import { fileURLToPath } from 'node:url';
import { defineConfig } from 'vitest/config';

// Two projects: `unit` is the per-module suites co-located under src/**/__tests__; `corpus` is the
// integration layer under test/ (fixture corpus, golden bytes, hostile inputs, SheetJS parity, memory).
// `npm test` runs both; `npm run test:corpus` runs the integration layer alone.
export default defineConfig({
  resolve: {
    // The node entry imports the core by its package name (the build keeps it external); tests run it from source.
    alias: [{ find: /^@jetstreamapp\/simple-excel$/, replacement: fileURLToPath(new URL('./src/index.ts', import.meta.url)) }],
  },
  test: {
    globals: false,
    environment: 'node',
    coverage: {
      reporter: ['text', 'json', 'html'],
      include: ['src/**/*.ts'],
      exclude: ['src/**/__tests__/**'],
    },
    projects: [
      {
        test: {
          name: 'unit',
          include: ['src/**/__tests__/**/*.test.ts'],
          environment: 'node',
          testTimeout: 10000,
        },
      },
      {
        test: {
          name: 'corpus',
          include: ['test/**/*.test.ts'],
          environment: 'node',
          testTimeout: 120000,
          hookTimeout: 120000,
        },
      },
    ],
  },
});
