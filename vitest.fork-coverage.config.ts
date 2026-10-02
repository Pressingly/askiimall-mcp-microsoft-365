/**
 * Upstream's vitest config plus per-file coverage limits, for this fork's CI
 * only (.github/workflows/fork-coverage.yml). `npm test` stays as upstream has
 * it.
 *
 * The limits cover the files our patches live in, about two points under their
 * statement coverage on 2026-10-02. They are a tripwire for upstream merges:
 * one fails when a merge adds untested code to the file or loses our tests.
 * src/server.ts was at 86.4 without the gateway tests (test/gateway-*.test.ts),
 * so its limit of 89 fails if those tests go.
 *
 * There is no limit for the whole of src/, on purpose: it would move with
 * every upstream change to files we never touch.
 */
import { defineConfig, mergeConfig } from 'vitest/config';
import baseConfig from './vitest.config.js';

export default mergeConfig(
  baseConfig,
  defineConfig({
    test: {
      coverage: {
        enabled: true,
        provider: 'v8',
        include: ['src/**'],
        exclude: ['src/generated/**'],
        reporter: ['text-summary'],
        // Still check the limits when a test fails, so a local run on macOS
        // (5 tests in attachment-split-listener need 127.0.0.2) shows them.
        reportOnFailure: true,
        thresholds: {
          'src/server.ts': { statements: 89 },
          'src/lib/microsoft-auth.ts': { statements: 93 },
          'src/graph-client.ts': { statements: 78 },
          'src/graph-tools.ts': { statements: 90 },
        },
      },
    },
  })
);
