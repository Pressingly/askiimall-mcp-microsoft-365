import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
// @ts-expect-error - generator module is plain JS with no type declarations
import { downloadGraphOpenAPI, BETA_OPENAPI_URL } from '../bin/modules/download-openapi.mjs';

// The tool schemas are generated from Microsoft's Graph spec at build time, and the
// generated client is not in git. Downloaded from a branch, the same commit of this
// repo built on two days can give two different tool sets. Downloaded from one
// msgraph-metadata commit, every build gets the same spec until we change that commit.

const PINNED_SPEC =
  /^https:\/\/raw\.githubusercontent\.com\/microsoftgraph\/msgraph-metadata\/([0-9a-f]{40})\/openapi\/(v1\.0|beta)\/openapi\.yaml$/;

describe('Graph spec download', () => {
  let dir: string;
  let requested: string[];

  beforeEach(() => {
    dir = mkdtempSync(path.join(tmpdir(), 'graph-spec-'));
    requested = [];
    vi.spyOn(globalThis, 'fetch').mockImplementation(async (input) => {
      requested.push(input instanceof Request ? input.url : String(input));
      return new Response('openapi: 3.0.4');
    });
    vi.spyOn(console, 'log').mockImplementation(() => {});
  });

  afterEach(() => {
    vi.restoreAllMocks();
    rmSync(dir, { recursive: true, force: true });
  });

  it('downloads the v1.0 and beta specs from one msgraph-metadata commit, not a branch', async () => {
    await downloadGraphOpenAPI(dir, path.join(dir, 'openapi.yaml'));
    await downloadGraphOpenAPI(dir, path.join(dir, 'openapi-beta.yaml'), BETA_OPENAPI_URL);

    expect(requested).toHaveLength(2);
    const [v1, beta] = requested.map((url) => url.match(PINNED_SPEC));
    expect(v1?.[2]).toBe('v1.0');
    expect(beta?.[2]).toBe('beta');
    expect(beta?.[1]).toBe(v1?.[1]);
  });
});
