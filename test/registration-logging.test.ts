/**
 * In HTTP mode every POST /mcp builds its own McpServer and registers the
 * tools again. Registration used to log at info for each tool it left out, so
 * with --enabled-tools one tools/list wrote a line per skipped tool (292 with
 * the OneDrive tool list), plus the whole filter pattern and a summary. Those
 * lines are at debug now: a request writes none at info, and LOG_LEVEL=debug
 * shows them all.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { createServer, type Server } from 'node:http';
import type { AddressInfo } from 'node:net';
import MicrosoftGraphServer from '../src/server.js';
import type AuthManager from '../src/auth.js';
import logger from '../src/logger.js';
import { clearSecretsCache } from '../src/secrets.js';

vi.mock('../src/logger.js', () => ({
  default: {
    info: vi.fn(),
    warn: vi.fn(),
    error: vi.fn(),
    debug: vi.fn(),
    verbose: vi.fn(),
  },
  enableConsoleLogging: vi.fn(),
}));

async function freePort(): Promise<number> {
  const holder = await new Promise<Server>((resolve) => {
    const s = createServer();
    s.listen(0, '127.0.0.1', () => resolve(s));
  });
  const port = (holder.address() as AddressInfo).port;
  await new Promise<void>((resolve) => holder.close(() => resolve()));
  return port;
}

function fakeAuthManager(): AuthManager {
  return {
    isOAuthModeEnabled: () => false,
    isMultiAccount: async () => false,
    listAccounts: async () => [],
    hasExpectedAccount: () => false,
  } as unknown as AuthManager;
}

describe('tool registration on a /mcp request', () => {
  const savedEnv = { ...process.env };
  let server: MicrosoftGraphServer;
  let base: string;

  beforeEach(async () => {
    clearSecretsCache();
    process.env.MS365_MCP_CLIENT_ID = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee';
    process.env.MS365_MCP_CLIENT_SECRET = 'secret';
    process.env.MS365_MCP_TENANT_ID = 'common';

    const port = await freePort();
    server = new MicrosoftGraphServer(fakeAuthManager(), {
      http: `127.0.0.1:${port}`,
      orgMode: true,
      enabledTools: '^list-drives$',
    });
    await server.initialize('0.0.0-test');
    await server.start();
    base = `http://127.0.0.1:${port}`;
  });

  afterEach(async () => {
    await server.stop();
    clearSecretsCache();
    process.env = { ...savedEnv };
  });

  it('writes no info line for tools/list, and keeps the registration summary at debug', async () => {
    vi.mocked(logger.info).mockClear();
    vi.mocked(logger.debug).mockClear();

    const response = await fetch(`${base}/mcp`, {
      method: 'POST',
      headers: {
        Authorization: 'Bearer a-graph-token',
        'Content-Type': 'application/json',
        Accept: 'application/json, text/event-stream',
      },
      body: JSON.stringify({ jsonrpc: '2.0', id: 1, method: 'tools/list', params: {} }),
    });

    expect(response.status).toBe(200);
    const { result } = await response.json();
    expect(result.tools.map((tool: { name: string }) => tool.name)).toEqual(['list-drives']);
    expect(logger.info).not.toHaveBeenCalled();
    expect(logger.debug).toHaveBeenCalledWith(expect.stringMatching(/\b1 registered\b/));
  });
});
