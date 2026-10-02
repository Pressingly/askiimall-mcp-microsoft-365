/**
 * An MCP gateway (mcpo) calls a tool for one of its users by posting a
 * JSON-RPC tools/call to /mcp with that user's Microsoft token as the bearer.
 * mcpo reads the answer in two ways (session_manager.py, gateway.py):
 *
 * - HTTP 401 means the token was rejected. mcpo expires the stored sign-in,
 *   and the user's next status check refreshes it or asks them to connect.
 * - A tool error inside HTTP 200 is shown as an error. mcpo reports it as an
 *   authentication error when its text says so ("unauthorized", among
 *   others), but keeps the sign-in.
 *
 * Many users share one server, so every Graph call must carry the token of
 * the request it serves, and no other.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { createServer, type Server } from 'node:http';
import type { AddressInfo } from 'node:net';
import MicrosoftGraphServer from '../src/server.js';
import type AuthManager from '../src/auth.js';
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

const GRAPH_DRIVES_URL = 'https://graph.microsoft.com/v1.0/me/drives';

// A Microsoft access token for a work account is a JWT; only its exp matters here.
function jwt(claims: Record<string, unknown>): string {
  const part = (value: object) => Buffer.from(JSON.stringify(value)).toString('base64url');
  return `${part({ typ: 'JWT', alg: 'RS256' })}.${part(claims)}.signature`;
}

const inAnHour = () => Math.floor(Date.now() / 1000) + 3600;
const TOKEN_A = jwt({ oid: 'user-a', exp: inAnHour() });
const TOKEN_B = jwt({ oid: 'user-b', exp: inAnHour() });

const DRIVES: Record<string, { value: { id: string; name: string }[] }> = {
  [TOKEN_A]: { value: [{ id: 'drive-a', name: 'OneDrive of user A' }] },
  [TOKEN_B]: { value: [{ id: 'drive-b', name: 'OneDrive of user B' }] },
};

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

function deferred() {
  let resolve!: () => void;
  const promise = new Promise<void>((res) => {
    resolve = res;
  });
  return { promise, resolve };
}

describe('/mcp tools/call as an MCP gateway makes it', () => {
  const savedEnv = { ...process.env };
  const realFetch = globalThis.fetch;
  let server: MicrosoftGraphServer;
  let origin: string;
  // The bearer token of each request that reached Graph, and how Graph answers.
  let graphReceived: string[];
  let graphAnswers: (token: string) => Promise<Response>;

  beforeEach(async () => {
    clearSecretsCache();
    process.env.MS365_MCP_CLIENT_ID = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee';
    process.env.MS365_MCP_CLIENT_SECRET = 'app-secret';
    process.env.MS365_MCP_TENANT_ID = 'common';

    graphReceived = [];
    graphAnswers = async (token) => Response.json(DRIVES[token]);
    vi.spyOn(globalThis, 'fetch').mockImplementation(async (input, init) => {
      const url = input instanceof Request ? input.url : String(input);
      if (url.startsWith(GRAPH_DRIVES_URL)) {
        const authorization = (init?.headers as Record<string, string>).Authorization;
        const token = authorization.replace(/^Bearer /, '');
        graphReceived.push(token);
        return graphAnswers(token);
      }
      return realFetch(input, init);
    });

    const port = await freePort();
    server = new MicrosoftGraphServer(fakeAuthManager(), {
      http: `127.0.0.1:${port}`,
      orgMode: true,
      enabledTools: '^list-drives$',
    });
    await server.initialize('0.0.0-test');
    await server.start();
    origin = `http://127.0.0.1:${port}`;
  });

  afterEach(async () => {
    vi.restoreAllMocks();
    await server.stop();
    clearSecretsCache();
    process.env = { ...savedEnv };
  });

  function callListDrives(bearer?: string): Promise<Response> {
    return realFetch(`${origin}/mcp`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Accept: 'application/json, text/event-stream',
        ...(bearer ? { Authorization: `Bearer ${bearer}` } : {}),
      },
      body: JSON.stringify({
        jsonrpc: '2.0',
        id: 1,
        method: 'tools/call',
        params: { name: 'list-drives', arguments: {} },
      }),
    });
  }

  async function toolResult(res: Response) {
    const body = (await res.json()) as {
      result: { content: { type: string; text: string }[]; isError?: boolean };
    };
    return body.result;
  }

  it('refuses a call with no bearer token with 401 and never calls Graph', async () => {
    const res = await callListDrives();

    expect(res.status).toBe(401);
    expect(await res.json()).toMatchObject({ error: 'invalid_token' });
    expect(graphReceived).toEqual([]);
  });

  it('refuses an expired token with 401 and never calls Graph, so the gateway refreshes it', async () => {
    const expired = jwt({ oid: 'user-a', exp: Math.floor(Date.now() / 1000) - 60 });

    const res = await callListDrives(expired);

    expect(res.status).toBe(401);
    expect(await res.json()).toMatchObject({ error: 'invalid_token' });
    expect(graphReceived).toEqual([]);
  });

  it("sends each caller's own token to Graph, also for two calls at the same time", async () => {
    // Graph holds user A's answer until user B's request has reached it, so the
    // two calls are in the server together.
    const bArrived = deferred();
    graphAnswers = async (token) => {
      if (token === TOKEN_A) await bArrived.promise;
      if (token === TOKEN_B) bArrived.resolve();
      return Response.json(DRIVES[token]);
    };

    const [resA, resB] = await Promise.all([callListDrives(TOKEN_A), callListDrives(TOKEN_B)]);

    expect(graphReceived.sort()).toEqual([TOKEN_A, TOKEN_B].sort());
    const [resultA, resultB] = [await toolResult(resA), await toolResult(resB)];
    expect(resultA.isError).toBeFalsy();
    expect(resultA.content[0].text).toContain('drive-a');
    expect(resultA.content[0].text).not.toContain('drive-b');
    expect(resultB.isError).toBeFalsy();
    expect(resultB.content[0].text).toContain('drive-b');
    expect(resultB.content[0].text).not.toContain('drive-a');
  });

  it('answers a Graph 401 with a tool error that says Unauthorized, inside HTTP 200', async () => {
    // Graph's own body names no word the gateway looks for; the status text
    // "Unauthorized" is what it matches on.
    graphAnswers = async () =>
      Response.json(
        {
          error: {
            code: 'InvalidAuthenticationToken',
            message: 'Access token has expired or is not yet valid.',
          },
        },
        { status: 401, statusText: 'Unauthorized' }
      );

    const res = await callListDrives(TOKEN_A);

    expect(res.status).toBe(200);
    const result = await toolResult(res);
    expect(result.isError).toBe(true);
    expect(result.content[0].text.toLowerCase()).toContain('unauthorized');
    expect(graphReceived).toEqual([TOKEN_A]);
  });
});
