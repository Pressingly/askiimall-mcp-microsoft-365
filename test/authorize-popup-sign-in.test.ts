/**
 * A web client (an MCP gateway's portal, a chat app) runs the sign-in in a
 * popup: it opens /authorize with window.open, Microsoft redirects the popup to
 * the client's callback, and the client watches popup.closed to know when the
 * user gave up. Two things on /authorize break that flow:
 *
 * - Cross-Origin-Opener-Policy: same-origin (helmet's default) on the
 *   /authorize response puts the popup in a new browsing context group, which
 *   cuts it from its opener. popup.closed then reads true a second after the
 *   popup opens, and the client reports a cancelled sign-in while the user is
 *   still typing their password.
 * - Without openid on the Microsoft request no id token comes back, so the
 *   client cannot tell who signed in (name, email, object id, tenant id) unless
 *   it happened to put the OIDC scopes in its own scope string.
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

const CLIENT_ID = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee';
const CALLBACK = 'http://localhost:8010/oauth/callback';

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

describe('/authorize in a sign-in popup', () => {
  const savedEnv = { ...process.env };
  let server: MicrosoftGraphServer | undefined;
  let base: string;

  beforeEach(async () => {
    clearSecretsCache();
    process.env.MS365_MCP_CLIENT_ID = CLIENT_ID;
    process.env.MS365_MCP_CLIENT_SECRET = 'secret';
    process.env.MS365_MCP_TENANT_ID = 'common';
    const port = await freePort();
    base = `http://127.0.0.1:${port}`;
    server = new MicrosoftGraphServer(fakeAuthManager(), { http: `127.0.0.1:${port}` });
    await server.initialize('0.0.0-test');
    await server.start();
  });

  afterEach(async () => {
    await server?.stop();
    server = undefined;
    clearSecretsCache();
    process.env = { ...savedEnv };
  });

  async function authorize(scope: string): Promise<Response> {
    const params = new URLSearchParams({
      response_type: 'code',
      client_id: 'mcp-client-1',
      redirect_uri: CALLBACK,
      state: 'abcdefghij',
      scope,
    });
    return fetch(`${base}/authorize?${params}`, { redirect: 'manual' });
  }

  it('sends no Cross-Origin-Opener-Policy, so the popup keeps its opener', async () => {
    const response = await authorize('Files.Read');

    expect(response.status).toBe(302);
    expect(response.headers.get('cross-origin-opener-policy')).toBeNull();
  });

  it('keeps Cross-Origin-Opener-Policy on the other routes', async () => {
    const response = await fetch(`${base}/.well-known/oauth-authorization-server`);

    expect(response.headers.get('cross-origin-opener-policy')).toBe('same-origin');
  });

  it('asks Microsoft for an id token with name and email, whatever scope the client sends', async () => {
    const response = await authorize('Files.Read User.Read');

    const location = new URL(response.headers.get('location')!);
    const scopes = (location.searchParams.get('scope') ?? '').split(' ');
    expect(scopes).toEqual(
      expect.arrayContaining(['Files.Read', 'User.Read', 'openid', 'profile', 'email'])
    );
  });
});
