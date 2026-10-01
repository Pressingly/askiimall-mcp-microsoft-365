/**
 * An MCP gateway (mcpo) adds this server as an OAuth service in two calls, and
 * stores what they return for the life of the service:
 *
 * - Discovery. mcpo reads the protected-resource document (the /mcp form
 *   first, then the bare one) and the authorization server metadata. It
 *   refuses a server with no registration_endpoint, turns PKCE on when S256 is
 *   listed, and stores scopes_supported as the scope string it asks for at
 *   every sign-in and refresh, taken from the protected-resource document
 *   first.
 * - Registration. mcpo posts its client metadata to /register as a public
 *   client and keeps the client_id and token_endpoint_auth_method it gets back.
 *
 * Production runs several copies of the server behind one address, so
 * registration must leave nothing behind that /authorize on another copy needs.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import crypto from 'node:crypto';
import { createServer, type Server } from 'node:http';
import type { AddressInfo } from 'node:net';
import MicrosoftGraphServer from '../src/server.js';
import type AuthManager from '../src/auth.js';
import type { CommandOptions } from '../src/cli.js';
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
const PUBLIC_URL = 'https://mcp-microsoft-365.example.com';

// The body askiimall-mcpo sends (oauth_client.py, _register_client_via_dcr).
const MCPO_REGISTRATION = {
  client_name: 'AskiiMall',
  client_uri: 'https://askiimall.com',
  redirect_uris: [CALLBACK],
  token_endpoint_auth_method: 'none',
  grant_types: ['authorization_code', 'refresh_token'],
  response_types: ['code'],
  scope: 'Mail.Read',
};

// One tool, so the advertised scopes are known: endpoints.json gives
// list-mail-messages the scope Mail.Read.
const OPTIONS = {
  orgMode: true,
  enableDynamicRegistration: true,
  enabledTools: '^list-mail-messages$',
} as const;

const OIDC_AND_OFFLINE = ['openid', 'profile', 'email', 'offline_access'];

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

describe('discovery and registration as an MCP gateway uses them', () => {
  const savedEnv = { ...process.env };
  let copies: MicrosoftGraphServer[] = [];

  async function startCopy(options: Partial<CommandOptions> = {}): Promise<string> {
    const port = await freePort();
    const server = new MicrosoftGraphServer(fakeAuthManager(), {
      http: `127.0.0.1:${port}`,
      ...OPTIONS,
      ...options,
    });
    await server.initialize('0.0.0-test');
    await server.start();
    copies.push(server);
    return `http://127.0.0.1:${port}`;
  }

  async function getJson(url: string): Promise<Record<string, unknown>> {
    const res = await fetch(url);
    expect(res.status).toBe(200);
    return (await res.json()) as Record<string, unknown>;
  }

  beforeEach(() => {
    clearSecretsCache();
    process.env.MS365_MCP_CLIENT_ID = CLIENT_ID;
    process.env.MS365_MCP_CLIENT_SECRET = 'secret';
    process.env.MS365_MCP_TENANT_ID = 'common';
    delete process.env.MS365_MCP_PUBLIC_URL;
    delete process.env.MS365_MCP_BASE_URL;
    delete process.env.MS365_MCP_ALLOWED_REDIRECT_URIS;
  });

  afterEach(async () => {
    await Promise.all(copies.map((server) => server.stop()));
    copies = [];
    clearSecretsCache();
    process.env = { ...savedEnv };
  });

  it('the server metadata has every field the gateway needs, on the server address', async () => {
    const origin = await startCopy();

    const metadata = await getJson(`${origin}/.well-known/oauth-authorization-server`);

    expect(metadata).toMatchObject({
      issuer: origin,
      authorization_endpoint: `${origin}/authorize`,
      token_endpoint: `${origin}/token`,
      registration_endpoint: `${origin}/register`,
    });
    expect(metadata.code_challenge_methods_supported).toContain('S256');
  });

  it.each(['/.well-known/oauth-protected-resource/mcp', '/.well-known/oauth-protected-resource'])(
    'the protected-resource document at %s names /mcp and this server',
    async (path) => {
      const origin = await startCopy();

      const resource = await getJson(`${origin}${path}`);

      expect(resource).toMatchObject({
        resource: `${origin}/mcp`,
        authorization_servers: [origin],
      });
    }
  );

  it('both documents advertise the tool scopes and leave the sign-in scopes to /authorize', async () => {
    const origin = await startCopy();

    const resource = await getJson(`${origin}/.well-known/oauth-protected-resource/mcp`);
    const metadata = await getJson(`${origin}/.well-known/oauth-authorization-server`);

    // The gateway stores this list as its scope string: the resource document's
    // copy wins, so the two must agree.
    expect(resource.scopes_supported).toEqual(metadata.scopes_supported);
    expect(resource.scopes_supported).toContain('Mail.Read');
    // /authorize adds these to every sign-in. Listing them here would change the
    // scope string of every service registered from then on.
    for (const scope of OIDC_AND_OFFLINE) {
      expect(resource.scopes_supported).not.toContain(scope);
    }
  });

  it('with MS365_MCP_PUBLIC_URL set, every advertised URL uses it', async () => {
    process.env.MS365_MCP_PUBLIC_URL = `${PUBLIC_URL}/`;
    const origin = await startCopy();

    const metadata = await getJson(`${origin}/.well-known/oauth-authorization-server`);
    const resource = await getJson(`${origin}/.well-known/oauth-protected-resource/mcp`);

    expect(metadata).toMatchObject({
      issuer: PUBLIC_URL,
      authorization_endpoint: `${PUBLIC_URL}/authorize`,
      token_endpoint: `${PUBLIC_URL}/token`,
      registration_endpoint: `${PUBLIC_URL}/register`,
    });
    expect(resource).toMatchObject({
      resource: `${PUBLIC_URL}/mcp`,
      authorization_servers: [PUBLIC_URL],
    });
  });

  it("registers the gateway's metadata as a public client", async () => {
    const origin = await startCopy();

    const res = await fetch(`${origin}/register`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', Accept: 'application/json' },
      body: JSON.stringify(MCPO_REGISTRATION),
    });

    expect(res.status).toBe(201);
    const client = (await res.json()) as Record<string, unknown>;
    expect(typeof client.client_id).toBe('string');
    expect(client.client_id).not.toBe('');
    expect(client.token_endpoint_auth_method).toBe('none');
  });

  it('a client registered on one copy can sign in through another', async () => {
    const copyA = await startCopy();
    const copyB = await startCopy();

    const registered = await fetch(`${copyA}/register`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(MCPO_REGISTRATION),
    });
    const { client_id } = (await registered.json()) as { client_id: string };

    const challenge = crypto
      .createHash('sha256')
      .update(crypto.randomBytes(32).toString('base64url'))
      .digest('base64url');
    const authorize = new URL(`${copyB}/authorize`);
    authorize.search = new URLSearchParams({
      response_type: 'code',
      client_id,
      redirect_uri: CALLBACK,
      scope: 'Mail.Read',
      state: 'state-1',
      code_challenge: challenge,
      code_challenge_method: 'S256',
    }).toString();
    const res = await fetch(authorize, { redirect: 'manual' });

    expect(res.status).toBe(302);
    const location = new URL(res.headers.get('location')!);
    expect(location.origin + location.pathname).toBe(
      'https://login.microsoftonline.com/common/oauth2/v2.0/authorize'
    );
    expect(location.searchParams.get('client_id')).toBe(CLIENT_ID);
    expect(location.searchParams.get('redirect_uri')).toBe(CALLBACK);
  });
});
