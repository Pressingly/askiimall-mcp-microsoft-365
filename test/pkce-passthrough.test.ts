/**
 * An MCP gateway (mcpo) signs its users in as a standard OAuth client: an S256
 * PKCE challenge on /authorize, the matching verifier on /token. If /authorize
 * keeps that challenge in memory and sends Microsoft a challenge of its own,
 * only the copy of the server that answered /authorize can finish the sign-in.
 * With two copies behind one service, /token reaches the other copy about every
 * second time; a restart between the two steps empties the memory the same
 * way. That copy sends the client's verifier, Microsoft holds the server's
 * challenge, and answers AADSTS501481.
 *
 * So a standard challenge goes to Microsoft as it is, and Microsoft checks the
 * client's verifier itself. Any other challenge still gets the two-leg mapping
 * (upstream issue #266).
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import crypto from 'node:crypto';
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
const MICROSOFT_TOKEN_URL = 'https://login.microsoftonline.com/common/oauth2/v2.0/token';

const s256 = (verifier: string) => crypto.createHash('sha256').update(verifier).digest('base64url');

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

/**
 * Microsoft's side of the sign-in: a code is bound to the challenge that came
 * with /authorize, and redeeming it needs the verifier for that challenge.
 */
class FakeMicrosoft {
  private challengeByCode = new Map<string, string | null>();

  // The user signs in at the URL /authorize redirected to; Microsoft issues a code.
  signIn(authorizeRedirect: URL): string {
    const code = `code-${this.challengeByCode.size + 1}`;
    this.challengeByCode.set(code, authorizeRedirect.searchParams.get('code_challenge'));
    return code;
  }

  redeem(form: URLSearchParams): Response {
    const challenge = this.challengeByCode.get(form.get('code') ?? '');
    const verifier = form.get('code_verifier');
    if (challenge && verifier && s256(verifier) === challenge) {
      return Response.json({
        access_token: 'graph-access-token',
        token_type: 'Bearer',
        scope: 'Files.ReadWrite User.Read',
        expires_in: 3599,
        refresh_token: 'graph-refresh-token',
      });
    }
    return Response.json(
      {
        error: 'invalid_grant',
        error_description:
          'AADSTS501481: The Code_Verifier does not match the code_challenge supplied in the authorization request.',
        error_codes: [501481],
      },
      { status: 400 }
    );
  }
}

describe('PKCE between an MCP gateway and Microsoft', () => {
  const savedEnv = { ...process.env };
  const realFetch = globalThis.fetch;
  let copies: MicrosoftGraphServer[] = [];
  let copyA: string;
  let copyB: string;
  let microsoft: FakeMicrosoft;

  async function startCopy(): Promise<string> {
    const port = await freePort();
    const server = new MicrosoftGraphServer(fakeAuthManager(), { http: `127.0.0.1:${port}` });
    await server.initialize('0.0.0-test');
    await server.start();
    copies.push(server);
    return `http://127.0.0.1:${port}`;
  }

  beforeEach(async () => {
    clearSecretsCache();
    process.env.MS365_MCP_CLIENT_ID = CLIENT_ID;
    process.env.MS365_MCP_CLIENT_SECRET = 'secret';
    process.env.MS365_MCP_TENANT_ID = 'common';

    microsoft = new FakeMicrosoft();
    vi.spyOn(globalThis, 'fetch').mockImplementation(async (input, init) => {
      const url = input instanceof Request ? input.url : String(input);
      if (url === MICROSOFT_TOKEN_URL) {
        return microsoft.redeem(new URLSearchParams(init?.body as URLSearchParams));
      }
      return realFetch(input, init);
    });

    copyA = await startCopy();
    copyB = await startCopy();
  });

  afterEach(async () => {
    vi.restoreAllMocks();
    await Promise.all(copies.map((server) => server.stop()));
    copies = [];
    clearSecretsCache();
    process.env = { ...savedEnv };
  });

  async function authorize(
    base: string,
    pkce: { code_challenge: string; code_challenge_method?: string }
  ): Promise<URL> {
    const params = new URLSearchParams({
      response_type: 'code',
      client_id: 'mcpo-client',
      redirect_uri: CALLBACK,
      scope: 'Files.ReadWrite',
      state: crypto.randomBytes(32).toString('base64url'),
      ...pkce,
    });
    const response = await fetch(`${base}/authorize?${params}`, { redirect: 'manual' });
    expect(response.status).toBe(302);
    return new URL(response.headers.get('location')!);
  }

  async function token(base: string, code: string, verifier: string): Promise<Response> {
    return fetch(`${base}/token`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        grant_type: 'authorization_code',
        code,
        redirect_uri: CALLBACK,
        client_id: 'mcpo-client',
        code_verifier: verifier,
      }),
    });
  }

  it('sends a standard S256 challenge to Microsoft as it is', async () => {
    const challenge = s256(crypto.randomBytes(32).toString('base64url'));

    const redirect = await authorize(copyA, {
      code_challenge: challenge,
      code_challenge_method: 'S256',
    });

    expect(redirect.searchParams.get('code_challenge')).toBe(challenge);
    expect(redirect.searchParams.get('code_challenge_method')).toBe('S256');
  });

  it('lets another copy of the server finish the sign-in', async () => {
    const verifier = crypto.randomBytes(32).toString('base64url');
    const redirect = await authorize(copyA, {
      code_challenge: s256(verifier),
      code_challenge_method: 'S256',
    });
    const code = microsoft.signIn(redirect);

    const response = await token(copyB, code, verifier);

    expect(response.status).toBe(200);
    expect((await response.json()).access_token).toBe('graph-access-token');
  });

  const standard = s256('a verifier');
  it.each([
    ['the plain method', { code_challenge: standard, code_challenge_method: 'plain' }],
    ['no method (plain by default)', { code_challenge: standard }],
    ['42 characters', { code_challenge: standard.slice(1), code_challenge_method: 'S256' }],
    ['44 characters', { code_challenge: `${standard}A`, code_challenge_method: 'S256' }],
    [
      'base64 rather than base64url',
      { code_challenge: `${standard.slice(2)}+/`, code_challenge_method: 'S256' },
    ],
  ])('keeps the two-leg mapping for a challenge with %s', async (_label, pkce) => {
    const redirect = await authorize(copyA, pkce);

    expect(redirect.searchParams.get('code_challenge')).not.toBe(pkce.code_challenge);
    expect(redirect.searchParams.get('code_challenge_method')).toBe('S256');
  });
});
