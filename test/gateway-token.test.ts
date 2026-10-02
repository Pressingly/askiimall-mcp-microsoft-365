/**
 * An MCP gateway (mcpo) signs its users in through /authorize and /token, and
 * refreshes their tokens through /token for as long as they stay connected.
 * What it gets back decides what it stores and whether the user stays signed
 * in:
 *
 * - The code exchange: mcpo stores the tokens and names the user from the
 *   id token.
 * - A refresh: mcpo stores the new access token and the new refresh token.
 *   When the refresh fails, mcpo reads the status: a 5xx or a 429 keeps the
 *   sign-in and is tried again later; any other 4xx marks the sign-in failed,
 *   and the user has to connect again (oauth_client.py, refresh_access_token).
 *
 * mcpo sends these as a public client: the client_id it got from /register,
 * no secret. This server swaps in its own Microsoft app.
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
const CLIENT_SECRET = 'app-secret';
const GATEWAY_CLIENT_ID = 'mcp-client-1700000000000';
const CALLBACK = 'https://mcpo.example.com/oauth/callback';
const MICROSOFT_TOKEN_URL = 'https://login.microsoftonline.com/common/oauth2/v2.0/token';

// What Microsoft answers to a code exchange with openid in the scope.
const MICROSOFT_CODE_ANSWER = {
  token_type: 'Bearer',
  scope: 'Files.ReadWrite User.Read openid profile email',
  expires_in: 3599,
  ext_expires_in: 3599,
  access_token: 'graph-access-1',
  refresh_token: 'refresh-1',
  id_token: 'header.eyJuYW1lIjoiRHp1bmcifQ.signature',
};

// What Microsoft answers to a refresh: a new access token and a new refresh token.
const MICROSOFT_REFRESH_ANSWER = {
  token_type: 'Bearer',
  scope: 'Files.ReadWrite User.Read openid profile email',
  expires_in: 3599,
  ext_expires_in: 3599,
  access_token: 'graph-access-2',
  refresh_token: 'refresh-2',
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

describe('/token and /authorize as an MCP gateway uses them', () => {
  const savedEnv = { ...process.env };
  const realFetch = globalThis.fetch;
  let server: MicrosoftGraphServer;
  let origin: string;
  // Each form posted to Microsoft's token endpoint, and how Microsoft answers.
  let microsoftReceived: URLSearchParams[];
  let microsoftAnswers: () => Promise<Response>;

  beforeEach(async () => {
    clearSecretsCache();
    process.env.MS365_MCP_CLIENT_ID = CLIENT_ID;
    process.env.MS365_MCP_CLIENT_SECRET = CLIENT_SECRET;
    process.env.MS365_MCP_TENANT_ID = 'common';
    delete process.env.MS365_MCP_ALLOWED_REDIRECT_URIS;

    microsoftReceived = [];
    microsoftAnswers = async () => Response.json({ error: 'not_set_up' }, { status: 500 });
    vi.spyOn(globalThis, 'fetch').mockImplementation(async (input, init) => {
      const url = input instanceof Request ? input.url : String(input);
      if (url === MICROSOFT_TOKEN_URL) {
        microsoftReceived.push(new URLSearchParams(init?.body as URLSearchParams));
        return microsoftAnswers();
      }
      return realFetch(input, init);
    });

    const port = await freePort();
    server = new MicrosoftGraphServer(fakeAuthManager(), {
      http: `127.0.0.1:${port}`,
      orgMode: true,
      enableDynamicRegistration: true,
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

  // The form mcpo posts, as a public client (oauth_client.py).
  function postToken(form: Record<string, string>): Promise<Response> {
    return realFetch(`${origin}/token`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        Accept: 'application/json',
      },
      body: new URLSearchParams({ ...form, client_id: GATEWAY_CLIENT_ID }),
    });
  }

  function refresh(): Promise<Response> {
    return postToken({ grant_type: 'refresh_token', refresh_token: 'refresh-1' });
  }

  describe('refresh', () => {
    it("redeems the gateway's refresh token with this server's app and returns Microsoft's answer as it is", async () => {
      microsoftAnswers = async () => Response.json(MICROSOFT_REFRESH_ANSWER);

      const res = await refresh();

      expect(res.status).toBe(200);
      expect(await res.json()).toEqual(MICROSOFT_REFRESH_ANSWER);
      expect(microsoftReceived).toHaveLength(1);
      const sent = microsoftReceived[0];
      expect(sent.get('grant_type')).toBe('refresh_token');
      expect(sent.get('refresh_token')).toBe('refresh-1');
      expect(sent.get('client_id')).toBe(CLIENT_ID);
      expect(sent.get('client_secret')).toBe(CLIENT_SECRET);
    });

    it('answers a refresh token Microsoft rejects with 400 invalid_grant, so the gateway asks the user to connect again', async () => {
      microsoftAnswers = async () =>
        Response.json(
          {
            error: 'invalid_grant',
            error_description:
              'AADSTS70008: The provided authorization code or refresh token has expired.',
            error_codes: [70008],
          },
          { status: 400 }
        );

      const res = await refresh();

      expect(res.status).toBe(400);
      expect(await res.json()).toMatchObject({ error: 'invalid_grant' });
    });

    it.each([
      [
        'a 5xx that is not JSON',
        async () => new Response('<html>502 Bad Gateway</html>', { status: 502 }),
      ],
      [
        'no answer at all',
        async (): Promise<Response> => {
          throw new TypeError('fetch failed');
        },
      ],
    ])(
      'answers %s from Microsoft with 500, so the gateway keeps the sign-in',
      async (_, answer) => {
        microsoftAnswers = answer;

        const res = await refresh();

        expect(res.status).toBe(500);
        expect(await res.json()).toMatchObject({ error: 'server_error' });
      }
    );
  });

  it("returns Microsoft's code exchange answer as it is, id token included", async () => {
    microsoftAnswers = async () => Response.json(MICROSOFT_CODE_ANSWER);
    const verifier = crypto.randomBytes(32).toString('base64url');

    const res = await postToken({
      grant_type: 'authorization_code',
      code: 'code-1',
      redirect_uri: CALLBACK,
      code_verifier: verifier,
    });

    expect(res.status).toBe(200);
    expect(await res.json()).toEqual(MICROSOFT_CODE_ANSWER);
    const sent = microsoftReceived[0];
    expect(sent.get('client_id')).toBe(CLIENT_ID);
    expect(sent.get('code')).toBe('code-1');
    expect(sent.get('redirect_uri')).toBe(CALLBACK);
    expect(sent.get('code_verifier')).toBe(verifier);
  });

  describe('/authorize with MS365_MCP_ALLOWED_REDIRECT_URIS set to the gateway callback', () => {
    function authorize(redirectUri: string): Promise<Response> {
      const url = new URL(`${origin}/authorize`);
      url.search = new URLSearchParams({
        response_type: 'code',
        client_id: GATEWAY_CLIENT_ID,
        redirect_uri: redirectUri,
        scope: 'Files.ReadWrite',
        state: 'state-1',
        code_challenge: crypto
          .createHash('sha256')
          .update(crypto.randomBytes(32).toString('base64url'))
          .digest('base64url'),
        code_challenge_method: 'S256',
      }).toString();
      return realFetch(url, { redirect: 'manual' });
    }

    beforeEach(() => {
      process.env.MS365_MCP_ALLOWED_REDIRECT_URIS = CALLBACK;
    });

    it('sends the gateway callback on to Microsoft', async () => {
      const res = await authorize(CALLBACK);

      expect(res.status).toBe(302);
      const location = new URL(res.headers.get('location')!);
      expect(location.origin).toBe('https://login.microsoftonline.com');
      expect(location.searchParams.get('redirect_uri')).toBe(CALLBACK);
    });

    it('refuses any other https callback and does not redirect', async () => {
      const res = await authorize('https://attacker.example.com/oauth/callback');

      expect(res.status).toBe(400);
      expect(res.headers.get('location')).toBeNull();
      expect(await res.json()).toMatchObject({ error: 'invalid_request' });
    });
  });
});
