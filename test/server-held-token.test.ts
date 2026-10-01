/**
 * In HTTP mode many users share one server, and each Graph call must use the
 * bearer token of the request it serves. A token the server holds itself would
 * answer for whoever has no request token: every user would read one account.
 *
 * This fork closes both ways to such a token:
 * - MS365_MCP_OAUTH_TOKEN (upstream's "bring your own token") is not read at
 *   all; setting it only logs a warning.
 * - In plain --http mode the Graph client gets its token from the call or the
 *   request context only, never from AuthManager, whatever AuthManager holds (a
 *   mounted MSAL cache, or a token set through setOAuthToken).
 * stdio mode and --trust-proxy-auth, where one account is the design, keep
 * AuthManager as their source.
 */
import type { Configuration } from '@azure/msal-node';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import AuthManager from '../src/auth.js';
import GraphClient from '../src/graph-client.js';
import logger from '../src/logger.js';
import { graphTokenSource } from '../src/lib/graph-token-source.js';
import { requestContext } from '../src/request-context.js';
import type { TokenCacheStorage } from '../src/token-cache-storage.js';

vi.mock('../src/logger.js', () => ({
  default: {
    info: vi.fn(),
    warn: vi.fn(),
    error: vi.fn(),
    debug: vi.fn(),
  },
}));

const msalConfig: Configuration = {
  auth: {
    clientId: 'test-client',
    authority: 'https://login.microsoftonline.com/common',
  },
};

function storage(): TokenCacheStorage {
  return {
    description: 'mock-storage',
    failClosed: true,
    load: vi.fn().mockResolvedValue(undefined),
    save: vi.fn().mockResolvedValue(undefined),
    delete: vi.fn().mockResolvedValue(undefined),
  };
}

// A real AuthManager with an empty MSAL cache.
function authManagerWithoutAccounts(): AuthManager {
  const auth = new AuthManager(msalConfig, ['User.Read'], undefined, storage());
  const tokenCache = { getAllAccounts: vi.fn().mockResolvedValue([]) };
  Object.assign(auth as unknown as Record<string, unknown>, {
    msalApp: { getTokenCache: () => tokenCache },
  });
  return auth;
}

describe('MS365_MCP_OAUTH_TOKEN', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  afterEach(() => {
    vi.unstubAllEnvs();
  });

  it('gives the server no token of its own', async () => {
    vi.stubEnv('MS365_MCP_OAUTH_TOKEN', 'server-held-token');

    const auth = authManagerWithoutAccounts();

    expect(auth.isOAuthModeEnabled()).toBe(false);
    await expect(auth.getToken()).rejects.toThrow('No valid token found');
    await expect(auth.getTokenForAccount()).rejects.toThrow();
  });

  it('logs a warning that it is ignored', () => {
    vi.stubEnv('MS365_MCP_OAUTH_TOKEN', 'server-held-token');

    authManagerWithoutAccounts();

    expect(logger.warn).toHaveBeenCalledWith(
      expect.stringContaining('MS365_MCP_OAUTH_TOKEN is ignored')
    );
  });

  it('logs nothing about it when it is not set', () => {
    vi.stubEnv('MS365_MCP_OAUTH_TOKEN', undefined);

    authManagerWithoutAccounts();

    expect(logger.warn).not.toHaveBeenCalledWith(expect.stringContaining('MS365_MCP_OAUTH_TOKEN'));
  });
});

describe('graphTokenSource', () => {
  const secrets = { clientId: 'client-id', tenantId: 'common', cloudType: 'global' as const };
  let serverToken: ReturnType<typeof vi.fn>;
  let fetchMock: ReturnType<typeof vi.fn>;

  // An AuthManager that holds a token of its own, as a mounted MSAL cache would.
  function authManagerHoldingAToken(): AuthManager {
    return { getToken: serverToken } as unknown as AuthManager;
  }

  function sentBearer(): string | undefined {
    const init = fetchMock.mock.calls[0]?.[1] as Parameters<typeof fetch>[1];
    return (init?.headers as Record<string, string> | undefined)?.Authorization;
  }

  beforeEach(() => {
    serverToken = vi.fn().mockResolvedValue('server-held-token');
    fetchMock = vi.fn().mockResolvedValue(Response.json({ id: 'me' }));
    vi.stubGlobal('fetch', fetchMock);
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  describe('in plain --http mode', () => {
    const options = { http: '127.0.0.1:3300' };

    it('refuses a Graph call that has no request token, and asks AuthManager nothing', async () => {
      const client = new GraphClient(
        graphTokenSource(authManagerHoldingAToken(), options),
        secrets
      );

      await expect(client.makeRequest('/me')).rejects.toThrow('No access token for this request');
      expect(serverToken).not.toHaveBeenCalled();
      expect(fetchMock).not.toHaveBeenCalled();
    });

    it("sends the request's own bearer token", async () => {
      const client = new GraphClient(
        graphTokenSource(authManagerHoldingAToken(), options),
        secrets
      );

      await requestContext.run({ accessToken: 'user-token' }, () => client.makeRequest('/me'));

      expect(sentBearer()).toBe('Bearer user-token');
    });
  });

  it.each([
    ['stdio mode', {}],
    ['--http with --trust-proxy-auth', { http: '127.0.0.1:3300', trustProxyAuth: true }],
  ])('keeps the server token as the source in %s', async (_mode, options) => {
    const client = new GraphClient(graphTokenSource(authManagerHoldingAToken(), options), secrets);

    await client.makeRequest('/me');

    expect(sentBearer()).toBe('Bearer server-held-token');
  });
});
