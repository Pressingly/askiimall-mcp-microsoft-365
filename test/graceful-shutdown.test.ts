/**
 * A rollout sends the server SIGTERM and waits a grace period before SIGKILL.
 * In a container the server is process 1, and Linux delivers SIGTERM to
 * process 1 only when the program handles it, so without a handler the server
 * kept taking requests until SIGKILL cut every one still running.
 *
 * Now SIGTERM, or SIGINT from Ctrl-C, closes the listeners, lets the requests
 * in flight finish, and exits. A second signal exits at once.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { EventEmitter } from 'node:events';
import { createServer, request, type Server } from 'node:http';
import type { AddressInfo } from 'node:net';
import MicrosoftGraphServer from '../src/server.js';
import type AuthManager from '../src/auth.js';
import { clearSecretsCache } from '../src/secrets.js';
import { installGracefulShutdown } from '../src/lib/graceful-shutdown.js';

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

function deferred<T = void>() {
  let resolve!: (value: T) => void;
  let reject!: (reason: unknown) => void;
  const promise = new Promise<T>((res, rej) => {
    resolve = res;
    reject = rej;
  });
  return { promise, resolve, reject };
}

describe('graceful shutdown on a signal', () => {
  it.each(['SIGTERM', 'SIGINT'] as const)(
    'on %s stops the server and exits 0 once stop has finished',
    async (signal) => {
      const signals = new EventEmitter();
      const exit = vi.fn();
      const stopped = deferred();
      const server = { stop: vi.fn(() => stopped.promise) };
      installGracefulShutdown(server, { signals, exit });

      signals.emit(signal, signal);

      expect(server.stop).toHaveBeenCalledTimes(1);
      await Promise.resolve();
      expect(exit).not.toHaveBeenCalled();

      stopped.resolve();
      await vi.waitFor(() => expect(exit).toHaveBeenCalledWith(0));
    }
  );

  it('exits 1 at once on a second signal, without waiting for stop', () => {
    const signals = new EventEmitter();
    const exit = vi.fn();
    const server = { stop: vi.fn(() => new Promise<void>(() => {})) };
    installGracefulShutdown(server, { signals, exit });

    signals.emit('SIGTERM', 'SIGTERM');
    signals.emit('SIGINT', 'SIGINT');

    expect(server.stop).toHaveBeenCalledTimes(1);
    expect(exit).toHaveBeenCalledWith(1);
  });

  it('exits 1 when stop fails', async () => {
    const signals = new EventEmitter();
    const exit = vi.fn();
    const server = { stop: vi.fn(async () => Promise.reject(new Error('close failed'))) };
    installGracefulShutdown(server, { signals, exit });

    signals.emit('SIGTERM', 'SIGTERM');

    await vi.waitFor(() => expect(exit).toHaveBeenCalledWith(1));
    expect(exit).not.toHaveBeenCalledWith(0);
  });
});

const CLIENT_ID = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee';
const MICROSOFT_TOKEN_URL = 'https://login.microsoftonline.com/common/oauth2/v2.0/token';

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
 * POST /token on a connection of its own that closes after the response
 * (agent: false), so the listener can close as soon as the request is done.
 */
function postToken(port: number): Promise<{ status: number; body: string }> {
  const body = new URLSearchParams({
    grant_type: 'authorization_code',
    code: 'a-code',
    redirect_uri: 'http://localhost:8010/oauth/callback',
    client_id: 'mcpo-client',
    code_verifier: 'a'.repeat(43),
  }).toString();
  return new Promise((resolve, reject) => {
    const req = request(
      {
        host: '127.0.0.1',
        port,
        path: '/token',
        method: 'POST',
        agent: false,
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      },
      (res) => {
        let text = '';
        res.setEncoding('utf8');
        res.on('data', (chunk) => (text += chunk));
        res.on('end', () => resolve({ status: res.statusCode ?? 0, body: text }));
      }
    );
    req.on('error', reject);
    req.end(body);
  });
}

describe('graceful shutdown of the HTTP server', () => {
  const savedEnv = { ...process.env };
  const realFetch = globalThis.fetch;
  let server: MicrosoftGraphServer;
  let port: number;
  let microsoftReceived: ReturnType<typeof deferred>;
  let microsoftAnswers: ReturnType<typeof deferred>;

  beforeEach(async () => {
    clearSecretsCache();
    process.env.MS365_MCP_CLIENT_ID = CLIENT_ID;
    process.env.MS365_MCP_CLIENT_SECRET = 'secret';
    process.env.MS365_MCP_TENANT_ID = 'common';

    // Microsoft holds the code exchange until the test lets it answer.
    microsoftReceived = deferred();
    microsoftAnswers = deferred();
    vi.spyOn(globalThis, 'fetch').mockImplementation(async (input, init) => {
      const url = input instanceof Request ? input.url : String(input);
      if (url === MICROSOFT_TOKEN_URL) {
        microsoftReceived.resolve();
        await microsoftAnswers.promise;
        return Response.json({
          access_token: 'graph-access-token',
          token_type: 'Bearer',
          expires_in: 3599,
        });
      }
      return realFetch(input, init);
    });

    port = await freePort();
    server = new MicrosoftGraphServer(fakeAuthManager(), { http: `127.0.0.1:${port}` });
    await server.initialize('0.0.0-test');
    await server.start();
  });

  afterEach(async () => {
    microsoftAnswers.resolve();
    vi.restoreAllMocks();
    await server.stop();
    clearSecretsCache();
    process.env = { ...savedEnv };
  });

  it('refuses new connections, finishes the request in flight, then exits 0', async () => {
    const signals = new EventEmitter();
    const exit = vi.fn();
    installGracefulShutdown(server, { signals, exit });

    const inFlight = postToken(port);
    await microsoftReceived.promise;

    signals.emit('SIGTERM', 'SIGTERM');

    const refused = await fetch(`http://127.0.0.1:${port}/.well-known/oauth-authorization-server`)
      .then(() => 'accepted')
      .catch((error: Error & { cause?: { code?: string } }) => error.cause?.code);
    expect(refused).toBe('ECONNREFUSED');
    expect(exit).not.toHaveBeenCalled();

    microsoftAnswers.resolve();
    const response = await inFlight;

    expect(response.status).toBe(200);
    expect(JSON.parse(response.body).access_token).toBe('graph-access-token');
    await vi.waitFor(() => expect(exit).toHaveBeenCalledWith(0));
  });
});
