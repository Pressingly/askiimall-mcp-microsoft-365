/**
 * Askii runs one copy of this server for all its Microsoft 365 products, and
 * the image's own CMD (Dockerfile) says which tools it serves. An MCP gateway
 * (mcpo) lists those tools for every Microsoft service in the AskiiMall, and
 * each service asks Microsoft only for the scope typed for it in the developer
 * portal, which /authorize passes on. So the CMD decides what Askii users see
 * in chat. It grows with each product: OneDrive first, then SharePoint (the
 * five site read tools of the `sharepoint` preset; a library's files are read
 * with the drive tools, which take a drive id).
 *
 * The test reads the CMD, parses it as the server does, starts the server and
 * asks it as mcpo does. A change to the CMD, or an upstream merge that adds,
 * renames or drops a tool in these presets, fails it.
 */
import { afterAll, afterEach, beforeAll, beforeEach, describe, expect, it, vi } from 'vitest';
import { readFileSync } from 'node:fs';
import { createServer, type Server } from 'node:http';
import type { AddressInfo } from 'node:net';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { parseArgs, type CommandOptions } from '../src/cli.js';
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

const DOCKERFILE = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', 'Dockerfile');

// Each of these replaces part of what the CMD sets: the tools, the scopes, or
// the scope /authorize asks for. A deployment must not set them.
const OVERRIDING_ENV = [
  'READ_ONLY',
  'ENABLED_TOOLS',
  'MS365_MCP_ALLOWED_SCOPES',
  'MS365_MCP_EXTRA_SCOPES',
  'MS365_MCP_ORG_MODE',
  'MS365_MCP_FORCE_WORK_SCOPES',
];

// Every OneDrive action, write, delete and sharing included. Excel's workbook
// tools (the `excel` preset) are left out for now (the user, 2026-10-02).
const ONEDRIVE_TOOLS = [
  'copy-drive-item',
  'create-drive-item-preview',
  'create-drive-item-share-link',
  'create-onedrive-folder',
  'create-upload-session',
  'delete-drive-item-permission',
  'delete-onedrive-file',
  'get-drive-delta',
  'get-drive-item',
  'get-drive-root-item',
  'list-drive-item-permissions',
  'list-drive-item-thumbnails',
  'list-drive-item-versions',
  'list-drives',
  'list-folder-files',
  'move-rename-onedrive-item',
  'search-onedrive-files',
  'share-drive-item',
  'upload-file-content',
];
// SharePoint: the `sharepoint` preset, five read tools to find a site and its
// document libraries (test/sharepoint-preset.test.ts pins the preset itself).
const SHAREPOINT_TOOLS = [
  'get-sharepoint-site',
  'get-sharepoint-site-by-path',
  'get-sharepoint-site-drive-by-id',
  'list-sharepoint-site-drives',
  'search-sharepoint-sites',
];
// download-bytes-to-file writes to the server's own disk and is hidden over HTTP.
const HELPER_TOOLS = ['download-bytes', 'get-download-url'];

function dockerInstruction(name: string): string[] | undefined {
  const line = readFileSync(DOCKERFILE, 'utf8')
    .split('\n')
    .find((l) => l.startsWith(`${name} `));
  return line ? (JSON.parse(line.slice(name.length + 1)) as string[]) : undefined;
}

// A Microsoft access token for a work account is a JWT; only its exp matters here.
function jwt(claims: Record<string, unknown>): string {
  const part = (value: object) => Buffer.from(JSON.stringify(value)).toString('base64url');
  return `${part({ typ: 'JWT', alg: 'RS256' })}.${part(claims)}.signature`;
}

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

describe("the image's CMD, as the server Askii runs", () => {
  const savedEnv = { ...process.env };
  const cmd = dockerInstruction('CMD');
  let options: CommandOptions;
  let server: MicrosoftGraphServer;
  let origin: string;

  beforeAll(() => {
    for (const name of OVERRIDING_ENV) delete process.env[name];
    const savedArgv = process.argv;
    process.argv = ['node', 'dist/index.js', ...(cmd ?? [])];
    try {
      options = parseArgs();
    } finally {
      process.argv = savedArgv;
    }
  });

  afterAll(() => {
    process.env = { ...savedEnv };
  });

  beforeEach(async () => {
    clearSecretsCache();
    process.env.MS365_MCP_CLIENT_ID = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee';
    process.env.MS365_MCP_CLIENT_SECRET = 'app-secret';
    process.env.MS365_MCP_TENANT_ID = 'common';
    delete process.env.MS365_MCP_ALLOWED_REDIRECT_URIS;

    const port = await freePort();
    server = new MicrosoftGraphServer(fakeAuthManager(), {
      ...options,
      http: `127.0.0.1:${port}`,
    });
    await server.initialize('0.0.0-test');
    await server.start();
    origin = `http://127.0.0.1:${port}`;
  });

  afterEach(async () => {
    await server.stop();
    clearSecretsCache();
  });

  async function getJson(url: string) {
    const res = await fetch(url);
    expect(res.status).toBe(200);
    return (await res.json()) as Record<string, unknown>;
  }

  it('starts the HTTP server on every interface and logs to stdout', () => {
    expect(dockerInstruction('ENTRYPOINT')).toEqual(['node', 'dist/index.js']);
    expect(cmd).toBeDefined();
    expect(options.http).toBe('0.0.0.0:3000');
    expect(options.v).toBe(true);
    expect(options.preset).toBe('onedrive,sharepoint');
    expect(options.orgMode).toBe(true);
  });

  it('lists every OneDrive and SharePoint tool to the gateway, and nothing else', async () => {
    const res = await fetch(`${origin}/mcp`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Accept: 'application/json, text/event-stream',
        Authorization: `Bearer ${jwt({ oid: 'user-a', exp: Math.floor(Date.now() / 1000) + 3600 })}`,
      },
      body: JSON.stringify({ jsonrpc: '2.0', id: 1, method: 'tools/list', params: {} }),
    });
    expect(res.status).toBe(200);
    const body = (await res.json()) as { result: { tools: { name: string }[] } };
    const names = body.result.tools.map((tool) => tool.name).sort();

    expect(names).toEqual([...ONEDRIVE_TOOLS, ...SHAREPOINT_TOOLS, ...HELPER_TOOLS].sort());
  });

  // Sites.Read.All is outside Microsoft's default user-consent policy, so on
  // most work tenants an admin approves the app once; Files.ReadWrite a user
  // approves alone. Org mode must add nothing else: the sensitivity-labels
  // tool's Files.Read.All stays out because that tool left the onedrive
  // preset. The gateway asks Microsoft for its service's scope, not this list
  // (next test); this list is what the server tells clients it can use.
  it('advertises exactly Files.ReadWrite and Sites.Read.All', async () => {
    const resource = await getJson(`${origin}/.well-known/oauth-protected-resource/mcp`);
    const metadata = await getJson(`${origin}/.well-known/oauth-authorization-server`);

    expect(resource.scopes_supported).toEqual(['Files.ReadWrite', 'Sites.Read.All']);
    expect(metadata.scopes_supported).toEqual(['Files.ReadWrite', 'Sites.Read.All']);
  });

  it("asks Microsoft for the scope the gateway's service stores, not the server's own", async () => {
    const params = new URLSearchParams({
      response_type: 'code',
      client_id: 'mcp-client-1',
      redirect_uri: 'https://mcpo.example.com/oauth/callback',
      state: 'abcdefghij',
      scope: 'Mail.Read',
    });
    const res = await fetch(`${origin}/authorize?${params}`, { redirect: 'manual' });

    expect(res.status).toBe(302);
    const scopes = new URL(res.headers.get('location')!).searchParams.get('scope')!.split(' ');
    expect(scopes).toContain('Mail.Read');
    expect(scopes).not.toContain('Files.ReadWrite');
  });
});
