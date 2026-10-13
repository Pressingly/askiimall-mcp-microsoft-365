/**
 * Graph keeps a copy of a mail sent with /me/sendMail in the sender's Sent Items
 * unless the request says `SaveToSentItems: false`. Microsoft's OpenAPI metadata
 * declares that field with `default: false`, and the tool's input schema passed
 * the default on. An MCP gateway that builds its request from the advertised
 * schema (mcpo) then sent false on every call: mail sent from Askii chat never
 * showed in the user's Sent Items.
 *
 * The test registers the tools from the generated client on a real McpServer and
 * reads send-mail's input schema from tools/list, as a gateway does.
 */
import { describe, expect, it, vi } from 'vitest';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { registerGraphTools } from '../src/graph-tools.js';
import type { GraphClient } from '../src/graph-client.js';

vi.mock('../src/logger.js', () => ({
  default: {
    info: vi.fn(),
    debug: vi.fn(),
    error: vi.fn(),
    warn: vi.fn(),
  },
}));

type JsonSchema = {
  properties?: Record<string, JsonSchema>;
  type?: string | string[];
  default?: unknown;
};

async function sendMailBodySchema(): Promise<JsonSchema> {
  const server = new McpServer({ name: 'test', version: '0.0.0' });
  registerGraphTools(server, {} as GraphClient, false);
  const listTools = (
    server.server as unknown as {
      _requestHandlers: Map<string, (req: unknown, extra: unknown) => Promise<unknown>>;
    }
  )._requestHandlers.get('tools/list')!;
  const result = (await listTools(
    { method: 'tools/list' },
    { signal: new AbortController().signal }
  )) as { tools: Array<{ name: string; inputSchema: JsonSchema }> };
  const tool = result.tools.find((t) => t.name === 'send-mail');
  expect(tool).toBeDefined();
  return tool!.inputSchema.properties!.body;
}

describe('send-mail and Sent Items', () => {
  it('advertises a copy in Sent Items when the caller does not say', async () => {
    const body = await sendMailBodySchema();

    expect(body.properties!.SaveToSentItems.default).toBe(true);
  });

  it('still lets the caller send without a copy', async () => {
    const body = await sendMailBodySchema();

    expect(body.properties!.SaveToSentItems.type).toContain('boolean');
  });
});
