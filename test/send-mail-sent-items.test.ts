/**
 * Graph keeps a copy of a mail sent with /me/sendMail in the sender's Sent Items
 * unless the request says `SaveToSentItems: false`. Microsoft's OpenAPI metadata
 * declares that field with `default: false`, and the tool's input schema passed
 * the default on. An MCP gateway that builds its request from the advertised
 * schema (mcpo) then sent false on every call: mail sent from Askii chat never
 * showed in the user's Sent Items.
 *
 * The tests register the tools from the generated client on a real McpServer and
 * talk to it through an MCP client, as a gateway does: tools/list for the
 * advertised schema, tools/call for what reaches Graph.
 */
import { describe, expect, it, vi } from 'vitest';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
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

const message = {
  subject: 'Hello',
  body: { contentType: 'text', content: 'Hi there' },
  toRecipients: [{ emailAddress: { address: 'user@example.com' } }],
};

async function connectedClient() {
  const graphRequest = vi.fn().mockResolvedValue({ content: [{ type: 'text', text: '{}' }] });
  const server = new McpServer({ name: 'test', version: '0.0.0' });
  registerGraphTools(server, { graphRequest } as unknown as GraphClient, false);
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  await server.connect(serverTransport);
  const client = new Client({ name: 'test-client', version: '0.0.0' });
  await client.connect(clientTransport);
  return { client, graphRequest };
}

describe('send-mail and Sent Items', () => {
  it('advertises a copy in Sent Items when the caller does not say', async () => {
    const { client } = await connectedClient();

    const { tools } = await client.listTools();
    const tool = tools.find((t) => t.name === 'send-mail');
    expect(tool).toBeDefined();
    const body = (tool!.inputSchema.properties as Record<string, JsonSchema>).body;

    expect(body.properties!.SaveToSentItems.default).toBe(true);
  });

  it('sends SaveToSentItems false to Graph when the caller asks for no copy', async () => {
    const { client, graphRequest } = await connectedClient();

    const result = await client.callTool({
      name: 'send-mail',
      arguments: { body: { Message: message, SaveToSentItems: false } },
    });

    expect(result.isError).toBeFalsy();
    expect(graphRequest).toHaveBeenCalledTimes(1);
    const [path, options] = graphRequest.mock.calls[0] as [string, { body: string }];
    expect(path).toBe('/me/sendMail');
    expect(JSON.parse(options.body)).toEqual({ Message: message, SaveToSentItems: false });
  });
});
