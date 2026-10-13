/**
 * Microsoft's OpenAPI spec gives almost every Graph list endpoint the same eight
 * OData query options ($top, $skip, $search, $filter, $count, $orderby, $select,
 * $expand), whatever the API behind it accepts. A model reads them as an
 * invitation: in the Askii chat tests gpt-4o sent $orderby to list-channel-messages
 * and $top to list-team-channels, and Graph answered 400 "Query option 'X' is not
 * allowed". On other tools Graph accepts an option and ignores it ($filter on file
 * versions), so the model gets a list it did not ask for.
 *
 * The expected options come from Microsoft's reference page of each API, kept only
 * where a live probe against Graph accepted them (2026-10-14). The tests register
 * the tools from the generated client on a real McpServer and talk to it through an
 * MCP client, as a gateway does.
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

const ODATA_OPTIONS = ['top', 'skip', 'search', 'filter', 'count', 'orderby', 'select', 'expand'];

async function connectedClient() {
  const graphRequest = vi.fn().mockResolvedValue({ content: [{ type: 'text', text: '{}' }] });
  const server = new McpServer({ name: 'test', version: '0.0.0' });
  registerGraphTools(server, { graphRequest } as unknown as GraphClient, false, undefined, true);
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  await server.connect(serverTransport);
  const client = new Client({ name: 'test-client', version: '0.0.0' });
  await client.connect(clientTransport);
  return { client, graphRequest };
}

async function offeredOptions(): Promise<Map<string, string[]>> {
  const { client } = await connectedClient();
  const { tools } = await client.listTools();
  return new Map(
    tools.map((t) => [
      t.name,
      Object.keys(t.inputSchema.properties ?? {})
        .map((name) => name.replace(/^\$/, '').toLowerCase())
        .filter((name) => ODATA_OPTIONS.includes(name))
        .sort(),
    ])
  );
}

describe('query options a tool offers', () => {
  it.each([
    ['list-drives', ['expand', 'orderby', 'select', 'top']],
    ['list-folder-files', ['expand', 'orderby', 'select', 'top']],
    ['search-onedrive-files', ['expand', 'orderby', 'select', 'top']],
    ['list-sharepoint-site-drives', ['expand', 'orderby', 'select', 'top']],
    ['get-drive-delta', ['expand', 'select', 'top']],
    ['list-drive-item-thumbnails', ['select']],
    ['list-drive-item-permissions', ['select']],
    ['list-drive-item-versions', ['select', 'top']],
    ['search-sharepoint-sites', ['orderby', 'search', 'select', 'top']],
    ['list-chats', ['expand', 'filter', 'orderby', 'top']],
    ['list-chat-members', []],
    ['list-chat-messages', ['filter', 'orderby', 'top']],
    ['list-team-channels', ['filter', 'select']],
    ['list-channel-messages', ['expand', 'top']],
    ['list-channel-message-replies', ['top']],
    ['get-team-channel', ['select']],
    ['get-channel-message', ['select']],
  ])('%s offers only the options Graph accepts there', async (tool, expected) => {
    const offered = await offeredOptions();

    expect(offered.get(tool)).toEqual(expected);
  });

  it('a mail list keeps all eight, since Graph accepts them on mail', async () => {
    const offered = await offeredOptions();

    expect(offered.get('list-mail-messages')).toEqual([
      'count',
      'expand',
      'filter',
      'orderby',
      'search',
      'select',
      'skip',
      'top',
    ]);
  });
});

describe("Graph's $top limit", () => {
  // mcpo builds its request model from the advertised JSON schema and reads neither
  // allOf nor maximum: an intersection there loses even the integer type.
  it.each([
    'list-chats',
    'list-chat-messages',
    'list-channel-messages',
    'list-channel-message-replies',
  ])('%s advertises $top as one plain integer schema with maximum 50', async (tool) => {
    const { client } = await connectedClient();

    const { tools } = await client.listTools();
    const top = tools.find((t) => t.name === tool)!.inputSchema.properties!.top as Record<
      string,
      unknown
    >;

    expect(top.type).toBe('integer');
    expect(top.maximum).toBe(50);
    expect(top).not.toHaveProperty('allOf');
  });
});

describe('calling a tool', () => {
  const channelArgs = { teamId: 'team-1', channelId: 'channel-1' };

  it('refuses an option the tool does not offer, naming it, without calling Graph', async () => {
    const { client, graphRequest } = await connectedClient();

    const result = await client.callTool({
      name: 'list-channel-messages',
      arguments: { ...channelArgs, top: 5, orderby: 'createdDateTime desc' },
    });

    expect(result.isError).toBe(true);
    expect(JSON.stringify(result.content)).toContain('$orderby');
    expect(graphRequest).not.toHaveBeenCalled();
  });

  it('refuses $top above the 50 Graph allows on channel messages, without calling Graph', async () => {
    const { client, graphRequest } = await connectedClient();

    const result = await client.callTool({
      name: 'list-channel-messages',
      arguments: { ...channelArgs, top: 60 },
    });

    expect(result.isError).toBe(true);
    expect(graphRequest).not.toHaveBeenCalled();
  });

  it('still sends $top up to 50 on channel messages', async () => {
    const { client, graphRequest } = await connectedClient();

    await client.callTool({
      name: 'list-channel-messages',
      arguments: { ...channelArgs, top: 50 },
    });

    const [path] = graphRequest.mock.calls[0] as [string];
    expect(path).toBe('/teams/team-1/channels/channel-1/messages?$top=50');
  });
});
