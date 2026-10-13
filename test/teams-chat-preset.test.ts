import { readFileSync } from 'fs';
import path from 'path';
import { describe, expect, it } from 'vitest';
import { buildAllowedScopeDiagnostics } from '../src/auth.js';
import { getCombinedPresetPattern, TOOL_CATEGORIES } from '../src/tool-categories.js';

// Contract tests for the teams-chat preset, Askii's Teams connector (FORK.md
// "Our patches"). In chat the model finds the user's chats, teams and
// channels, reads their messages and a channel message's replies, and sends a
// message or a channel reply.
// Upstream's teams preset holds 69 tools, many with admin-only scopes
// (meeting transcripts and recordings, presence of other users, deleting
// channels, managing members), and every Microsoft connector lists every tool
// the server runs, so the preset is exactly these 15: no meetings, no
// presence, no edits or deletes, no new chats, no channel or member
// management. Replies exist only in channels: Graph v1.0 has no
// /chats/{id}/messages/{id}/replies (404 "Requested API is not supported", see
// test/teams-write-preset.test.ts), so the two chat-reply tools stay out. A
// change that grows the preset must consciously update this contract.

interface Endpoint {
  toolName: string;
  pathPattern: string;
  method: string;
  presets?: string[];
  scopes?: string[] | string[][];
  workScopes?: string[] | string[][];
}

const endpoints = JSON.parse(
  readFileSync(path.join(__dirname, '../src/endpoints.json'), 'utf8')
) as Endpoint[];

const READ_TOOLS = [
  'get-channel-message',
  'get-chat',
  'get-chat-message',
  'get-team',
  'get-team-channel',
  'list-channel-message-replies',
  'list-channel-messages',
  'list-chat-members',
  'list-chat-messages',
  'list-chats',
  'list-joined-teams',
  'list-team-channels',
];
const SEND_TOOLS = ['reply-to-channel-message', 'send-channel-message', 'send-chat-message'];
const CHAT_TOOLS = [...READ_TOOLS, ...SEND_TOOLS].sort();

const presetEndpoints = endpoints.filter((e) => e.presets?.includes('teams-chat'));

function scopeGroups(raw: string[] | string[][] | undefined): string[][] {
  if (!raw || raw.length === 0) return [];
  return Array.isArray(raw[0]) ? (raw as string[][]) : [raw as string[]];
}

describe('teams-chat preset contract', () => {
  it('contains exactly the chat and channel read and send tools', () => {
    expect(presetEndpoints.map((e) => e.toolName).sort()).toEqual(CHAT_TOOLS);
  });

  it('exposes the 15 tools and the universal byte readers, nothing else', () => {
    const pattern = TOOL_CATEGORIES['teams-chat'].pattern;
    for (const tool of [...CHAT_TOOLS, 'download-bytes', 'download-bytes-to-file']) {
      expect(tool).toMatch(pattern);
    }
    for (const tool of [
      'parse-teams-url', // meeting URLs only; meetings are out
      'get-download-url', // rides with the drive-backed presets; onedrive brings it
      'list-chat-message-replies', // not a Graph v1.0 API (header)
      'reply-to-chat-message', // not a Graph v1.0 API (header)
      'create-chat',
      'update-chat-message',
      'unpin-chat-message',
      'set-chat-message-reaction',
      'delete-team-channel',
      'create-team-channel',
      'add-team-member',
      'list-team-members',
      'list-meeting-transcripts',
      'get-meeting-recording-content',
      'get-user-presence',
      'set-my-presence',
      'send-my-activity-notification',
      'list-mail-messages',
    ]) {
      expect(tool).not.toMatch(pattern);
    }
  });

  it('reads with GET and sends with POST only', () => {
    for (const endpoint of presetEndpoints) {
      const expected = SEND_TOOLS.includes(endpoint.toolName) ? 'post' : 'get';
      expect(endpoint.method.toLowerCase(), endpoint.toolName).toBe(expected);
    }
  });

  it('carries work-account scopes only, and requires org mode', () => {
    for (const endpoint of presetEndpoints) {
      expect(endpoint.scopes, `${endpoint.toolName} has personal-account scopes`).toBeUndefined();
      expect(scopeGroups(endpoint.workScopes).length, endpoint.toolName).toBeGreaterThan(0);
    }
    expect(TOOL_CATEGORIES['teams-chat'].requiresOrgMode).toBe(true);
  });

  // What the auth layer computes from the declarations: the scope a login for
  // the preset alone would request (User.Read and offline_access are injected
  // at the OAuth layer, not here). ChannelMessage.Read.All always needs an
  // admin's consent. Chat.Read covers list-chats' and get-chat's
  // Chat.ReadBasic, so the auth layer folds that one away.
  it('computes exactly the seven chat and channel scopes', () => {
    const diagnostics = buildAllowedScopeDiagnostics({
      enabledTools: getCombinedPresetPattern(['teams-chat']),
      orgMode: true,
    });
    expect(diagnostics.effectivePermissions).toEqual([
      'Channel.ReadBasic.All',
      'ChannelMessage.Read.All',
      'ChannelMessage.Send',
      'Chat.Read',
      'ChatMember.Read',
      'ChatMessage.Send',
      'Team.ReadBasic.All',
    ]);
    expect(diagnostics.disabledTools).toEqual([]);
  });

  // Graph v1.0 reads chat messages, listed or one by one, only with Chat.Read
  // or Chat.ReadWrite in delegated mode; a token with ChatMessage.Read gets 403
  // "API requires one of 'Chat.Read, Chat.ReadWrite'" (seen live in Askii's
  // Teams connector). Upstream's teams preset hides this because
  // list-pinned-chat-messages brings Chat.Read along; this preset has no such
  // tool, so the two readers must declare it themselves.
  it('reads chat messages only with Chat.Read, the scope Graph requires', () => {
    const diagnostics = buildAllowedScopeDiagnostics({
      enabledTools: getCombinedPresetPattern(['teams-chat']),
      orgMode: true,
      allowedScopes:
        'Channel.ReadBasic.All ChannelMessage.Read.All ChannelMessage.Send Chat.ReadBasic ' +
        'ChatMember.Read ChatMessage.Read ChatMessage.Send Team.ReadBasic.All',
    });
    expect(diagnostics.disabledTools.map((tool) => tool.toolName).sort()).toEqual([
      'get-chat-message',
      'list-chat-messages',
    ]);
    expect(diagnostics.missingAllowedScopesForTools).toEqual(['Chat.Read']);
  });
});
