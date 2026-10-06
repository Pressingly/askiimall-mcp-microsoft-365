import { readFileSync } from 'fs';
import path from 'path';
import { describe, expect, it } from 'vitest';
import { buildAllowedScopeDiagnostics } from '../src/auth.js';
import { getCombinedPresetPattern, TOOL_CATEGORIES } from '../src/tool-categories.js';

// Contract tests for the sharepoint preset, Askii's SharePoint connector
// (FORK.md "Our patches"). In chat the model finds a site and lists its
// document libraries with these tools, and reads a library's files with the
// onedrive preset's drive tools, which take a drive id. So the preset is
// exactly the five site read tools: no lists, no OneNote, no site items, no
// site delta, no writes. A change that grows it must consciously update this
// contract.
//
// The second describe guards the other half of the same change: the onedrive
// preset no longer holds extract-drive-item-sensitivity-labels, whose only
// scope, Files.Read.All (admin consent), --org-mode would otherwise add to
// what the OneDrive server advertises, and which no AskiiMall service asks for.

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

const SITE_TOOLS = [
  'get-sharepoint-site',
  'get-sharepoint-site-by-path',
  'get-sharepoint-site-drive-by-id',
  'list-sharepoint-site-drives',
  'search-sharepoint-sites',
].sort();

const presetEndpoints = endpoints.filter((e) => e.presets?.includes('sharepoint'));

function scopeGroups(raw: string[] | string[][] | undefined): string[][] {
  if (!raw || raw.length === 0) return [];
  return Array.isArray(raw[0]) ? (raw as string[][]) : [raw as string[]];
}

describe('sharepoint preset contract', () => {
  it('contains exactly the five site read tools', () => {
    expect(presetEndpoints.map((e) => e.toolName).sort()).toEqual(SITE_TOOLS);
  });

  it('exposes the five tools and the universal byte readers, nothing else', () => {
    const pattern = TOOL_CATEGORIES.sharepoint.pattern;
    for (const tool of [...SITE_TOOLS, 'download-bytes', 'download-bytes-to-file']) {
      expect(tool).toMatch(pattern);
    }
    for (const tool of [
      'get-download-url', // rides with the drive-backed presets; onedrive brings it
      'list-sharepoint-site-lists',
      'list-sharepoint-site-items',
      'get-sharepoint-sites-delta',
      'list-sharepoint-site-onenote-notebooks',
      'create-sharepoint-list-item',
      'list-folder-files',
      'get-drive-item',
    ]) {
      expect(tool).not.toMatch(pattern);
    }
  });

  it('is read-only, and every tool asks first for Sites.Read.All, a work-account scope', () => {
    for (const endpoint of presetEndpoints) {
      expect(endpoint.method.toLowerCase(), endpoint.toolName).toBe('get');
      expect(endpoint.scopes, `${endpoint.toolName} has personal-account scopes`).toBeUndefined();
      const [primary] = scopeGroups(endpoint.workScopes);
      expect(primary, endpoint.toolName).toEqual(['Sites.Read.All']);
    }
  });

  it('requires org mode', () => {
    expect(TOOL_CATEGORIES.sharepoint.requiresOrgMode).toBe(true);
  });

  // What the auth layer computes from the declarations: the scope a login for
  // the preset alone would request (User.Read and offline_access are injected
  // at the OAuth layer, not here).
  it('computes exactly Sites.Read.All as the token the preset asks for', () => {
    const diagnostics = buildAllowedScopeDiagnostics({
      enabledTools: getCombinedPresetPattern(['sharepoint']),
      orgMode: true,
    });
    expect(diagnostics.effectivePermissions).toEqual(['Sites.Read.All']);
    expect(diagnostics.disabledTools).toEqual([]);
  });
});

describe('the onedrive preset in org mode', () => {
  it('leaves extract-drive-item-sensitivity-labels to the files and personal presets', () => {
    const tool = endpoints.find((e) => e.toolName === 'extract-drive-item-sensitivity-labels');
    expect(tool?.presets).toContain('files');
    expect(tool?.presets).toContain('personal');
    expect(tool?.presets).not.toContain('onedrive');
  });

  it('asks for Files.ReadWrite alone, with and without --org-mode', () => {
    for (const orgMode of [false, true]) {
      const diagnostics = buildAllowedScopeDiagnostics({
        enabledTools: getCombinedPresetPattern(['onedrive']),
        orgMode,
      });
      expect(diagnostics.effectivePermissions, `orgMode=${orgMode}`).toEqual(['Files.ReadWrite']);
    }
  });
});
