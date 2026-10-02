/**
 * This repository is a fork. Upstream's workflows come with every upstream
 * merge, and GitHub runs them here too: release.yml runs on each push to main
 * and would run semantic-release (tags and a GitHub release in this repo), then
 * try to publish @softeria/ms-365-mcp-server to npm and push
 * ghcr.io/softeria/ms-365-mcp-server.
 *
 * So every job of every upstream workflow must be guarded to run only in
 * upstream's repository. The workflows this fork does want to run are listed in
 * FORK_WORKFLOWS. A new upstream workflow, or an upstream merge that drops a
 * guard, fails this test.
 */
import { describe, expect, it } from 'vitest';
import { readdirSync, readFileSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import yaml from 'js-yaml';

const WORKFLOWS_DIR = path.join(
  path.dirname(fileURLToPath(import.meta.url)),
  '..',
  '.github',
  'workflows'
);

// Workflows that run in this fork as well (CI on pull requests): upstream's
// build.yml, and fork-coverage.yml, which is this fork's own.
const FORK_WORKFLOWS = ['build.yml', 'fork-coverage.yml'];

const UPSTREAM_GUARD = "github.repository == 'Softeria/ms-365-mcp-server'";

interface Workflow {
  jobs: Record<string, { if?: string }>;
}

function upstreamWorkflows(): string[] {
  return readdirSync(WORKFLOWS_DIR)
    .filter((file) => /\.ya?ml$/.test(file))
    .filter((file) => !FORK_WORKFLOWS.includes(file));
}

function jobsOf(file: string): Workflow['jobs'] {
  const workflow = yaml.load(readFileSync(path.join(WORKFLOWS_DIR, file), 'utf8')) as Workflow;
  return workflow.jobs;
}

function stripExpressionBraces(condition: string): string {
  const trimmed = condition.trim();
  const match = /^\$\{\{([\s\S]*)\}\}$/.exec(trimmed);
  return (match ? match[1] : trimmed).trim();
}

// True when "||" appears outside parentheses and string literals. With a
// top-level "||", "guard && a || b" runs whenever b is true, guard or not.
function hasTopLevelOr(expression: string): boolean {
  let depth = 0;
  let inString = false;
  for (let i = 0; i < expression.length; i++) {
    const char = expression[i];
    if (char === "'") inString = !inString;
    if (inString) continue;
    if (char === '(') depth++;
    if (char === ')') depth--;
    if (depth === 0 && expression.startsWith('||', i)) return true;
  }
  return false;
}

function isGuarded(condition: string | undefined): boolean {
  if (condition === undefined) return false;
  const expression = stripExpressionBraces(condition);
  if (expression === UPSTREAM_GUARD) return true;
  return expression.startsWith(`${UPSTREAM_GUARD} && `) && !hasTopLevelOr(expression);
}

describe('upstream workflows in this fork', () => {
  it('finds the upstream release workflow', () => {
    expect(upstreamWorkflows()).toContain('release.yml');
  });

  it('runs no upstream job outside upstream', () => {
    const unguarded = upstreamWorkflows().flatMap((file) =>
      Object.entries(jobsOf(file))
        .filter(([, job]) => !isGuarded(job.if))
        .map(([name, job]) => `${file} job "${name}": if: ${job.if ?? '(none)'}`)
    );

    expect(unguarded).toEqual([]);
  });
});

describe('isGuarded', () => {
  it('accepts the guard alone or joined with && and a parenthesised rest', () => {
    expect(isGuarded(UPSTREAM_GUARD)).toBe(true);
    expect(isGuarded(`${UPSTREAM_GUARD} && github.event_name == 'push'`)).toBe(true);
    expect(isGuarded(`\${{ ${UPSTREAM_GUARD} && !cancelled() && (a != '' || b != '') }}`)).toBe(
      true
    );
  });

  it('rejects a missing guard, a guard not in front, or a top-level ||', () => {
    expect(isGuarded(undefined)).toBe(false);
    expect(isGuarded("github.event_name == 'push'")).toBe(false);
    expect(isGuarded(`github.event_name == 'push' && ${UPSTREAM_GUARD}`)).toBe(false);
    expect(isGuarded(`${UPSTREAM_GUARD} && a != '' || b != ''`)).toBe(false);
  });
});
