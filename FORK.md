# About this fork

This repository is Pressingly's fork of
[Softeria/ms-365-mcp-server](https://github.com/Softeria/ms-365-mcp-server).
We run it as the Microsoft 365 MCP server for Askii, behind our MCP gateway
(mcpo). One deployment serves every Microsoft 365 product we offer; the
image's `CMD` in `Dockerfile` decides which tools it serves.

`README.md` and the other upstream files stay as upstream wrote them, so
upstream merges stay simple. Our notes live in this file only.

**Last upstream merge:** `v0.157.2` (`004002c`, 2026-09-30).

## Rules that no tool checks

1. **This repository is public.** Never commit secrets, tokens, tenant IDs or
   customer data. That covers commit messages, branch names and pull request
   text too.
2. **Changes reach `main` only by pull request, merged with a merge commit.**
   Never rebase, squash or force-push `main`. GitHub enforces part of this:
   the merge commit is the only merge method allowed, and the ruleset on
   `main` blocks force pushes and deletion and requires the `build (22.x)`
   and `coverage` checks.
3. **Keep our patches additive**: a new file under `src/lib/`, a new test
   file or a flag, rather than an edited upstream line. Fewer edited upstream
   lines mean fewer merge conflicts.
4. **Every patch has a test that fails if an upstream merge loses it**, and a
   row in "Our patches" below.
5. **Branches and commits** follow the team's practice: branch
   `<type>/<issue-key>-<summary>`, for example
   `fix/agenticmem-63-token-errors`; commit subject in Conventional Commits
   with a scope and the issue key, for example
   `fix(http): pass Microsoft's 429 through /token [AGENTICMEM-63]`.
6. **Our release tags start with `askii-`**: `askii-v1.0.0-rc.1` for sandbox,
   `askii-v1.0.0` for production. Upstream tags each release `v0.x.y` and may
   reach `v1.0.0` itself; the prefix keeps the two apart. Dev deploys from
   `main` and needs no tag.
7. **Issues are off here.** Work is planned in the team's own tracker.

## Our patches

| What it does                                                                                          | Commit    | Upstream files we edit                              | Test that guards it                    |
| ----------------------------------------------------------------------------------------------------- | --------- | --------------------------------------------------- | -------------------------------------- |
| Sign-in popups stay connected (COOP), and an ID token is always requested                             | `0fb252b` | `src/server.ts`                                     | `test/authorize-popup-sign-in.test.ts` |
| Upstream's release jobs run only in upstream's repository                                             | `64cfd9f` | `.github/workflows/release.yml`                     | `test/fork-upstream-workflows.test.ts` |
| Standard PKCE challenges go straight to Microsoft, so any copy of the server can answer `/token`      | `04cc03e` | `src/server.ts`                                     | `test/pkce-passthrough.test.ts`        |
| HTTP mode uses only the caller's token, never `MS365_MCP_OAUTH_TOKEN`                                 | `7caf3de` | `src/auth.ts`, `src/server.ts`                      | `test/server-held-token.test.ts`       |
| Requests in flight finish on SIGTERM, then the server exits                                           | `28775e7` | `src/index.ts`                                      | `test/graceful-shutdown.test.ts`       |
| Tool registration logs at debug, not on every `/mcp` request                                          | `b70f551` | `src/graph-tools.ts`, 16 upstream test files        | `test/registration-logging.test.ts`    |
| The image runs as the `node` user                                                                     | `8e42aba` | `Dockerfile`                                        | none: check `USER node` after a merge  |
| The Graph spec comes from a fixed `msgraph-metadata` commit                                           | `7208ce1` | `bin/modules/download-openapi.mjs`, `.dockerignore` | `test/graph-spec-pinned.test.ts`       |
| `/token` passes Microsoft's 429 and 5xx through, so a gateway keeps the sign-in                       | `0bcea0b` | `src/lib/microsoft-auth.ts`                         | `test/gateway-token.test.ts`           |
| The image starts as the OneDrive server; its tool list is pinned                                      | `ff23874` | `Dockerfile`                                        | `test/gateway-tool-list.test.ts`       |
| A read-only `sharepoint` preset of the five site tools; the sensitivity-labels tool leaves `onedrive` | `82c3220` | `src/endpoints.json`, `src/tool-categories.ts`      | `test/sharepoint-preset.test.ts`       |
| The image starts as the OneDrive and SharePoint server, in org mode                                   | `5da8472` | `Dockerfile`                                        | `test/gateway-tool-list.test.ts`       |

Our own files, which upstream does not have:

- `test/gateway-*.test.ts`: discovery, `/register`, `/token` and `/mcp` as
  our gateway uses them (`f28bafe`, `edaa53b`, `e7e26c0`).
- `.github/workflows/fork-coverage.yml` and `vitest.fork-coverage.config.ts`:
  a pull request fails when one of our patched files loses coverage
  (`515b497`).
- `src/lib/pkce.ts`, `src/lib/graph-token-source.ts`,
  `src/lib/graceful-shutdown.ts`, and this file.

## Merging upstream

We do not merge upstream on a schedule. When we do, use upstream's latest
release tag, not the tip of its `main`:

```bash
git fetch upstream --tags
git tag -l 'v*' --sort=-v:refname | head -1      # upstream's latest, e.g. v0.158.0
git switch -c chore/merge-upstream-v0.158.0 origin/main
git merge v0.158.0
npm run verify
npx vitest run --config vitest.fork-coverage.config.ts
git push -u origin HEAD
```

- Resolve conflicts with "Our patches" open: each row names the upstream
  files we edit.
- `test/fork-upstream-workflows.test.ts` fails on a new upstream workflow:
  guard its jobs to upstream's repository, or add it to `FORK_WORKFLOWS` if
  it should run here.
- `test/gateway-tool-list.test.ts` fails when upstream adds, renames or drops
  a tool we serve: check the change, then update the pinned list.
- Check that `Dockerfile` still says `USER node`.
- Open a PR titled `chore: merge upstream v0.158.0`, merge it with a merge
  commit, then update "Last upstream merge" and the patch table.
- Push the branch only. Never push upstream's tags to this repository.

## Adding a Microsoft 365 product

The image's `CMD` in `Dockerfile` sets the tools the server serves (today
`--preset onedrive,sharepoint --org-mode`). To add a product, change the
`CMD` and update the pinned tool list in `test/gateway-tool-list.test.ts` in
the same PR.

A product whose tools carry `workScopes` only (SharePoint, Teams, shared
mailboxes) registers only with `--org-mode`. Give it its own preset:
`requiresOrgMode: true` in `PRESET_META` (`src/tool-categories.ts`), the
preset's name in each tool's `presets` array in `src/endpoints.json`, and a
contract test like `test/sharepoint-preset.test.ts`. Before the first
`--org-mode`, compare `--list-permissions` with and without it for the
presets already served: a tool that org mode adds with an admin-consent
scope must leave those presets first, as
`extract-drive-item-sensitivity-labels` left `onedrive`.

mcpo lists every tool the `CMD` serves to every AskiiMall Microsoft service,
whatever scope the service asks for: since SharePoint, the OneDrive connector
lists the five site tools, and a call with a OneDrive token fails with a 403.
So a product adds its tools to every Microsoft connector; keep its preset to
what the product needs.

## Moving the Graph spec forward

`npm run generate` downloads Microsoft's Graph spec from a fixed
`msgraph-metadata` commit (`GRAPH_SPEC_COMMIT` in
`bin/modules/download-openapi.mjs`), so every build gets the same tools. To
move forward, change that constant in its own commit, then run
`npm run generate -- --force` and `npm run verify`.

## Testing on your machine

`npm run verify` must pass before a PR, with two known exceptions. Both also
fail on untouched upstream code and pass on GitHub's runner:

- macOS: 5 tests in `test/attachment-split-listener.test.ts` (macOS has no
  `127.0.0.2`).
- Docker: the "uppercase LOCALHOST bind" test in `test/loopback-http.test.ts`
  (Docker resolves `LOCALHOST` to `::1` first).
