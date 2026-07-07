# AGENTS.md

Instructions for AI coding agents (Claude Code, Cursor, Aider, Codex, …)
working on this repository.

If you only read one other file, read [`LEARNINGS.md`](./LEARNINGS.md). It
documents the long sequence of dead ends we hit before landing on the current
architecture. Skim it before proposing alternative approaches — most of the
"obvious" ideas have already been tried and won't work in a locked-down
enterprise tenant.

---

## TL;DR — what this project is

A small Node.js CLI that talks to the Outlook REST API by **capturing the
Bearer JWT that Outlook Web (OWA) already uses for its own API calls**. It
does this by launching Playwright Chromium against the user's persistent
profile, watching the network for the first authed request, and snapshotting
the `Authorization` + routing headers to disk. Subsequent CLI calls reuse
the cached headers from Node `fetch` without re-opening the browser.

## Multiple resources (Outlook / Graph / Substrate)

Each Microsoft API is a separate OAuth resource with its own token audience;
a token is only valid against its own resource. `resources.mjs` is the
registry. Tokens are captured per-resource (`captureAllTokens` classifies
every Bearer OWA/Teams/Copilot emit by audience), cached per-resource
(`auth-<resource>.json`), and routed per-resource (`call(auth, path, {
resource })`, `getAuth({ resource })`). `outlook graph <path>` is the generic
Graph passthrough; `outlook token-audit` reports what's reachable. Reaching
Teams/Copilot is a *capture* problem (get the right-audience token), **not** a
consent problem — see [`LEARNINGS.md`](./LEARNINGS.md) §13. Non-default
resources never auto-launch the browser mid-command; they error toward `auth
--all`.

## TL;DR — hard constraints you must respect

1. **Never add OAuth flows that require admin consent.** The whole point of
   this project is to avoid them. Many corporate tenants block them. See
   [`LEARNINGS.md`](./LEARNINGS.md) §1–4.
2. **Never assume headless Chromium will work.** Microsoft Conditional Access
   stalls silent SSO when the UA looks headless. We always launch headed.
   See [`LEARNINGS.md`](./LEARNINGS.md) §10.
3. **Never read tokens from the macOS keychain in Office identity slots.**
   They're ACL-protected to Microsoft-signed binaries; CLI tools just get
   silent empty reads. See [`LEARNINGS.md`](./LEARNINGS.md) §5.
4. **Never authenticate with cookies alone against the REST API.** The
   `/api/v2.0/me/messages` endpoint returns 401 on cookie-only calls — it
   requires a Bearer token with audience `https://outlook.office.com/`.
5. **Don't commit anything from `~/.cache/outlook-spike/` or `~/.local/share/
   outlook-spike/`** into the repo. These contain real credentials. The
   `.gitignore` already protects against the obvious filenames; double-check
   anything you stage.

---

## Setup

```bash
npm install
npm run setup       # `playwright install chromium`
```

State directories created at runtime:

| Path | Purpose | In repo? |
| --- | --- | --- |
| `~/.local/share/outlook-spike/browser-profile/` | Chromium user-data-dir (cookies, IndexedDB) | **No** |
| `~/.cache/outlook-spike/auth.json` | Cached Bearer + routing headers | **No** |

Both honour `$XDG_DATA_HOME` / `$XDG_CACHE_HOME` and the per-path overrides
`OUTLOOK_PROFILE` / `OUTLOOK_TOKEN_CACHE`.

## Run / test commands

```bash
./src/cli.mjs --help            # commander-generated help
./src/cli.mjs auth              # interactive sign-in (headed Chromium)
./src/cli.mjs list -n 3         # smoke test — should return JSON inbox
./src/cli.mjs --debug list      # verbose diagnostics on stderr
node src/diagnose.mjs           # sniff OWA endpoints (dev tool)
npm test                        # full unit + integration suite, no network
```

The test suite lives in `test/` and uses Node's built-in `node:test`
runner. Integration tests spawn the CLI as a subprocess against a local
mock HTTP server (no Playwright, no network). When you add new
functionality, **add a test that exercises it via the CLI surface** — that
catches both wire format and exit-code regressions in one shot. See
`docs/development.md` for examples.

---

## File layout

```
src/
  cli.mjs        # commander entry point, signal handling, error formatting
  auth.mjs       # token cache: load/save/clear + getAuth() cache-first wrapper
  capture.mjs    # Playwright launcher + Bearer-header capture (lazy-loaded)
  client.mjs     # pure-fetch call() helper, OUTLOOK_API_BASE override
  odata.mjs      # mail filter / query builders (--unread, --from, --folder, …)
  calendar.mjs   # calendar time parsing, calendarView URL, event filters
  resources.mjs  # registry of Microsoft resources (outlook/graph/substrate): base URL, audiences, classification
  audit.mjs      # token-audit: decode cached tokens + group scopes into capability areas (offline)
  learn.mjs      # persistent learnings (load/add/forget/clear)
  jwt.mjs        # tiny JWT payload decoder (no signature validation)
  output.mjs     # printJson / debug / info / errorBlock (stdout vs stderr)
  errors.mjs     # AppError class, error codes, exit-code mapping
  paths.mjs      # XDG-compliant cache and data directories (lazy-evaluated); per-resource cache files
  diagnose.mjs   # dev-only: sniff OWA endpoints to find new APIs
  sniff.mjs      # dev-only: PII-safe live capture — logs method+path+audience+JSON shape while you click; saves per-resource tokens
skill/
  outlook/SKILL.md  # agent skill — installs to ~/.claude/skills/outlook/ via skill/install.mjs
  commands/*.md     # slash commands (/outlook, /inbox, /unread, /draft, /agenda)
  install.mjs       # symlink/copy the skill into Claude / Codex / Cursor
test/
  helpers.mjs    # mock HTTP server, fake JWT, CLI subprocess runner
  *.test.mjs     # unit + integration tests (run via `npm test`)
```

Touch `cli.mjs` to add a subcommand. Touch `client.mjs` when extending what
the CLI does with the captured headers (new endpoints, new HTTP verbs,
retry/timeout policy). Touch `auth.mjs` only for cache-lifecycle changes —
the capture mechanism itself lives in `capture.mjs`, which must only ever
be imported lazily (it drags in Playwright; a static import from `cli.mjs`
slows down every command).

---

## Code conventions

- **ESM only.** Files are `.mjs`; package is `"type": "module"`.
- **No TypeScript.** Plain JS with JSDoc where useful. The project is small
  enough that a build step is overhead.
- **stdout = data, stderr = diagnostics.** Never write progress or debug
  text to stdout — it breaks `outlook list | jq …`. Use `info()` / `debug()`
  / `errorBlock()` from `output.mjs`.
- **Errors are `AppError` instances** with a stable `code` (one of the `E.*`
  constants in `errors.mjs`) and an actionable `hint`. Map new failure modes
  to a code and add it to the table in `docs/troubleshooting.md`.
- **Exit codes follow `EXIT.*`** in `errors.mjs`: 0 OK, 1 generic, 2 auth, 3
  HTTP/network, 64 usage, 130 SIGINT. Don't invent new ones without updating
  the docs.
- **Recovery is automatic where safe.** A 401 clears the cache, recaptures
  once, and retries; 429/503 retry with `Retry-After` backoff; network
  errors retry GETs only. Don't add per-command retry loops on top of this.
- **Dependencies stay minimal.** Right now only `commander` and `playwright`.
  Adding a dep requires a clear justification (writing your own would be
  significant work).
- **No emojis in code.** Comments stay plain ASCII.
- **Comment style: explain *why*, not *what*.** Anyone can read the code to
  see what it does; pick a few key spots and explain the reasoning.

---

## What to do if Microsoft breaks the integration

The Outlook REST v2 endpoint at `/api/v2.0/me/...` is undocumented and being
sunset in favour of Microsoft Graph. If `list` starts returning `404` or
`Gone`, the path is:

1. Run `node src/diagnose.mjs` and let it sniff OWA's outgoing requests for
   ~60 seconds.
2. Look for the new endpoint OWA uses to fetch the inbox (likely something
   under `/owa/service.svc?action=…` or a Graph-style URL).
3. Update `REST_BASE` and the relevant query strings in `client.mjs` /
   `cli.mjs` to match.

Conditional Access policies can also change — if `auth` starts redirecting
to login indefinitely, see [`docs/troubleshooting.md`](./docs/troubleshooting.md).

---

## When you edit the user-facing skill

If you change the CLI surface (new subcommand, renamed flag, changed exit
code), update `skill/outlook/SKILL.md` in the same patch. That file is what
Claude Code / Codex / Cursor see when they decide whether and how to drive
the CLI on the user's behalf. Stale skill descriptions cause agents to call
non-existent flags.

Quick install during dev:

```bash
npm run skill:install              # user scope (symlinked, edits live)
npm run skill:install:project      # repo scope only
```

Then ask Claude Code to "check my unread mail" or similar — if your changes
to the skill landed, the agent should drive the new behaviour correctly.

## Pull request checklist

If you're producing a patch:

- [ ] `npm test` is green.
- [ ] `./src/cli.mjs list -n 3` returns HTTP 200 against a live mailbox.
- [ ] `./src/cli.mjs --help` and `./src/cli.mjs --version` still work.
- [ ] Any new error path throws an `AppError` with a `code` and `hint`.
- [ ] New stdout writes only contain command output; diagnostics go to stderr.
- [ ] Any new subcommand has at least one integration test in `test/cli.test.mjs`.
- [ ] CLI surface changes are reflected in `skill/outlook/SKILL.md`.
- [ ] `.gitignore` covers any new file types you introduced.
- [ ] If you tried an alternative approach and rejected it, add the result
      to [`LEARNINGS.md`](./LEARNINGS.md) so the next agent doesn't repeat it.
