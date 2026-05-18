# Contributing

Thanks for the interest. This is a personal spike that grew into a real
tool; contributions are welcome but the project has a few firm
constraints worth understanding before you open a PR.

---

## Before you start

Read these, in order, and skim the rest as needed:

1. **[`LEARNINGS.md`](./LEARNINGS.md)** — the 11 dead ends we hit before
   the current architecture worked. Most "what about Graph API?" / "what
   about reading the keychain?" / "what about headless?" ideas have
   already been tried and have a documented reason they don't work in
   locked-down enterprise tenants. Save yourself an afternoon.
2. **[`docs/PRD.md`](./docs/PRD.md)** — what this tool is for, who it's
   for, and what's deliberately out of scope. PRs that conflict with the
   non-goals will get pushed back.
3. **[`AGENTS.md`](./AGENTS.md)** — if you're using an AI assistant to
   help write the change, the hard constraints listed there apply to
   both of you.

---

## What kinds of contributions are welcome

**Welcome:**

- Bug fixes, especially anything that breaks the cache-hit fast path or
  causes a regression in the test suite.
- New mail / calendar filter flags that compose with the existing ones.
- New subcommands that fit the established surface (commander, JSON
  output, AppError-based failures, exit-code conventions).
- Test coverage improvements — particularly integration tests against
  the mock server for edge cases.
- Doc improvements. Especially [`docs/troubleshooting.md`](./docs/troubleshooting.md)
  entries when you hit something weird and figure it out.
- Improvements to `src/diagnose.mjs` — making endpoint rediscovery
  easier when Microsoft changes things.

**Probably welcome, ask first:**

- New AI-agent skill files / slash commands beyond what's in `skill/`.
- Daemon-mode work (long-lived browser process exposing a local socket
  to avoid the cache-miss flash).
- MCP server wrapper.

**Out of scope** (see PRD §3 and §7):

- Gmail / IMAP / non-Microsoft mail support.
- Multi-mailbox / multi-account features.
- A GUI.
- Anything that requires Azure AD app registration or admin consent.
  The whole point of this project is to avoid those.

---

## Development setup

```bash
git clone git@github.com:muhammaddadu/outlook-cli.git
cd outlook-cli
npm install
npm run setup           # downloads the Chromium binary Playwright uses (~150 MB)

./src/cli.mjs auth      # sign in once with your own M365 account
./src/cli.mjs list -n 3 # smoke test
npm test                # 100+ tests, no network, no Playwright (~7s)
```

Optional:

```bash
npm link                # so `outlook` works from any directory
npm run skill:install   # install the agent skill into Claude/Codex/Cursor
```

---

## Code style

The bar is **boring is best**.

- **ESM only.** Files are `.mjs`; package is `"type": "module"`.
- **No TypeScript.** Plain JS with JSDoc where the type isn't obvious.
- **No semicolons-on / semicolons-off war**: just match what's already in
  the file you're editing.
- **2-space indent, single quotes, trailing commas in multi-line literals.**
- **Comments explain *why*, not *what*.** Anyone can read the code to
  see what it does; comment the reasoning a future maintainer needs.
- **No emojis in source files.** Comments stay plain ASCII.
- **Dependencies stay minimal.** Currently `commander` and `playwright`.
  Adding a dep needs a clear justification — writing it yourself would
  need to be significant work.
- **Errors are `AppError`** instances with a stable `code` (one of the
  `E.*` constants in `src/errors.mjs`) and an actionable `hint` field
  written for the user, not the developer.
- **stdout = data, stderr = diagnostics.** Always. Never write progress
  or debug text to stdout — it breaks `outlook list | jq …`.

---

## Testing

```bash
npm test
```

- Tests live in `test/` and use Node's built-in `node:test` runner.
- **Unit tests** import modules directly and assert on pure functions
  (errors, paths, JWT decoding, OData composition, calendar parsing,
  the learnings store).
- **Integration tests** spawn `src/cli.mjs` as a subprocess against a
  local mock HTTP server (`test/helpers.mjs::startMockServer`). No
  real network, no Playwright, no real mailbox. This is what catches
  wire-format bugs and exit-code regressions in one shot.
- **Every new subcommand** needs at least one integration test in
  `test/cli.test.mjs`. See the existing tests for the pattern (seed a
  token cache, start a mock server, run `runCli([…])`, assert on
  exit code + stdout JSON + observed request URL/body).
- **Every new error path** needs an `AppError({ code, message, hint })`
  and a test that asserts on the exit code + `code` in stderr.

What NOT to add to the test suite:

- Tests that hit `outlook.office.com` for real. Always use the mock.
- Tests that exercise Playwright. The auth-capture path is e2e-only —
  verify manually with `./src/cli.mjs auth`.
- Tests that depend on a specific real JWT lifetime. Use
  `seedTokenCache({ secondsFromNow })` for synthetic, predictable
  tokens.

---

## Commit & PR conventions

**Commits:**

- One logical change per commit. Squash noise locally.
- Subject line in imperative mood, under 70 chars
  (`Add agenda command, not Added agenda command`).
- Body explains the why if it's not obvious from the diff.

**PRs:**

- Title same shape as a commit subject.
- Body should answer:
  1. What does this do?
  2. Why does it do it that way?
  3. What changed in observable behaviour (new flags, new exit codes,
     new files, etc.)?
- Reference any open question from [`docs/PRD.md`](./docs/PRD.md) §10 that
  this resolves.

**PR checklist** (also in `AGENTS.md` for AI contributors):

- [ ] `npm test` is green (no regressions, < 10s).
- [ ] `./src/cli.mjs list -n 3` returns HTTP 200 against a live mailbox
      (verify manually before pushing).
- [ ] `./src/cli.mjs --help` and `./src/cli.mjs --version` still work.
- [ ] Any new error path throws an `AppError` with a `code` and `hint`.
- [ ] New stdout writes only contain command output; diagnostics go to
      stderr.
- [ ] Any new subcommand has an integration test in `test/cli.test.mjs`.
- [ ] CLI surface changes are reflected in `skill/outlook/SKILL.md`
      and the slash commands in `skill/commands/`.
- [ ] `.gitignore` covers any new file types you introduced.
- [ ] If you tried an alternative approach and rejected it, add the
      result to [`LEARNINGS.md`](./LEARNINGS.md) so the next contributor
      doesn't repeat it.
- [ ] No PII (real names, real email addresses, tenant IDs, project
      names) in any file. The repo `.gitignore` blocks the obvious
      filenames; the audit is on you for content.

---

## Updating the agent skill

When you change the CLI surface (new subcommand, renamed flag, changed
exit code), update [`skill/outlook/SKILL.md`](./skill/outlook/SKILL.md)
in the same patch. That file is what Claude Code / Codex / Cursor see
when deciding whether and how to drive the CLI on the user's behalf.
**Stale skill descriptions cause agents to call non-existent flags.**

Quick re-install during dev:

```bash
npm run skill:install              # user scope, all detected agents (symlinked)
npm run skill:install:project      # repo scope only
```

Then ask Claude Code to "check my unread mail" or similar — if your
changes landed, the agent drives the new behaviour correctly. If it
hallucinates a flag that doesn't exist, your SKILL.md is out of sync.

---

## Security & PII

This project handles real mail and calendar data. Treat the codebase
accordingly.

- **Never commit anything that looks like a token, cookie, JWT, or
  mailbox identifier.** `.gitignore` blocks the obvious filenames
  (`auth.json`, `*.token.json`, `.env*`) but the audit is on you for
  inline content.
- **Never commit real names, real email addresses, real tenant /
  organisation IDs, real subject lines, real project names.** Use
  placeholders like `alice@example.com`, `Jane Doe`, `Project Alpha`,
  `tenant-uuid`.
- Before pushing: `git diff --cached | grep -iE "<your tenant>|<your
  domain>|<your name>"` and visually confirm.
- The `~/.cache/outlook-spike/auth.json` file is a credential. Don't
  share it in screenshots, gists, or bug reports.

If you find a security issue, please open a GitHub issue tagged
`security` rather than emailing — the maintainer doesn't have a private
disclosure channel and the project is too small to need one. Be careful
not to include real captured tokens in the issue body.

See [`docs/security.md`](./docs/security.md) for the full posture.

---

## Reporting issues

Use the GitHub issue tracker. Useful issues include:

- **What you tried** (exact command).
- **What you expected** to happen.
- **What actually happened** (full `--debug` output, with any tokens
  redacted).
- **Your environment** — OS, Node version (`node --version`), CLI
  version (`outlook --version`), tenant type (M365 enterprise vs
  personal).
- **Whether you've checked [`docs/troubleshooting.md`](./docs/troubleshooting.md).**

Bonus points for a minimal repro using the mock server from
`test/helpers.mjs`.

---

## Licensing

This project is [MIT-licensed](./LICENSE). By submitting a contribution
you agree it can be distributed under the same terms.
