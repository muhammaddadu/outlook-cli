# outlook-experiment

A local CLI that reads, searches and sends mail through your existing **Outlook
Web (OWA) session** — no Azure AD app registration, no admin consent, no
third-party SaaS.

```bash
$ outlook list -n 3 | jq '.value[] | {From: .From.EmailAddress.Name, Subject}'
{ "From": "Doe, Jane",      "Subject": "Re: Q3 planning sync" }
{ "From": "Garcia, Alex",   "Subject": "Vendor proposal" }
{ "From": "Lee, Sam",       "Subject": "Quarterly review prep" }
```

> [!WARNING]
> This is an end-user developer experiment, not a sanctioned production tool.
> It piggybacks on undocumented Outlook Web endpoints and may break without
> notice. Read [`docs/security.md`](./docs/security.md) before using it,
> especially on a managed work device.

---

## Quickstart

One command installs the CLI globally, downloads Chromium, and registers
the Claude Code agent skill:

```bash
npm install -g github:muhammaddadu/outlook-cli && outlook setup --with-skill
outlook auth                # one-time interactive sign-in (handles MFA)
outlook list -n 5           # smoke test
```

After that, the binary is on your `$PATH` everywhere:

```bash
outlook unread --from boss@example.com         # unread from your boss
outlook list --folder Sent --since 7d          # sent in the last week
outlook search "deploy" --has-attachments      # full-text + filter
```

`outlook setup --with-skill` auto-detects every AI agent on your machine
and installs the skill into all of them:

| Agent | Skill location | Slash commands |
| --- | --- | --- |
| Claude Code | `~/.claude/skills/outlook/SKILL.md` | `~/.claude/commands/*.md` |
| Codex CLI | `~/.codex/skills/outlook/SKILL.md` | `~/.codex/prompts/*.md` |
| Cursor | `~/.cursor/rules/outlook.md` | (Cursor uses rules, not slash commands) |

In any agent session you can then either:

- **Ask in plain English** — the skill auto-loads when intent matches:
  *"Check my unread mail"*, *"Find emails from Alice this week"*,
  *"Send a reply to the message from Bob"*.
- **Use slash commands** (Claude Code and Codex CLI only):
  - `/outlook <request>` — generic Outlook task
  - `/inbox` — summarise the inbox
  - `/unread` — triage unread mail

> **Restart your agent CLI** after installing — skills and commands are
> loaded at session start, not on the fly.

To target only specific agents:

```bash
node skill/install.mjs --target claude              # just Claude Code
node skill/install.mjs --target claude,codex        # Claude + Codex
node skill/install.mjs --target all                 # all three regardless of detection
node skill/install.mjs --uninstall                  # remove from everywhere
```

### From source (for development)

```bash
git clone git@github.com:muhammaddadu/outlook-cli.git
cd outlook-cli
npm install
node src/cli.mjs setup --with-skill
node src/cli.mjs auth
```

Or use the helper scripts:

```bash
npm run skill:install              # user scope: ~/.claude/skills/outlook/
npm run skill:install:project      # this repo only: .claude/skills/outlook/
```

---

## How it works (30 seconds)

1. The CLI launches Chromium with a persistent profile and points it at
   `outlook.office.com`.
2. OWA performs its own silent OAuth handshake. The CLI watches the network
   and **captures the `Authorization: Bearer …` JWT** that OWA uses to call
   its own APIs.
3. The token (plus a handful of routing headers OWA always sends) is cached
   on disk. Subsequent CLI calls reuse it and never re-open the browser
   until the JWT expires.

No new permission grant is created. The scopes are exactly the ones you
consented to the first time you signed into Outlook web.

Full explanation in [`docs/architecture.md`](./docs/architecture.md).

---

## Documentation

| Doc | What it covers |
| --- | --- |
| [`docs/usage.md`](./docs/usage.md) | Every command, every flag, real examples |
| [`docs/architecture.md`](./docs/architecture.md) | Bearer-capture mechanism, file layout, design decisions |
| [`docs/security.md`](./docs/security.md) | Threat model, what's safe, what to be careful about |
| [`docs/troubleshooting.md`](./docs/troubleshooting.md) | Error codes, recovery procedures |
| [`docs/development.md`](./docs/development.md) | Working on the codebase, diagnosing new endpoints |
| [`AGENTS.md`](./AGENTS.md) | Instructions for AI agents editing this code |
| [`LEARNINGS.md`](./LEARNINGS.md) | Approaches that didn't work and why — read before exploring alternatives |
| [`skill/outlook/SKILL.md`](./skill/outlook/SKILL.md) | Agent skill that teaches Claude / Codex / Cursor how to drive the CLI |

Tests live in [`test/`](./test) and run with `npm test` (~5 seconds, no
network, no Playwright).

---

## Requirements

- **Node.js ≥ 20**
- **macOS, Linux, or WSL** (untested on plain Windows)
- An Outlook Web account you can sign into at `outlook.office.com`

---

## License

[MIT](./LICENSE).
