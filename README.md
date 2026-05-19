# outlook-experiment

[![CI](https://github.com/muhammaddadu/outlook-cli/actions/workflows/ci.yml/badge.svg)](https://github.com/muhammaddadu/outlook-cli/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](./LICENSE)
[![Node](https://img.shields.io/badge/node-%E2%89%A520-brightgreen)](./package.json)

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
# Mail
outlook unread --from boss@example.com         # unread from your boss
outlook list --folder Sent --since 7d          # sent in the last week
outlook search "deploy" --has-attachments      # full-text + filter
outlook draft-reply <id> '{...}'               # AI-composed draft for review
outlook draft-reply <id> --attach ~/file.png   # attach files to a draft or send

# Calendar
outlook agenda --days 7                        # next week
outlook agenda --from today --to tomorrow      # just today
outlook free-busy alice@example.com bob@example.com   # availability lookup
outlook accept <event-id> -c "see you there"   # RSVP
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
| [`docs/PRD.md`](./docs/PRD.md) | Product Requirements Document — what this is for, who it's for, what's in / out of scope |
| [`docs/reverse-engineering.md`](./docs/reverse-engineering.md) | How to apply the same approach to other web apps |
| [`CONTRIBUTING.md`](./CONTRIBUTING.md) | How to send PRs, what's welcome, what's out of scope, security / PII rules |

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
