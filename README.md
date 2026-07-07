# outlook-experiment

[![CI](https://github.com/muhammaddadu/outlook-cli/actions/workflows/ci.yml/badge.svg)](https://github.com/muhammaddadu/outlook-cli/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](./LICENSE)
[![Node](https://img.shields.io/badge/node-%E2%89%A520-brightgreen)](./package.json)

A local CLI for **mail, calendar, and Microsoft Teams** that piggybacks on
your existing **Outlook Web (OWA) session** — no Azure AD app registration, no
admin consent, no third-party SaaS. It captures the OAuth tokens OWA already
uses for its own APIs and replays them from Node `fetch`.

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

## What it can do

| Area | Commands | Notes |
| --- | --- | --- |
| **Mail — read** | `list`, `unread`, `search`, `read`, `folders` | Friendly filters (`--unread`, `--from`, `--since 7d`, `--folder Sent`, `--has-attachments`, …) or raw OData |
| **Mail — write** | `draft`, `draft-reply`, `draft-reply-all`, `draft-forward`, `discard-draft`, `send` | Review-first drafts preferred over direct `send`; `--attach` files (≤3 MB) |
| **Calendar** | `agenda`, `events`, `event-read`, `calendars`, `event-create`, `event-update`, `event-cancel`, `accept`/`decline`/`tentative`, `free-busy` | Natural time ranges (`today`, `+7d`, `--days 14`); RSVP; availability lookup |
| **Teams** (via Graph) | `teams`, `teams-channels`, `teams-chats`, `teams-members`, `teams-messages`, `teams-send` | Needs `outlook auth --all`. Listing + sending work; reading message bodies needs `Chat.Read` (may be tenant-blocked) |
| **Auth / diagnostics** | `auth [--all]`, `refresh`, `logout`, `token-audit`, `graph`, `context`, `learn` | `token-audit` shows which Microsoft resources you can reach; `graph` is a raw authenticated Graph/Substrate passthrough |

Mail and calendar use your **Outlook** token; Teams uses a **Microsoft Graph**
token (a separate audience) captured by `outlook auth --all`. **Copilot is not
supported** — its conversation runs over a streaming WebSocket, which needs a
dedicated client (see [`LEARNINGS.md`](./LEARNINGS.md) §14).

Every API call has **built-in recovery**: a rejected token auto-recaptures and
retries once, throttling (429/503) is retried with `Retry-After` backoff, and
requests are timeout-bounded so the CLI never hangs. All output is JSON on
stdout (pipe-friendly); diagnostics go to stderr.

---

## Quickstart

One command installs the CLI globally, downloads Chromium, and registers
the Claude Code agent skill:

```bash
npm install -g github:muhammaddadu/outlook-cli && outlook setup --with-skill
outlook auth                # one-time interactive sign-in (handles MFA)
outlook list -n 5           # smoke test
```

For Teams (and other Microsoft Graph resources), do the broader sign-in once —
it captures a token per resource in the same session:

```bash
outlook auth --all          # also opens Teams + Copilot to capture their tokens
outlook token-audit         # confirm which resources are now reachable
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

# Teams (needs `outlook auth --all`)
outlook teams                                  # teams you belong to
outlook teams-chats -n 20                      # your 1:1 and group chats
outlook teams-members <chat-id>                # who's in a chat
outlook teams-send <chat-id> "on my way"       # post a message (sends immediately)
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
  *"Send a reply to the message from Bob"*, *"What Teams am I in?"*,
  *"Send a Teams message to the Eng chat"*.
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
   `outlook.office.com` (and, with `--all`, Teams + Copilot too).
2. Those web apps perform their own silent OAuth handshakes. The CLI watches
   the network and **captures the `Authorization: Bearer …` JWTs** they use to
   call their own APIs.
3. Each token is classified by audience and cached per resource
   (`outlook.office.com`, `graph.microsoft.com`, `substrate.office.com`).
   Subsequent CLI calls reuse the matching cached token and never re-open the
   browser until it expires — routing each command to the right API base.

No new permission grant is created. The scopes are exactly the ones you
consented to the first time you signed into those web apps. Because each
Microsoft API validates the token audience, mail can't call Graph and vice
versa — which is why Teams needs its own captured token from `auth --all`.

Full explanation in [`docs/architecture.md`](./docs/architecture.md); the
dead ends we ruled out first are in [`LEARNINGS.md`](./LEARNINGS.md).

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

Tests live in [`test/`](./test) and run with `npm test` (no network, no
Playwright).

---

## Requirements

- **Node.js ≥ 20**
- **macOS, Linux, or WSL** (untested on plain Windows)
- An Outlook Web account you can sign into at `outlook.office.com`

---

## License

[MIT](./LICENSE).
