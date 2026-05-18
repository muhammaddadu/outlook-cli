# Product Requirements Document — outlook-cli

| Field | Value |
| --- | --- |
| **Status** | Active spike — experimental, single-author |
| **Owner** | Repo maintainers |
| **Last updated** | 2026-05-18 |
| **Code home** | https://github.com/muhammaddadu/outlook-cli |

---

## 1. Problem statement

End users in locked-down Microsoft 365 tenants need programmatic access
to their own Outlook mailbox and calendar, but every standard path is
gated:

- **Microsoft Graph API** requires an Azure AD app registration with
  admin consent on `Mail.*` / `Calendars.*` scopes. Many enterprise
  tenants either block self-service app registration outright or require
  admin approval for these scopes — both stopping end-user automation
  dead.
- **Outlook desktop add-ins** require sideloading or tenant-wide
  deployment, both of which need IT sign-off.
- **Third-party SaaS bridges** (Nylas, etc.) hit the same admin consent
  wall plus introduce a data-egress hop that many security teams block.
- **Outlook for Mac AppleScript automation** works but is brittle, slow
  on large mailboxes, and language-locks you to JXA / AppleScript.

The result: users who can perfectly well sign into Outlook web and act
on their mailbox manually cannot script the same actions, even when
those actions would be allowed by their organisation's policy.

## 2. Goals

1. **Zero-admin setup.** Anything a user can do in a normal browser as
   themselves should be scriptable without touching Azure AD, IT, or a
   third-party service.
2. **AI-agent first.** The primary consumer is an AI agent
   (Claude Code, Codex CLI, Cursor) acting on the user's behalf. The
   CLI surface is JSON-first, has stable exit codes, and ships with an
   agent skill file that teaches the agent how to drive it.
3. **Safe-by-default mutations.** AI-composed mail uses the Drafts
   folder so the user reviews before sending. Calendar mutations on
   attendee-bearing events require explicit user confirmation.
4. **Adapt over time.** A persistent learnings store lets the agent
   accumulate user preferences (sign-off, aliases, default tone) so the
   experience improves session-over-session without manual config.
5. **Low friction.** One install command on a Mac with Node 20+ gets
   the CLI, the Chromium browser Playwright needs, and the agent skill
   installed across every supported AI CLI on the machine.

## 3. Non-goals

- Multi-user / multi-mailbox support. This is single-user automation
  for the signed-in account.
- A graphical UI. Output is JSON; humans use `jq` and AI agents
  consume it directly.
- Bypassing any genuine access control. If the user can't sign into
  Outlook web, the CLI can't help them. If the tenant requires admin
  consent and the user genuinely shouldn't have mail access, this tool
  isn't an end-run around that.
- Replacement for IT-sanctioned automation platforms (Power Automate,
  Logic Apps). Those exist for a reason in enterprise contexts; this
  is for personal productivity scripts.
- Wide cross-mailbox-provider support. Gmail and IMAP are explicitly
  out of scope.

## 4. Users / personas

1. **The power-user developer.** Has shell fluency. Wants `outlook
   unread | jq` style ad-hoc queries and the ability to script weekly
   digests, automated triage rules, and PR-from-email workflows.
2. **The AI-agent operator.** Runs Claude Code (or Codex / Cursor) for
   day-to-day work. Says things like "draft a reply to Alice's email
   about the deploy" or "what's on my calendar this afternoon" and
   expects the agent to know what to do.
3. **The locked-down enterprise user.** Wants the things above but
   their IT department has blocked everything obvious. This tool is
   purpose-built for them.

## 5. Functional requirements

### 5.1 Mail

| ID | Requirement |
| --- | --- |
| MAIL-1 | List, search, and read messages with full filter support (sender, date window, attachments, folder, importance, subject, raw OData). |
| MAIL-2 | Browse non-Inbox folders by friendly name (`Sent`, `Drafts`, `Archive`, …) or folder Id. |
| MAIL-3 | Create review-first drafts: new draft, reply, reply-all, forward, discard. Quoted thread MUST be preserved on replies / forwards. |
| MAIL-4 | Send mail directly (no review) — explicit `send` command, surfaced as a separate operation. |
| MAIL-5 | List mail folders with counts. |
| MAIL-6 | Output is pipe-safe JSON to stdout; diagnostics to stderr; exit codes follow sysexits conventions. |

### 5.2 Calendar

| ID | Requirement |
| --- | --- |
| CAL-1 | Show agenda for a configurable time window (today, tomorrow, this week, custom range). Recurring events MUST be expanded into individual instances. |
| CAL-2 | Read full event detail by Id (body, attendees, location, recurrence, online meeting URL). |
| CAL-3 | Create / update / cancel events. Operations on events with `Attendees` MUST be flagged as immediate (no draft workflow exists). |
| CAL-4 | RSVP (accept / decline / tentative) with optional comment and a `--no-respond` flag for local-only response. |
| CAL-5 | Free/busy lookup across multiple mailboxes with configurable interval. |
| CAL-6 | List user's calendars (primary + shared + subscribed). |

### 5.3 Agent integration

| ID | Requirement |
| --- | --- |
| AGT-1 | Ship a single skill file (`SKILL.md`) with the right frontmatter so Claude Code / Codex / Cursor auto-load it on intent matching. |
| AGT-2 | Provide slash-command shortcuts (`/outlook`, `/inbox`, `/unread`, `/draft`, `/agenda`) for explicit invocation. |
| AGT-3 | Auto-detect installed agents and install skill + commands to each in a single command. |
| AGT-4 | Expose a session-start context call (`outlook context`) that returns user identity + accumulated learnings, with no network traffic. |
| AGT-5 | Provide a learnings store (`outlook learn add/forget/clear`) the agent maintains across sessions. |

### 5.4 Operational

| ID | Requirement |
| --- | --- |
| OPS-1 | One-shot install: `npm i -g github:muhammaddadu/outlook-cli && outlook setup --with-skill && outlook auth`. |
| OPS-2 | Token cache + browser profile live under XDG paths outside the repo. |
| OPS-3 | Cache-hit calls (the common case) must complete in < 1 second and must NOT open a browser window. |
| OPS-4 | The CLI must fail closed on auth issues: 401 → clear cache, exit 2, tell user to re-run. |
| OPS-5 | Test suite runs `node --test test/` with no network, no Playwright, no real mailbox — only mock servers and synthetic JWTs. |

## 6. Non-functional requirements

| Area | Requirement |
| --- | --- |
| **Performance** | Cache hit: sub-second (P95 < 1s). Cache miss: ≤ 5 seconds end-to-end including silent SSO. |
| **Reliability** | All 4xx/5xx surface as `E_HTTP` with the server's error body in the hint. 401 specifically clears the token cache so the next call refreshes. |
| **Privacy** | Token cache and learnings file are local-only, never transmitted. Plain-text markdown for learnings so the user can audit/edit. |
| **Security** | No code-signed binaries needed; no keychain reads of Microsoft-protected slots; no proxy/MITM. Captured Bearer is scoped to the user's existing OWA consent — no scope elevation. |
| **Portability** | macOS confirmed working; Linux likely works (untested); WSL likely works; Windows native not yet supported. |
| **Compatibility** | Node ≥ 20. ESM only. No TypeScript build step. Two production dependencies (`commander`, `playwright`). |
| **Documentation** | Architecture, usage, security, troubleshooting, development docs maintained alongside the code. Plus a public skill file and slash commands for AI consumers. |

## 7. Out of scope (now)

- Calendar attachments helper (you can pass them in raw JSON today,
  but no convenience flag).
- `findMeetingTimes` wrapper (free/busy is supported, scheduling
  suggestions are not).
- Categories / flags / mark-as-read mutations on messages.
- Conversation-thread views (the API supports threading; the CLI
  doesn't yet surface it as a first-class concept).
- Multi-account support (one signed-in user per machine).
- Daemon mode for sub-100ms cold-call latency (architecture supports
  it; not built).
- MCP server wrapper for first-class agent tool integration.
- Windows-native install path (works in WSL today).

## 8. Success metrics

For a personal-spike project, success is qualitative:

1. **Adoption signal.** Owner uses the CLI > 5 times/day in normal work
   for two consecutive weeks without falling back to Outlook web.
2. **Agent quality signal.** Claude Code / Codex sessions that involve
   mail or calendar handle the task end-to-end (read → draft → user
   reviews → user sends) ≥ 80% of the time without the user having to
   correct the agent's plan.
3. **Stability signal.** Less than one Microsoft-side breakage per
   quarter requiring a re-discovery of endpoints via
   `node src/diagnose.mjs`.
4. **Privacy hygiene.** Zero PII leaks in committed code (audited at
   every commit; verified via grep at the time of writing).

## 9. Risks and mitigations

| Risk | Mitigation |
| --- | --- |
| Microsoft retires the Outlook REST v2 endpoint (sunset is announced). | Diagnostic sniffer in `src/diagnose.mjs` lets us re-discover the new endpoints. Architecture supports porting to Graph by re-pointing `REST_BASE`. |
| Tenant tightens Conditional Access and breaks even the silent-SSO path. | User runs `outlook auth` interactively; this is the same MFA flow they already do in their browser. |
| Token-cache file leaks (lost laptop, accidental sync to iCloud Drive). | `docs/security.md` documents file-permission posture; `outlook logout` clears the cache; signing out of OWA invalidates further refreshes. |
| AI agent sends mail or creates a meeting with attendees without confirmation. | `SKILL.md` codifies a strict "always confirm recipients + subject + body before send / event-create / event-cancel on attendee-bearing events" rule. Drafts (`draft-reply`) are preferred over `send` for AI-composed mail. |
| Endpoint change silently degrades behaviour (e.g. new required header). | `diagnose.mjs` + integration tests against a mock server catch shape changes early. Real-world breaks surface as `E_HTTP` with the server's error body. |
| EDR / DLP on a managed device flags the Bearer-capture pattern. | `docs/security.md` explicitly recommends telling your security team in advance. Behaviour matches MITRE T1528 description; this is transparency, not stealth. |

## 10. Open questions

- **Daemon mode UX.** Background browser process vs on-demand short
  lifecycles — which gives the better failure mode when the user's
  laptop sleeps mid-call?
- **MCP wrapper priority.** Slash commands cover the same surface
  today; is MCP worth the additional surface area for the latency win?
- **Learnings TTL.** Should observations age out automatically (e.g.
  after 90 days untouched) or stay until the user prunes? Currently
  the latter.
- **HTML reply bodies.** Today `Comment` in createReply is plain text.
  Should we add `--html` to do a GET-then-merge with the server-composed
  HTML thread for users who want rich formatting?
- **Calendar timezone handling.** Today users supply
  `{DateTime, TimeZone}` themselves. Worth adding a `--timezone`
  default + sugar like `--start "Friday 2pm"`?

## 11. Decision log

Key decisions made during the spike, in chronological order. The full
chronicle of failed approaches is in [`LEARNINGS.md`](../LEARNINGS.md).

| Date | Decision | Rationale |
| --- | --- | --- |
| 2026-05-18 | Use Bearer-from-the-wire capture instead of any OAuth flow | Tenant blocks admin consent on Mail scopes; this is the only path that doesn't need IT. |
| 2026-05-18 | Always run Chromium headed (never headless) | Conditional Access stalls headless silent-SSO indefinitely. |
| 2026-05-18 | Cache captured headers to disk keyed on JWT `exp` | Sub-second cache-hit latency; only flash a browser window once per ~24h. |
| 2026-05-18 | Default to draft-* over send for AI-composed mail | Safety: review-first reduces the cost of agent mistakes. |
| 2026-05-18 | Pass reply text through `Comment` in createReply, never PATCH `Body` | PATCH wipes the server-composed quoted thread. Comment preserves it. |
| 2026-05-18 | Add a learnings store the agent reads/writes between sessions | Self-improvement over time without per-session re-config. |
| 2026-05-18 | Calendar mutations have no draft equivalent — SKILL enforces user confirmation | API doesn't expose a draft workflow; we work around with policy at the agent layer. |

---

For implementation detail see [`architecture.md`](./architecture.md).
For the why-not-other-approaches list see [`../LEARNINGS.md`](../LEARNINGS.md).
For the playbook to apply this technique to other web apps see
[`reverse-engineering.md`](./reverse-engineering.md).
