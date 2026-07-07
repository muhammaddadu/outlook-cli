# Architecture

How the CLI manages to talk to Outlook without an app registration or admin
consent — and the design decisions that fall out of that constraint.

## The core trick: capture OWA's own Bearer token

Outlook Web (OWA) is itself a Microsoft-blessed first-party app
(`client_id=9199bf20-a13f-4107-85dc-02114787ef48`). When you load
`outlook.office.com/mail/`, the page runs an in-browser silent OAuth
handshake against `login.microsoftonline.com` and ends up with a JWT whose
audience is `https://outlook.office.com/`. OWA then attaches that JWT —
plus a constellation of routing headers — to every request it makes to
`outlook.office.com/api/...`.

That JWT is exactly the token the Outlook REST v2 endpoint expects. So we:

1. Launch Chromium ourselves with a persistent profile.
2. Navigate to OWA. OWA does its handshake and starts making API calls.
3. Watch the page's outgoing requests. Snapshot the first request that has
   `Authorization: Bearer …` going to `outlook.office.com`.
4. Save the Authorization header **and the routing headers OWA always
   includes** to disk.
5. From that point on, make API calls from Node `fetch` with those headers.
   Skip Chromium entirely.

No new permission is created. No admin consent screen ever fires. The
scopes are whatever the user already consented to when they first signed
into Outlook web.

```
┌────────────────────────┐   1. Silent OAuth   ┌───────────────────────┐
│ Chromium (persistent)  │ ──────────────────▶ │ login.microsoftonline │
│ outlook.office.com/mail│ ◀────── JWT ─────── │ Bearer token issued   │
└─────────┬──────────────┘                     └───────────────────────┘
          │
          │ 2. Network interception
          ▼
┌────────────────────────┐  3. Save to disk    ┌───────────────────────┐
│ captureAuth()          │ ──────────────────▶ │ ~/.cache/outlook-spike│
│ extracts headers       │                     │ /auth.json            │
└────────────────────────┘                     └───────────────────────┘
                                                          │
                                                          │ 4. Reuse until exp
                                                          ▼
                                               ┌───────────────────────┐
                                               │ Node fetch() with     │
                                               │ cached headers →      │
                                               │ outlook.office.com/   │
                                               │ api/v2.0/me/messages  │
                                               └───────────────────────┘
```

## The full set of headers we capture

OWA sends more than just `Authorization` on every authed request. The
backend uses several other headers for routing and consistency. Forwarding
all of them isn't strictly required (Authorization + x-anchormailbox is the
minimum), but matching OWA exactly avoids throttling and version-skew
issues.

| Header | Why we forward it |
| --- | --- |
| `authorization: Bearer <jwt>` | The actual credential. Audience is `https://outlook.office.com/`. |
| `x-anchormailbox` | Routes the request to the mailbox's Exchange server shard. |
| `x-routingparameter-sessionkey` | Sticky-session affinity for the same shard. |
| `x-tenantid` | Helps the backend resolve cross-tenant access faster. |
| `x-ms-appname` | OWA reports `owa-reactmail`; matching avoids "unknown client" branches. |
| `x-clientid`, `x-owa-sessionid`, `x-client-version` | Telemetry / consistency. |
| `prefer: IdType="ImmutableId", exchange.behavior="…"` | Selects ID format + feature flags. |

The list lives at `FORWARDED_HEADERS` in `src/capture.mjs`.

## Two layers of state

| Layer | Lives at | Lifetime | Why it's separate |
| --- | --- | --- | --- |
| Browser profile | `~/.local/share/outlook-spike/browser-profile/` | Until cookies expire (weeks–months) | Re-creating it requires real interactive sign-in including MFA. Treat as data. |
| Token cache | `~/.cache/outlook-spike/auth.json` | JWT `exp` claim (~24h in our tested tenant) | Cheap to recreate from the browser profile. Treat as a cache. |

Both honour `$XDG_DATA_HOME` / `$XDG_CACHE_HOME`. Per-path overrides via
`OUTLOOK_PROFILE` and `OUTLOOK_TOKEN_CACHE`.

## Multiple resources (Outlook / Graph / Substrate)

Each Microsoft API is a distinct OAuth *resource* with its own token
audience, and a token is only accepted by its own resource — an
`outlook.office.com` token gets `401 InvalidAuthenticationToken` from Graph.
So auth is keyed by resource (`src/resources.mjs`):

| Resource | Audience | Base URL | Powers |
| --- | --- | --- | --- |
| `outlook` | `https://outlook.office.com` | `…/api/v2.0/me` | mail, calendar |
| `graph` | `https://graph.microsoft.com` | `graph.microsoft.com/v1.0` | Teams, Files, People, Groups |
| `substrate` | `https://substrate.office.com` | `substrate.office.com` | Copilot, unified search |

`captureAllTokens()` watches every Bearer the browser emits, classifies each
by audience, and keeps one header set per resource — so a single `auth --all`
session (which also opens Teams + Copilot tabs) can populate several. Each
resource caches to its own `auth-<resource>.json` (the default `outlook`
keeps the original `auth.json`). `call(auth, path, { resource })` and
`getAuth({ resource })` route to the right base with the right token;
`outlook token-audit` decodes the caches and reports what's reachable. Only
the default `outlook` resource auto-launches the browser on a cache miss —
Graph/Substrate error toward `auth --all` instead so a `graph` call never
pops Chromium unexpectedly. See [`LEARNINGS.md`](../LEARNINGS.md) §13.

When the token cache misses, we re-open Chromium against the persistent
profile. Cookies still valid → silent SSO completes in 2–3 seconds, we
capture a fresh JWT, save it, continue. Cookies expired → page redirects
to login and our 60-second timeout fires; the CLI exits with `E_AUTH_REQUIRED`
and tells you to run `outlook auth`.

## Why headed (and never headless)

Microsoft Conditional Access fingerprints headless Chromium and refuses
silent SSO. The page sits on `login.microsoftonline.com/.../authorize`
forever. We tried:

- Realistic UA strings.
- `--disable-blink-features=AutomationControlled`.
- Playwright `--headless=new`.
- Various stealth args.

None of these work in a tenant with default-strength CA. The eventual fix
will be a long-lived daemon (open the browser once, keep it alive, expose
a local socket the CLI talks to) — but for now headed-on-demand with the
on-disk cache covering ~99% of calls is good enough.

See [`LEARNINGS.md`](../LEARNINGS.md) §10 for the full investigation.

## The `call()` helper

`client.mjs` exports a `fetch` wrapper with built-in resilience:

```js
call(auth, path, { resource = 'outlook', ...init })
```

- `auth` is the header bag (`{ authorization, x-anchormailbox, … }`).
- `path` is appended to the chosen resource's base URL (default
  `https://outlook.office.com/api/v2.0/me`; `graph` → `graph.microsoft.com/v1.0`).
- Every request is bounded by `OUTLOOK_HTTP_TIMEOUT_MS` (default 30s).
- Transient failures retry with backoff: `429`/`503` (honouring
  `Retry-After`) for any method; network errors (DNS/refused/timeout) for
  GET only. Exhausted network retries raise `AppError(E.NETWORK)`.
- Always parses JSON when possible; otherwise returns the raw text.

`cli.mjs` wraps `call()` with `runApi()`, which:

- Catches 401 → clears that resource's cache, re-captures a token, and
  retries once; a second 401 raises `AppError(E.AUTH_REQUIRED)`.
- Throws `AppError(E.HTTP)` for any other 4xx/5xx, extracting the OData
  `error.code: message` for the hint.
- Returns the parsed body on success.

## Error model

All command failures raise an `AppError` from `src/errors.mjs` with a stable
code:

| Code | Meaning | Exit code |
| --- | --- | --- |
| `E_AUTH_REQUIRED` | No usable session; run `outlook auth` | 2 |
| `E_AUTH_BLOCKED` | Browser loaded OWA but never produced a Bearer | 2 |
| `E_HTTP` | API returned 4xx/5xx | 3 |
| `E_ARGS` | Bad input (e.g. invalid JSON to `send`) | 64 |
| `E_UNEXPECTED` | Anything we didn't anticipate | 1 |

Every `AppError` carries a `hint` field — short, actionable, written for the
end user, not the developer.

## CLI surface (commander)

Single binary, subcommand-based:

```
outlook auth        # interactive sign-in (10-min timeout)
outlook refresh     # force-refresh the cached Bearer
outlook logout      # clear the local token cache
outlook list        # list inbox
outlook search      # KQL search
outlook read <id>   # full message
outlook folders     # list mail folders
outlook send [json] # send mail (JSON arg or STDIN)
```

Global `--debug` flag wires through `OUTLOOK_DEBUG=1` to enable `debug()`
output to stderr. Global `--version` derived from `package.json`.

## Calendar surface

Calendar lives on the same backend as mail — same Bearer, same routing
headers, just different paths. The CLI splits queries between two
endpoints depending on the use case:

| Endpoint | Wrapped by | When |
| --- | --- | --- |
| `GET /me/calendarView?startDateTime=&endDateTime=` | `outlook agenda` | "What's on my calendar this week" — the server expands recurring events into individual instances within the window. |
| `GET /me/events` | `outlook events` | Generic event queries (filter by organizer, all-day, cancelled, etc.). Returns recurring **master series**, not instances. |
| `POST /me/events` | `outlook event-create` | Create event. With Attendees → invitations send IMMEDIATELY. No draft workflow exists for meetings. |
| `PATCH /me/events/{id}` | `outlook event-update` | Modify event. Attendees → update notification sent. |
| `DELETE /me/events/{id}` | `outlook event-cancel` | Cancel event. Attendees → cancellation notice sent. |
| `POST /me/events/{id}/{accept,decline,tentativelyAccept}` | `outlook accept` / `decline` / `tentative` | RSVP with optional comment. |
| `POST /me/getSchedule` | `outlook free-busy` | Free/busy lookup across multiple mailboxes; returns an `AvailabilityView` per person. |
| `GET /me/calendars` | `outlook calendars` | List user's calendars (primary + shared + holiday subs). |

`src/calendar.mjs` owns time-window parsing (`today`/`tomorrow`/`+7d`/
ISO), the calendarView URL builder, and an OData filter composer for
event-specific flags. The CLI command layer in `cli.mjs` is a thin
wrapper that translates flags into REST calls via `runApi`.

## Future work

- **Daemon mode.** Keep one Chromium process alive, expose a local socket,
  shave the ~3-second cache-miss latency to ~50ms.
- **MCP server wrapper.** Surface mail + calendar operations as Model
  Context Protocol tools so Claude Code / other AI agents can call them
  as first-class operations (vs. shelling out via Bash).
- **Calendar attachments.** Currently `event-create` JSON doesn't include
  attachment helpers — would be a small extension to `src/calendar.mjs`.
- **`find-meeting-times`.** Outlook exposes `/me/findMeetingTimes` to
  suggest slots when attendees are mutually free. Worth wrapping.
- **Graph fallback.** If Microsoft fully retires the v2 endpoint, port to
  Graph using the same Bearer (audience-bridging via a second silent
  token acquisition in the page).
