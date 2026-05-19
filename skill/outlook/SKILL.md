---
name: outlook
description: Read, search, triage, draft, and send mail PLUS view and manage calendar events through the user's Outlook / Microsoft 365 mailbox. Use when the user asks to check inbox, summarise unread mail, find a specific email, draft a reply they can review, send mail, view their calendar / agenda, find what meetings they have today / this week, create or update calendar events, RSVP to invites, or check free/busy. Prefer drafts (`draft`, `draft-reply`, `draft-reply-all`, `draft-forward`) over direct send so the user can review in Outlook first. Confirm with the user before creating or cancelling events with attendees (invitations / cancellations send IMMEDIATELY). Requires the `outlook` CLI to be installed on the user's PATH; if it is not, tell the user and stop. All output is JSON to stdout; pipe through `jq` for transformation. This skill does NOT cover Gmail, IMAP, or other mail providers.
---

# Using the `outlook` CLI

A locally-installed CLI that wraps the Outlook REST API by piggybacking on
the user's existing Outlook Web (OWA) session. There is no API key, no
account to configure — if it's installed, it Just Works.

If `outlook --version` fails, stop and tell the user the CLI isn't
installed. Do not attempt to install it autonomously.

## Start every session with `outlook context`

Run `outlook context` **before your first mail-related action** in a
session. It's free (no network, no Chromium) and returns:

```json
{
  "user": { "email": "…", "name": "…", "tenant_id": "…" },
  "tokenMinutesUntilExpiry": 1408,
  "learnings": [
    "2026-05-18 | Signs off as Sam",
    "2026-05-18 | Prefers terse one-line replies",
    "2026-05-18 | 'the team' usually means team-eng@example.com"
  ],
  "learningsFile": "/Users/…/.local/share/outlook-spike/learnings.md"
}
```

Use `user.email` to know whose mailbox you're acting on. Use `learnings`
to adapt your tone, recipient resolution, and default behaviour to what
the user has shown they prefer.

## Adapt over time with `outlook learn`

When you observe something durable, non-obvious, and useful about the
user's mail habits, append it:

```bash
outlook learn add "Signs off as 'Sam'"
outlook learn add "Drafts to vendors should be HTML with the legal footer"
outlook learn add "When user says 'my boss' they mean alice@example.com"
```

Good learnings have all three properties:
- **Durable** — true today and likely true next week. Not "the user is
  in a hurry right now."
- **Non-obvious** — a fact you couldn't trivially re-derive next time.
  Skip "user has an inbox."
- **Useful** — informs how you'd act next time. "User likes pizza" is
  cute but doesn't change your behaviour.

Don't pollute. A good rule: < 30 entries total. If learnings get noisy,
suggest the user run `outlook learn forget <substring>` to prune.

**Never record sensitive content without asking** (financial details, HR
context, anything the user might not want a future agent to see). If in
doubt, ask: "Want me to remember this for next time?"

## Quick reference

```
# Reading
outlook list    [filters]            # list messages (default: 10 newest from Inbox)
outlook unread  [filters]            # list --unread --top 25
outlook search  "<query>" [filters]  # KQL full-text search
outlook read    <id>                 # full message JSON
outlook folders                       # list mail folders

# Drafting (review-first — preferred for AI-generated mail)
outlook draft             <json>  [--attach <path>…]   # new draft (reads STDIN if no arg)
outlook draft-reply       <id>  [json] [--attach <path>…]   # draft reply to a message
outlook draft-reply-all   <id>  [json] [--attach <path>…]   # draft reply-all
outlook draft-forward     <id>  [json] [--attach <path>…]   # draft a forward
outlook discard-draft     <id>              # delete a draft

# Sending (skips review — only after explicit user confirmation)
outlook send                       [--attach <path>…]   # send mail directly (STDIN JSON)

# Lifecycle
outlook auth                          # interactive sign-in (USER-driven only)
outlook logout                        # clear cached token

# Calendar — reading
outlook agenda    [--from --to --days]     # what's on the calendar (expanded recurring instances)
outlook events    [filters]                # generic event list (recurring masters, not instances)
outlook event-read <id>                    # full event detail
outlook calendars                           # list user's calendars

# Calendar — writing  (events with attendees send invites IMMEDIATELY)
outlook event-create   <json>              # create event
outlook event-update   <id> <json>         # PATCH event
outlook event-cancel   <id>                # delete event

# Calendar — RSVP
outlook accept     <id> [-c "<comment>"]   # RSVP yes
outlook decline    <id> [-c "<comment>"]   # RSVP no
outlook tentative  <id> [-c "<comment>"]   # RSVP maybe

# Calendar — availability
outlook free-busy <email> [<email>…] [--from --to --interval]

# Self-learning (call at session start, then append observations)
outlook context                       # user info + accumulated learnings (no network)
outlook learn                         # list current learnings
outlook learn add "<observation>"     # record a durable, useful observation
outlook learn forget "<substring>"    # remove matching learnings
outlook learn clear                   # wipe everything
```

## Filters (shared by `list`, `unread`, and `search`)

| Flag | Meaning |
| --- | --- |
| `--unread` | Only unread messages |
| `--from <addr>` | Sender match (exact email address) |
| `--to <addr>` | Recipient match |
| `--since <when>` | Received after — `15m`, `24h`, `7d`, `2w`, or ISO date |
| `--until <when>` | Received before — same syntax as `--since` |
| `--has-attachments` | Only messages with attachments |
| `--importance Low\|Normal\|High` | Filter by importance |
| `--subject "<text>"` | Subject contains substring (case-insensitive) |
| `--folder <name>` | `Inbox`, `Sent`, `Drafts`, `Deleted`, `Junk`, `Outbox`, `Archive`, or a folder Id |
| `--filter "<odata>"` | Raw OData $filter — escape hatch for anything above |
| `-n, --top <N>` | How many results (defaults vary per command) |
| `-s, --skip <N>` | Pagination offset |
| `--select "Id,Subject"` | Override the default field set |
| `--order-by "<expr>"` | OData $orderby (default: `ReceivedDateTime desc`) |

Flags are ANDed. The order on the command line does not matter.

## Reading mail

```bash
outlook list -n 5                                  # 5 most recent
outlook unread --from boss@example.com             # unread from boss
outlook list --since 24h --has-attachments         # mail with attachments today
outlook list --folder Sent --since 7d              # what the user sent this week
outlook search "deploy" --since 30d -n 10          # full-text + filter
```

Each result item in `.value[]` has at minimum: `Id`, `Subject`, `From`,
`ReceivedDateTime`, `BodyPreview`, `IsRead`, `HasAttachments`,
`Importance`. The `BodyPreview` is ~150 chars — use it for summarisation,
but call `read <id>` if the user wants the full content.

```bash
outlook read AAMkADQxZDQ1.../=  | jq -r '.Body.Content'
```

## Searching effectively

`search` uses Outlook's full-text engine (KQL-style). Quote multi-word
queries:

```bash
outlook search "quarterly review"
```

For sender or attachment filtering, prefer the flag form — it's faster and
more reliable than putting `from:` into the query string:

```bash
# Good
outlook list --from alice@example.com --since 7d

# Works but slower / less precise
outlook search "from:alice"
```

## Drafting mail (review-first — STRONGLY PREFERRED)

For anything you draft on the user's behalf, use `draft-*` and let them
review in Outlook before sending. Don't go straight to `send` unless the
user has explicitly approved the recipient, subject, AND body.

### Voice & tone — sound like the user, not like an AI

The point of a draft is for the user to send it as-is. If they have to
rewrite your output to make it sound like them, you've failed. Rules:

- **Preserve the user's exact phrasing when they dictate.** If they say
  "draft a reply saying X", X goes in verbatim. Don't "correct"
  grammar, expand contractions, or rephrase for clarity. "login" stays
  "login". "I'm I" stays "I'm I" unless they ask you to clean it up.
- **Keep contractions and the user's register.** Don't expand "I'm" to
  "I am", "can't" to "cannot", "we'll" to "we will". Match how the user
  actually writes in this thread.
- **Mirror the recipient's style.** Read the last 1-2 messages in the
  thread before drafting. Terse thread → terse reply. Chatty thread →
  matching warmth. A two-line "Thanks, looks good" is often the right
  answer, not a three-paragraph essay.
- **Avoid AI tells.** No "I hope this email finds you well." No "Just
  to clarify" openers. No em-dash-heavy hedging. No bulleted breakdowns
  for a one-sentence ask. No closing "Please let me know if you have
  any questions." unless the user already writes that way.
- **When in doubt, shorter.** Most work email is too long. If the
  draft can lose a sentence without losing meaning, lose it.

This applies to `draft`, `draft-reply`, `draft-reply-all`,
`draft-forward`, and any `send` you compose for the user.

### Drafting a reply

The Outlook REST endpoint does the heavy lifting — it creates a draft
already addressed to the right recipient with the original thread quoted
underneath. You only need to provide your new reply body.

Under the hood the CLI sends your text via the API's `Comment` field so
the server can compose `<your reply>\n\n<quoted thread>` correctly. **Do
not try to PATCH `Body` directly — that wipes the thread.** If you need
to set non-body fields (Cc, Subject changes), pass them alongside `Body`
in the override JSON; the CLI splits them and routes each correctly.

```bash
# Find the message you're replying to.
ID=$(outlook search "from:alice@example.com deploy" -n 1 | jq -r '.value[0].Id')

# Create the draft with your AI-generated body.
outlook draft-reply "$ID" '{
  "Body": { "ContentType": "Text", "Content": "Thanks Alice — confirming the rollout for Tuesday 10am." }
}'
```

Output includes the draft `Id` and a `WebLink` you can give the user to
open the draft in Outlook web (or just tell them "check your Drafts
folder").

Use `draft-reply-all` to address everyone on the thread, `draft-forward`
to forward to someone new (you'll need to add `ToRecipients` in the
override JSON).

### Attaching files

Pass `--attach <path>` (repeatable) on `draft`, `draft-reply`,
`draft-reply-all`, `draft-forward`, or `send`. Files are read from the
local filesystem, base64-encoded, and uploaded as Outlook
`FileAttachment` resources. The cap is ~3 MB per file (Outlook
inline-attachment ceiling); larger files need OneDrive.

```bash
# Reply with two screenshots
outlook draft-reply "$ID" \
  --attach ~/Pictures/screenshot1.png \
  --attach ~/Pictures/screenshot2.png \
  '{ "Body": { "ContentType": "Text", "Content": "See attached." } }'

# Send a quick note with a PDF
cat <<'EOF' | outlook send --attach ./report.pdf
{ "Subject": "Q1 report", "Body": {"ContentType":"Text","Content":"Attached."},
  "ToRecipients": [{"EmailAddress":{"Address":"boss@example.com"}}] }
EOF
```

Bad paths fail fast with `E_ARGS` (exit 64) *before* the draft is
created, so a typo will not leave an orphan empty draft in the user's
mailbox.

### Drafting a new message from scratch

```bash
cat <<EOF | outlook draft
{
  "Subject": "<subject>",
  "Body": { "ContentType": "Text", "Content": "<body>" },
  "ToRecipients": [{ "EmailAddress": { "Address": "to@example.com" } }]
}
EOF
```

The draft lands in the Drafts folder. The user reviews and hits Send in
Outlook.

### Discarding a draft

If the user changes their mind:

```bash
outlook discard-draft <draftId>
```

## Calendar

### Reading the calendar

```bash
outlook agenda                          # next 7 days, primary calendar
outlook agenda --days 14                # next 2 weeks
outlook agenda --from today --to tomorrow   # just today
outlook agenda --from "+7d" --to "+14d"     # week 2 from now
outlook agenda --calendar <calendarId>      # a specific (non-primary) calendar
outlook agenda --organizer alice@example.com  # only meetings Alice runs
outlook agenda --subject "standup"           # filter by subject substring
```

`agenda` uses `/calendarView` under the hood, so recurring events are
expanded into individual instances — which is what humans expect when
they say "what's on my calendar." Use `outlook events` for filter-heavy
non-time-bound queries (recurring master series, cancelled events, etc.).

Each result row has at least: `Id`, `Subject`, `Start`, `End`, `Location`,
`Organizer`, `Attendees`, `IsAllDay`, `ShowAs`, `IsCancelled`,
`IsOnlineMeeting`, `OnlineMeetingUrl`, `ResponseStatus`. Call
`outlook event-read <id>` for the full body / attachments / reminders.

### Creating calendar events

The full event payload is the Outlook Event resource. Minimum viable
example (no attendees → calendar block only):

```bash
cat <<EOF | outlook event-create
{
  "Subject": "Focus block",
  "Start": { "DateTime": "2026-05-25T14:00:00", "TimeZone": "Pacific Standard Time" },
  "End":   { "DateTime": "2026-05-25T15:30:00", "TimeZone": "Pacific Standard Time" },
  "ShowAs": "Busy"
}
EOF
```

With attendees → **invitations send immediately**. Always confirm with the
user first:

```bash
cat <<EOF | outlook event-create
{
  "Subject": "1:1 with Alice",
  "Start": { "DateTime": "2026-05-25T10:00:00", "TimeZone": "Pacific Standard Time" },
  "End":   { "DateTime": "2026-05-25T10:30:00", "TimeZone": "Pacific Standard Time" },
  "Attendees": [
    { "EmailAddress": { "Address": "alice@example.com" }, "Type": "Required" }
  ],
  "IsOnlineMeeting": true,
  "OnlineMeetingProvider": "teamsForBusiness"
}
EOF
```

### Updating / cancelling events

```bash
outlook event-update <id> '{"Subject":"renamed", "Location":{"DisplayName":"Room 12"}}'
outlook event-cancel <id>     # if event has attendees, sends cancellation notices
```

### RSVP

```bash
outlook accept    <id>                       # quiet accept
outlook accept    <id> -c "looking forward"  # accept with comment
outlook decline   <id> --no-respond           # decline without notifying organiser
outlook tentative <id>
```

### Free/busy lookup

```bash
outlook free-busy alice@example.com bob@example.com --from today --to tomorrow
outlook free-busy alice@example.com --interval 15
```

Returns an `AvailabilityView` string per person — characters indicate
free (`0`), tentative (`1`), busy (`2`), oof (`3`), working-elsewhere
(`4`) at each `interval`-minute slot. Use this to suggest meeting times.

### Calendar safety rules

- **Events with attendees send invitations / updates / cancellations
  immediately.** No "draft" workflow exists for meetings. Always
  confirm recipient list + time + subject + body with the user before
  calling `event-create` or `event-update` on something that has
  `Attendees`.
- **Cancelling an event with attendees notifies them.** Confirm before
  `event-cancel`.
- For personal calendar blocks (no Attendees), feel free to create
  without confirmation if the user clearly asked for one ("block 2-3pm
  Friday for deep work").

## Sending mail directly (skips review)

**Only use `send` when the user has explicitly confirmed the recipient,
subject, and body in plain English. Otherwise prefer the `draft-*`
commands above.**

`send` reads the message JSON from STDIN — same shape as `draft`:

```bash
cat <<'EOF' | outlook send
{
  "Subject": "<subject>",
  "Body": { "ContentType": "Text", "Content": "<body>" },
  "ToRecipients": [{ "EmailAddress": { "Address": "to@example.com" } }]
}
EOF
```

For HTML: `"ContentType": "HTML"` with HTML in `Content`. The server
returns 202 and the CLI prints `{ "sent": true }`.

## Output conventions

- stdout: JSON only — pipe-safe.
- stderr: diagnostics; ignore unless something failed.
- Empty `.value` array = no matches; this is **not** an error.

When piping into `jq`, always redirect stderr to `/dev/null` if you want
silent operation:

```bash
outlook list --unread -n 50 2>/dev/null | jq '.value | length'
```

## Exit codes

| Code | Meaning | What to do |
| ---: | --- | --- |
| 0 | Success | — |
| 2 | `E_AUTH_REQUIRED` / `E_AUTH_BLOCKED`. Session is gone. | **Ask the user to run `outlook auth` manually.** You cannot run `outlook auth` yourself — it opens a browser window the user must interact with for MFA. |
| 3 | `E_HTTP`. API returned 4xx/5xx. The stderr block includes the response body. | Read the body, decide whether to retry or surface the error. |
| 64 | `E_ARGS`. Bad input. | Fix the flag / JSON and retry. |
| 130 | User pressed Ctrl-C. | Stop. |

When a command exits with 2, stop and surface the recovery instruction
clearly. Do not loop trying to re-authenticate.

## When to use this skill

- "Did anyone email me about X?" → `outlook search "X"`
- "What's in my inbox?" → `outlook list -n 10`
- "Anything unread from <person>?" → `outlook unread --from <addr>`
- (always) Start with `outlook context` to load identity + learnings
- "Draft a reply to <message>" → `outlook draft-reply <id> '<json with Body>'`
- "Send a reply to <message>" → confirm details, then `outlook draft-reply` (preferred) or `outlook send` if user demands direct send
- "When did <person> last email me?" → `outlook list --from <addr> -n 1`
- "Find emails from this week with attachments" → `outlook list --since 7d --has-attachments`
- "What did I send to <person> last month?" → `outlook list --folder Sent --to <addr> --since 30d`

## When NOT to use this skill

- Gmail / IMAP / non-Microsoft mail (this CLI is Outlook-specific).
- Calendar events (the CLI doesn't expose /events yet).
- Bulk mail operations on thousands of messages without explicit user
  consent — pagination makes this possible but you should confirm scope
  first.
- Any send operation where the recipient, subject, or body wasn't
  explicitly confirmed by the user. **Never send mail without user
  confirmation of all three.**

## Don'ts

- Don't attempt `outlook auth` autonomously. It needs the user.
- Don't dump full message `Body.Content` into the conversation without
  asking — emails can be tens of KB. Summarise first.
- Don't `outlook send` from an unattended script flow. Always confirm
  recipient + subject + body with the user, in plain English, before
  sending. When in doubt, use `draft-reply` / `draft` and let the user
  send from Outlook.
- Don't suggest installation or configuration steps. The CLI is either on
  PATH or it isn't.

## Examples of good learnings

Things worth recording with `outlook learn add "…"`:

- "Signs off as 'Sam' (no full name in informal replies)"
- "Prefers terse 1-2 sentence replies; only goes long for vendor / external"
- "'The team' = team-eng@example.com (mailing list)"
- "Boss is alice@example.com — replies to her are usually short and direct"
- "External vendor replies should be HTML with the standard legal footer"
- "Always BCC archive@example.com on customer-facing mail"
- "Don't send mail before 9am or after 7pm without explicit confirmation"

Things NOT worth recording:

- "User checked their inbox today"  (transient, useless next time)
- "User has 12 unread emails right now"  (will be wrong tomorrow)
- "User likes pizza"  (true but unrelated to mail)
- "User is upset"  (transient, sensitive)

## Verifying the CLI is available

A single command tells you everything:

```bash
outlook --version 2>/dev/null || echo "not installed"
```

If output is a semver, you're good. If it's "not installed" or the command
errors, stop and tell the user.
