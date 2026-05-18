---
name: outlook
description: Read, search, triage, draft and send mail through the user's Outlook / Microsoft 365 mailbox. Use when the user asks to check their inbox, look up a specific email, summarise unread messages, find what someone sent, draft a reply to a specific message that they can review before sending, or send a message via Outlook. Prefer drafts (`draft`, `draft-reply`, `draft-reply-all`, `draft-forward`) over direct send so the user can review in Outlook first. Requires the `outlook` CLI to be installed on the user's PATH; if it is not, tell the user and stop. All output is JSON to stdout; pipe through `jq` for transformation. This skill does NOT cover Gmail, IMAP, or other mail providers.
---

# Using the `outlook` CLI

A locally-installed CLI that wraps the Outlook REST API by piggybacking on
the user's existing Outlook Web (OWA) session. There is no API key, no
account to configure — if it's installed, it Just Works.

If `outlook --version` fails, stop and tell the user the CLI isn't
installed. Do not attempt to install it autonomously.

## Quick reference

```
# Reading
outlook list    [filters]            # list messages (default: 10 newest from Inbox)
outlook unread  [filters]            # list --unread --top 25
outlook search  "<query>" [filters]  # KQL full-text search
outlook read    <id>                 # full message JSON
outlook folders                       # list mail folders

# Drafting (review-first — preferred for AI-generated mail)
outlook draft             <json>            # new draft (reads STDIN if no arg)
outlook draft-reply       <id>  [json]      # draft reply to a message
outlook draft-reply-all   <id>  [json]      # draft reply-all
outlook draft-forward     <id>  [json]      # draft a forward
outlook discard-draft     <id>              # delete a draft

# Sending (skips review — only after explicit user confirmation)
outlook send                          # send mail directly (STDIN JSON)

# Lifecycle
outlook auth                          # interactive sign-in (USER-driven only)
outlook logout                        # clear cached token
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

### Drafting a reply

The Outlook REST endpoint does the heavy lifting — it creates a draft
already addressed to the right recipient with the original thread quoted
underneath. You only need to provide your new reply body:

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

## Verifying the CLI is available

A single command tells you everything:

```bash
outlook --version 2>/dev/null || echo "not installed"
```

If output is a semver, you're good. If it's "not installed" or the command
errors, stop and tell the user.
