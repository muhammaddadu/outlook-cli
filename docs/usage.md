# Usage

Every command, every flag, real examples. Run `outlook --help` for the
machine-generated version.

## Global flags

| Flag | Effect |
| --- | --- |
| `--debug` | Verbose diagnostics on stderr (`OUTLOOK_DEBUG=1`) |
| `-V`, `--version` | Print the CLI version |
| `-h`, `--help` | Print help (works on any subcommand too) |

Environment variables:

| Variable | Effect |
| --- | --- |
| `OUTLOOK_DEBUG=1` | Same as `--debug` |
| `OUTLOOK_PROFILE=/path` | Override the Chromium user-data-dir |
| `OUTLOOK_TOKEN_CACHE=/path` | Override the cached-headers JSON file |
| `XDG_DATA_HOME` / `XDG_CACHE_HOME` | Standard XDG roots |

## Subcommands

### `outlook auth`

Interactive sign-in. Opens a Chromium window pointed at OWA and waits up to
**10 minutes** for you to complete SSO + MFA. Exits once your inbox actually
renders (not just when the URL settles). Caches the captured Bearer token.

Run this once per machine; re-run when the persistent profile's cookies
expire (typically weeks–months) or after a tenant-driven re-auth event.

### `outlook refresh`

Force a fresh Bearer capture even if the cache is still valid. Opens
Chromium briefly. Use when the on-disk token looks broken or after a CA
policy change that invalidated the existing JWT.

### `outlook logout`

Delete the local token cache. Does **not** sign you out of OWA — the
browser profile cookies remain. Next `outlook list` will re-capture a token
silently from the existing session.

### `outlook list [filters…]`

List messages from a folder. Default: 10 most recent from the Inbox.

```bash
outlook list                              # 10 most recent
outlook list -n 25                        # 25 most recent
outlook list --unread                     # unread only
outlook list --from alice@example.com     # specific sender
outlook list --since 7d                   # last 7 days (also: 24h, 2w, ISO date)
outlook list --folder Sent -n 5           # 5 most recent from Sent Items
outlook list --has-attachments --since 30d
outlook list --importance High
outlook list --subject "deploy"           # subject contains "deploy"
outlook list --skip 50 -n 25              # pagination
outlook list --select "Id,Subject"        # custom field set
outlook list --filter "Categories/any(c: c eq 'red')"  # raw OData escape hatch
```

Filter flags are ANDed together. The `--filter` raw expression is the
escape hatch for anything not covered by a friendly flag.

**Folder names** (case-insensitive, spaces stripped): `Inbox`, `Drafts`,
`Sent` / `Sent Items`, `Deleted` / `Deleted Items`, `Junk` / `Junk Email`,
`Outbox`, `Archive`. You can also pass a folder Id from `outlook folders`.

**Time syntax** for `--since` / `--until`:
- Relative: `15m`, `24h`, `7d`, `2w`
- Absolute: any ISO-parseable string — `2026-01-15`, `2026-01-15T10:00:00Z`

### `outlook unread`

Shortcut for `outlook list --unread --top 25`. All other list flags are
available too (`--folder`, `--from`, `--since`, …).

```bash
outlook unread                        # 25 most recent unread
outlook unread --from boss@example.com    # unread from your boss
outlook unread --folder Archive       # unread in Archive
```

### `outlook search <query> [filters…]`

Search messages with a KQL-style query. The query is automatically wrapped
in OData literal quotes. Accepts the same filter flags as `list`.

```bash
outlook search "quarterly review"
outlook search "from:jane.doe@example.com" -n 5
outlook search "deploy" --since 30d --has-attachments
outlook search "contract" --folder Archive
```

> `$search` and `$orderby` are mutually exclusive server-side. When you
> search, results come back scored, not date-sorted; the CLI drops
> `--order-by` automatically.

### `outlook read <messageId>`

Fetch a single message in full — body, attachments metadata, headers, etc.

```bash
MSG=$(outlook list -n 1 | jq -r '.value[0].Id')
outlook read "$MSG" | jq '.Body.Content' -r
```

### `outlook folders`

List mail folders (Inbox, Sent Items, Drafts, custom folders…).

```bash
outlook folders | jq '.value[] | {DisplayName, UnreadItemCount}'
```

### `outlook draft [json]`

Create a **new draft** message. Same JSON shape as `send` but the message
lands in the Drafts folder for you to review and send from Outlook.

```bash
cat <<'JSON' | outlook draft
{
  "Subject": "Q3 plan",
  "Body": { "ContentType": "Text", "Content": "Draft body…" },
  "ToRecipients": [{ "EmailAddress": { "Address": "team@example.com" } }]
}
JSON
```

Output includes the new `DraftId` and a `WebLink` you can click to open
the draft in Outlook web.

### `outlook draft-reply <id> [json]`

Create a draft reply to an existing message. Outlook fills in the quoted
thread automatically; the optional JSON override sets your new body or
extra recipients:

```bash
outlook draft-reply $ID '{
  "Body": { "ContentType": "Text", "Content": "Looks good — confirming." }
}'
```

### `outlook draft-reply-all <id> [json]`

Same as `draft-reply` but addresses every original recipient.

### `outlook draft-forward <id> [json]`

Create a draft forward. Usually you'll want the JSON override to add
`ToRecipients`:

```bash
outlook draft-forward $ID '{
  "ToRecipients": [{ "EmailAddress": { "Address": "boss@example.com" } }],
  "Body": { "ContentType": "Text", "Content": "FYI" }
}'
```

### `outlook discard-draft <id>`

Permanently delete a draft. No confirmation prompt — call only when you're
sure.

## Calendar

### `outlook agenda [options]`

Show events in a time window. Uses `/calendarView` so recurring events
are expanded into individual instances.

| Flag | Default | Effect |
| --- | --- | --- |
| `--from <when>` | `now` | Window start (ISO, `today`, `tomorrow`, `-1d`, etc.) |
| `--to <when>` | start + days | Window end |
| `--days <n>` | `7` | Shortcut for `--to: start + n days` |
| `--calendar <id>` | primary | Non-primary calendar Id |
| `--organizer <addr>` | — | Only events organised by this person |
| `--subject <text>` | — | Subject contains substring |
| `--show-as <state>` | — | `Free` / `Tentative` / `Busy` / `Oof` / `WorkingElsewhere` |
| `-n, --top <N>` | `50` | Max events to fetch |
| `-s, --skip <N>` | — | Pagination offset |
| `--filter <odata>` | — | Raw $filter (ANDed with the friendly flags) |
| `--order-by <expr>` | `Start/DateTime asc` | OData $orderby |
| `--select <fields>` | sensible default | CSV of fields to return |

```bash
outlook agenda                          # next 7 days
outlook agenda --days 1                 # today + a bit
outlook agenda --from today --to tomorrow
outlook agenda --organizer boss@example.com
outlook agenda --subject "1:1"
```

### `outlook events [options]`

Generic event list — does NOT expand recurring instances. Use for
filter-heavy queries against the raw `/events` collection (recurring
masters, cancelled events, etc.). Same filter flags as `agenda` minus the
time-range ones, plus `--all-day` and `--cancelled`.

### `outlook event-read <id>`

Full event detail — body, attachments metadata, recurrence rule, etc.

### `outlook calendars`

List the user's calendars (primary, secondary, shared).

### `outlook event-create [json]`

Create an event. The JSON is the Outlook Event resource:

```bash
cat <<JSON | outlook event-create
{
  "Subject": "Focus block",
  "Start": { "DateTime": "2026-05-25T14:00:00", "TimeZone": "America/New_York" },
  "End":   { "DateTime": "2026-05-25T15:30:00", "TimeZone": "America/New_York" },
  "ShowAs": "Busy"
}
JSON
```

> **⚠ Invitations send immediately when `Attendees` is non-empty.** There
> is no "draft" workflow for meetings. Confirm with attendees / yourself
> before adding people to the payload.

### `outlook event-update <id> [json]`

PATCH the event with the provided partial Event JSON.

### `outlook event-cancel <id>`

Cancel/delete the event. Sends cancellation notices to attendees if any.

### `outlook accept|decline|tentative <id> [options]`

RSVP to a meeting.

| Flag | Effect |
| --- | --- |
| `-c, --comment <text>` | Include a response comment |
| `--no-respond` | Don't notify the organiser of your response |

```bash
outlook accept    $ID -c "looking forward to it"
outlook decline   $ID --no-respond
outlook tentative $ID
```

### `outlook free-busy <emails…> [options]`

Look up free/busy schedule for one or more people.

```bash
outlook free-busy alice@example.com bob@example.com --from today --to tomorrow
outlook free-busy alice@example.com --interval 15
```

| Flag | Default | Effect |
| --- | --- | --- |
| `--from <when>` | `now` | Window start |
| `--to <when>` | `+24h` | Window end |
| `--interval <minutes>` | `30` | Slot granularity |

Returns `AvailabilityView` strings — `0`=free, `1`=tentative, `2`=busy,
`3`=oof, `4`=working-elsewhere. Use to suggest meeting times.

---

### `outlook send [json]`

Send a message **immediately** (no review step). Accepts the message JSON
as an argument or via STDIN.

> Prefer the `draft-*` commands when an AI agent is composing on your
> behalf — they let you review in Outlook before anything leaves your
> mailbox.

The payload shape is the Outlook REST v2 `Message` resource. Minimum
viable example:

```bash
cat <<'JSON' | outlook send
{
  "Subject": "hello from a script",
  "Body": { "ContentType": "Text", "Content": "sent via outlook-experiment" },
  "ToRecipients": [{ "EmailAddress": { "Address": "me@example.com" } }]
}
JSON
```

HTML body:

```bash
cat <<'JSON' | outlook send
{
  "Subject": "weekly digest",
  "Body": { "ContentType": "HTML", "Content": "<h1>This week</h1><p>…</p>" },
  "ToRecipients": [{ "EmailAddress": { "Address": "team@example.com" } }],
  "CcRecipients": [{ "EmailAddress": { "Address": "boss@example.com" } }]
}
JSON
```

On success, the API returns 202 with no body. The CLI prints `{ "sent":
true }`.

## Output conventions

- All command output goes to **stdout** as JSON.
- All diagnostics, progress, and errors go to **stderr**.
- Pretty-printed when stdout is a TTY, compact when piped — so `outlook
  list | jq …` Just Works.

## Exit codes

| Code | Constant | When |
| ---: | --- | --- |
| 0 | `EXIT.OK` | Success |
| 1 | `EXIT.GENERAL` | Generic / unexpected error |
| 2 | `EXIT.AUTH` | `E_AUTH_REQUIRED` or `E_AUTH_BLOCKED` |
| 3 | `EXIT.HTTP` | `E_HTTP` (4xx/5xx from the API) |
| 64 | `EXIT.USAGE` | `E_ARGS` (bad input) |
| 130 | `EXIT.SIGINT` | Interrupted with Ctrl-C |

Scripts can branch on `$?`:

```bash
outlook list -n 1 > /dev/null 2>&1
case $? in
  0)  echo "ok" ;;
  2)  echo "need to re-auth: outlook auth" ;;
  3)  echo "API problem; check connectivity" ;;
  *)  echo "unknown failure ($?)" ;;
esac
```

## Recipes

### Triage unread mail

```bash
outlook unread -n 50 \
  | jq '.value[] | {Subject, From: .From.EmailAddress.Name}'
```

### Save last 100 message bodies for offline search

```bash
mkdir -p ~/mailbox-export
outlook list -n 100 | jq -r '.value[].Id' | while read -r id; do
  outlook read "$id" > ~/mailbox-export/"$id".json
done
```

### Periodic reminder if a specific person hasn't replied

```bash
HITS=$(outlook list --from alice@example.com --since 24h -n 1 | jq '.value | length')
[ "$HITS" -eq 0 ] && say "Alice hasn't replied yet"
```

### VIP-only digest

```bash
for vip in boss@example.com cofounder@example.com; do
  echo "=== $vip ==="
  outlook list --from "$vip" --since 7d -n 5 \
    | jq -r '.value[] | "\(.ReceivedDateTime[:10])  \(.Subject)"'
done
```

### Show messages with attachments from this month

```bash
outlook list --has-attachments --since 30d -n 50 \
  | jq '.value[] | {Subject, From: .From.EmailAddress.Name, ReceivedDateTime}'
```

### Pagination — iterate through the whole inbox 50 at a time

```bash
skip=0
while true; do
  page=$(outlook list -n 50 --skip $skip)
  count=$(echo "$page" | jq '.value | length')
  [ "$count" -eq 0 ] && break
  echo "$page" | jq -c '.value[]'
  skip=$((skip + count))
done
```

### Pipe Claude Code's output straight into a draft

```bash
claude code "draft an email to the team about the deploy" \
  | jq -Rs '{Subject: "deploy update", Body: {ContentType: "Text", Content: .}, ToRecipients: [{EmailAddress: {Address: "team@example.com"}}]}' \
  | outlook send
```
