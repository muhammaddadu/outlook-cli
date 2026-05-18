# Troubleshooting

If something's broken, run with `--debug` first — most issues are visible
in the stderr stream:

```bash
outlook --debug list 2>&1 1>/dev/null
```

## Error codes

Every CLI failure prints `Error (E_…): message\nHint: …` and exits with a
code from the table below.

| Code | Exit | Meaning | Fix |
| --- | ---: | --- | --- |
| `E_AUTH_REQUIRED` | 2 | No usable session; OWA redirected to login. | `outlook auth` |
| `E_AUTH_BLOCKED` | 2 | Browser loaded OWA but never produced a Bearer token. | See "Bearer never observed" below |
| `E_HTTP` | 3 | Outlook API returned 4xx/5xx. The hint contains the response body. | Depends on body — see "HTTP errors" below |
| `E_ARGS` | 64 | Bad input (e.g. invalid JSON for `send`). | Fix the input |
| `E_UNEXPECTED` | 1 | Something we didn't anticipate. | Re-run with `--debug` and read the stack trace |

## Common failure modes

### "Chromium opens, sits on a Microsoft login page, eventually times out"

Cookies in the persistent profile have expired (or the tenant forced
re-auth). The browser can't complete silent SSO. **Fix:**

```bash
outlook auth
```

This opens the same browser in interactive mode with a 10-minute timeout.
Complete sign-in + MFA. Subsequent calls reuse the refreshed cookies.

### "Bearer never observed within 60000ms"

`E_AUTH_BLOCKED`. The page loaded outlook.office.com but no
`Authorization: Bearer …` request fired before our deadline. Causes seen in
the wild:

1. **Network instability.** OWA is mid-handshake when the timeout fires.
   Re-run.
2. **CA policy change forcing a step-up.** OWA renders, but every API call
   bounces to login.microsoftonline.com for re-auth. Run `outlook auth` to
   complete the step-up interactively.
3. **OWA blocked by an extension or proxy.** Inspect Chromium's actual
   network in headed mode: `node src/diagnose.mjs` and watch the network
   tab manually.

### "HTTP 401 on every API call after a clean `outlook auth`"

Possibilities:

- **Captured the wrong header set.** If `client.mjs`' capture filter is
  off (e.g. you edited `FORWARDED_HEADERS`), the request OWA fires first
  may have been a non-API request and we got its headers. Re-run with
  `--debug` and inspect the captured headers in
  `~/.cache/outlook-spike/auth.json`. The `authorization` value should be a
  ~4000-character JWT.
- **Token audience mismatch.** Decode the JWT (paste it into jwt.io) and
  check `aud`. It should be `https://outlook.office.com/`. Anything else
  means OWA acquired a different audience than expected — likely Microsoft
  has changed the silent acquisition flow.

### "HTTP 404 / Gone on `/api/v2.0/me/messages`"

Microsoft is sunsetting Outlook REST v2 in favour of Graph. If your tenant
has been migrated, the v2 endpoint is gone.

**Fix (involves code changes):**

1. Run `node src/diagnose.mjs` to see the current endpoints OWA uses.
2. Likely targets: `https://graph.microsoft.com/v1.0/me/messages` (Graph)
   or `https://outlook.office.com/owa/service.svc?action=FindItem` (legacy
   OWA-internal).
3. Adjust `REST_BASE` and the query strings in `client.mjs` / `cli.mjs`.
4. The Bearer's audience matters — if you switch to Graph you need a Graph-
   audience token. OWA does acquire one internally; the capture filter will
   pick it up automatically once you point the request URL at Graph.

### "Chromium flashes on every call"

Token cache isn't being read. Check:

```bash
cat ~/.cache/outlook-spike/auth.json | jq .expiresAt
# Should be a Unix timestamp in milliseconds far in the future.
```

If `expiresAt` is `null` or in the past, the JWT couldn't be parsed.
Decode the bearer manually:

```bash
jq -r .headers.authorization < ~/.cache/outlook-spike/auth.json \
  | sed 's/^Bearer //' | cut -d. -f2 \
  | base64 -d 2>/dev/null | jq .exp
```

If the captured value isn't a real JWT (wrong header type, missing
segments), the capture in `client.mjs` is grabbing the wrong request.

### "Playwright: 'Executable doesn't exist'"

You skipped `npm run setup`. Run it now:

```bash
npm run setup
```

That downloads the Chromium binary Playwright drives (~150 MB into
`~/Library/Caches/ms-playwright/`).

### "AADSTS50105: Your administrator has configured…"

This is a tenant-level block on a specific Microsoft client ID, surfaced
during `outlook auth`. It does **not** apply to the OWA client we
piggyback on (`9199bf20-…`). If you see this, something is wrong with the
Playwright launch — it's somehow trying to authenticate against a different
client. Check that `HOME_URL` in `client.mjs` still points at
`https://outlook.office.com/mail/`.

## Calendar gotchas

### "Events look like they're on the wrong day"

Outlook stores event times **without a timezone offset** in `Start.DateTime`
/ `End.DateTime` — the timezone lives in a sibling field
(`Start.TimeZone`, e.g. `"Pacific Standard Time"`). When you read events,
look at both fields together. When you create events, supply both:

```json
{ "Start": { "DateTime": "2026-05-25T14:00:00", "TimeZone": "America/New_York" } }
```

Don't pass an ISO string with a `Z` suffix — Outlook treats it
literally as the wall-clock time and silently strips the offset.

### "I created an event and people got invites I didn't expect"

There is no "draft" workflow for events. The moment you `POST /events`
or `PATCH /events/{id}` with an `Attendees` array, the server sends
invitations or updates immediately. If you want a personal block
nobody's invited to, drop the `Attendees` field entirely or set it to
`[]`.

### "Cancelled an event and got an angry message about it"

`outlook event-cancel <id>` deletes the event and, if there were
attendees, sends cancellation notices. To remove a meeting from *your*
calendar only without notifying anyone (e.g. one declined on your
behalf), use `decline <id> --no-respond` instead — it removes the entry
without bothering the organiser.

### "agenda returned nothing but I know I have events"

`agenda` defaults to `now → now + 7 days`. If your events are in the
past, pass `--from "-30d"`. If they're further out, pass `--days 60`.

### "I want recurring master series, not instances"

`agenda` uses `/calendarView` which expands recurring events. To see
the master series (one row per recurring pattern), use `outlook events`
instead — it queries `/events` directly.

## Inspecting state

```bash
# Cached headers (DO NOT SHARE)
cat ~/.cache/outlook-spike/auth.json | jq

# Browser profile size
du -sh ~/.local/share/outlook-spike/browser-profile/

# Token expiry in human time
jq -r .expiresAt < ~/.cache/outlook-spike/auth.json | xargs -I{} date -r {}
```

## Nuclear options

```bash
# Clear just the token cache (keeps cookies)
outlook logout

# Clear cookies but keep cache (you'll fail silent SSO until you re-auth)
rm -rf ~/.local/share/outlook-spike/browser-profile/

# Burn it all down
rm -rf ~/.local/share/outlook-spike/ ~/.cache/outlook-spike/
outlook auth
```
