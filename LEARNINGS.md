# LEARNINGS

A chronicle of approaches we tried, in roughly the order we tried them, and
why each one failed in a locked-down enterprise tenant (Microsoft 365 with
Conditional Access + admin-consent-required for mail scopes).

The point of this file is to save the next person — agent or human — from
re-walking dead ends. If you propose an alternative to the current
architecture, check whether it appears below first.

The successful approach is at the bottom (§12).

---

## 1. Direct Microsoft Graph via end-user-registered Azure app

**Plan:** Register an Entra ID app yourself, add delegated `Mail.Read`,
consent for your own account, use it with MSAL / device-code flow.

**Result:** Tenant blocks self-service app registration entirely from the
portal — the App Registrations blade returned `401`. Even with the right
Graph delegated permissions baked into the token, creating new apps via
`az ad app create` returned `Insufficient privileges to complete the
operation`, despite the token claiming `Application.ReadWrite.All`.

**Takeaway:** "Can the user register applications?" is a tenant-level
directory toggle that overrides individual role permissions. Don't assume
roles in JWT claims correspond to what you can actually do.

---

## 2. Pre-consented Microsoft first-party clients (Graph CLI, Azure CLI)

**Plan:** Microsoft ships first-party clients (`14d82eec-…` Graph CLI,
`04b07795-…` Azure CLI) that are pre-consented in most tenants. Use the
Graph CLI client ID with a device-code flow requesting `Mail.Read
offline_access`.

**Result:** `AADSTS50105: Your administrator has configured the application
'Microsoft Graph Command Line Tools' to block users unless they are
specifically granted access`. The tenant admin had locked first-party
clients to assigned users only.

**Takeaway:** Even Microsoft-blessed first-party clients can be blocked at
the tenant level. The `azure-cli` client succeeded at `az login` but its
default scope set on Mac was `Application.ReadWrite.All`,
`Directory.AccessAsUser.All`, etc. — no `Mail.*`. Token scopes ≠ token
audience ≠ token-redemption scopes.

---

## 3. Modifying an existing user-owned app registration

**Plan:** The user owned three Copilot Studio-created app registrations.
Their portal blade wasn't accessible but `az ad app show` worked. Hijack
one: add `Mail.Read` and `offline_access` delegated permissions via
`az ad app update`, mark it a public client with `http://localhost` redirect,
then run device code against it.

**Result:** Permissions and redirect URI both updated cleanly via the
Graph API (we owned the app). Device-code flow returned a token endpoint
URL. But the in-browser consent screen showed **"Approval required"** with
no "Accept" button — the tenant requires admin consent on `Mail.Read` even
for apps the user already owns and has full edit rights on.

**Takeaway:** Admin-consent policy is enforced at the scope level on every
single grant, regardless of who owns the app. Owning the app is not
sufficient to consent to its scopes.

---

## 4. Self-grant via `DelegatedPermissionGrant.ReadWrite.All`

**Plan:** The Azure CLI token had `DelegatedPermissionGrant.ReadWrite.All`
in its scope list. Use it to POST directly to `/oauth2PermissionGrants` and
create the consent record server-side, bypassing the in-browser consent
screen.

**Result:** `Authorization_RequestDenied: Insufficient privileges to
complete the operation`. Microsoft hard-codes Mail.* into a protected scope
category that requires admin consent regardless of how the grant is created.

**Takeaway:** `DelegatedPermissionGrant.ReadWrite.All` is not a universal
override. Certain scope categories (Mail, Calendar, Files in some tenants)
are gated separately.

---

## 5. FOCI refresh-token harvest from the keychain

**Plan:** Microsoft Office apps use a "Family of Client IDs" (FOCI) — their
refresh tokens are mutually redeemable across Office, Outlook, Teams,
OneDrive, etc. Read the local MSAL/OneAuth keychain cache, find a refresh
token with `foci: 1`, redeem it under a sibling client ID for Graph scopes.

**Result:** The Office identity items in `login.keychain-db` (account name
`Microsoft Office Identities Cache 3`) are stored under strict ACLs that
restrict reads to Microsoft-signed binaries. `security
find-generic-password` returned a one-byte payload (just a newline) — no
prompt, no token, no error. macOS silently denies the read because the
calling binary isn't on the ACL.

We also looked at:
- `~/Library/Group Containers/UBF8T346G9.com.microsoft.oneauth/BlobStore/` —
  contains avatar PNGs and identity-provider metadata, no tokens.
- `~/Library/Containers/com.microsoft.Outlook/Data/.../RCTAsyncLocalStorage_V1/` —
  only three unrelated config keys.

**Takeaway:** OneAuth deliberately hides tokens from non-Microsoft processes
on macOS. Don't waste time on keychain spelunking unless you're going to
sign and notarise a binary that masquerades as Outlook (which you should
not, for many reasons).

---

## 6. Outlook for Mac desktop cookies

**Plan:** The new Outlook for Mac is a partial WebView wrapper. Read its
cookies from `~/Library/Containers/com.microsoft.Outlook/Data/Library/
Cookies/Cookies.binarycookies` and replay them against the REST API.

**Result:** Only two cookies in the file (`ClientId`, `fpc`) — no session,
no SSO. The desktop app's main UI is not a WebView; only add-in iframes
(KnowBe4, addin.insights.static.microsoft) use the WebKit cookie store. The
main app talks to Exchange directly via MSAL tokens which are kept out of
that cookie file.

**Takeaway:** "Web-based desktop app" ≠ "all auth lives in cookies." Modern
Microsoft apps split UI WebViews from auth flows; tokens live in keychain,
not browser storage.

---

## 7. Browser cookie harvest (Chrome / Arc)

**Plan:** Read the encrypted cookie database from a Chromium-family browser,
decrypt it with the AES key from the "Arc Safe Storage" keychain entry, hit
`/api/v2.0/me/messages` with the resulting cookies.

**Result:** Decryption worked perfectly. We pulled all 58
outlook.office.com / login.microsoftonline.com cookies including
`ESTSAUTH`, `ESTSAUTHPERSISTENT`, `OIDC`, `CCState`. `outlook list`
returned **HTTP 401** because…

---

## 8. Outlook REST API v2 with cookie auth

**Plan:** `/api/v2.0/me/messages` used to accept cookie-authenticated calls
from the OWA session. The endpoint is documented as supporting both Bearer
and cookie auth.

**Result:** As of 2024+, Microsoft hardened the endpoint. Cookie-only calls
return `401 Forbidden`. The endpoint now requires an `Authorization: Bearer
<jwt>` header with audience `https://outlook.office.com/` — exactly the
token OWA itself acquires via its silent OAuth handshake.

**Takeaway:** Cookies get you onto the OWA web app, but the web app's API
endpoints want a Bearer. The web app acquires that Bearer in JavaScript
after the page loads.

---

## 9. Nylas CLI (third-party email API SaaS)

**Plan:** Nylas is a SaaS that abstracts over Graph/IMAP/Exchange. Their
docs claim "no Azure AD app registration, no admin consent" for end users.

**Result:** Nylas does avoid app registration — it uses their own
multi-tenant Entra app. But the consent screen the user hits is for the
**Nylas Entra app**, which requests `Mail.ReadWrite`, `Mail.Send`,
`offline_access` — exactly the scopes that triggered "Approval required"
in §3. Same wall, different logo. And it adds a SaaS hop that
enterprise egress policies usually block independently.

**Takeaway:** SaaS email-API middlemen do not bypass tenant consent policy;
they just shift which app's consent screen you see. Stop suggesting them
when admin consent is the blocker.

---

## 10. Playwright headless mode

**Plan:** Drive Chromium in headless mode for an invisible CLI experience.
Use it to open OWA, capture the Bearer, close.

**Result:** OWA's silent SSO redirect (`login.microsoftonline.com/.../
authorize` → `outlook.office.com/mail/`) stalls indefinitely. The page sits
on the authorize URL and never completes. Pinning a realistic User-Agent
(`Chrome/131.0.0.0 Safari/537.36`) did not help.

**Takeaway:** Microsoft Conditional Access fingerprints headless Chromium
(probably via `navigator.webdriver`, missing window dimensions, or absence
of expected browser extensions). Headed mode works. Don't waste time on UA
pinning, stealth plugins, or `--headless=new` until you've tried a
long-lived daemon process — that's the right next optimisation if browser
flashes are a problem.

---

## 11. Outlook for Mac AppleScript automation

**Plan:** Drive the legacy Outlook for Mac via AppleScript / `osascript`.
Zero auth concerns — it uses whatever session the desktop app already has.

**Result:** Works. The user rejected it as "not great" because:

- The Outlook desktop app must be running.
- AppleScript performance is poor on large mailboxes (sequential single-item
  fetches).
- The user wanted a programmable Node.js surface, not shell-out scripts.

**Takeaway:** Valid fallback if everything else fails. Keep it as plan Z.

---

## 12. ✅ The approach that worked: Bearer-from-the-wire capture

**Plan:** Launch Playwright Chromium against a persistent user-data-dir.
Navigate to `outlook.office.com/mail/`. Let OWA do its silent OAuth dance.
Watch the outgoing network for the first request to `outlook.office.com`
with `Authorization: Bearer …`. Snapshot the Authorization header and the
routing headers OWA sends with every API call (`x-anchormailbox`,
`x-routingparameter-sessionkey`, `x-tenantid`, `x-ms-appname`, `prefer`,
…). Save to `~/.cache/outlook-spike/auth.json` with the JWT's `exp`
decoded for cache invalidation. Subsequent CLI calls use Node `fetch`
with the cached headers and skip the browser entirely.

**Why it works:**

- The token's scopes are whatever OWA was originally consented to — no new
  grant is created at any point.
- The Outlook REST v2 endpoint accepts the OWA-issued JWT (correct audience).
- The browser handles all CA / MFA / SSO complexity for us; we just read
  the result off the wire.
- Token TTL is ~24 hours in the tested tenant. Cache-hit calls are
  sub-second and never re-open the browser.

**Trade-offs:**

- Headed Chromium flashes briefly on every cache miss (~once a day).
- Endpoints are undocumented; Microsoft can change them.
- Token cache file is sensitive; treat it like a password.

This is the architecture currently in `src/`.

---

## 13. Reaching Graph / Teams / Copilot — the audience wall, and how it differs from the consent wall

**Question:** the OWA-captured token has a huge scope list — `Chat.ReadWrite.All`,
`Channel.*`, `Team.ReadBasic.All`, `Files.ReadWrite.All`,
`OutlookCopilot-Internal.ReadWrite`, `SubstrateSearch-Internal.ReadWrite`.
Can we use it to call Microsoft Graph and drive Teams / Copilot?

**Finding:** Not with *that* token. Its `aud` is `https://outlook.office.com`.
Every Microsoft API validates the token audience and rejects a foreign one
with `401 InvalidAuthenticationToken` — a Graph call with an Outlook-audience
token fails regardless of scopes. (We confirmed the audience from the JWT
directly; a live probe was inconclusive only because the cached token had
expired.)

**But this is NOT the §1–4 consent wall.** Those dead ends were about
`Mail.*` requiring *admin consent* that the tenant refuses to grant. Here the
scopes are **already consented** to the "One Outlook Web" first-party app
(appid `9199bf20-…`). The blocker is purely the token audience, which is a
capture problem, not a consent problem.

**Approach:** the §12 wire-capture mechanism generalises. Each resource
(Graph, Substrate) has its own audience; driving the web app that mints that
token (Teams web → Graph + Teams tokens; the M365 Copilot surface →
Substrate tokens) lets us capture it the same way. Implemented as:

- `src/resources.mjs` — resource registry (base URL + audiences + host).
- `captureAllTokens()` in `capture.mjs` — classifies every Bearer seen by
  audience and keeps one header set per resource; `auth --all` opens
  Teams/Copilot tabs (best-effort) so their tokens get minted in one session.
- Per-resource token cache (`auth-<resource>.json`), `getAuth({resource})`.
- `outlook token-audit` — decodes cached tokens offline and reports audience
  + grouped scopes + which resources are reachable.
- `outlook graph <path>` — authenticated passthrough to Graph (or `--resource
  substrate`).

**Resolved (§14):** a live `auth --all` in the target tenant *does* emit
Graph + Substrate tokens. Teams-via-Graph works; Copilot does not (streaming).

---

## 14. Live reverse-engineering of Teams + Copilot (what `sniff.mjs` found)

Ran `src/sniff.mjs` (headed browser + PII-safe network sniffer — logs method
+ redacted path + token audience + JSON *shape*, never values/bodies) while
signing in and clicking through Teams and Copilot. Findings:

**Tokens the OWA/Teams/Copilot session mints (all captured live):**
- `https://outlook.office.com` — mail/calendar (the original).
- `https://graph.microsoft.com` — **Graph works.** Teams-via-Graph is real.
- `https://substrate.office.com` and sub-audiences `…/search` and
  `…/sydney` — Substrate. `…/sydney` is the Copilot ("Sydney") backend.
- Unmodelled: `https://ic3.teams.office.com`, `https://presence.teams.microsoft.com`
  — Teams real-time services (see below).

**Teams — what the captured Graph token can and can't do.** The Graph token's
scopes were `Chat.ReadBasic`, `ChatMessage.Send`, `Team.ReadBasic.All`,
`Channel.ReadBasic.All`, `User.Read.All`, … but **not** `Chat.Read` /
`ChannelMessage.Read.All`. So, verified live:
- ✅ `/me/joinedTeams`, `/teams/{id}/channels`, `/me/chats`,
  `/chats/{id}/members` — list/metadata work.
- ✅ `/chats/{id}/messages` POST — sending works (ChatMessage.Send).
- ❌ `/chats/{id}/messages` GET — **403**, "requires one of 'Chat.Read,
  Chat.ReadWrite'". Reading message *bodies* is not available from this token.

The Teams **web client doesn't read messages via Graph at all** — it uses the
undocumented chat-aggregator service:
`GET teams.microsoft.com/api/chatsvc/amer/v1/users/ME/conversations/{id}/messages`,
authenticated with a **skype token** (`Authentication: skypetoken=…`, not a
Bearer). Reading Teams messages programmatically therefore needs either a
broader-scoped Graph token (tenant won't grant Chat.Read here — cf. §2) or the
skypetoken + chatsvc path. Deferred.

**Copilot is a streaming protocol, not a REST call.** The conversation
endpoint is `POST m365.cloud.microsoft/chat` — observed 8× with **no request
body** (WebSocket/SSE upgrade; BizChat streams tokens). There is no simple
JSON request/response to wrap. A `copilot` command needs a real streaming
client (WS frames, which `sniff.mjs` does not capture) — a separate project.
The one plain read endpoint, `substrate.office.com/m365Copilot/GetGptList`,
needs the `…/sydney`-audience token specifically.

**Substrate/Graph tokens are short-lived** (minutes, not the ~24h of the
Outlook token) and the CLI deliberately does not auto-relaunch the browser for
non-default resources — so `graph`/`teams-*` calls can return `E_AUTH_REQUIRED`
soon after `auth --all`; re-run `auth --all` to refresh.

**Shipped:** `teams`, `teams-channels`, `teams-chats`, `teams-members`,
`teams-messages` (works where the tenant grants Chat.Read), `teams-send`.
Copilot: documented here + resource plumbing only; the streaming client is
future work.

---

## Rules summarised

1. Admin consent on `Mail.*` is a wall. There is no clever way around it.
   Stop looking.
2. OneAuth tokens are not accessible from a normal binary on macOS. Don't
   bother with the keychain.
3. Cookies alone won't authenticate to Outlook REST v2. You need the Bearer.
4. Headless Chromium gets blocked by Conditional Access at silent SSO. Run
   headed.
5. The only durable mechanism to get a usable Bearer without admin consent
   is to **capture OWA's own token from network traffic in a real browser
   session**. Everything else is a dead end in a locked-down tenant.
