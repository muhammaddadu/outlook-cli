# Reverse-engineering web apps for CLI automation

A general method, abstracted from how we built `outlook-cli`. Useful when:

- You want programmatic access to a web app you already log into.
- The app has no public API, or its public API requires admin consent /
  enterprise approval / a paid plan you can't get.
- You're acting **only on your own account**, with credentials you
  already have. (See "Ethics and legal" at the bottom — this technique
  is for personal automation, not scraping someone else's data.)

The end result is a Node script (or whatever) that talks to the web app's
own backend using the same auth your browser already negotiated, with no
extra approval from anyone.

---

## The core idea in one diagram

```
┌─────────────────────────────┐    capture     ┌─────────────────────────┐
│ Real browser (Playwright)   │ ─────────────▶ │ Auth credential         │
│ pointed at app.example.com  │  from network  │ (Bearer / cookies / …)  │
└─────────────────────────────┘                └──────────┬──────────────┘
                                                          │ cache to disk
                                                          ▼
                                              ┌─────────────────────────┐
                                              │ Node fetch() with same  │
                                              │ headers → app's own API │
                                              └─────────────────────────┘
```

The web app does its own login dance. You watch what credential it
acquired and what URL it talked to, then make those same calls yourself
from a script.

This works because:
- The user has already consented (when they signed in normally).
- The web app's backend was designed to accept exactly the headers the
  web app sends — no extra validation that "the caller is the real web
  app." Cryptographic proof of that would require client-side code-signing
  and almost no web apps bother.
- A persistent browser profile + cached credential means you only have
  to do this dance occasionally; most calls are headers-from-cache.

---

## Method, step by step

### Step 1 — Confirm the web app works in a normal browser

Sign into the target app in any browser. Verify you can see / do what
you eventually want to automate. **If you can't do it as a user, you
can't do it from a script.** This step also gives you the consent record
the backend will accept.

### Step 2 — Survey the network

Open DevTools → Network in the browser, with the web app loaded. Scroll
the relevant feature (compose mail, open a chart, whatever) and watch
what requests fire.

Look at each one:

| What to inspect | What you're learning |
| --- | --- |
| URL | The endpoint you'll call yourself |
| Method | GET / POST / PATCH / DELETE |
| `Authorization` header | Bearer JWT? Basic? Missing? |
| `Cookie` header | Session cookies the backend recognises |
| `X-CSRF-Token`, `X-Something` | Custom headers the server expects |
| Response body | Schema, error shape, pagination tokens |

Quick rules of thumb:

- **Single-page apps that talk to a JSON API at `/api/…` or `…api.app.com`** → typically Bearer JWT. Worth capturing.
- **Older server-rendered apps with form posts** → typically session cookies + CSRF token.
- **Apps that work in a private window once you log in** → at least one of cookies / bearer is enough.
- **Apps that break in a private window** → likely depend on localStorage / IndexedDB state too; harder.

### Step 3 — Decide your capture strategy

Three common options, ordered by how invasive they are:

**A. Cookie harvest from the browser profile**

If the app uses cookie auth only, you can read the cookies from the
browser's local storage. Chromium and Firefox encrypt these at rest;
several libraries decrypt automatically:

- Node: `chrome-cookies-secure`, `tough-cookie-file-store`.
- Python: `browser_cookie3`, `pycookiecheat`.

Cheap and works without Playwright, but breaks if the app rotates session
cookies frequently or if cookies have additional fingerprint binding.

**B. Playwright network interception (preferred)**

Launch your own Chromium with a persistent profile, navigate to the
app, and watch outgoing requests. Snapshot the first authenticated call
to capture exactly the headers the app uses.

```js
import { chromium } from 'playwright';

const ctx = await chromium.launchPersistentContext('/path/to/profile', {
  headless: false,                    // headless = blocked by many CDNs / CA
  userAgent: REAL_CHROME_UA,           // headless UA gets flagged
});
const page = await ctx.newPage();

let captured = null;
page.on('request', (req) => {
  if (captured) return;
  if (!req.url().startsWith('https://api.example.com/')) return;
  const h = req.headers();
  if (!h.authorization?.startsWith('Bearer ')) return;
  captured = {
    authorization: h.authorization,
    'x-anchor': h['x-anchor'],
    // …copy every x-* header you saw in DevTools that's required.
  };
});

await page.goto('https://app.example.com/home');
// Wait for the app to make its first authed call.
while (!captured) await page.waitForTimeout(150);

await ctx.close();
// captured ← now use this from Node fetch()
```

**C. Add-in / bookmarklet inside the page**

If the app's CSP allows it, run a small script *inside* the page that
fetches its own API and writes the token to a known place (clipboard,
localStorage, an HTTP listener you spawn locally). Sometimes the cleanest
when MSAL.js or similar exposes a `getAccessToken()` global.

We did **(B)** for outlook-cli because Conditional Access blocks
cookie-only auth on the API endpoints, and (A) wasn't enough.

### Step 4 — Replay the call from Node

Once you have headers, the rest is plain `fetch`:

```js
const res = await fetch('https://api.example.com/me/things', {
  headers: { ...captured, Accept: 'application/json' },
});
const data = await res.json();
```

Make sure to forward **every** header the browser was sending that
looks app-specific (`x-anchor-*`, `x-tenant-*`, `prefer`, …). Servers
often degrade behaviour if these are missing — different feature flags,
stricter throttling, "client too old" 426 responses.

### Step 5 — Cache the credential

Capturing fresh headers requires launching a real browser, which is slow
and visible. Cache them to disk and re-use until they expire.

For JWTs, decode the `exp` claim:

```js
function jwtExp(bearer) {
  const [, payload] = bearer.replace(/^Bearer /, '').split('.');
  const json = JSON.parse(Buffer.from(payload, 'base64').toString());
  return new Date(json.exp * 1000);
}
```

Refresh proactively with a few-minute buffer to avoid mid-call expiry.

For cookie-only auth, set up a TTL based on observed behaviour (most
session cookies last hours-to-days).

### Step 6 — Handle expiry and re-auth

Two failure modes:

1. **Token rejected mid-call (401 or 403).** Clear the cache, re-capture,
   retry once. If the second call also 401s, the user's session itself
   has lapsed — surface this and ask them to sign in again.
2. **Browser profile lost cookies.** When you launch the persistent
   profile, the redirect lands on the login page instead of the app.
   Detect this and instruct the user to run an interactive sign-in.

### Step 7 — Hide the browser flash

The first capture-on-cache-miss opens a visible Chromium window for ~3
seconds. That's tolerable as a once-a-day cost. To avoid it entirely:

- **Long-lived daemon**: keep one Chromium process alive, expose a local
  socket, route every CLI call through it.
- **Headless plus stealth**: usually fails against Conditional Access /
  Cloudflare. Try first, but don't be surprised if it stalls on the
  silent SSO redirect (this is what blocked headless mode for us — see
  `LEARNINGS.md` §10).
- **Virtual display (`xvfb`)**: render Chromium headed but invisibly.
  Works on Linux, awkward on Mac.

---

## Choosing your endpoint

Most apps have *multiple* backends you could replay against:

| Pattern | Pros | Cons |
| --- | --- | --- |
| **Public REST API** (`api.example.com/v2/…`) | Documented, stable, predictable | May require admin consent / paid plan you can't get |
| **Internal/private API** the web UI calls | No consent needed; same auth as the web UI | Undocumented, breaks without warning |
| **GraphQL gateway** | Often used by modern SPAs; introspection helps | Schemas can be huge / restricted |
| **Legacy/EWS-style RPC** | Sometimes still works on older M365/Salesforce/etc. | Brittle, often deprecated |

In `outlook-cli` we use the **internal** Outlook REST v2 surface
(`outlook.office.com/api/v2.0/me/…`) — undocumented but works without
admin consent. Microsoft Graph would have needed admin approval.

When you have a choice, prefer the most stable surface you can use
without elevated permissions. Document which endpoint you picked and why
in a `LEARNINGS.md`-style file so future maintainers know.

---

## The diagnostic script you'll write 10 times

For any new app, the very first script to write is the network sniffer.
Ours is `src/diagnose.mjs` — ~50 lines that opens the app and dumps every
URL it touches plus the auth-header shape:

```js
import { chromium } from 'playwright';

const ctx = await chromium.launchPersistentContext(PROFILE_DIR, { headless: false });
const page = await ctx.newPage();

const seen = new Map();
page.on('request', (req) => {
  const url = req.url();
  if (!/api\.example\.com|app\.example\.com\/(api|graphql)/.test(url)) return;
  const key = `${req.method()} ${url.split('?')[0]}`;
  if (seen.has(key)) return;
  const h = req.headers();
  seen.set(key, {
    method: req.method(),
    url,
    auth: h.authorization ? `Bearer ${h.authorization.slice(7, 27)}…` : 'cookie-only',
    extra: Object.fromEntries(
      Object.entries(h).filter(([k]) =>
        /^(x-|action|prefer|content-type)/i.test(k),
      ),
    ),
  });
});

await page.goto('https://app.example.com/');
await page.waitForSelector('[data-loaded]', { timeout: 60_000 });
await page.waitForTimeout(5_000);

for (const r of seen.values()) {
  console.log(`${r.method} ${r.url}`);
  console.log(`  auth: ${r.auth}`);
  for (const [k, v] of Object.entries(r.extra)) console.log(`  ${k}: ${v}`);
}

await ctx.close();
```

Copy-paste this into your project, change the URL filter, run it. The
output tells you immediately:

1. Which endpoints power the feature you're after.
2. What auth shape each endpoint expects.
3. Which extra headers are required.

From there, build a thin wrapper that calls those endpoints with the
captured headers — exactly what `src/client.mjs` does for `outlook-cli`.

---

## Common gotchas

### Headless detection

CDNs (Cloudflare, Akamai) and identity providers (Microsoft CA, Okta,
Auth0) fingerprint headless Chromium and either block it outright or stall
the redirect chain. Symptoms:

- Browser sits on a `/authorize` URL forever.
- `403` responses with a captcha-style page.
- `Sec-Fetch-Dest` header missing or wrong.

Fix: launch headed (with a real UA), and accept that the first
cache-miss flashes a window. Or invest in a daemon (see Step 7).

### Token audience mismatch

If you capture a token meant for `api.example.com` and try to use it
against `other.example.com`, you'll get 401. Decode the JWT (`jwt.io`
or a one-liner) and check the `aud` claim. Different parts of the same
product can use different audiences.

### Required routing headers

A token alone often isn't enough. Apps that shard by user / tenant / region
pass headers like `x-anchor-mailbox`, `x-tenant-id`, `x-realm`. The
backend uses these to route to the right shard. Drop them and you get
weird "user not found" or 500 errors.

**Always replay every header the browser sent, not just `Authorization`.**

### CSRF tokens

For form-post style apps, the server requires an `X-CSRF-Token` that
changes per session (often embedded in a `<meta>` tag or set as a
JS-readable cookie). Read it in the same Playwright pass that captures
the session and forward it.

### Conditional-access step-up

Some flows trigger an additional auth step (re-prompt for MFA, device
compliance check). When you next launch your persistent profile,
Chromium may redirect to that step instead of the app. Detect by URL
inspection and surface to the user.

### Endpoint instability

Internal APIs change. Today's `/api/v2/messages` might be tomorrow's
`/graphql` with a totally different schema. Don't write a permanent
business on top of one — write a personal automation tool, and keep your
diagnostic script handy so you can re-discover the new shape in 30
seconds when it breaks.

---

## Project structure that generalises

```
your-cli/
  src/
    cli.mjs        # command surface (commander.js)
    capture.mjs    # Playwright launcher + header capture
    client.mjs     # plain fetch wrapper around the captured headers
    auth.mjs       # token cache: load / save / clear / expiry
    diagnose.mjs   # the sniffer from above
    errors.mjs     # AppError + exit codes
    paths.mjs      # XDG-compliant cache + data dirs
    output.mjs     # stdout (data) vs stderr (diagnostics) discipline
  test/
    helpers.mjs    # mock HTTP server, fake JWT, runCli subprocess helper
    *.test.mjs
  skill/
    <name>/SKILL.md   # AI-agent skill (if you want Claude/Codex to drive)
    commands/*.md     # slash commands
  docs/
    architecture.md
    troubleshooting.md
    security.md
  LEARNINGS.md     # write down each dead end you hit — the next person
                   # (or your future self) saves an afternoon
```

The shape stays the same regardless of target. Swap `outlook` for
`linear`, `notion`, `airtable`, `your-corp-intranet`, …; the auth code
is 90% identical.

---

## Ethics and legal

This pattern is for **personal automation of your own account**. It's a
fancy way of saying "I want to script things I can already do in a
browser." That's almost always allowed.

It is **not** appropriate for:

- Acting on accounts you don't own (scraping other users' data).
- Bypassing rate limits, fair-use policies, or paywalls.
- Anything your terms-of-service explicitly prohibits ("you may not
  access the service except through approved clients").
- Workplace environments where you've been told **not** to automate the
  service. Even if technically possible, this can be a fireable offence
  in regulated industries.

Specifically for enterprise tools, check:

1. **Acceptable Use Policy.** Some companies prohibit unsanctioned
   tooling that accesses corporate data programmatically, even by the
   data owner.
2. **Conditional Access policies.** If the policy says "compliant
   devices only," the captured token already encodes that. You're not
   bypassing CA — you're using it. But if you ever try to move the token
   to an unmanaged device, the next API call will 401.
3. **Data Loss Prevention.** Client-side DLP plugins that hook into the
   web app's UI (e.g. content scanning before send) won't run when you
   call the backend directly. Server-side DLP still applies.
4. **EDR / SIEM monitoring.** Token capture from network traffic is the
   exact technique attackers use (MITRE T1528). On a managed device,
   your EDR may flag it. Best practice: tell your security team
   proactively — "I'm using my own captured session for a personal
   script, here's what it does." They'd rather have the conversation
   in advance than after an alert.

When in doubt, ask. The best automation in the world isn't worth a
disciplinary review.
