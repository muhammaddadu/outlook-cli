// Bearer-token + routing-header capture via Playwright Chromium.
//
// Split out from client.mjs so unit tests for the rest of the codebase
// (call(), token cache, errors, …) can import without paying the cost of
// loading Playwright + dragging in the Chromium binary requirement.
//
// The core trick (see LEARNINGS.md §12): a real browser session performs
// Microsoft's silent OAuth dance for us; we read the resulting Bearer tokens
// off the wire. Each Microsoft API is a separate OAuth resource with its own
// token audience, so we classify every captured token by audience and keep
// one header set per resource. Driving different web apps (OWA, Teams,
// Copilot) in the same window mints tokens for their respective resources.

import { chromium } from 'playwright';
import { profileDir } from './paths.mjs';
import { AppError, E } from './errors.mjs';
import { debug, info } from './output.mjs';
import { decodePayload } from './jwt.mjs';
import { classifyToken, DEFAULT_RESOURCE } from './resources.mjs';

export const HOME_URL = 'https://outlook.office.com/mail/';

// Extra web surfaces we can open (best-effort) during a broad `--all`
// capture so their resource tokens get minted without the user hunting for
// URLs. Teams mints Graph + Teams tokens; the M365 Copilot chat surface
// mints Substrate tokens. Failures here are non-fatal.
export const EXTRA_SURFACES = Object.freeze([
  'https://teams.microsoft.com/v2/',
  'https://m365.cloud.microsoft/chat/',
]);

// Pin a real Chrome UA — Playwright's default contains "HeadlessChrome",
// which Conditional Access flags even in headed mode on first navigation.
const DESKTOP_UA =
  'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 ' +
  '(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36';

// Headers to forward from the captured request into our own calls. The
// bearer + routing values are the auth substrate; the rest match OWA's
// version pinning so the server doesn't downgrade us into a stricter
// throttling bucket.
export const FORWARDED_HEADERS = [
  'authorization',
  'x-anchormailbox',
  'x-routingparameter-sessionkey',
  'x-tenantid',
  'x-ms-appname',
  'x-clientid',
  'x-owa-sessionid',
  'x-client-version',
  'prefer',
];

const LOGIN_URL_RE = /login\.(microsoft|microsoftonline)\.com/;

export async function openContext({ headless = false } = {}) {
  return chromium.launchPersistentContext(profileDir(), {
    headless,
    viewport: { width: 1280, height: 900 },
    userAgent: DESKTOP_UA,
    args: ['--disable-blink-features=AutomationControlled'],
  });
}

/** Playwright throws generic Errors when the user closes the window. */
function isClosedError(err) {
  return /Target (page|context|browser).*closed|has been closed/i.test(
    err?.message ?? '',
  );
}

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

/**
 * Open the browser and capture Bearer tokens for every Microsoft resource
 * observed on the wire. Returns a `{ resourceKey: headersBag }` map.
 *
 * @param {object} [opts]
 * @param {number}  [opts.timeoutMs=60000]  overall wait budget
 * @param {boolean} [opts.interactive=false] tolerate the login page (user is signing in / navigating)
 * @param {boolean} [opts.waitForAll=false]  sniff the whole window instead of stopping at the first outlook token
 * @param {string}  [opts.target=HOME_URL]   URL to open first
 * @param {boolean} [opts.openExtraSurfaces=false] also open Teams/Copilot tabs (best-effort) to mint their resource tokens
 */
export async function captureAllTokens({
  timeoutMs = 60_000,
  interactive = false,
  waitForAll = false,
  target = HOME_URL,
  openExtraSurfaces = false,
} = {}) {
  const extraSurfaces = openExtraSurfaces ? EXTRA_SURFACES : [];
  const context = await openContext({ headless: false });
  /** @type {Record<string, Record<string,string>>} */
  const found = {};

  try {
    const page = context.pages()[0] ?? (await context.newPage());

    // Listen context-wide so tokens are captured no matter which tab/popup
    // fires them (Teams and Copilot open their own surfaces).
    context.on('request', (req) => {
      const h = req.headers();
      if (!h.authorization?.startsWith('Bearer ')) return;
      let host;
      try {
        host = new URL(req.url()).host;
      } catch {
        return;
      }
      const claims = decodePayload(h.authorization);
      const key = classifyToken({ aud: claims?.aud, host });
      if (!key) return;
      const bag = {};
      for (const k of FORWARDED_HEADERS) if (h[k]) bag[k] = h[k];
      found[key] = bag; // keep the most recent headers for each resource
    });

    await page.goto(target, { waitUntil: 'domcontentloaded', timeout: timeoutMs });

    // Best-effort: open extra surfaces to trigger their token acquisition.
    for (const url of extraSurfaces) {
      try {
        const extra = await context.newPage();
        await extra.goto(url, { waitUntil: 'domcontentloaded', timeout: timeoutMs });
      } catch (e) {
        debug(`extra surface ${url} did not load: ${e.message}`);
      }
    }

    const deadline = Date.now() + timeoutMs;
    let sawLogin = false;
    while (Date.now() < deadline) {
      if (page.isClosed()) {
        if (found[DEFAULT_RESOURCE]) break; // got what we needed before it closed
        throw new AppError({
          code: E.AUTH_BLOCKED,
          message: 'Browser window was closed before a token was captured.',
          hint: 'Re-run the command and leave the window open until it closes itself.',
        });
      }
      sawLogin = sawLogin || LOGIN_URL_RE.test(page.url());
      if (sawLogin && !interactive && !found[DEFAULT_RESOURCE]) {
        throw new AppError({
          code: E.AUTH_REQUIRED,
          message: `OWA redirected to login (${page.url().split('?')[0]}).`,
          hint: 'Run `outlook auth` to sign in interactively (MFA included).',
        });
      }
      // Fast path: single-resource auto-capture stops as soon as outlook is in.
      if (!waitForAll && found[DEFAULT_RESOURCE]) break;
      await sleep(150);
    }

    if (Object.keys(found).length === 0) {
      throw new AppError({
        code: sawLogin ? E.AUTH_REQUIRED : E.AUTH_BLOCKED,
        message: sawLogin
          ? `Sign-in was not completed within ${timeoutMs}ms.`
          : `No Bearer token observed within ${timeoutMs}ms.`,
        hint: 'Try `outlook auth` to sign in interactively — your session may have lapsed.',
      });
    }

    debug(`captured tokens for: ${Object.keys(found).join(', ')}`);
    if (waitForAll) {
      info(`Captured tokens for: ${Object.keys(found).join(', ')}.`);
    }
    return found;
  } catch (err) {
    if (err instanceof AppError) throw err;
    if (isClosedError(err)) {
      if (found[DEFAULT_RESOURCE] || Object.keys(found).length > 0) return found;
      throw new AppError({
        code: E.AUTH_BLOCKED,
        message: 'Browser window was closed before a token was captured.',
        hint: 'Re-run the command and leave the window open until it closes itself.',
        cause: err,
      });
    }
    throw err;
  } finally {
    // Brief delay so the browser flushes any pending profile writes before
    // shutdown — otherwise an interactive sign-in can fail to persist
    // cookies/IndexedDB.
    await sleep(500);
    await context.close().catch(() => {});
  }
}

/**
 * Capture just the Outlook (default-resource) token. Thin back-compat
 * wrapper over captureAllTokens for callers that only want mail/calendar.
 */
export async function captureAuth(opts = {}) {
  const map = await captureAllTokens(opts);
  return map[DEFAULT_RESOURCE] ?? Object.values(map)[0];
}
