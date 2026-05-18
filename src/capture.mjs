// Bearer-token + routing-header capture via Playwright Chromium.
//
// Split out from client.mjs so unit tests for the rest of the codebase
// (call(), token cache, errors, …) can import without paying the cost of
// loading Playwright + dragging in the Chromium binary requirement.

import { chromium } from 'playwright';
import { profileDir } from './paths.mjs';
import { AppError, E } from './errors.mjs';
import { debug } from './output.mjs';

const HOME_URL = 'https://outlook.office.com/mail/';

// Pin a real Chrome UA — Playwright's default contains "HeadlessChrome",
// which Conditional Access flags even in headed mode on first navigation.
const DESKTOP_UA =
  'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 ' +
  '(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36';

// Headers to forward from the captured request into our own calls. The
// bearer + routing values are the auth substrate; the rest match OWA's
// version pinning so the server doesn't downgrade us into a stricter
// throttling bucket.
const FORWARDED_HEADERS = [
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

export async function openContext({ headless = false } = {}) {
  return chromium.launchPersistentContext(profileDir(), {
    headless,
    viewport: { width: 1280, height: 900 },
    userAgent: DESKTOP_UA,
    args: ['--disable-blink-features=AutomationControlled'],
  });
}

/**
 * Open OWA, wait for the first Bearer-authed request, return its headers.
 *
 * Throws AppError(E.AUTH_REQUIRED) if the browser is redirected to a login
 * page (cached cookies are gone — interactive `outlook auth` needed).
 * Throws AppError(E.AUTH_BLOCKED) if the navigation succeeds but no Bearer
 * token is observed before the timeout (rare; usually means OWA hung).
 */
export async function captureAuth({ timeoutMs = 60_000 } = {}) {
  const context = await openContext({ headless: false });
  /** @type {Record<string,string>|null} */
  let captured = null;

  try {
    const page = context.pages()[0] ?? (await context.newPage());

    page.on('request', (req) => {
      if (captured) return;
      const h = req.headers();
      if (!h.authorization?.startsWith('Bearer ')) return;
      if (!req.url().startsWith('https://outlook.office.com/')) return;
      const bag = {};
      for (const key of FORWARDED_HEADERS) {
        if (h[key]) bag[key] = h[key];
      }
      captured = bag;
    });

    await page.goto(HOME_URL, { waitUntil: 'domcontentloaded', timeout: timeoutMs });

    if (/login\.(microsoft|microsoftonline)\.com/.test(page.url())) {
      throw new AppError({
        code: E.AUTH_REQUIRED,
        message: `OWA redirected to login (${page.url().split('?')[0]}).`,
        hint: 'Run `outlook auth` to sign in interactively (MFA included).',
      });
    }

    const deadline = Date.now() + timeoutMs;
    while (!captured && Date.now() < deadline) {
      await page.waitForTimeout(150);
    }

    if (!captured) {
      throw new AppError({
        code: E.AUTH_BLOCKED,
        message: `No Bearer token observed within ${timeoutMs}ms.`,
        hint: 'Try `outlook auth` to sign in interactively — your session may have lapsed.',
      });
    }

    debug('captured bearer token');
    return captured;
  } finally {
    // Brief delay so the browser flushes any pending profile writes before
    // shutdown — otherwise an interactive sign-in can fail to persist
    // cookies/IndexedDB.
    await new Promise((r) => setTimeout(r, 500));
    await context.close();
  }
}
