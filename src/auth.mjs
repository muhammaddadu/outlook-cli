// Auth lifecycle: load from cache, capture via browser when stale, save back.
//
// The captured headers include the Bearer JWT. We decode its `exp` claim to
// know when the cache becomes stale; we refresh proactively with a 5-minute
// buffer so a long-running command never trips a mid-call expiry.
//
// Auth is per-resource: each Microsoft API (outlook, graph, substrate) needs
// its own token with the matching audience. The default resource is
// "outlook", whose cache stays at the original `auth.json` for backward
// compatibility; other resources sit alongside as `auth-<resource>.json`.

import {
  readFileSync,
  writeFileSync,
  mkdirSync,
  existsSync,
  unlinkSync,
} from 'node:fs';
import { dirname } from 'node:path';
import { tokenCacheFile } from './paths.mjs';
import { debug, info } from './output.mjs';
import { decodePayload } from './jwt.mjs';
import { AppError, E } from './errors.mjs';
import { DEFAULT_RESOURCE, resource } from './resources.mjs';

// Lazy import — capture.mjs pulls in Playwright (≈80MB of resident memory
// and a Chromium executable check). Cache-hit calls never hit this code
// path and shouldn't pay that cost.
//
// Two env seams control the browser launch:
//   OUTLOOK_CAPTURE_FIXTURE=/path  read a { resourceKey: headers } map (or a
//                                  bare outlook headers object) from a JSON
//                                  file instead of launching Chromium (tests).
//   OUTLOOK_NO_CAPTURE=1           fail with E_AUTH_REQUIRED instead of
//                                  launching Chromium (CI, headless boxes).
async function captureAll(opts) {
  if (process.env.OUTLOOK_CAPTURE_FIXTURE) {
    debug('using capture fixture', process.env.OUTLOOK_CAPTURE_FIXTURE);
    const parsed = JSON.parse(
      readFileSync(process.env.OUTLOOK_CAPTURE_FIXTURE, 'utf8'),
    );
    // Accept either a resource->headers map or a bare outlook headers object
    // (the shape the older single-token fixture used).
    return parsed.authorization ? { outlook: parsed } : parsed;
  }
  if (process.env.OUTLOOK_NO_CAPTURE) {
    throw new AppError({
      code: E.AUTH_REQUIRED,
      message: 'A fresh token is needed but browser capture is disabled (OUTLOOK_NO_CAPTURE is set).',
      hint: 'Run `outlook auth` in an environment with a display, or unset OUTLOOK_NO_CAPTURE.',
    });
  }
  const mod = await import('./capture.mjs');
  return mod.captureAllTokens(opts);
}

const MIN_LIFETIME_MS = 5 * 60 * 1000;

function decodeJwtExp(bearer) {
  const payload = decodePayload(bearer);
  return typeof payload?.exp === 'number' ? payload.exp * 1000 : null;
}

export function loadCachedAuth(res = DEFAULT_RESOURCE) {
  const file = tokenCacheFile(res);
  if (!existsSync(file)) return null;
  let entry;
  try {
    entry = JSON.parse(readFileSync(file, 'utf8'));
  } catch {
    return null;
  }
  if (!entry?.headers?.authorization) return null;

  const exp = decodeJwtExp(entry.headers.authorization);
  if (!exp || exp - Date.now() < MIN_LIFETIME_MS) {
    debug(`cached ${res} token is stale or unparseable`);
    return null;
  }
  return { headers: entry.headers, expiresAt: exp };
}

export function saveAuth(headers, res = DEFAULT_RESOURCE) {
  const file = tokenCacheFile(res);
  mkdirSync(dirname(file), { recursive: true });
  const exp = decodeJwtExp(headers.authorization ?? '');
  if (!exp) {
    // Without a decodable expiry the cache is treated as stale on every
    // load, which silently degrades into a Chromium relaunch per command.
    // Warn loudly so the failure mode is visible instead of mysterious.
    info(
      `Warning: captured ${res} token has no decodable expiry; it will not be ` +
        'reused and the browser will reopen on the next command. See docs/troubleshooting.md.',
    );
  }
  writeFileSync(
    file,
    JSON.stringify({ headers, expiresAt: exp, savedAt: Date.now() }, null, 2),
  );
}

export function clearAuth(res = DEFAULT_RESOURCE) {
  const file = tokenCacheFile(res);
  if (existsSync(file)) unlinkSync(file);
}

/** Persist a { resourceKey: headers } capture map; returns the keys saved. */
export function saveAllAuth(map) {
  const saved = [];
  for (const [res, headers] of Object.entries(map)) {
    if (!headers?.authorization) continue;
    saveAuth(headers, res);
    saved.push(res);
  }
  return saved;
}

/**
 * Capture fresh tokens and persist them. Backs `auth` / `refresh` and the
 * automatic 401 recovery path. Lives here (not cli.mjs) so the Playwright
 * import stays lazy for every command.
 *
 * Returns the headers for `wantResource` (default outlook) so callers that
 * only care about one resource keep their old shape. Every resource observed
 * during the capture window is persisted regardless.
 */
export async function refreshAuth({ wantResource = DEFAULT_RESOURCE, ...opts } = {}) {
  const map = await captureAll(opts);
  const saved = saveAllAuth(map);
  debug(`captured tokens for: ${saved.join(', ') || '(none)'}`);
  return map[wantResource] ?? null;
}

/**
 * Get headers for an authed API call to `res`. Uses the on-disk cache when
 * the JWT still has > MIN_LIFETIME_MS left; otherwise opens the browser to
 * capture fresh tokens via silent SSO and persists them.
 *
 * For non-default resources we never silently launch the browser mid-command
 * (a `graph` call shouldn't pop Chromium unexpectedly) unless the caller opts
 * in with `capture: true` — the outlook workhorse keeps its auto-capture.
 */
export async function getAuth({ resource: res = DEFAULT_RESOURCE, forceRefresh = false, capture } = {}) {
  if (!forceRefresh) {
    const cached = loadCachedAuth(res);
    if (cached) {
      const minsLeft = Math.floor((cached.expiresAt - Date.now()) / 60_000);
      debug(`using cached ${res} token (~${minsLeft}m until expiry)`);
      return cached.headers;
    }
  }

  const autoCapture = capture ?? res === DEFAULT_RESOURCE;
  if (!autoCapture) {
    const r = resource(res);
    throw new AppError({
      code: E.AUTH_REQUIRED,
      message: `No usable ${res} token cached.`,
      hint: `Run \`outlook auth --all\` to capture a ${r.label} token, then retry.`,
    });
  }

  debug(`refreshing ${res} token via browser`);
  const headers = await refreshAuth({ wantResource: res });
  if (!headers) {
    const r = resource(res);
    throw new AppError({
      code: E.AUTH_REQUIRED,
      message: `Signed in, but no ${res} token was emitted by the browser session.`,
      hint: `Your OWA session may not mint a ${r.label} token. Run \`outlook auth --all\` after opening the matching web app (Teams/Copilot), or check \`outlook token-audit\`.`,
    });
  }
  return headers;
}
