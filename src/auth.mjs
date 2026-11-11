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
  readdirSync,
} from 'node:fs';
import { dirname, basename, join } from 'node:path';
import { tokenCacheFile } from './paths.mjs';
import { debug, info } from './output.mjs';
import { decodePayload } from './jwt.mjs';
import { AppError, E } from './errors.mjs';
import { DEFAULT_RESOURCE, RESOURCES, resource } from './resources.mjs';
import { keepalivePid, pokeKeepalive, keepaliveOwnsProfileError } from './keepalive.mjs';

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
      hint: 'Run `outlook auth --all` in an environment with a display, or unset OUTLOOK_NO_CAPTURE.',
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
  saveVariant(headers, res);
}

// ---------------------------------------------------------------------------
// Per-app variants. Different web apps get tokens for the *same* resource
// with different scopes (e.g. Graph from Teams web has ChannelMessage.Read.All
// but not Chat.Read; Graph from Outlook web has Chat.Read but not
// ChannelMessage.Read.All). The primary cache file holds one; every token is
// also kept as `<cache>@<appid>.json` so a 403 "missing scope" can retry with
// a sibling that has the scope.

function variantFile(res, appid) {
  const primary = tokenCacheFile(res);
  return join(dirname(primary), `${basename(primary, '.json')}@${appid}.json`);
}

function scopesOf(bearer) {
  const scp = decodePayload(bearer)?.scp;
  return new Set(typeof scp === 'string' ? scp.split(' ') : []);
}

/** Save `headers` as the variant for its app id, if it outlives the one there. */
export function saveVariant(headers, res = DEFAULT_RESOURCE) {
  const appid = decodePayload(headers?.authorization ?? '')?.appid;
  const exp = decodeJwtExp(headers?.authorization ?? '');
  if (!appid || !exp) return false;
  const file = variantFile(res, appid);
  try {
    const cur = JSON.parse(readFileSync(file, 'utf8'));
    if (cur?.headers?.authorization === headers.authorization) return false;
    if ((cur?.expiresAt ?? 0) > exp) return false;
  } catch {}
  mkdirSync(dirname(file), { recursive: true });
  writeFileSync(file, JSON.stringify({ headers, expiresAt: exp, savedAt: Date.now() }, null, 2));
  return true;
}

/** Every still-valid cached token for `res` (variants), newest-expiring first. */
export function loadVariants(res = DEFAULT_RESOURCE) {
  const primary = tokenCacheFile(res);
  const dir = dirname(primary);
  const prefix = `${basename(primary, '.json')}@`;
  let names = [];
  try {
    names = readdirSync(dir).filter((n) => n.startsWith(prefix) && n.endsWith('.json'));
  } catch {
    return [];
  }
  const out = [];
  for (const n of names) {
    try {
      const entry = JSON.parse(readFileSync(join(dir, n), 'utf8'));
      const exp = decodeJwtExp(entry?.headers?.authorization ?? '');
      if (!exp || exp - Date.now() < MIN_LIFETIME_MS) continue;
      out.push({ headers: entry.headers, expiresAt: exp, scopes: scopesOf(entry.headers.authorization) });
    } catch {}
  }
  return out.sort((a, b) => b.expiresAt - a.expiresAt);
}

/** Scopes named in a Graph 403 like "API requires one of 'A, B'". */
export function requiredScopesFrom(body) {
  const msg = typeof body === 'string' ? body : (body?.error?.message ?? '');
  const m = /requires one of '([^']+)'/i.exec(msg);
  return m ? m[1].split(',').map((s) => s.trim()).filter(Boolean) : [];
}

/** A cached variant of `res` holding one of `required` scopes, other than `exclude`. */
export function variantWithScopes(res, required, exclude) {
  return (
    loadVariants(res).find(
      (v) => v.headers.authorization !== exclude && required.some((s) => v.scopes.has(s)),
    ) ?? null
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
 * only care about one resource keep their old shape; pass `wantResource: null`
 * to get the whole { resource: headers } map back. Every resource observed
 * during the capture window is persisted regardless.
 */
export async function refreshAuth({ wantResource = DEFAULT_RESOURCE, ...opts } = {}) {
  // A running `outlook keepalive` owns the browser profile (Chromium locks
  // it) and keeps the cache fresh. Ask it to reload and read the cache
  // instead of launching a second browser.
  if (!process.env.OUTLOOK_CAPTURE_FIXTURE && keepalivePid()) {
    return waitForKeepalive(wantResource, opts);
  }
  const map = await captureAll(opts);
  const saved = saveAllAuth(map);
  debug(`captured tokens for: ${saved.join(', ') || '(none)'}`);
  if (wantResource === null) return map; // caller wants every resource captured
  return map[wantResource] ?? null;
}

const KEEPALIVE_WAIT_MS = 90_000;

async function waitForKeepalive(wantResource, { interactive } = {}) {
  const pid = keepalivePid();
  if (interactive) {
    throw new AppError({
      code: E.AUTH_REQUIRED,
      message: `\`outlook keepalive\` is running (pid ${pid}) and owns the sign-in browser.`,
      hint: 'Sign in in its window if it is asking, or stop it first with `outlook keepalive --stop`.',
    });
  }
  pokeKeepalive();
  const wanted = wantResource === null ? Object.keys(RESOURCES) : [wantResource];
  const deadline = Date.now() + KEEPALIVE_WAIT_MS;
  for (;;) {
    const map = {};
    for (const r of wanted) {
      const c = loadCachedAuth(r);
      if (c) map[r] = c.headers;
    }
    if (wantResource === null ? Object.keys(map).length > 0 : map[wantResource]) {
      return wantResource === null ? map : map[wantResource];
    }
    if (Date.now() > deadline) throw keepaliveOwnsProfileError(pid);
    await new Promise((r) => setTimeout(r, 2_000));
  }
}

/**
 * Get headers for an authed API call to `res`. Uses the on-disk cache when
 * the JWT still has > MIN_LIFETIME_MS left; otherwise runs a *silent*
 * (headless) capture that succeeds only if the saved SSO session signs in by
 * itself, and persists every token it sees. No window is ever shown here, so
 * scheduled/unattended callers are safe; when the session has lapsed this
 * fails fast with E_AUTH_REQUIRED and the user runs `outlook auth`.
 *
 * Pass `capture: false` to skip the silent capture entirely.
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

  const autoCapture = capture ?? true;
  if (!autoCapture) {
    const r = resource(res);
    throw new AppError({
      code: E.AUTH_REQUIRED,
      message: `No usable ${res} token cached.`,
      hint: `Run \`outlook auth --all\` to capture a ${r.label} token, then retry.`,
    });
  }

  debug(`refreshing ${res} token via silent browser capture`);
  const headers = await refreshAuth({
    wantResource: res,
    silent: true,
    // Non-outlook tokens are minted by the Teams/Copilot surfaces.
    waitForAll: res !== DEFAULT_RESOURCE,
    openExtraSurfaces: res !== DEFAULT_RESOURCE,
  });
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
