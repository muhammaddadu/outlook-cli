// Auth lifecycle: load from cache, capture via browser when stale, save back.
//
// The captured headers include the Bearer JWT. We decode its `exp` claim to
// know when the cache becomes stale; we refresh proactively with a 5-minute
// buffer so a long-running command never trips a mid-call expiry.

import {
  readFileSync,
  writeFileSync,
  mkdirSync,
  existsSync,
  unlinkSync,
} from 'node:fs';
import { dirname } from 'node:path';
import { tokenCacheFile } from './paths.mjs';
import { debug } from './output.mjs';

// Lazy import — capture.mjs pulls in Playwright (≈80MB of resident memory
// and a Chromium executable check). Cache-hit calls never hit this code
// path and shouldn't pay that cost.
async function captureAuth(opts) {
  const mod = await import('./capture.mjs');
  return mod.captureAuth(opts);
}

const MIN_LIFETIME_MS = 5 * 60 * 1000;

function decodeJwtExp(bearer) {
  const jwt = bearer.replace(/^Bearer\s+/i, '');
  const parts = jwt.split('.');
  if (parts.length < 2) return null;
  const b64 = parts[1].replace(/-/g, '+').replace(/_/g, '/');
  const padded = b64 + '='.repeat((4 - (b64.length % 4)) % 4);
  try {
    const payload = JSON.parse(Buffer.from(padded, 'base64').toString('utf8'));
    return typeof payload.exp === 'number' ? payload.exp * 1000 : null;
  } catch {
    return null;
  }
}

export function loadCachedAuth() {
  if (!existsSync(tokenCacheFile())) return null;
  let entry;
  try {
    entry = JSON.parse(readFileSync(tokenCacheFile(), 'utf8'));
  } catch {
    return null;
  }
  if (!entry?.headers?.authorization) return null;

  const exp = decodeJwtExp(entry.headers.authorization);
  if (!exp || exp - Date.now() < MIN_LIFETIME_MS) {
    debug('cached token is stale or unparseable');
    return null;
  }
  return { headers: entry.headers, expiresAt: exp };
}

export function saveAuth(headers) {
  mkdirSync(dirname(tokenCacheFile()), { recursive: true });
  const exp = decodeJwtExp(headers.authorization ?? '');
  writeFileSync(
    tokenCacheFile(),
    JSON.stringify({ headers, expiresAt: exp, savedAt: Date.now() }, null, 2),
  );
}

export function clearAuth() {
  if (existsSync(tokenCacheFile())) unlinkSync(tokenCacheFile());
}

/**
 * Get headers for an authed API call. Uses the on-disk cache when the JWT
 * still has > MIN_LIFETIME_MS left; otherwise opens OWA to capture a fresh
 * token via silent SSO and persists it.
 */
export async function getAuth({ forceRefresh = false } = {}) {
  if (!forceRefresh) {
    const cached = loadCachedAuth();
    if (cached) {
      const minsLeft = Math.floor((cached.expiresAt - Date.now()) / 60_000);
      debug(`using cached token (~${minsLeft}m until expiry)`);
      return cached.headers;
    }
  }
  debug('refreshing token via OWA');
  const auth = await captureAuth();
  saveAuth(auth);
  return auth;
}
