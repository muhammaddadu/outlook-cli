// Token audit: decode every cached resource token and report what it can
// reach. Pure/offline — reads only the on-disk caches, never the network —
// so it works even against an expired token (it will say so). This is the
// instrument for answering "do we have access to Graph / Teams / Copilot?".

import { existsSync, readFileSync } from 'node:fs';
import { tokenCacheFile } from './paths.mjs';
import { decodePayload } from './jwt.mjs';
import { RESOURCES } from './resources.mjs';

// Scope prefixes → capability area. A token's `scp` claim lists the
// delegated permissions consented for that resource; grouping them tells a
// human (or agent) which feature families are unlocked without reading 80
// raw scope strings. First matching group wins; order matters.
const SCOPE_GROUPS = [
  ['mail', /^Mail\.|FocusedInbox/i],
  ['calendar', /^Calendars|OnlineMeetings|^Place|Locations/i],
  ['teams-chat', /^Chat\.|^Channel|^Team\.|^Team$|^Group\.|Collab/i],
  ['files', /^Files\./i],
  ['copilot', /Copilot/i],
  ['search', /Search|^Signals?\b|^Signal\./i],
  ['contacts-people', /^Contacts|^People/i],
  ['tasks-notes', /^Tasks|^Todo|^Notes/i],
  ['directory-users', /^Directory|^User[.\-]/i],
];

/** Group a space-delimited `scp` string into capability areas. */
export function groupScopes(scp) {
  const scopes = (scp ?? '').split(/\s+/).filter(Boolean);
  /** @type {Record<string,string[]>} */
  const groups = {};
  const other = [];
  for (const s of scopes) {
    const hit = SCOPE_GROUPS.find(([, re]) => re.test(s));
    if (hit) (groups[hit[0]] ??= []).push(s);
    else other.push(s);
  }
  if (other.length) groups.other = other;
  return groups;
}

/** Audit one resource's cached token. Never throws — missing/garbage → status. */
export function auditResource(key) {
  const meta = RESOURCES[key];
  const file = tokenCacheFile(key);
  const out = { resource: key, label: meta.label, cacheFile: file };

  if (!existsSync(file)) return { ...out, status: 'absent' };

  let entry;
  try {
    entry = JSON.parse(readFileSync(file, 'utf8'));
  } catch {
    return { ...out, status: 'unreadable' };
  }
  const bearer = entry?.headers?.authorization;
  const claims = decodePayload(bearer);
  if (!claims) return { ...out, status: 'no-jwt' };

  const expMs = typeof claims.exp === 'number' ? claims.exp * 1000 : null;
  const minutesLeft = expMs ? Math.round((expMs - Date.now()) / 60_000) : null;

  return {
    ...out,
    status: expMs && expMs > Date.now() ? 'live' : 'expired',
    audience: claims.aud ?? null,
    appId: claims.appid ?? null,
    appName: claims.app_displayname ?? null,
    user: claims.upn ?? claims.unique_name ?? null,
    tenantId: claims.tid ?? null,
    expiresAt: expMs ? new Date(expMs).toISOString() : null,
    minutesUntilExpiry: minutesLeft,
    capabilities: groupScopes(claims.scp),
    scopeCount: (claims.scp ?? '').split(/\s+/).filter(Boolean).length,
  };
}

/** Audit every known resource. */
export function auditAll() {
  const resources = Object.keys(RESOURCES).map(auditResource);
  return {
    resources,
    reachable: resources.filter((r) => r.status === 'live').map((r) => r.resource),
    note:
      'A token is only usable against its own resource (audience). "live" ' +
      'means a non-expired token is cached; run `outlook auth --all` to add ' +
      'missing resources.',
  };
}
