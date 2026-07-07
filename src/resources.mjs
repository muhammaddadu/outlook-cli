// Registry of the Microsoft resources this CLI can talk to.
//
// Each Microsoft API is a distinct OAuth *resource* with its own token
// audience. A token minted for one resource is rejected (401
// InvalidAuthenticationToken) by every other resource — this is why the
// mail token cannot call Graph. To reach Teams / Copilot / Files we capture
// a separate token per resource from the browser and route each call to the
// matching base URL with the matching token.
//
// `aud` values: Microsoft issues audiences either as the resource URL
// (e.g. "https://graph.microsoft.com") or as the resource's app GUID
// (e.g. Graph's "00000003-0000-0000-c000-000000000000"). We match both.

/**
 * @typedef {Object} Resource
 * @property {string} key        short id used in cache filenames + CLI flags
 * @property {string} base       default API base URL (no trailing slash)
 * @property {string} envBase    env var that overrides `base` (tests)
 * @property {string[]} audiences JWT `aud` values that map to this resource
 * @property {string} host       request host that identifies this resource on the wire
 * @property {string} label      human name for audit output
 */

/** @type {Record<string, Resource>} */
export const RESOURCES = Object.freeze({
  outlook: {
    key: 'outlook',
    base: 'https://outlook.office.com/api/v2.0/me',
    envBase: 'OUTLOOK_API_BASE',
    audiences: ['https://outlook.office.com', 'https://outlook.office.com/'],
    host: 'outlook.office.com',
    label: 'Outlook REST v2 (mail, calendar)',
  },
  graph: {
    key: 'graph',
    base: 'https://graph.microsoft.com/v1.0',
    envBase: 'OUTLOOK_GRAPH_BASE',
    audiences: [
      'https://graph.microsoft.com',
      'https://graph.microsoft.com/',
      '00000003-0000-0000-c000-000000000000',
    ],
    host: 'graph.microsoft.com',
    label: 'Microsoft Graph (Teams, Files, People, Groups)',
  },
  substrate: {
    key: 'substrate',
    base: 'https://substrate.office.com',
    envBase: 'OUTLOOK_SUBSTRATE_BASE',
    audiences: ['https://substrate.office.com', 'https://substrate.office.com/'],
    host: 'substrate.office.com',
    label: 'Substrate (Copilot, unified search)',
  },
});

export const DEFAULT_RESOURCE = 'outlook';

/** Look up a resource by key; throws a friendly list if unknown. */
export function resource(key) {
  const r = RESOURCES[key];
  if (!r) {
    const known = Object.keys(RESOURCES).join(', ');
    throw new Error(`Unknown resource "${key}". Known resources: ${known}.`);
  }
  return r;
}

/** Base URL for a resource, honouring its env override (used by tests). */
export function resourceBase(key) {
  const r = resource(key);
  return process.env[r.envBase] ?? r.base;
}

/**
 * Classify a token by its `aud` claim and/or the host it was sent to.
 * Returns the resource key, or null if it belongs to no resource we model.
 */
export function classifyToken({ aud, host } = {}) {
  if (aud) {
    for (const r of Object.values(RESOURCES)) {
      if (r.audiences.includes(aud)) return r.key;
    }
  }
  if (host) {
    for (const r of Object.values(RESOURCES)) {
      if (host === r.host) return r.key;
    }
  }
  return null;
}
