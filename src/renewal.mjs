// Which tokens the keepalive renews on purpose, and when.
//
// OWA renews its own tokens all day. Teams web's Graph token, the Substrate
// token `outlook recap` needs and the SharePoint token `outlook transcript`
// needs are only minted when a page asks for them, so the keepalive watches
// their expiry in the cache and re-opens the tab that mints each one before
// it lapses (seen 2026-11-16: reloading every 30 min left all three expired
// for hours). Pure helpers; keepalive.mjs and auth.mjs do the I/O.

export const TEAMS_APPID = '5e3ce6c0-2b1f-4285-8d4b-75ee78787346';
export const RECAP_AUDIENCES = new Set(['https://substrate.office.com', 'https://substrate.office.com/']);

/** Re-open a tab when one of its tokens has less than this left. */
export const RENEW_MARGIN_MS = 10 * 60_000;
/** ...but not more often than this per tab (MSAL may reuse a token that is
 * still valid, so the next try, closer to expiry, gets a new one). */
export const RENEW_RETRY_MS = 5 * 60_000;

/**
 * Tokens the keepalive keeps fresh. `tabs` names the keepalive tabs whose
 * pages mint the token; `match` picks it among the cached tokens of
 * `resource`. SharePoint: live on 2026-11-16 the usable token came from
 * Teams web (its appid), and re-opening OneDrive alone never renewed it.
 */
export const RENEWED_TOKENS = Object.freeze([
  { name: 'teams-graph', resource: 'graph', tabs: ['teams'], match: (c) => c.appid === TEAMS_APPID },
  { name: 'recap-substrate', resource: 'substrate', tabs: ['teams'], match: (c) => RECAP_AUDIENCES.has(c.aud) },
  { name: 'sharepoint', resource: 'sharepoint', tabs: ['teams', 'onedrive'], match: () => true },
]);

/**
 * Latest expiry (ms) per renewed token, from cached tokens given as
 * { resource: [claims, ...] }. Missing tokens map to null.
 */
export function tokenExpiries(claimsByResource) {
  const out = {};
  for (const t of RENEWED_TOKENS) {
    const exps = (claimsByResource[t.resource] ?? [])
      .filter((c) => c && typeof c.exp === 'number' && t.match(c))
      .map((c) => c.exp * 1000);
    out[t.name] = exps.length ? Math.max(...exps) : null;
  }
  return out;
}

/** Names of renewed tokens with less than `margin` left (or none cached). */
export function staleTokens(expiries, now, margin = RENEW_MARGIN_MS) {
  return RENEWED_TOKENS.filter((t) => {
    const exp = expiries[t.name];
    return exp == null || exp - now < margin;
  }).map((t) => t.name);
}

/**
 * Tabs to re-open now: those that exist (`tabs`), own a stale token, and were
 * last opened at least RENEW_RETRY_MS ago (`lastOpened`: tab -> ms).
 */
export function tabsToRenew({ expiries, now, tabs, lastOpened = {} }) {
  const stale = new Set(staleTokens(expiries, now));
  const due = new Set();
  for (const t of RENEWED_TOKENS) {
    if (!stale.has(t.name)) continue;
    for (const tab of t.tabs) {
      if (tabs.has(tab) && now - (lastOpened[tab] ?? 0) >= RENEW_RETRY_MS) due.add(tab);
    }
  }
  return [...due];
}
