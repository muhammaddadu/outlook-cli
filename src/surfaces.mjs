// Web surfaces opened to mint per-tenant tokens.
//
// The OneDrive web app is what calls the tenant's "-my.sharepoint.com" host
// with a SharePoint-audience Bearer (it gets one from SharePoint's own token
// endpoint), so opening it lets the token sniffer capture the token
// `outlook transcript` needs. Its URL is per user and per tenant, so we ask
// Graph for it with a cached Graph token instead of hard-coding a host.

/** The user's OneDrive web URL, or null if no Graph token is cached or Graph fails. */
export async function oneDriveSurface() {
  if (process.env.OUTLOOK_ONEDRIVE_URL) return process.env.OUTLOOK_ONEDRIVE_URL;
  const { loadCachedAuth, loadVariants } = await import('./auth.mjs');
  const { call } = await import('./client.mjs');
  const auth = loadCachedAuth('graph')?.headers ?? loadVariants('graph')[0]?.headers;
  if (!auth) return null;
  try {
    const res = await call(auth, '/me/drive?$select=webUrl', { resource: 'graph' });
    return res.status === 200 && typeof res.body?.webUrl === 'string' ? res.body.webUrl : null;
  } catch {
    return null;
  }
}

/**
 * Teams web URL of a meeting recap page for a recording driveItem (needs
 * id, webUrl, parentReference.driveId and source). Uses the web client's
 * own route (`/_#/l/...`): the public `/l/meetingrecap` link often stops at
 * the "use the desktop app?" launcher page.
 */
export function recapUrl(item) {
  const s = item?.source ?? {};
  if (!item?.id || !item.webUrl || !s.threadId) return null;
  const q = new URLSearchParams({
    driveId: item.parentReference?.driveId ?? '',
    driveItemId: item.id,
    sitePath: item.webUrl,
    fileUrl: item.webUrl,
    iCalUid: s.iCalUid ?? '',
    threadId: s.threadId,
    organizerId: s.meetingOrganizerId ?? '',
    tenantId: s.meetingOrganizerTenantId ?? '',
    callId: s.externalId ?? '',
    threadType: 'meeting',
    subType: 'RecapSharingLink_RecapCore',
  });
  return `https://teams.microsoft.com/_#/l/meetingrecap?${q}`;
}

/**
 * Recap page of the newest recording in the user's OneDrive that has a
 * transcript, or null. Opening a recap is what makes Teams web mint the
 * Substrate-audience token that `outlook recap` needs; nothing else in
 * Teams web requests one.
 */
export async function recapSurface() {
  if (process.env.OUTLOOK_RECAP_SURFACE) return process.env.OUTLOOK_RECAP_SURFACE;
  const { loadCachedAuth, loadVariants } = await import('./auth.mjs');
  const { call } = await import('./client.mjs');
  const auth = loadCachedAuth('graph')?.headers ?? loadVariants('graph')[0]?.headers;
  if (!auth) return null;
  try {
    const res = await call(
      auth,
      '/me/drive/root:/Recordings:/children?$top=50&$orderby=lastModifiedDateTime desc&$select=id,name,webUrl,parentReference,source,media',
      { resource: 'graph' },
    );
    if (res.status !== 200) return null;
    const item = (res.body?.value ?? []).find(
      (it) => it.media?.mediaSource?.contentType === 'Transcript' && it.source?.threadId,
    );
    return item ? recapUrl(item) : null;
  } catch {
    return null;
  }
}
