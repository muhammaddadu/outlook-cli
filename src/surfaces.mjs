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
