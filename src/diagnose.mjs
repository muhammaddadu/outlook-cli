#!/usr/bin/env node
// Sniff the network for ~10 seconds while OWA boots, then print the API
// endpoints it actually calls. This tells us which URL pattern + headers to
// replay in outlook.mjs.

import { openContext, HOME_URL } from './capture.mjs';

const context = await openContext({ headless: false });
const page = await context.newPage();
page.on('console', (m) => console.log('  [page]', m.text().slice(0, 120)));

const seen = new Map(); // method+url -> { count, sampleHeaders }

page.on('request', (req) => {
  const url = req.url();
  // Filter to plausible "fetch the inbox" endpoints
  // Look at anything that smells like a data API (not static assets).
  if (/\.(js|css|png|jpg|svg|woff2?|ico|map)(\?|$)/.test(url)) return;
  if (!/outlook\.office\.com|graph\.microsoft\.com|substrate\.office\.com|office\.com/.test(url)) {
    return;
  }
  const key = `${req.method()} ${url.split('?')[0]}`;
  const existing = seen.get(key);
  if (existing) {
    existing.count++;
    return;
  }
  // Capture which auth shape this endpoint uses, but redact the secrets.
  const headers = req.headers();
  const authShape = headers.authorization
    ? `Bearer ${headers.authorization.slice(7, 27)}…(${headers.authorization.length} chars)`
    : 'cookie-only';
  seen.set(key, {
    count: 1,
    method: req.method(),
    url,
    auth: authShape,
    extraHeaders: Object.fromEntries(
      Object.entries(headers).filter(([k]) =>
        /^(x-|action|prefer|client-request-id|content-type)/i.test(k),
      ),
    ),
  });
});

await page.goto(HOME_URL, { waitUntil: 'domcontentloaded' });
console.log('Waiting for inbox to render (up to 60s)…');
try {
  // OWA renders a mail list with role="listbox" once the inbox loads. If that
  // never shows, fall through after the timeout so we can still see what was
  // captured.
  await page.waitForSelector('[role="listbox"], [aria-label*="Message list" i]', {
    timeout: 60_000,
  });
  console.log('Inbox visible. Sniffing 5 more seconds for follow-up calls…');
  await page.waitForTimeout(5_000);
} catch {
  console.log('Inbox selector never appeared; printing what we saw anyway.');
}
console.log(`\nFinal page URL: ${page.url()}\n`);

const sorted = [...seen.values()].sort((a, b) => b.count - a.count);
for (const r of sorted) {
  console.log(`\n[${r.count}x] ${r.method} ${r.url.slice(0, 140)}`);
  console.log(`  auth: ${r.auth}`);
  for (const [k, v] of Object.entries(r.extraHeaders)) {
    const display = v.length > 80 ? v.slice(0, 80) + '…' : v;
    console.log(`  ${k}: ${display}`);
  }
}

await context.close();
