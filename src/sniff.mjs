#!/usr/bin/env node
// Live, PII-safe capture + endpoint-learning session.
//
// Opens a headed browser against a persistent profile. You sign in and click
// around (Outlook, Teams, Copilot); this tool watches the network and, for
// every Microsoft API request, records ONLY:
//   - method + host + path with all IDs/emails/tokens redacted
//   - the token audience (so we learn which resource each surface uses)
//   - for POST/PATCH bodies: the JSON *shape* (keys + value types), never the
//     values — so we learn request schemas without capturing message text,
//     names, or Copilot prompts.
//
// It also saves any Bearer token whose audience maps to a known resource
// (outlook/graph/substrate) into the normal per-resource cache, so `outlook
// graph …` works the moment a token appears. Response bodies are never read.
//
// Usage: node src/sniff.mjs [startUrl]   (Ctrl-C to stop; prints a summary)

import { openContext, FORWARDED_HEADERS, EXTRA_SURFACES } from './capture.mjs';
import { decodePayload } from './jwt.mjs';
import { classifyToken } from './resources.mjs';
import { saveAuth } from './auth.mjs';

const START_URL = process.argv[2] ?? 'https://outlook.office.com/mail/';

const MS_HOST_RE =
  /(^|\.)(microsoft\.com|office\.com|office365\.com|cloud\.microsoft|skype\.com|microsoftonline\.com|office\.net|teams\.microsoft\.us|sharepoint\.com|live\.com)$/i;
const ASSET_RE = /\.(js|css|png|jpe?g|svg|gif|woff2?|ico|map|webp|mp4|wasm)(\?|$)/i;

function redactPath(pathname) {
  return pathname
    .replace(/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/gi, '{guid}')
    .replace(/[^/]+@[^/]+/g, '{email}')
    .replace(/\b[0-9A-Fa-f]{16}\b/g, '{puid}')
    .replace(/[A-Za-z0-9_\-=+]{28,}/g, '{id}'); // long opaque ids (base64-ish)
}

/** Structure of a JSON value with primitives replaced by their type — no values. */
function shapeOf(v, depth = 0) {
  if (depth > 4) return '…';
  if (Array.isArray(v)) return v.length ? [shapeOf(v[0], depth + 1)] : [];
  if (v === null) return 'null';
  if (typeof v === 'object') {
    const o = {};
    for (const k of Object.keys(v).slice(0, 40)) o[k] = shapeOf(v[k], depth + 1);
    return o;
  }
  return typeof v; // 'string' | 'number' | 'boolean'
}

const endpoints = new Map(); // "METHOD host/path" -> { count, auds:Set, qkeys:Set, bodyShape }
const audiences = new Map(); // aud -> { hostSample, count }
const savedResources = new Map(); // resource -> aud (last saved)
let sinceFlush = 0;

const context = await openContext({ headless: false });
const page = context.pages()[0] ?? (await context.newPage());

context.on('request', (req) => {
  let url;
  try {
    url = new URL(req.url());
  } catch {
    return;
  }
  if (!MS_HOST_RE.test(url.hostname)) return;
  if (ASSET_RE.test(url.pathname)) return;

  const h = req.headers();
  const bearer = h.authorization?.startsWith('Bearer ') ? h.authorization : null;
  const claims = bearer ? decodePayload(bearer) : null;
  const aud = claims?.aud ?? null;

  // Learn every audience we see (including Teams-specific ones we don't model).
  if (aud) {
    const a = audiences.get(aud) ?? { hostSample: url.hostname, count: 0 };
    a.count++;
    audiences.set(aud, a);
    // Save tokens for resources we know how to route to.
    const key = classifyToken({ aud, host: url.hostname });
    if (key && savedResources.get(key) !== aud) {
      const bag = {};
      for (const k of FORWARDED_HEADERS) if (h[k]) bag[k] = h[k];
      try {
        saveAuth(bag, key);
        savedResources.set(key, aud);
        console.log(`\n  ✓ saved ${key} token (aud=${aud})`);
        sinceFlush++;
      } catch (e) {
        console.log(`  ! could not save ${key} token: ${e.message}`);
      }
    }
  }

  const method = req.method();
  const epKey = `${method} ${url.hostname}${redactPath(url.pathname)}`;
  const ep =
    endpoints.get(epKey) ??
    { count: 0, auds: new Set(), qkeys: new Set(), bodyShape: null };
  ep.count++;
  if (aud) ep.auds.add(aud);
  for (const k of url.searchParams.keys()) ep.qkeys.add(k);
  if (!ep.bodyShape && method !== 'GET' && method !== 'DELETE') {
    const raw = req.postData();
    if (raw) {
      try {
        ep.bodyShape = shapeOf(JSON.parse(raw));
      } catch {
        ep.bodyShape = '(non-JSON body)';
      }
    }
  }
  const wasNew = !endpoints.has(epKey);
  endpoints.set(epKey, ep);
  if (wasNew) sinceFlush++;
});

console.log('Sniffer running. A browser window is opening.');
console.log('1) Sign in (MFA) if prompted.');
console.log('2) Click into Teams — open a chat, open a channel.');
console.log('3) Open Copilot / M365 Chat and send one prompt.');
console.log('Tokens save automatically as they appear. Ctrl-C when done.\n');

await page.goto(START_URL, { waitUntil: 'domcontentloaded' }).catch(() => {});
for (const extra of EXTRA_SURFACES) {
  context.newPage().then((p) => p.goto(extra, { waitUntil: 'domcontentloaded' }).catch(() => {})).catch(() => {});
}

// Periodically print a delta so an observer tailing stdout learns live.
const timer = setInterval(() => {
  if (sinceFlush === 0) return;
  sinceFlush = 0;
  console.log(
    `\n--- discovered so far: ${endpoints.size} endpoints, ${audiences.size} audiences, saved: [${[...savedResources.keys()].join(', ') || 'none'}] ---`,
  );
  printSummary();
}, 4000);

function printSummary() {
  const byHost = [...endpoints.entries()].sort((a, b) =>
    a[0].split(' ')[1].localeCompare(b[0].split(' ')[1]),
  );
  for (const [k, v] of byHost) {
    const audShort = [...v.auds]
      .map((a) => a.replace(/^https?:\/\//, '').replace(/\/$/, ''))
      .join(',');
    console.log(`  [${v.count}x] ${k}${audShort ? `  aud=${audShort}` : ''}`);
    if (v.qkeys.size) console.log(`        ?params: ${[...v.qkeys].join(', ')}`);
    if (v.bodyShape) console.log(`        body: ${JSON.stringify(v.bodyShape)}`);
  }
  console.log('\n  audiences seen:');
  for (const [a, info] of audiences) console.log(`    ${a}  (${info.count}x, e.g. ${info.hostSample})`);
}

function finish() {
  clearInterval(timer);
  console.log('\n========== FINAL SUMMARY ==========');
  printSummary();
  console.log(`\nSaved resource tokens: ${[...savedResources.keys()].join(', ') || 'none'}`);
  context.close().catch(() => {}).finally(() => process.exit(0));
}

process.on('SIGINT', finish);
process.on('SIGTERM', finish);
