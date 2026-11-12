// `outlook keepalive`: keep one signed-in browser running so tokens stay fresh.
//
// Why: Entra keeps the SSO session in ESTSAUTH, a session cookie that only
// lives as long as the browser process (tested 2026-11-11: even with
// "continue where you left off" the bundled Chromium doesn't write it to
// disk). An everyday Chrome stays signed in because it never quits. This does
// the same: it opens OWA, Teams and Copilot once (you sign in there, MFA and
// all), minimises the window, and leaves the web apps running. Their own MSAL
// clients renew access tokens as they always do; we read the Bearer headers
// off the wire and write them to the normal token cache, so every other
// `outlook` command keeps working from the cache with no browser of its own.
//
// Nothing here reads or replays refresh tokens or cookies: the browser does
// all the authentication, exactly as it would for the user.

import {
  readFileSync,
  writeFileSync,
  existsSync,
  unlinkSync,
  mkdirSync,
} from 'node:fs';
import { dirname, resolve } from 'node:path';
import { homedir } from 'node:os';
import { cacheDir } from './paths.mjs';
import { debug, info } from './output.mjs';
import { AppError, E } from './errors.mjs';
import { decodePayload } from './jwt.mjs';

export const KEEPALIVE_SURFACES = Object.freeze([
  'https://outlook.office.com/mail/',
  'https://teams.microsoft.com/v2/',
  'https://m365.cloud.microsoft/chat/',
]);

export const LAUNCHD_LABEL = 'com.outlook-cli.keepalive';

export function statusFile() {
  return process.env.OUTLOOK_KEEPALIVE_STATUS ?? resolve(cacheDir(), 'keepalive.json');
}

/**
 * Last status the keepalive wrote, plus whether its process is alive. Lets
 * unattended callers tell "keepalive down" from "sign-in needed".
 * state: signing-in | signed-in | stopped
 */
export function readKeepaliveStatus() {
  let status = null;
  try {
    status = JSON.parse(readFileSync(statusFile(), 'utf8'));
  } catch {}
  const pid = keepalivePid();
  return { running: !!pid, pid, ...(status ?? {}), ...(pid ? {} : { state: 'stopped' }) };
}

export function pidFile() {
  return process.env.OUTLOOK_KEEPALIVE_PID ?? resolve(cacheDir(), 'keepalive.pid');
}

/** PID of a live keepalive process, or null. Cleans up a stale pid file. */
export function keepalivePid() {
  const file = pidFile();
  if (!existsSync(file)) return null;
  const pid = Number.parseInt(readFileSync(file, 'utf8').trim(), 10);
  if (!Number.isInteger(pid) || pid <= 0) return null;
  try {
    process.kill(pid, 0);
    return pid;
  } catch (err) {
    if (err.code === 'EPERM') return pid; // alive, owned by someone else
    try {
      unlinkSync(file);
    } catch {}
    return null;
  }
}

/**
 * Working-hours check for the periodic reload. `days` uses Date#getDay
 * (0 = Sunday); `start`/`end` are "HH:MM" local time.
 */
export function isWithinHours(date, { days = [1, 2, 3, 4, 5], start = '07:55', end = '19:00' } = {}) {
  if (!days.includes(date.getDay())) return false;
  const mins = date.getHours() * 60 + date.getMinutes();
  const toMins = (s) => {
    const [h, m] = s.split(':').map(Number);
    return h * 60 + m;
  };
  return mins >= toMins(start) && mins < toMins(end);
}

/** True when `incoming` carries a different, later-expiring token than `current`. */
export function isNewerToken(current, incoming) {
  const inc = incoming?.authorization;
  if (!inc) return false;
  const cur = current?.authorization;
  if (!cur) return true;
  if (cur === inc) return false;
  const expOf = (b) => decodePayload(b)?.exp ?? 0;
  return expOf(inc) >= expOf(cur);
}

/**
 * Ask a running keepalive to reload its tabs now (SIGUSR1). Returns false
 * if none is running.
 */
export function pokeKeepalive() {
  const pid = keepalivePid();
  if (!pid) return false;
  try {
    process.kill(pid, 'SIGUSR1');
    return true;
  } catch {
    return false;
  }
}

/** launchd agent: start at login, and again at 07:55 on weekdays if not running. */
export function launchdPlist({ nodePath, cliPath, logFile }) {
  const intervals = [1, 2, 3, 4, 5]
    .map(
      (d) =>
        `      <dict><key>Weekday</key><integer>${d}</integer><key>Hour</key><integer>7</integer><key>Minute</key><integer>55</integer></dict>`,
    )
    .join('\n');
  return `<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
  <dict>
    <key>Label</key><string>${LAUNCHD_LABEL}</string>
    <key>ProgramArguments</key>
    <array>
      <string>${nodePath}</string>
      <string>${cliPath}</string>
      <string>keepalive</string>
    </array>
    <key>RunAtLoad</key><true/>
    <key>StartCalendarInterval</key>
    <array>
${intervals}
    </array>
    <key>StandardOutPath</key><string>${logFile}</string>
    <key>StandardErrorPath</key><string>${logFile}</string>
  </dict>
</plist>
`;
}

export function launchdPlistPath() {
  return resolve(homedir(), 'Library', 'LaunchAgents', `${LAUNCHD_LABEL}.plist`);
}

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const LOGIN_URL_RE = /login\.(microsoft|microsoftonline)\.com/;

/**
 * Run the keepalive until the browser is closed or the process is signalled.
 * @param {object} opts
 * @param {number} [opts.reloadMinutes=30] reload the tabs this often (within hours)
 * @param {object} [opts.hours] see isWithinHours
 */
export async function runKeepalive({ reloadMinutes = 30, hours } = {}) {
  const existing = keepalivePid();
  if (existing) {
    info(`keepalive is already running (pid ${existing}).`);
    return;
  }
  const { openContext, attachTokenSniffer } = await import('./capture.mjs');
  const { loadCachedAuth, saveAuth, saveVariant } = await import('./auth.mjs');

  mkdirSync(dirname(pidFile()), { recursive: true });
  writeFileSync(pidFile(), String(process.pid));
  const cleanupPid = () => {
    try {
      if (readFileSync(pidFile(), 'utf8').trim() === String(process.pid)) unlinkSync(pidFile());
    } catch {}
  };

  const status = { pid: process.pid, startedAt: new Date().toISOString(), state: 'signing-in', lastReloadAt: null, lastError: null, tokens: {} };
  const writeStatus = (patch = {}) => {
    Object.assign(status, patch, { updatedAt: new Date().toISOString() });
    try {
      writeFileSync(statusFile(), JSON.stringify(status, null, 2));
    } catch (e) {
      debug(`status write failed: ${e.message}`);
    }
  };
  writeStatus();

  const context = await openContext({ headless: false });
  let stopping = false;
  let reloadRequested = false;
  const stop = async () => {
    if (stopping) return;
    stopping = true;
    writeStatus({ state: 'stopped' });
    await context.close().catch(() => {});
    cleanupPid();
    process.exit(0);
  };
  process.on('SIGTERM', stop);
  process.on('SIGINT', stop);
  process.on('SIGUSR1', () => {
    debug('reload requested (SIGUSR1)');
    reloadRequested = true;
  });

  attachTokenSniffer(context, (key, bag) => {
    // loadCachedAuth returns null for stale tokens, so a stale cache is
    // always replaced.
    const savedVariant = saveVariant(bag, key); // keep every app's token (scopes differ per app)
    const current = loadCachedAuth(key)?.headers;
    const savedPrimary = isNewerToken(current, bag);
    if (savedPrimary) saveAuth(bag, key);
    if (savedVariant || savedPrimary) {
      const claims = decodePayload(bag.authorization) ?? {};
      const exp = claims.exp;
      status.tokens[`${key}@${claims.appid ?? '?'}`] = { savedAt: new Date().toISOString(), expiresAt: exp ? new Date(exp * 1000).toISOString() : null };
      writeStatus();
      const mins = Math.round(((decodePayload(bag.authorization)?.exp ?? 0) * 1000 - Date.now()) / 60_000);
      debug(`keepalive saved ${key} token (~${mins}m)`);
    }
  });

  // OneDrive mints the SharePoint token transcripts need; its URL comes from
  // Graph, so it is only added once a Graph token is cached (the next
  // keepalive start after the first sign-in at the latest).
  const { oneDriveSurface } = await import('./surfaces.mjs');
  const oneDrive = await oneDriveSurface();
  if (!oneDrive) debug('no OneDrive URL yet (no Graph token cached); SharePoint token will not be captured this run');
  const surfaces = oneDrive ? [...KEEPALIVE_SURFACES, oneDrive] : KEEPALIVE_SURFACES;

  const pages = [];
  for (const url of surfaces) {
    const page = pages.length === 0 ? (context.pages()[0] ?? (await context.newPage())) : await context.newPage();
    pages.push(page);
    page.on('crash', () => {
      writeStatus({ lastError: `${new Date().toISOString()} tab crashed: ${url}` });
      info(`keepalive: tab crashed (${url}); reloading.`);
      page.reload().catch(() => {});
    });
    page.goto(url, { waitUntil: 'domcontentloaded', timeout: 120_000 }).catch((e) => debug(`goto ${url}: ${e.message}`));
  }

  // Window state via CDP: minimise once signed in, bring back if a login
  // page needs the user.
  let cdp = null;
  let windowId = null;
  try {
    cdp = await context.newCDPSession(pages[0]);
    ({ windowId } = await cdp.send('Browser.getWindowForTarget'));
  } catch (e) {
    debug(`window control unavailable: ${e.message}`);
  }
  let windowState = 'normal';
  const setWindow = async (state) => {
    if (!cdp || windowId == null || state === windowState) return;
    try {
      await cdp.send('Browser.setWindowBounds', { windowId, bounds: { windowState: state } });
      windowState = state;
    } catch (e) {
      debug(`setWindowBounds ${state}: ${e.message}`);
    }
  };

  info('keepalive running. Sign in in the browser window if asked; it minimises itself once signed in. Leave it open.');
  let lastReload = Date.now();
  let lastTick = Date.now();
  let toldToSignIn = false;
  for (;;) {
    // A gap much longer than the 5s tick means the Mac slept. The web apps'
    // tokens have likely expired meanwhile: reload so they renew now.
    if (Date.now() - lastTick > 60_000) {
      debug(`woke after ${Math.round((Date.now() - lastTick) / 1000)}s; reloading`);
      reloadRequested = true;
    }
    lastTick = Date.now();
    if (pages.every((p) => p.isClosed())) {
      info('keepalive: browser window closed; exiting.');
      return stop();
    }
    const onLogin = pages.some((p) => !p.isClosed() && LOGIN_URL_RE.test(p.url()));
    const signedIn = !!loadCachedAuth('outlook');
    if (onLogin) {
      await setWindow('normal');
      if (status.state !== 'signing-in') writeStatus({ state: 'signing-in' });
      if (!toldToSignIn) {
        info('keepalive: Microsoft is asking you to sign in — complete it in the browser window.');
        toldToSignIn = true;
      }
    } else if (signedIn) {
      await setWindow('minimized');
      if (status.state !== 'signed-in') writeStatus({ state: 'signed-in' });
      toldToSignIn = false;
    }

    const due = Date.now() - lastReload >= reloadMinutes * 60_000 && isWithinHours(new Date(), hours);
    if ((due || reloadRequested) && !onLogin) {
      reloadRequested = false;
      lastReload = Date.now();
      for (const p of pages) {
        if (!p.isClosed()) p.reload({ waitUntil: 'domcontentloaded', timeout: 120_000 }).catch((e) => debug(`reload: ${e.message}`));
      }
      writeStatus({ lastReloadAt: new Date().toISOString() });
      debug('keepalive reloaded tabs');
    }
    await sleep(5_000);
  }
}

/** Error for commands that would need the browser while keepalive owns the profile. */
export function keepaliveOwnsProfileError(pid) {
  return new AppError({
    code: E.AUTH_REQUIRED,
    message: `The keepalive browser (pid ${pid}) is using the sign-in profile, and it has no fresh token yet.`,
    hint: 'If its window is asking you to sign in, complete it there. Otherwise wait a minute and retry, or run `outlook keepalive --stop` first.',
  });
}
