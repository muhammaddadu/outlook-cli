// Test helpers — JWT minting, fake token-cache seeding, mock HTTP server,
// and CLI subprocess runner.
//
// Everything here is deliberately small and dep-free. We rely on Node's
// built-in `node:test`, `node:http`, and `node:child_process`.

import { createServer } from 'node:http';
import { spawn } from 'node:child_process';
import { mkdtempSync, writeFileSync, mkdirSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const HERE = dirname(fileURLToPath(import.meta.url));
export const CLI_PATH = resolve(HERE, '..', 'src', 'cli.mjs');

/**
 * Build a JWT-shaped string. The signature segment is placeholder text —
 * the CLI never validates the signature, only decodes the payload to read
 * `exp`.
 */
export function makeFakeJwt({ expSeconds, claims = {} } = {}) {
  const enc = (obj) =>
    Buffer.from(JSON.stringify(obj))
      .toString('base64')
      .replace(/=+$/, '')
      .replace(/\+/g, '-')
      .replace(/\//g, '_');
  const header = enc({ typ: 'JWT', alg: 'none' });
  const payload = enc({
    aud: 'https://outlook.office.com/',
    exp: expSeconds,
    ...claims,
  });
  return `${header}.${payload}.test-signature-not-validated`;
}

/**
 * Create a fresh tmp directory and write a token cache file inside it whose
 * Bearer JWT expires `secondsFromNow` from now. Returns the file path.
 */
export function seedTokenCache({
  secondsFromNow = 3600,
  headers = {},
  claims = {},
} = {}) {
  const dir = mkdtempSync(join(tmpdir(), 'outlook-test-'));
  const file = join(dir, 'auth.json');
  const exp = Math.floor(Date.now() / 1000) + secondsFromNow;
  const jwt = makeFakeJwt({ expSeconds: exp, claims });
  mkdirSync(dir, { recursive: true });
  writeFileSync(
    file,
    JSON.stringify({
      headers: {
        authorization: `Bearer ${jwt}`,
        'x-anchormailbox': 'PUID:test@test',
        ...headers,
      },
      expiresAt: exp * 1000,
      savedAt: Date.now(),
    }),
  );
  return file;
}

/**
 * Write a captured-headers JSON file for the OUTLOOK_CAPTURE_FIXTURE seam —
 * what capture.mjs would have produced from a real browser run. Lets tests
 * exercise the automatic 401-recapture path without Playwright.
 */
export function seedCaptureFixture({ secondsFromNow = 3600, claims = {} } = {}) {
  const dir = mkdtempSync(join(tmpdir(), 'outlook-fixture-'));
  const file = join(dir, 'captured-headers.json');
  const exp = Math.floor(Date.now() / 1000) + secondsFromNow;
  writeFileSync(
    file,
    JSON.stringify({
      authorization: `Bearer ${makeFakeJwt({ expSeconds: exp, claims })}`,
      'x-anchormailbox': 'PUID:fixture@test',
    }),
  );
  return file;
}

/** Return a unique tmp path for an isolated learnings file. */
export function freshLearningsPath() {
  return join(mkdtempSync(join(tmpdir(), 'outlook-learn-')), 'learnings.md');
}

/**
 * Run an HTTP server on a random port. `handler` receives (req, res) and is
 * responsible for the entire response. Returns `{ url, close }`.
 */
export function startMockServer(handler) {
  return new Promise((resolveStart) => {
    const server = createServer(handler);
    server.listen(0, '127.0.0.1', () => {
      const { port } = server.address();
      resolveStart({
        url: `http://127.0.0.1:${port}`,
        close: () =>
          new Promise((r) => {
            server.closeAllConnections?.();
            server.close(() => r());
          }),
      });
    });
  });
}

/**
 * Spawn the CLI as a subprocess. Returns `{ code, stdout, stderr }`.
 * Pass `env` to add/override environment variables for that run.
 *
 * OUTLOOK_NO_CAPTURE is set by default so no test can ever fall through to
 * a real Playwright/Chromium launch; override it (or set
 * OUTLOOK_CAPTURE_FIXTURE) in tests that exercise the recapture path.
 */
export function runCli(args, { env = {}, stdin = null, timeoutMs = 15_000 } = {}) {
  return new Promise((resolveRun, reject) => {
    const child = spawn(process.execPath, [CLI_PATH, ...args], {
      env: { OUTLOOK_NO_CAPTURE: '1', ...process.env, ...env },
      stdio: ['pipe', 'pipe', 'pipe'],
    });
    let stdout = '';
    let stderr = '';
    child.stdout.on('data', (d) => (stdout += d));
    child.stderr.on('data', (d) => (stderr += d));

    const timer = setTimeout(() => {
      child.kill('SIGKILL');
      reject(new Error(`CLI run timed out after ${timeoutMs}ms`));
    }, timeoutMs);

    child.on('close', (code) => {
      clearTimeout(timer);
      resolveRun({ code, stdout, stderr });
    });
    child.on('error', (err) => {
      clearTimeout(timer);
      reject(err);
    });

    if (stdin !== null) {
      child.stdin.write(stdin);
    }
    child.stdin.end();
  });
}
