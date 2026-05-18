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
export function seedTokenCache({ secondsFromNow = 3600, headers = {} } = {}) {
  const dir = mkdtempSync(join(tmpdir(), 'outlook-test-'));
  const file = join(dir, 'auth.json');
  const exp = Math.floor(Date.now() / 1000) + secondsFromNow;
  const jwt = makeFakeJwt({ expSeconds: exp });
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
 */
export function runCli(args, { env = {}, stdin = null, timeoutMs = 15_000 } = {}) {
  return new Promise((resolveRun, reject) => {
    const child = spawn(process.execPath, [CLI_PATH, ...args], {
      env: { ...process.env, ...env },
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
