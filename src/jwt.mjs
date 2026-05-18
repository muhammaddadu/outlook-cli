// Tiny JWT payload decoder. We never validate signatures — the token is
// already trusted (we captured it from OWA's own traffic, on the user's
// own machine, and we only ever send it back to Microsoft).

export function decodePayload(bearer) {
  if (!bearer) return null;
  const jwt = bearer.replace(/^Bearer\s+/i, '');
  const parts = jwt.split('.');
  if (parts.length < 2) return null;
  const b64 = parts[1].replace(/-/g, '+').replace(/_/g, '/');
  const padded = b64 + '='.repeat((4 - (b64.length % 4)) % 4);
  try {
    return JSON.parse(Buffer.from(padded, 'base64').toString('utf8'));
  } catch {
    return null;
  }
}
