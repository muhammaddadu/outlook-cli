// stdout-vs-stderr discipline:
//   - All command results (JSON) go to stdout — pipeable, jq-friendly.
//   - Diagnostics, progress, and errors go to stderr — never pollute stdout.
//
// Debug output is gated on `--debug` (sets OUTLOOK_DEBUG) or DEBUG env var.

const isTty = process.stdout.isTTY;

/** Print a JSON value to stdout. Pretty-prints when running in a terminal. */
export function printJson(value, { pretty = isTty } = {}) {
  process.stdout.write(JSON.stringify(value, null, pretty ? 2 : 0) + '\n');
}

/** Diagnostic to stderr, gated on debug mode. */
export function debug(...args) {
  if (process.env.OUTLOOK_DEBUG || process.env.DEBUG) {
    process.stderr.write(`[debug] ${args.map(String).join(' ')}\n`);
  }
}

/** Informational message to stderr. */
export function info(msg) {
  process.stderr.write(`${msg}\n`);
}

/** Error block to stderr, with optional actionable hint. */
export function errorBlock(code, message, hint) {
  process.stderr.write(`Error (${code}): ${message}\n`);
  if (hint) process.stderr.write(`Hint: ${hint}\n`);
}
