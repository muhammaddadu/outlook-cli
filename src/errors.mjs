// Trackable, actionable errors + sysexits-style exit codes.
//
// Every AppError carries a stable `code` (E_AUTH, E_HTTP, …) that callers can
// grep for, plus a `hint` string telling the user what to do next.

export const E = Object.freeze({
  AUTH_REQUIRED: 'E_AUTH_REQUIRED',
  AUTH_BLOCKED: 'E_AUTH_BLOCKED',
  HTTP: 'E_HTTP',
  ARGS: 'E_ARGS',
  UNEXPECTED: 'E_UNEXPECTED',
});

// Loosely follows BSD sysexits.h. 0 success, 1 generic, 2 auth, 3 HTTP, 64 usage.
export const EXIT = Object.freeze({
  OK: 0,
  GENERAL: 1,
  AUTH: 2,
  HTTP: 3,
  USAGE: 64,
  SIGINT: 130,
});

export class AppError extends Error {
  /**
   * @param {object} opts
   * @param {string} opts.code - One of the E.* codes.
   * @param {string} opts.message - Short, factual description.
   * @param {string} [opts.hint] - What the user should do next.
   * @param {unknown} [opts.cause] - Optional underlying error.
   */
  constructor({ code, message, hint, cause }) {
    super(message);
    this.name = 'AppError';
    this.code = code;
    this.hint = hint;
    if (cause !== undefined) this.cause = cause;
  }
}

export function exitCodeFor(err) {
  if (err instanceof AppError) {
    switch (err.code) {
      case E.AUTH_REQUIRED:
      case E.AUTH_BLOCKED:
        return EXIT.AUTH;
      case E.HTTP:
        return EXIT.HTTP;
      case E.ARGS:
        return EXIT.USAGE;
      default:
        return EXIT.GENERAL;
    }
  }
  return EXIT.GENERAL;
}
