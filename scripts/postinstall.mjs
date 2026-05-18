#!/usr/bin/env node
// Friendly post-install hint.
//
// We deliberately do NOT auto-download the Chromium binary (it's ~150 MB
// and many install environments are bandwidth-constrained or non-interactive).
// Instead, point the user at `outlook setup`.
//
// CI / scripted installs can skip the hint entirely:
//   OUTLOOK_QUIET_POSTINSTALL=1 npm i -g github:muhammaddadu/outlook-cli

if (process.env.OUTLOOK_QUIET_POSTINSTALL || process.env.CI) process.exit(0);

const lines = [
  '',
  '  outlook-cli installed.',
  '',
  '  Next steps:',
  '    1. outlook setup           # download Chromium + (optionally) install the agent skill',
  '    2. outlook auth            # one-time interactive sign-in with MFA',
  '    3. outlook list -n 5       # smoke test',
  '',
  '  Docs:  https://github.com/muhammaddadu/outlook-cli',
  '',
];

process.stderr.write(lines.join('\n'));
