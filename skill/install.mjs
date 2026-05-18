#!/usr/bin/env node
// Install the `outlook` agent skill for Claude Code / Codex / Cursor / etc.
//
// Default: symlinks `skill/outlook/SKILL.md` into the user's Claude Code
// skills directory so any Claude Code session can load it.
//
//   node skill/install.mjs                          # default: Claude Code user-scope
//   node skill/install.mjs --target codex
//   node skill/install.mjs --target claude --scope project
//
// Supported targets:
//   claude   ~/.claude/skills/outlook/SKILL.md      (user scope)
//            <cwd>/.claude/skills/outlook/SKILL.md  (project scope, --scope project)
//   codex    ~/.codex/skills/outlook/SKILL.md       (best-effort; convention may vary)
//   cursor   ~/.cursor/rules/outlook.md             (Cursor rules file)
//   print    just print the source path             (for piping/inspection)

import { mkdirSync, symlinkSync, unlinkSync, existsSync, lstatSync, copyFileSync } from 'node:fs';
import { homedir } from 'node:os';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const HERE = dirname(fileURLToPath(import.meta.url));
const SKILL_SOURCE = resolve(HERE, 'outlook', 'SKILL.md');

const args = process.argv.slice(2);
const flag = (name, defaultValue) => {
  const i = args.indexOf(`--${name}`);
  return i >= 0 ? args[i + 1] : defaultValue;
};
const target = flag('target', 'claude');
const scope = flag('scope', 'user');
const copy = args.includes('--copy');

function destinationFor(target, scope) {
  switch (target) {
    case 'claude':
      return scope === 'project'
        ? resolve(process.cwd(), '.claude', 'skills', 'outlook', 'SKILL.md')
        : resolve(homedir(), '.claude', 'skills', 'outlook', 'SKILL.md');
    case 'codex':
      return resolve(homedir(), '.codex', 'skills', 'outlook', 'SKILL.md');
    case 'cursor':
      return resolve(homedir(), '.cursor', 'rules', 'outlook.md');
    case 'print':
      console.log(SKILL_SOURCE);
      process.exit(0);
    default:
      console.error(`Unknown target: ${target}`);
      console.error('Supported: claude, codex, cursor, print');
      process.exit(64);
  }
}

const dest = destinationFor(target, scope);
mkdirSync(dirname(dest), { recursive: true });

// If something already lives at the destination, replace it. Symlinks and
// real files are both common after multiple installs / dev iterations.
if (existsSync(dest) || (() => { try { return !!lstatSync(dest); } catch { return false; } })()) {
  unlinkSync(dest);
}

if (copy) {
  copyFileSync(SKILL_SOURCE, dest);
  console.log(`Copied skill → ${dest}`);
} else {
  symlinkSync(SKILL_SOURCE, dest);
  console.log(`Linked skill → ${dest}`);
  console.log(`             ↳ ${SKILL_SOURCE}`);
  console.log('Edits to the source file take effect immediately. Pass --copy for a static install.');
}
