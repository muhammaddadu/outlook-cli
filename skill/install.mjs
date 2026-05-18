#!/usr/bin/env node
// Install the `outlook` agent skill + slash commands for Claude Code /
// Codex / Cursor.
//
//   node skill/install.mjs                              # default: Claude user-scope
//   node skill/install.mjs --target codex
//   node skill/install.mjs --target claude --scope project
//   node skill/install.mjs --copy                       # static copy instead of symlink
//   node skill/install.mjs --skill-only                 # skip slash commands
//   node skill/install.mjs --commands-only              # skip skill
//
// Skills are loaded by the agent automatically when the user's intent
// matches the description in SKILL.md's frontmatter ("check my unread
// mail"). Slash commands are explicit user invocations like `/outlook` or
// `/unread`. We install both so the user has the choice.

import {
  mkdirSync,
  symlinkSync,
  unlinkSync,
  lstatSync,
  copyFileSync,
  readdirSync,
} from 'node:fs';
import { homedir } from 'node:os';
import { dirname, resolve, basename } from 'node:path';
import { fileURLToPath } from 'node:url';

const HERE = dirname(fileURLToPath(import.meta.url));
const SKILL_SOURCE = resolve(HERE, 'outlook', 'SKILL.md');
const COMMANDS_SOURCE = resolve(HERE, 'commands');

const args = process.argv.slice(2);
const flag = (name, defaultValue) => {
  const i = args.indexOf(`--${name}`);
  return i >= 0 ? args[i + 1] : defaultValue;
};
const target = flag('target', 'claude');
const scope = flag('scope', 'user');
const copy = args.includes('--copy');
const skillOnly = args.includes('--skill-only');
const commandsOnly = args.includes('--commands-only');

if (target === 'print') {
  console.log(SKILL_SOURCE);
  console.log(COMMANDS_SOURCE);
  process.exit(0);
}

function destinations(target, scope) {
  switch (target) {
    case 'claude': {
      const root =
        scope === 'project'
          ? resolve(process.cwd(), '.claude')
          : resolve(homedir(), '.claude');
      return {
        skill: resolve(root, 'skills', 'outlook', 'SKILL.md'),
        commandsDir: resolve(root, 'commands'),
      };
    }
    case 'codex':
      return {
        skill: resolve(homedir(), '.codex', 'skills', 'outlook', 'SKILL.md'),
        commandsDir: resolve(homedir(), '.codex', 'commands'),
      };
    case 'cursor':
      // Cursor has rules, not slash commands; we drop just the skill.
      return {
        skill: resolve(homedir(), '.cursor', 'rules', 'outlook.md'),
        commandsDir: null,
      };
    default:
      console.error(`Unknown target: ${target}`);
      console.error('Supported: claude, codex, cursor, print');
      process.exit(64);
  }
}

function place(src, dest) {
  mkdirSync(dirname(dest), { recursive: true });
  try {
    // lstat will succeed for both regular files and symlinks; we replace
    // either with the new install.
    lstatSync(dest);
    unlinkSync(dest);
  } catch {
    // dest doesn't exist; fine.
  }
  if (copy) {
    copyFileSync(src, dest);
    console.log(`Copied  → ${dest}`);
  } else {
    symlinkSync(src, dest);
    console.log(`Linked  → ${dest}`);
  }
}

const { skill, commandsDir } = destinations(target, scope);

if (!commandsOnly) {
  place(SKILL_SOURCE, skill);
}

if (!skillOnly && commandsDir) {
  for (const cmdFile of readdirSync(COMMANDS_SOURCE)) {
    if (!cmdFile.endsWith('.md')) continue;
    place(resolve(COMMANDS_SOURCE, cmdFile), resolve(commandsDir, cmdFile));
  }
} else if (!skillOnly && !commandsDir) {
  console.log(`(${target} doesn't support slash commands — skill only.)`);
}

console.log('');
console.log('Done. If your agent is already running, restart it so the new');
console.log('skill and commands get picked up.');
