#!/usr/bin/env node
// Install the `outlook` agent skill (and slash commands) for every AI
// agent CLI we can detect on this machine.
//
// By default we install to every supported agent that has a config
// directory under $HOME *or* a binary on $PATH:
//
//   Claude Code  →  ~/.claude/skills/outlook/SKILL.md
//                   ~/.claude/commands/{outlook,inbox,unread}.md
//   Codex CLI    →  ~/.codex/skills/outlook/SKILL.md
//                   ~/.codex/prompts/{outlook,inbox,unread}.md   (if `~/.codex/prompts` exists or codex on PATH)
//   Cursor       →  ~/.cursor/rules/outlook.md                    (rules only — Cursor has no slash commands)
//
//   node skill/install.mjs                    # default: every detected agent
//   node skill/install.mjs --target claude    # one specific agent
//   node skill/install.mjs --target claude,codex
//   node skill/install.mjs --target all       # force install to all known agents
//   node skill/install.mjs --copy             # static copy instead of symlink
//   node skill/install.mjs --skill-only       # skip slash commands
//   node skill/install.mjs --commands-only    # skip the skill itself
//   node skill/install.mjs --print            # just print source paths
//   node skill/install.mjs --uninstall        # remove everything we installed

import {
  mkdirSync,
  symlinkSync,
  unlinkSync,
  lstatSync,
  existsSync,
  copyFileSync,
  readdirSync,
} from 'node:fs';
import { execSync } from 'node:child_process';
import { homedir } from 'node:os';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const HERE = dirname(fileURLToPath(import.meta.url));
const SKILL_SOURCE = resolve(HERE, 'outlook', 'SKILL.md');
const COMMANDS_SOURCE = resolve(HERE, 'commands');

// ---------------------------------------------------------------------------
// CLI args

const args = process.argv.slice(2);
const flag = (name, fallback) => {
  const i = args.indexOf(`--${name}`);
  return i >= 0 ? args[i + 1] : fallback;
};
const has = (name) => args.includes(`--${name}`);

const copy = has('copy');
const skillOnly = has('skill-only');
const commandsOnly = has('commands-only');
const uninstall = has('uninstall');
const printOnly = has('print');
const target = flag('target', 'auto');

if (printOnly) {
  console.log(`skill   : ${SKILL_SOURCE}`);
  console.log(`commands: ${COMMANDS_SOURCE}`);
  process.exit(0);
}

// ---------------------------------------------------------------------------
// Agent registry

/**
 * Each agent declares:
 *   id      - canonical key for --target
 *   label   - human display name
 *   probe() - return true if this agent is "installed" (has config dir or binary)
 *   layout() - return { skill, commandsDir | null } for where to drop files
 */
const AGENTS = [
  {
    id: 'claude',
    label: 'Claude Code',
    probe: () => dirExists('.claude') || onPath('claude'),
    layout: () => ({
      skill: resolve(homedir(), '.claude', 'skills', 'outlook', 'SKILL.md'),
      commandsDir: resolve(homedir(), '.claude', 'commands'),
    }),
  },
  {
    id: 'codex',
    label: 'Codex CLI',
    probe: () => dirExists('.codex') || onPath('codex'),
    layout: () => ({
      skill: resolve(homedir(), '.codex', 'skills', 'outlook', 'SKILL.md'),
      // Codex CLI uses `~/.codex/prompts/` for custom prompts (analogous to
      // Claude's slash commands).
      commandsDir: resolve(homedir(), '.codex', 'prompts'),
    }),
  },
  {
    id: 'cursor',
    label: 'Cursor',
    probe: () => dirExists('.cursor') || onPath('cursor'),
    layout: () => ({
      skill: resolve(homedir(), '.cursor', 'rules', 'outlook.md'),
      commandsDir: null, // Cursor has rules, not slash commands.
    }),
  },
];

function dirExists(rel) {
  try {
    return lstatSync(resolve(homedir(), rel)).isDirectory();
  } catch {
    return false;
  }
}

function onPath(bin) {
  try {
    execSync(`command -v ${bin}`, { stdio: 'ignore' });
    return true;
  } catch {
    return false;
  }
}

// ---------------------------------------------------------------------------
// Resolve which agents to install to

function selectAgents(targetSpec) {
  if (targetSpec === 'all') return AGENTS;
  if (targetSpec === 'auto') {
    const detected = AGENTS.filter((a) => a.probe());
    if (detected.length === 0) {
      console.error('No supported agent CLI detected.');
      console.error('Supported: claude, codex, cursor.');
      console.error('Use `--target all` to install everywhere anyway, or');
      console.error('`--target claude,codex` to pick specific agents.');
      process.exit(1);
    }
    return detected;
  }
  const wanted = targetSpec.split(',').map((s) => s.trim().toLowerCase());
  const selected = AGENTS.filter((a) => wanted.includes(a.id));
  const unknown = wanted.filter((w) => !AGENTS.some((a) => a.id === w));
  if (unknown.length) {
    console.error(`Unknown target(s): ${unknown.join(', ')}`);
    console.error(`Supported: ${AGENTS.map((a) => a.id).join(', ')}`);
    process.exit(64);
  }
  return selected;
}

// ---------------------------------------------------------------------------
// File ops

function place(src, dest) {
  mkdirSync(dirname(dest), { recursive: true });
  try {
    lstatSync(dest);
    unlinkSync(dest);
  } catch {
    /* dest didn't exist */
  }
  if (copy) {
    copyFileSync(src, dest);
  } else {
    symlinkSync(src, dest);
  }
}

function remove(dest) {
  try {
    lstatSync(dest);
    unlinkSync(dest);
    return true;
  } catch {
    return false;
  }
}

function installToAgent(agent) {
  const { skill, commandsDir } = agent.layout();
  const installed = [];

  if (!commandsOnly) {
    place(SKILL_SOURCE, skill);
    installed.push(skill);
  }

  if (!skillOnly && commandsDir) {
    for (const file of readdirSync(COMMANDS_SOURCE)) {
      if (!file.endsWith('.md')) continue;
      const dest = resolve(commandsDir, file);
      place(resolve(COMMANDS_SOURCE, file), dest);
      installed.push(dest);
    }
  }
  return installed;
}

function uninstallFromAgent(agent) {
  const { skill, commandsDir } = agent.layout();
  const removed = [];

  if (remove(skill)) removed.push(skill);

  if (commandsDir && existsSync(commandsDir)) {
    for (const file of readdirSync(COMMANDS_SOURCE)) {
      if (!file.endsWith('.md')) continue;
      const dest = resolve(commandsDir, file);
      if (remove(dest)) removed.push(dest);
    }
  }
  return removed;
}

// ---------------------------------------------------------------------------
// Main

const agents = selectAgents(target);

console.log(
  `${uninstall ? 'Uninstalling' : 'Installing'} outlook skill${
    skillOnly ? '' : commandsOnly ? ' (commands only)' : ' + slash commands'
  } for: ${agents.map((a) => a.label).join(', ')}\n`,
);

for (const agent of agents) {
  console.log(`▸ ${agent.label}`);
  const changed = uninstall ? uninstallFromAgent(agent) : installToAgent(agent);
  if (changed.length === 0) {
    console.log('    (nothing to do)');
  } else {
    for (const path of changed) {
      console.log(`    ${uninstall ? '✗' : copy ? '↧' : '→'} ${path}`);
    }
  }
  console.log('');
}

if (!uninstall) {
  console.log('Done. Restart your agent CLI(s) so the new skill is picked up.');
  if (!copy) {
    console.log('(Symlinked — edits to the source files take effect immediately.)');
  }
}
