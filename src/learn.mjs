// Persistent learnings the AI agent accumulates about the user's mail
// habits, preferences, and context. The agent reads the whole file at
// session start (via `outlook context`) and appends observations via
// `outlook learn add "<note>"`.
//
// The file is plain-text markdown so the user can edit it directly. Each
// learning is one line, prefixed with "- " and a YYYY-MM-DD date so we
// have a rough sense of when it was recorded.

import {
  readFileSync,
  writeFileSync,
  mkdirSync,
  existsSync,
  unlinkSync,
} from 'node:fs';
import { dirname, resolve } from 'node:path';
import { dataDir } from './paths.mjs';

const HEADER =
  '# outlook-cli learnings\n' +
  '\n' +
  '<!-- Maintained by the outlook CLI. The AI agent reads this file at\n' +
  '     session start and may append observations. Edit freely — anything\n' +
  '     here will be loaded next time. -->\n' +
  '\n';

export function learningsFile() {
  return (
    process.env.OUTLOOK_LEARNINGS ?? resolve(dataDir(), 'learnings.md')
  );
}

/** @returns {string[]} observation lines, oldest first. */
export function loadLearnings() {
  const file = learningsFile();
  if (!existsSync(file)) return [];
  return readFileSync(file, 'utf8')
    .split('\n')
    .filter((line) => line.startsWith('- '))
    .map((line) => line.slice(2).trim())
    .filter(Boolean);
}

export function saveLearnings(items) {
  const file = learningsFile();
  mkdirSync(dirname(file), { recursive: true });
  const body = items.length ? items.map((i) => `- ${i}`).join('\n') + '\n' : '';
  writeFileSync(file, HEADER + body);
}

/**
 * Append a learning. Returns true if it was new, false if a near-duplicate
 * already exists (case-insensitive trailing match).
 */
export function addLearning(text) {
  const trimmed = text.trim();
  if (!trimmed) return false;
  const existing = loadLearnings();
  if (existing.some((e) => e.toLowerCase().endsWith(trimmed.toLowerCase()))) {
    return false;
  }
  const stamp = new Date().toISOString().slice(0, 10);
  existing.push(`${stamp} | ${trimmed}`);
  saveLearnings(existing);
  return true;
}

/** Remove every learning containing `substring` (case-insensitive). Returns the count removed. */
export function removeLearning(substring) {
  const q = substring.trim().toLowerCase();
  if (!q) return 0;
  const existing = loadLearnings();
  const remaining = existing.filter((e) => !e.toLowerCase().includes(q));
  saveLearnings(remaining);
  return existing.length - remaining.length;
}

export function clearLearnings() {
  const file = learningsFile();
  if (existsSync(file)) unlinkSync(file);
}
