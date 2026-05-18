<!--
Thanks for the PR! Fill in the sections below, then run through the
checklist. If anything is unclear, see CONTRIBUTING.md.
-->

## What does this change?

<!-- 1–2 sentences. What user-visible behaviour changes? -->

## Why this way?

<!-- Brief context. Reference an issue, a PRD requirement (MAIL-N /
CAL-N / AGT-N / OPS-N), or an open question (PRD §10) when relevant. -->

## Observable changes

<!-- New flags? New exit codes? New files? Anything a user or AI agent
will notice. Skip if internal-only. -->

## Verification

<!-- How did you test this beyond `npm test`? Live mailbox smoke?
Specific edge case? -->

---

### PR checklist

- [ ] `npm test` is green (all tests pass).
- [ ] `./src/cli.mjs --help` and `./src/cli.mjs --version` still work.
- [ ] Live smoke test (`./src/cli.mjs list -n 3`, or the command you
      touched) returns the expected output.
- [ ] Any new error path throws an `AppError` with a stable `code` and
      an actionable `hint`.
- [ ] New stdout writes only contain command output; diagnostics go to
      stderr.
- [ ] Any new subcommand has an integration test in `test/cli.test.mjs`.
- [ ] CLI surface changes are mirrored in `skill/outlook/SKILL.md`
      (and `skill/commands/*.md` if a slash command applies).
- [ ] No PII (real names / emails / tenant IDs / project names) in any
      diff. Verified via `git diff --cached | grep` against your
      personal terms.
- [ ] `.gitignore` covers any new file types introduced.
- [ ] If you tried an alternative approach and rejected it, captured the
      result in `LEARNINGS.md`.
