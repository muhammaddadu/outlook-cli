---
description: Run an Outlook query or send mail via the local outlook CLI.
argument-hint: <natural-language-request, e.g. "show unread from Alice this week">
---

The user invoked `/outlook` with: $ARGUMENTS

Use the `outlook` CLI (skill: `outlook`) to fulfil the request. If
`$ARGUMENTS` is empty, ask the user what they want to do (read inbox,
search, summarise unread, send a reply, …).

Reminders:
- All output is JSON to stdout; pipe through `jq` for transformation.
- Prefer filter flags (`--unread`, `--from`, `--since`, `--folder`) over
  long KQL queries.
- Never send mail without confirming recipient + subject + body with the
  user in plain English first.
- If the CLI returns exit code 2 (E_AUTH_REQUIRED), stop and ask the user
  to run `outlook auth` manually — you can't run it yourself.
