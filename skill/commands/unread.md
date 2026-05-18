---
description: List or summarise the user's unread Outlook mail.
argument-hint: [optional filters, e.g. "from boss" or "this week"]
---

Use the `outlook` CLI (skill: `outlook`) to triage unread mail.

Default: run `outlook unread -n 25` and produce a concise list (sender,
subject, age). Group by sender or topic if there's a clear pattern.

If `$ARGUMENTS` has extra context (`from <person>`, `since <time>`,
`important`, …), translate it into the matching `outlook unread` flags
before running.

After the summary, offer next actions: open one of the messages
(`outlook read <id>`), reply, or archive.
