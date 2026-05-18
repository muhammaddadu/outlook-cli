---
description: Show a quick summary of the user's Outlook inbox.
argument-hint: [optional filters, e.g. "from Alice this week"]
---

Use the `outlook` CLI (skill: `outlook`) to fetch the user's recent inbox
and summarise it.

If `$ARGUMENTS` is empty:
1. Run `outlook list -n 10` to get the 10 most recent messages.
2. Summarise sender + subject in a short list.

If `$ARGUMENTS` has filter context (sender name, time window, etc.):
1. Translate the natural-language intent into `outlook list` flags
   (`--from`, `--since`, `--has-attachments`, `--folder`, …).
2. Run the command, summarise, and offer follow-up actions (read full
   body, reply, archive — anything achievable via the CLI).

Always show **just** sender + subject + age in the summary. Don't dump
full message bodies unless the user asks.
