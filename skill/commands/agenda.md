---
description: Show the user's calendar agenda for a given time window.
argument-hint: [optional window, e.g. "today" / "this week" / "next 14 days"]
---

Use the `outlook` CLI (skill: `outlook`) to surface the user's calendar.

If `$ARGUMENTS` is empty, run `outlook agenda` (next 7 days) and produce a
concise list grouped by day:

```
Wed 2026-05-21
  09:30  Standup (Teams)
  14:00  1:1 with Alice (Room 4)
  16:00  Focus block — busy, no attendees
Thu 2026-05-22
  10:00  Project Beta sync (4 attendees)
  …
```

If `$ARGUMENTS` includes natural-language time hints, translate them into
`outlook agenda` flags:

| User says | Flags |
| --- | --- |
| "today" | `--from today --to tomorrow` |
| "tomorrow" | `--from tomorrow --to "+2d"` |
| "this week" | `--from today --days 7` |
| "next 14 days" / "next two weeks" | `--days 14` |
| "May 25th" | `--from 2026-05-25 --to 2026-05-26` |

After the listing, offer next actions: read full detail
(`outlook event-read <id>`), check who's busy at a slot
(`outlook free-busy`), or RSVP (`outlook accept/decline/tentative <id>`).

**Do NOT create or cancel events from this command.** Use `/outlook` or
ask the user explicitly.
