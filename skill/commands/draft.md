---
description: Draft a reply (or new message) in Outlook for the user to review and send.
argument-hint: <natural-language-instruction, e.g. "reply to Alice's deploy email saying confirmed">
---

The user invoked `/draft` with: $ARGUMENTS

Use the `outlook` CLI (skill: `outlook`) to create a draft the user can
review in their Outlook Drafts folder before sending.

Workflow:
1. If `$ARGUMENTS` references a specific email (sender, subject, keyword),
   run `outlook search` or `outlook list` to find its `Id`.
2. Read the original with `outlook read <id>` to get context.
3. Compose the reply body based on the user's instruction.
4. Create the draft with `outlook draft-reply <id> '{"Body": {"ContentType": "Text", "Content": "<body>"}}'`.
5. Return the `DraftId` and `WebLink` to the user. Tell them the draft is
   in their Drafts folder ready for review.

If the user wants a brand-new message (not a reply), use `outlook draft`
with the full Message JSON.

**Do not send mail directly via `outlook send` from this command — the
whole point of `/draft` is the review step in Outlook.**
