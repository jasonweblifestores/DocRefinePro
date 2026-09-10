# ClickUp API quirks — things that look broken but are not

> **Provenance.** This began as one of Jason Diaz's local Claude Code memory files
> (`~/.claude/projects/.../memory/clickup-api-quirks.md`) and was moved into the repo on 2026-09-10 as part of
> his handoff. It is preserved close to verbatim — the numbers in here were measured, often
> expensively, and paraphrasing them would lose the point. Dates are absolute. Treat it as a
> record of what was true when written, not a guarantee about today's code: **verify any file,
> function or flag it names still exists before acting on it.**

---

**Markdown tables in ClickUp comments DO render in the UI. The API's `comment_text` shows the
literal string `undefined` in place of each table block.** Confirmed by Jason on 2026-08-18 after I
posted an 8-item comment with six tables to `86eyaantb` and read it back: every table came back as
`undefined`, and I flagged it as possibly broken. It was fine on screen — the plain-text
serialization simply has no representation for a table block.

**Why:** it reads exactly like data loss, and the wrong reaction is to rewrite the comment as
plain-text lists or repost it — which double-notifies the assignee and makes the message worse.

**How to apply:** after posting a table-heavy comment, do NOT treat `undefined` in
`clickup_get_task_comments` as evidence of a problem. It is expected. If you genuinely need to
confirm rendering, ask the user to glance at it rather than poking their authenticated session.
An older comment of ours on the same task shows the same `undefined`, so it is long-standing.

**Mentions placed inline mid-body DO render inline. The API returns them concatenated at the END of
`comment_text`, with blanks left where they belonged.** Confirmed by Jason on 2026-08-31 on
`86eypep2z` reply `90180251134504`: three `type: "tag"` objects set mid-body read back as
`@Kunchana Godahewa@Ramez Sedra@Ramez Sedra@Laleesha Wijeratne` appended after the sign-off, with
gaps at their real positions (" — following my note above"). On screen they were inline and correct.

**Why:** it reads as a mangled comment, and the tempting fix — delete and repost with the mentions
in a leading block — double-notifies three people to repair a defect that does not exist.

**How to apply:** inlining mentions mid-body is fine. The `clickup-rich-comment` skill's "mentions at
the start" example is one working pattern, NOT a requirement. Never judge mention placement from
`comment_text`; ask the user to glance instead.

**The two comment endpoints serialize differently — do not port assumptions between them.** The
threaded-reply endpoint `POST /api/v2/comment/{comment_id}/reply` returned markdown tables as
literal markdown, where the task-comment endpoint returns `undefined` for the same thing. Threaded
replies are not covered by `clickup-rich-comment` or by `weblife-clickup-ops` (which has
`comment_post_with_mentions` but no reply equivalent); the same structured `comment` array works
against the reply URL.

**Attachments:** `clickup_attach_task_file` takes base64 only up to ~200KB. For anything larger,
`clickup_request_attachment_upload` returns a short-lived ticket — POST `multipart/form-data` with
the file in a field named `attachment` and the ticket in an `X-Upload-Ticket` header. Build the
request in Python rather than shelling out, since paths can contain shell metacharacters. A
345KB PNG went through this way fine. Attachments land on the TASK, not the comment, so reference
them by filename in the comment text.

**Assigning a comment** (`assignee` = user id) puts it in that person's inbox rather than relying on
them reading the task — worth doing when the comment is a request for decisions. Look ids up with
`clickup_find_member_by_name`, which needs the FULL name: "Kunchana" returned null, "Kunchana
Godahewa" returned id `95574619`.

See [the release protocol](release-protocol.md) for the separate rule that release subtasks on `86ex00r23` are
posted automatically by CI and must not be duplicated by hand.
