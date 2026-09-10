# ClickUp map

Every task, brief, SOP and comment referenced anywhere in this pack. Task links work for anyone with WebLife
workspace access. **Comment IDs are given as plain IDs, not links** — fetch them with
`clickup_get_task_comments` (or `clickup_get_threaded_comments` for replies) rather than guessing a URL format.

## The rebranding deliveries, oldest first

| Task | What it is | State |
|---|---|---|
| [86evrpk8a](https://app.clickup.com/t/86evrpk8a) (parent [86evnee38](https://app.clickup.com/t/86evnee38)) | **Batch 1** — the hand-made Canva pilot, 19 PDFs. "Add Budget Mailboxes branding (co-branded or Budget Mailboxes only)"; delivered BM-only. | Complete |
| [86ewmngk3](https://app.clickup.com/t/86ewmngk3) | **Batch 2** — Kunchana → Jason, Shehara signed off. "178 product PDFs… according to the approved template"; 178 unique → **2,277** SKU copies. Jason hit a bug where the app would not redistribute rebranded files back to their original locations — *that* is why Batch 1/2 kept original filenames. | Complete |
| [86ey7j4wx](https://app.clickup.com/t/86ey7j4wx) | **Batch 4 / "BM PDF Rebranding"** — the first app-driven delivery. ~3,314 unique PDFs across 33 brands in the brief; shipped **2,174**. QA chain Catalog → Shehara → Adam. | Signed off + uploaded 2026-08-13 |
| [86eyaantb](https://app.clickup.com/t/86eyaantb) | **MailboxWorks** — 9,087 files deduplicated to **970 documents**. Delivered 2026-08-21. | Delivered |
| [86ey9rdk5](https://app.clickup.com/t/86ey9rdk5) | The **full MBW brief**, authored by Kunchana. Cites the SOP below as the current standard. | Reference |

### Key comments on those tasks

| Task | Comment ID | What it says |
|---|---|---|
| 86eyaantb | `90180246790629` | Asked Kunchana to decide the deliverable shape — 970 files vs 9,087 copies. |
| 86eyaantb | `90180247128578` | **All MBW decisions settled** (2026-08-18): 970 documents as "BM Batch 4" shape, keep original filenames, tagline OFF (MailboxWorks has none), no SKU exclusions. |
| 86eyaantb | `90180247866426` | The eight numbered asks that unblocked the MBW run. |
| 86eyaantb | `90180248874809` | The last open item before delivery — folding bare `IMI` into a canonical name. |
| 86eyaantb | `90180249203557` | **The MBW delivery comment.** 970 files, 470 branded / 500 left, 425 with attribution. |
| 86ey7j4wx | `90180249202161` | **Batch 4 disclosure** (2026-08-21) — the corrupt-source finding, posted after sign-off. Read this before trusting any Batch 4 file blindly. |

## The open job — Sustainable Hearth

**[86eypep2z](https://app.clickup.com/t/86eypep2z)** — "Update vendor name from European home to sustainable
hearth in system". List `901801595963` (Issues Board), folder `90180996821` (DP - Catalog Team), space
`90182513037`. Owner Amjad Ali, high priority.

Subtasks: [86eypgw4b](https://app.clickup.com/t/86eypgw4b) (coordinate) ·
[86eyquxar](https://app.clickup.com/t/86eyquxar) (BM/BPITU, done/review) ·
[86eyquxc9](https://app.clickup.com/t/86eyquxc9) (MBW, revision)

The thread in order — this is the whole history of how the PDF work landed on Jason:

| Comment ID | Who | What |
|---|---|---|
| `90180248466694` | Amjad Ali | Opening post, with screenshots. |
| `90180249509756` | Ramez Sedra | SEO done both sites — brand page URL changed, content mentions updated, 301s in place. Asks to be told the moment any product URL changes so he can verify Yoast's automatic redirect rather than assume it. |
| `90180250019667` | Kunchana Godahewa | Brand image live on both sites. Raises the two leftovers: **the Download Library PDFs are still branded European Home inside**, and two copy mentions on MBW. Asks Amjad whether the PDFs are a Catalog job or Jason's. |
| ↳ `90180250371252` | Amjad Ali | "Please assign the PDF work to Jason." |
| ↳ `90180250039658` | Ramez Sedra | Copy leftovers updated. |
| `90180250616508` | Kunchana Godahewa | **Assigns the PDF rebranding to Jason**, described as "the same rebrand treatment as your BM batches". Offers to pull the affected file list. |
| ↳ `90180251107035` | Jason Diaz | "Please submit through the TA Task Request form." |
| ↳ `90180251134504` | Jason Diaz | **The scoping reply** — seven things the TA request must specify, and why this is not the same job. Full text at [deliveries/SH_scoping_comment.md](deliveries/SH_scoping_comment.md). |
| `90180251105272` | Laleesha Wijeratne | Asks Amjad what is blocking the task from being complete. |

## Project and process references

| Reference | What it is |
|---|---|
| [86ex00r23](https://app.clickup.com/t/86ex00r23) | **The DocRefine Pro project task** — list `901815817512` ("Project Managment"), folder "TA - Operations". Every release adds a subtask here **automatically** via CI. Do not post manually; you will duplicate it. Check *subtasks*, not comments — it posts no comments and never will. |
| `86eyhpkn5`, `86eyhpp2j` | Examples of those auto-created release subtasks (v147, v148). |
| `8cr937k-390278` | Kunchana's **Canva rebranding guide** (ClickUp doc, Nov 2026). **Superseded** — this is the pilot process Batch 1/2 followed. Do not treat it as current. |
| `team/kunchana/designer-sidekick/knowledge-sources/vp_proc_sitecontent_rebranding-asset-management.md` | **The current SOP**, cited by the MBW brief. This is the standard: attribution line, tagline, `Version 1.0` + `Last Updated`, disclaimer, and the naming pattern are all real requirements. Batch 1/2 are the outliers, not the rule. One known mismatch: the SOP puts manufacturer attribution in a **small footer on every page**, which is where v147 moved it. |

## Who is who

ClickUp user IDs, useful when posting comments with real @mentions (the plain MCP comment tool cannot make a
mention that notifies — see [knowledge/clickup-api-quirks.md](knowledge/clickup-api-quirks.md)).

| Person | User ID | Role in this work |
|---|---|---|
| Kunchana Godahewa | `95574619` | Design lead on the rebrand briefs. Every manufacturer-name and template decision went through him. He/him. |
| Amjad Ali | `89486905` | Owner of the Sustainable Hearth task; DP Catalog. |
| Ramez Sedra | `101449743` | SEO — owns URLs, 301s and site copy. Tell him before any filename or URL change. |
| Laleesha Wijeratne | `95462717` | Tracking / delivery oversight. |
| Shehara Meadows | `95453086` | QA sign-off, between Catalog and Adam. |
| Adam | — | Final sign-off. Has a stated communication preference — there is a `writing-for-adam` skill for drafting to him. |
| Sean | — | TA lead on the MBW brief. |
| Nouman Khan | `89486906` | Subtask assignee (BM/BPITU site updates). |
| Hammad Rafique | `89486907` | Subtask assignee (MBW site updates). |
| Jason Diaz | `101512919` | Built the app, ran both deliveries. Left October 2026. |

**Name lookups:** `clickup_find_member_by_name` needs the **full** name — "Kunchana" returns null, "Kunchana
Godahewa" returns `95574619`.
