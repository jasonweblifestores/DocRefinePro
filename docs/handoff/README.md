# DocRefine Pro — handoff pack

**Start here.** This folder exists because Jason Diaz (jason@weblifestores.com), who built DocRefine Pro
and ran every PDF rebranding delivery through it, left WebLife in October 2026. Everything he knew about
this system that was not already in the code is written down here.

If you are picking this up cold with Claude Code, read [00-offboarding.md](00-offboarding.md) first —
some of it will already be broken by the time you arrive, and that document says what and why.

## What this system is

A local Windows/macOS desktop app (Python + PySide6) that batch-processes documents: dedup, flatten, OCR,
export — and, since v136, **rebrands PDFs**. The rebrand feature is the part that matters commercially:
it takes a folder of vendor PDFs and wraps each page in WebLife retailer branding (cover, header/footer
strips, watermark, page stamps) without touching the original page content, then produces a delivery set.

Two large deliveries have shipped through it:

| Delivery | ClickUp | Files | Outcome |
|---|---|---:|---|
| Budget Mailboxes Batch 4 | [86ey7j4wx](https://app.clickup.com/t/86ey7j4wx) | 2,174 | Signed off + uploaded 2026-08-13 |
| MailboxWorks | [86eyaantb](https://app.clickup.com/t/86eyaantb) | 970 | Delivered 2026-08-21 |

Earlier pilots (Batch 1 and 2) were done by hand in Canva, before the app could do it. They matter because
they set precedents that conflict with the written SOP — see the playbook.

## Read in this order

| # | Document | Why you need it |
|---|---|---|
| 1 | [00-offboarding.md](00-offboarding.md) | **What breaks when Jason's accounts are disabled, and who has to fix it.** Includes the repo-ownership problem. |
| 2 | [01-clickup-map.md](01-clickup-map.md) | Every ClickUp task, brief, SOP and comment ID referenced anywhere in this pack, with links. |
| 3 | [02-sustainable-hearth.md](02-sustainable-hearth.md) | The one open job. If "European Home → Sustainable Hearth PDFs" comes back, this is the runbook. |
| 4 | [knowledge/project-setup.md](knowledge/project-setup.md) | How to get the app running and verified on a new machine. |
| 5 | [knowledge/rebrand-playbook.md](knowledge/rebrand-playbook.md) | **The big one (74KB).** Design decisions, every release v136–v158, both delivered batches, and every defect found by QA'ing real output. |
| 6 | [knowledge/release-protocol.md](knowledge/release-protocol.md) | The ordered release checklist. ClickUp updates itself — do not post manually. |
| 7 | [knowledge/vision-pass.md](knowledge/vision-pass.md) | Local vision classification: the two-pass design and the 8GB-VRAM measurements that constrain it. |
| 8 | [knowledge/known-issues.md](knowledge/known-issues.md) | Bundled-binary quirks, quarantine data model, template bundling. |
| 9 | [knowledge/clickup-api-quirks.md](knowledge/clickup-api-quirks.md) | Things that look like data loss and are not. Read before "fixing" a comment. |
| 10 | [knowledge/sustainable-hearth-findings.md](knowledge/sustainable-hearth-findings.md) | The measured scope behind document 3. |
| 11 | [knowledge/people-and-pronouns.md](knowledge/people-and-pronouns.md) | Pronouns colleagues have stated. Don't guess from names. |

## What else is in this repo because of the handoff

- **`brandkits/`** — `budget-mailboxes-brand.json` and `mailboxworks-brand.json`. **These are the single most
  irreplaceable files here.** They hold the canonical manufacturer names and 56 aliases that Kunchana
  Godahewa approved one at a time over two deliveries. The Batch 4 kit was deleted from disk once already
  (2026-08-14) and the artwork was re-downloadable from Drive but *these decisions were not* — they had to be
  reconstructed by hand. Do not let them live in one place again.
- **`verify/`** — 59 verification scripts (~440KB). The suite ran 563/563 across 23 scripts as of v157, plus
  the per-batch QA helpers (`deliver_qa.py`, `survivors.py`, `scan_empty_pages.py`, `pixel_truth.py`,
  `review_branded_titles.py`, …). These lived in `Documents\DocRefinePro_Data\verify\` — outside version
  control, on one laptop — until this handoff. `batch_paths.py` centralises per-batch paths via the
  `DRP_BATCH` environment variable.
- **`docs/handoff/deliveries/`** — the review sheets and QA records for both shipped batches: the Batch 4 and
  MBW rebrand plans (`.xlsx`), the vision title-review JSON, `MBW_no-attribution-line.csv`, and the delivery
  comment that was posted. This is the audit trail for what was shipped and why each file was branded or left.

## Three things worth knowing before you touch anything

1. **Every substantive defect in this project was found by QA'ing real output, never by the unit tests** —
   which were green throughout. A negative page box that silently blanked a file, 255 files that would have
   shipped as `installation-guide-...-255.pdf`, 108 wrong manufacturer attributions, a corrupt source we
   branded and shipped. Always inspect the actual deliverable after a run. The playbook repeats this warning
   several times because it kept being the lesson.
2. **Wording is data, not code.** Taglines, disclaimers, version labels, attribution format and manufacturer
   aliases all live in the kit's `brand.json`. A blank field prints nothing and the run log names the gap.
   **The loader looks for the literal filename `brand.json`** — anything else is ignored and the kit silently
   falls back to defaults naming *Budget Mailboxes*, which will put the wrong company on every cover while
   looking like it worked.
3. **The app never edits the vendor's own page content.** It extends pages and overlays our branding. Any
   request phrased as "change what the document says" — as the Sustainable Hearth job is — is a different
   operation and needs scoping before it is accepted. Document 3 explains why that distinction cost real
   analysis to establish.
