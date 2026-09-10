# Sustainable Hearth findings — measured scope of the European Home rename

> **Provenance.** This began as one of Jason Diaz's local Claude Code memory files
> (`~/.claude/projects/.../memory/docrefine-sustainable-hearth-job.md`) and was moved into the repo on 2026-09-10 as part of
> his handoff. It is preserved close to verbatim — the numbers in here were measured, often
> expensively, and paraphrasing them would lose the point. Dates are absolute. Treat it as a
> record of what was true when written, not a guarantee about today's code: **verify any file,
> function or flag it names still exists before acting on it.**

---

Vendor **European Home is renaming to Sustainable Hearth**. ClickUp task **`86eypep2z`** (DP - Catalog Team → Issues Board, Amjad Ali owner). Site branding, copy, URLs and 301s were finished by Ramez + Kunchana around 2026-08-21/24. On **2026-08-24** Kunchana assigned the **PDF leg to Jason** (comment `90180250616508`, at Amjad's direction in `90180250371252`), framed as "the same rebrand treatment as your BM batches". **It has NOT come through the TA Intake Form yet** — this is pre-work. See [the rebrand playbook](rebrand-playbook.md).

**It is not the same job, and that is the thing to say first.** Our pipeline wraps an untouched original page in *retailer* branding (cover/strips/watermark/stamps). This asks to change the **vendor's own branding printed inside the page**. Different operation, different risk.

Scope pulled from the two delivered review sheets on 2026-08-28 — **13 documents**: 10 in Batch 4 (4 `rebrand` / 6 `leave`), 3 in MBW (all `rebrand`). List at scratchpad `european_home_rows.csv`.

What is actually cheap vs. what is not:
- **Cheap (ours):** 6 of the 7 branded files print our own stamp `Manufactured by European Home | Sold by [Budget Mailboxes|MailboxWorks]`. That is a canonical-name + alias edit in the brand kits (`European Home`, alias `EuropeanHome`) plus a re-run of those files. Note `european-home-mail-slot-installation-instructions.pdf` (MBW) is credited to **"The Mailbox Works"** — a misclassification off the retailer footer in the scan; fix in the same pass.
- **Not cheap (the vendor's own artwork), and effectively blocked three different ways:**
  1. **Subset fonts with scrambled encodings** — `Large/Small-Mail-Slot-Installation-Sheet.pdf` extract as `-PH#SJEHF3PBE` for "Log Bridge Road". No safe find/replace, and replacement glyphs may not exist in the subset.
  2. **Image-only pages** — all 3 MBW files are a single full-page raster, 0 text chars. The EH house-mark + wordmark is pixels. (MailboxWorks already swapped the *footer* to their own; the EH logo at the top was left.)
  3. **Text converted to outlines** — 5 Batch 4 files have 0 text and 0 images (`Centauri-`, `Galaxy-Large-`, `Galaxy-Small-Mail-Slot-Dimensions`, `mailsloteurohome`, and `Vega-` partly).
- **It is not just a name.** The docs carry a full identity block — name, address, phone, fax, `www.europeanhome.com` — across **three different addresses** (Middleton MA; 376 Washington St, Malden MA; 501-B Main Street, Saugus MA). Renaming alone ships a stale address and a dead domain.

**Recommendation: ask Sustainable Hearth for re-issued source PDFs.** They are mid-rebrand and will have re-cut collateral; that turns this back into a normal ingest + rebrand run instead of destructive editing.

Two scope facts worth carrying into the intake:
- The live **BM brand-page Download Library has 13 entries** (MageWorx ids 161–172 + 361), so matching on our `manufacturer` column **under-counts** — items like "Mailbox Stand Info Flyer", "Torgen", "View Point" are in the library but did not match. Get the id→filename mapping (Kunchana offered to pull the affected file list).
- **Two Download Library entry titles still read "European Home"** (ids 163, 168) and their thumbnail alt text too. That is site metadata, not the PDFs — belongs to Ramez/Kunchana, not this job.
- MBW's Sustainable Hearth group page carries no PDFs; MBW docs sit on product pages.
