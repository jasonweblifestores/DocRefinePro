# The rebrand playbook — design decisions, every version v136-v158, and the two delivered batches

> **Provenance.** This began as one of Jason Diaz's local Claude Code memory files
> (`~/.claude/projects/.../memory/docrefine-rebrand-feature.md`) and was moved into the repo on 2026-09-10 as part of
> his handoff. It is preserved close to verbatim — the numbers in here were measured, often
> expensively, and paraphrasing them would lose the point. Dates are absolute. Treat it as a
> record of what was true when written, not a guarantee about today's code: **verify any file,
> function or flag it names still exists before acting on it.**

---

Planned feature: automated document rebranding for Budget Mailboxes (brand: navy + orange/blue floral mark). Replaces a manual per-document Canva PDF-editor workflow. Newest batch to rebrand: 3,000+ PDFs. See [project setup](project-setup.md).

Brand assets + samples live in `…\PROJECTS\DocRefine Pro\BrandKit\Sample Files\` — `Portrait\` and `Landscaape\` (their spelling) subfolders each with Cover, Back Cover, Header, Footer, Watermark PNGs (RGBA). Portrait assets sized for Letter@300dpi (2550px wide); landscape for ~A4@300dpi (3508px). Samples: `1570_FCBU_Installation_Instructions.pdf` (16pp Letter portrait, has text layer) and `1570-4T5-BM.pdf` (1pp 44x34in landscape drawing, has text layer). An image-domain POC lives at scratchpad `rebrand_poc.py`; previews in `BrandKit\Previews\`. User's initial verdict on the look: "looks great."

Design decisions (confirmed with user):
- **Drop flatten as a required step.** It was only needed because Canva was destructive. Do rebranding in the VECTOR domain (pypdf: build a taller page, place original page as-is, overlay strip/watermark/cover images). This preserves original quality AND the text layer.
- Because text is preserved, **OCR also becomes optional** — only needed for scanned/image-only PDFs with no text. Most docs are searchable for free.
- **Composable pipeline**: blocks (Ingest/Dedupe · Flatten · Rebrand · OCR · Export) the user arranges into a recipe, each folder-in/folder-out, with sensible guardrails (not fully arbitrary). Rebrand job = point at folder → review titles → Rebrand → export.
- **Header/footer = page extension** (strips added above/below; original never overlapped/cropped). Watermark over content.
- **Preserve original page dimensions.** Covers keep their own native size; pages keep theirs (+ strips). Neither is forced to match the other (POC's cover-to-width scaling will be removed).
- **Cover title**: rendered text, Poppins Bold (calibrate to the screenshot in `BrandKit`; title shows "Assembly Instructions" white, centered, ~11% from cover top). Poppins is OFL/free to bundle. Title source = **auto-suggest, user approves** via a spreadsheet round-trip (`titles.csv`).
- **Title engine = local Ollama** (user has it installed, v0.30.10, server on :11434). Keeps it 100% local — no API/key/cost. NO model pulled yet; recommended `llama3.2:3b`. Plan: extract page-1 text → local model → concise title + generic fallback.
- **Cover sizing (revised)**: cover PAGE = document page size; cover ART scaled to FIT preserving aspect (no stretch), centered, leftover margin filled with cover's own bg color. (User's native-size idea backfired: A4 cover looked tiny next to a 44x34in doc.)

Status: VECTOR proof validated on both samples (scratchpad `rebrand_vector.py`, previews in `BrandKit\Previews\VECTOR_*`). Confirmed: original text stays searchable (extracted 98/983 chars from branded pages), quality/native sizes preserved, title renders, landscape cover-fit works. Uses reportlab (overlays/covers) + pypdf (compose). reportlab added to venv.

DONE in prototype (all hard parts proven):
- **Poppins-Bold.ttf** sourced from Google Fonts → `BrandKit\Fonts\Poppins-Bold.ttf` (user okayed). Title now real Poppins.
- **Title sizing** made proportional to cover height (TITLE_FRAC≈0.027 of drawn cover ht) — legible on both Letter and the 44x34 landscape.
- **Local Ollama titling works**: `llama3.2:3b` pulled; rebrand calls http://localhost:11434/api/generate with page-1 text → content-based title + "Product Documentation" fallback. Demo titles: manual→"Installing 1570 F Series Mailboxes", drawing→"Cluster Box Unit Product Documentation". First call ~50s (model load), warm calls a few sec.

Open items / build tasks (prototype phase complete; ready to scope the real in-app build):
- **File size/speed**: still 39MB & slow for 16pp — overlay images re-embedded per page. FIX in build (share image XObjects once / downsample). User confirmed: fix as part of build.
- Build the in-app **"Rebrand" stage** + composable pipeline + brand-kit picker + title-review spreadsheet (`titles.csv`) + Ollama integration + bundle Poppins.
- User idea (long-term): a **font selector/manager** in the app.
- Handle /Rotate.

Brief (ClickUp task 86ey7j4wx "BM PDF Rebranding", ~3,314 unique PDFs / 33 brands; source + template in Google Drive; QA: Catalog→Shehara→Adam sign-off):
- **Classify, don't rebrand-all**: REBRAND installation guides, spec sheets, manuals. LEAVE AS-IS: CAD/technical drawings, UL/manufacturer certifications, non-PDF. The local LLM classifies rebrand-vs-skip + extracts product / asset-type / manufacturer / title → all human-approved in one review spreadsheet. (The landscape sample drawing = a "leave as-is" case.)
- Stamp on rebranded files: BM navy logo header/footer, URL, phone, "Trusted by the Nation", manufacturer attribution "Manufactured by [X] | Sold by Budget Mailboxes" (X per-doc), Version 1.0 + Last Updated [month year], disclaimer. Navy/black on light bg. PRESERVE technical content, safety warnings, compliance/cert marks.
- Hard limits: **under 50 MB/file**, embed fonts, PDF metadata Author=Budget Mailboxes.
- Output: `product-asset-type-budget-mailboxes.pdf` (lowercase, hyphens, ≤60 chars) into a `_rebranded` tree mirroring source; unique files rebranded ONCE, mapped to SKUs via `_manifest.csv`.

Decisions settled: mixed-orientation → dominant orientation; brand kit → swappable folder per batch (BM is the last set on this kit); technical drawings → classified out (skipped).

Plan: published artifact "Rebranding — Build Plan" (v2, aligned to brief) at claude.ai/code/artifact/62672e5d-5d50-462b-9429-ec942ff82989. Target release v136. 5 phases: 0 engine+<50MB, 1 one-click rebrand MVP, 2 classify+tag review, 3 flexible pipeline, 4 scale/ship.

PHASE 0 DONE (verified, not yet committed): built `docrefine/rebrand.py` (BrandKit + rebrand_pdf; vector compose, header/footer page-extension, watermark, fitted titled covers via reportlab+pypdf; branding images embedded ONCE and shared). Bundled `docrefine/assets/fonts/Poppins-Bold.ttf`; added reportlab to requirements.txt; DocRefinePro.spec bundles docrefine/assets + hiddenimports docrefine.rebrand/reportlab. Verify harness (scratchpad verify_phase0.py) = 12/12: portrait 39MB→15.3MB (<50MB), searchable text kept, metadata Author=Budget Mailboxes, fonts embedded, covers clean. Covers JPEG'd, caps COVER_MAX_PX=2200/WATERMARK_MAX_PX=1400. Note: ~2s/page (speed tuning deferred to Phase 4); residual size is the doc's own content images (preserved). Changes are in the working copy only — not committed/pushed yet.

PHASE 1 DONE (verified, not committed): "🎨 Rebrand a Folder" button (main_window) → RebrandDialog (dialogs.py: source + brand-kit folder pickers, remembers CFG.last_brand_kit) → app_qt.launch_rebrand → worker.run_rebrand. run_rebrand walks source for PDFs, mirrors into `<src>_rebranded` tree, names `output_filename()` = slug + `-budget-mailboxes.pdf` (≤60 chars), title `title_from_filename()`, smart-skips existing, single-threaded (parallelism deferred to Phase 4), progress via existing worker/monitor. Config field last_brand_kit added. Verify (scratchpad verify_phase1.py) 14/14: mirrored tree + naming + searchable + <50MB + author + skip-on-rerun + progress/DONE/notification; app dry-run boots; AppController+RebrandDialog construct offscreen. Still not committed to git.

PHASE 2 DONE (verified, not committed): `docrefine/classify.py` — local Ollama (llama3.2:3b, JSON-forced) classifies each PDF rebrand-vs-leave + product/asset_type/manufacturer/title/confidence, graceful fallback if no Ollama/no text. worker.run_rebrand_analyze → `_rebrand_plan.csv` (cols file,action,doc_type,product,asset_type,manufacturer,title,pages,confidence,source,notes); worker.run_rebrand_apply reads the (edited) CSV → brands "rebrand" rows (cover subtitle "Manufactured by [X] | Sold by Budget Mailboxes", filename via output_filename_from_fields = product-asset_type-budget-mailboxes.pdf) and copies "leave" rows byte-identical (original name). rebrand.py: added slugify, output_filename_from_fields, subtitle on cover. RebrandDialog now two-step (Analyze / Apply + CSV picker); app_qt routes by d.mode. spec hiddenimport docrefine.classify. Verify (scratchpad verify_phase2.py) 11/11: real Ollama classified manual→rebrand, drawing→leave; sheet round-trip; branded named from fields + manufacturer line on cover + searchable + <50MB; leave copied byte-identical; app boots; dialog wired.

FEATURE FUNCTIONALLY COMPLETE for the job (Analyze→review→Apply produces the deliverable).

Phases 0-2 committed + pushed (commit f697426 on main; NOT version-bumped/released).

PHASE 3 DONE (verified, not committed): composable pipeline. "⚙ Process a Folder" button → PipelineDialog (toggle Flatten/Rebrand/OCR + brand-kit picker) → worker.run_pipeline. Runs selected steps in fixed safe order Flatten→Rebrand→OCR (OCR always LAST), folder-to-folder via temp chain → `<src>_processed`. Rebrand step (worker._folder_rebrand) uses `_rebrand_plan.csv` from source if present else filename-based; Flatten/OCR via worker._folder_pdf_op (reuses PdfProcessor; OCR only on files with no text). worker._load_plan helper. Verify (scratchpad verify_phase3.py) 6/6: full Flatten→Rebrand→OCR chain restores searchability on a flattened file (OCR-last proven), rebrand-only preserves text, boots, dialog wired. Scope note: NOT full arbitrary reordering / Dedupe+Export as blocks (larger refactor, deferred).

Phase 3 committed (e2ff037).

PHASE 4 DONE + RELEASED: parallelism — worker._auto_workers (RAM-scaled: <8GB→1, <16GB→2, else 4) + worker._parallel_map (ThreadPoolExecutor). run_rebrand_apply and pipeline stages (_folder_pdf_op/_folder_rebrand) now fan out; shared state (rebrand._ensure_font + BrandKit.readers cache) warmed once before the pool to avoid races. /Rotate handled in rebrand_pdf (transfer_rotation_to_content); broken PDFs caught per-file; multiple kits = swappable folder. All phase tests re-passed after refactor (P1 14/14, P2 11/11, P3 6/6). Committed 0acbf40, tagged **v136** (pushed → CI building run 30525474719). config v136, changelog+README updated.

POST-BUILD AUDIT + v137 (2026-07-30): ran a fresh-eyes code audit + full regression (18/18, original engine untouched — rebranding was purely additive). Fixed: **F1** reportlab ImageReader thread-safety — BrandKit now caches encoded bytes (`specs()`); `rebrand_pdf` builds per-call readers via `live_readers()` (no sharing across threads). **F4** output filename collisions — `_assign_output_names` (apply) and a pre-pass in `_folder_rebrand` assign unique numbered names single-threaded before parallel dispatch (collision- AND resume-safe). **F12** no-text PDFs default to action=leave (safer for scanned certs). Removed dead `worker.run_rebrand`; trimmed imports; normalized path separators. Added **"Rebrand Unique Masters"** (Export tab Option D → app_qt.launch_rebrand_job → RebrandDialog default_source = job's `01_Master_Files`) so deduplicated jobs feed rebranding directly. Verified: collision 4/4, phase2 11/11, phase3 6/6, regression 18/18, boot+wiring OK. Released v137.
Deferred (noted, not blocking): pipeline temp uses ~4x corpus (F8); analyze step still single-threaded LLM per file (F10); PdfReader handles rely on GC (F11); prog_sub `_last_update` dict unsynchronized but GIL/ per-tid safe (F3).

v138 (2026-07-30) — Ollama detection fix: root cause of "could not detect Ollama" was the server being installed-but-not-running (an HTTP call doesn't start it; only `ollama` CLI commands do). classify.py now uses 127.0.0.1 (localhost→IPv6 ::1 caused false refusals on Windows) and adds find_ollama_exe/server_up/list_models/has_model/start_server (auto-starts `ollama serve`)/ollama_status/pull_model. worker.run_rebrand_analyze auto-starts the server + logs clear guidance; worker.run_pull_model downloads a model with progress. app_qt._start_analyze gates Analyze: auto-start, else a dialog offering to download Ollama (opens ollama.com/download) or the ~2GB model. Verified auto-start on a stopped server here. Model present: llama3.2:3b.

Behavior notes: rebrand does NOT create a job-list item (only dedup ingest emits JOB_DATA); output goes to `<source>_rebranded` (for "Rebrand Unique Masters" that's `<ws>/01_Master_Files_rebranded`).

v139 (2026-08-03) — RELEASED, both prior open items closed:
- **Complete set option** (was the open non-PDF question): `run_rebrand_apply(..., complete_set)` + `_copy_extras` mirrors every non-PDF source file into the output byte-identical (parallel, resume-safe, nested structure kept); un-analyzed PDFs are now logged instead of silently absent. Same flag on `run_pipeline`/`_folder_rebrand`/`_folder_pdf_op` (off = PDFs only). Checkbox in RebrandDialog + PipelineDialog, remembered as `CFG.rebrand_complete_set`, **default ON**.
- **Review sheet relocated**: new `docrefine/reviews.py`. Sheets now go to `Documents\DocRefinePro_Data\Rebrand Reviews\<parent>__<folder>_rebrand_plan.csv` (parent in the name because every dedup job's masters folder is `01_Master_Files`). Chosen over "beside the output" so Complete-set can't sweep the sheet into the upload tree. `find_plan()` order: Reviews folder → beside source → pre-v139 in-source (old sheets still work); `plan_csv` arg is now optional and the dialog pre-fills it. Re-analyze keeps the reviewed copy as `*.previous.csv`.
- Also dropped the `stats.json` timing file from rebrand/pipeline OUTPUT folders (nothing reads it there; it would have shipped in the deliverable) — elapsed time is logged. Workspace stages that the GUI does read are untouched.
- Verified: v139 27/27, regression 18/18, P1 14/14, P2 12/12, P3 6/6, collision 4/4, boot/wiring 12/12, dry-run OK. Commit d08167c, tag v139, CI run 30835329828 green (Win + Mac), release has both assets.
- NOTE: `verify_phase1.py` in scratchpad was stale since v137 (called the deleted `run_rebrand`) — retargeted to the pipeline rebrand-only path, same 14 assertions. Verify scripts now live in scratchpad `…\92046270-f2d5-4759-82bd-207e565559eb\scratchpad\` (verify_v139/boot/phase1-3/collision/regression); run as `python verify_X.py <repo> "<…>\BrandKit\Sample Files"`.

v140 (2026-08-04) — RELEASED (commit b860e50, tag v140, CI run 30849435813 green Win+Mac, both assets published).
- **Cover titles now match the hand-made Batch 1 set** (found at `G:\Shared drives\[VP] Venia Products KMS\Shared_Services\DP\DP_CONT\04_Operational\Operational - Task Related Docs & Sheets\BM Downloadable re-branding\Batch 1` — 19 manually rebranded PDFs). House style = UPPERCASE, 3-4 words, NO model numbers: "PRODUCT CARE & CLEANING", "FIVE YEAR PRODUCT WARRANTY". So `rebrand.title_for()` derives from asset_type via `ASSET_TYPE_TITLES`; LLM prompt forbids part numbers; `fit_title()` wraps to 2 lines + shrinks + char-truncates; `title_from_filename` keeps digit tokens (4C11D not 4C11d). NOTE: Batch 1 covers have NO manufacturer subtitle — ours adds "Manufactured by X | Sold by BM" per the Batch 4 brief. Open question for the user.
- **Excel review sheets**: `reviews.write_plan/read_plan` do .xlsx + .csv; xlsx has a "review?" triage column, PDF hyperlinks, action/asset_type dropdowns, freeze+autofilter. Apply reads either. openpyxl added to requirements + spec.
- **No-text instruction flip**: `classify.filename_suggests_instructions()` flips no-text PDFs whose filename says instructions to action=rebrand (55/1056 on the real batch, 0 false positives). Gotcha: NO `\b` before INS — filenames run it onto a part number (206550INS-1400.pdf).
- Rebrand audit fixes: atomic writes (rebrand_pdf + `_atomic_copy`); `BrandKit.has()` now requires cover/back art to resolve (was folder-existence only → kit missing art failed on every doc); removed duplicate `ollama_available` that shadowed v138's auto-start; extract_text reads 4 pages; kits carry `brand.json` (name/slug).
- **MAIN-APP audit fixes (A1-A6)**: Stop used to permanently lock the UI — run_inventory/run_organize/run_distribute/run_full_export had exits that skipped DONE, and only DONE re-enables the buttons. run_full_export with no manifest failed silently. `_atomic_write_json` for manifest/stats/status. `processing._staged` for flatten/OCR/resize/img2pdf/sanitize. Office sanitize was re-zipping ZIP_STORED (measured 225x inflation) → ZIP_DEFLATED, and core.xml now explicit UTF-8. `Image.MAX_IMAGE_PIXELS` now follows CFG.max_pixels (the Settings field was dead).
- Verify suite now 159 checks: engine 21, v140 45, v139 27, regression 18, boot 12, collision 4, phase1 14, phase2 12, phase3 6. verify_engine re-runs the AST sweep so a future DONE-skipping exit path fails.
- **DEFERRED (A7, user's call)**: Standard-mode dedup hashes extracted TEXT + page count, not bytes — two PDFs with identical text but different drawings collapse into one master. It's the DEFAULT ingest mode and produced Batch 4's 3,271 → 2,174. Advised spot-checking duplicates_report.csv.

Batch 4 corpus facts (analyzed 2026-08-04, 2,174 rows): source is `C:\Users\WORK\Documents\Batch 4\` with 36 brand subfolders each holding `_unique-to-rebrand\` (3,271 PDFs total), PLUS a flattened cross-brand-deduped `Batch 4\_unique-to-rebrand\` (2,174) which is what was analyzed. **49% (1,056) have no text layer at all** (confirmed with poppler, not a pypdf bug) → 901 CAD drawings (correctly left), 55 instruction guides (now flipped), ~97 unclassified. Manufacturer blank on 1,574 rows but **91% recoverable by matching filenames back to the brand folder names** (not yet implemented — the flattening discarded the brand). Analyze throughput ~1 PDF/sec (240 in 4 min) with llama3.2:3b at 100% GPU on the RTX 5060 (8GB VRAM, model 2.6GB). Some source files ALREADY carry BM branding (`_BM` suffix, e.g. 120RCS_BM.pdf) — double-branding risk, only 2 flagged by filename but others may be unlabelled.

v141-v144 (2026-08-04/05) — all released, all found by analysing the REAL batch output rather than by testing in the abstract:
- **v141**: `reviews.needs_review` flagged 70% of rows (useless as a filter); an unreadable file is now only flagged when we propose to BRAND it → 1,517→525 flagged. Plus `classify.filename_suggests_drawing()` — a low-confidence (<0.9) "rebrand" of a drawing-named file (`tech-*`, `-cut-sheet`, `-bolt-pattern`, `-elevation`) becomes "leave" with a note. Moved 211 rows; left all 215 confident rebrands alone.
- **v142**: 40% of rebrand rows have a BLANK product, so 255 files would have shipped as `installation-guide-budget-mailboxes.pdf` … `-255.pdf`. `rebrand.delivery_filename()` falls back to the source stem; when several documents share product+asset_type, ALL of them use their source name (a counting pass first — previously sheet order decided who kept the clean name). `_clip_words` truncates on hyphens; `numbered_filename` keeps `-N` inside 60 chars. Colliding files 481→53, all 874 names unique.
- **v143**: cover attribution line is now a TOGGLE, **default OFF** (`CFG.rebrand_show_attribution`). User decided to omit it despite the brief, for consistency with the signed-off Batch 1/2. It was wrong on 109/874 covers anyway (`Florencemailboxes.com`, `WebLife Stores LLC`, `Manufactured by Budget Mailboxes | Sold by Budget Mailboxes`). If ever turned back on, the manufacturer column needs normalising first.
- **v144**: dedup (audit item A7) — `Worker._page_artwork_signature()` mixes embedded-image dimensions + raw stream size into the smart hash so documents that read alike but look different stay separate. NOTE pypdf strips `/Length` from image dicts; use width/height + `len(o._data)`. Also fixed A8 (silent fallback to byte hashing now logged). Measured on the real corpus: 63/65 text-bearing duplicate pairs merged before AND after — zero regressions; 0 of 30 text-identical groups differ in artwork, so this batch never hit the bug. It's insurance.

**Full-batch run completed successfully (v142 build)**: 874 rebranded + 1,300 copied, 0 missing, 874 unique names, largest 28.2MB (0 over the 50MB cap), Author metadata correct, page count = source+2, text preserved, leave-rows byte-identical. Output 3.25GB from 2.0GB source. Covers verified visually against Batch 1 house style. **The user is re-running with v143+ to drop the attribution line** — delete `_unique-to-rebrand_rebranded` first, because Apply skips existing outputs and filenames are identical either way. **Re-running Apply does NOT need a re-Analyze** — the sheet is unaffected by Apply-time settings.

Supplied-data audit (their `_unique-to-rebrand` worklist was deduped by the client, not by us): 3,271 per-brand files → 2,921 distinct names → 2,174 in the flat worklist. 298 name-merges were byte-identical bar one (`usa-measurements-2017.pdf`). Sampled content-merges were genuine duplicates. Verdict: the supplied worklist is trustworthy.

BRIEF GAPS still unresolved (ClickUp 86ey7j4wx lists these under "APPLY to each"): **"Trusted by the Nation"**, **Version 1.0 + Last Updated [Month Year]**, and a **disclaimer** are in NEITHER the template art NOR the output — and NOT in Batch 1/2 either. Precedent supports omitting them but Adam/Shehara should confirm; adding them means rendering text onto every page (a real feature, not a tweak).

~~OPEN FEATURE REQUEST~~ **BUILT in v147** (user chose "a section under the job list", asked that naming differentiate it from ingest jobs): **rebrand runs are invisible in the dashboard.** The job list is built from ingest-created workspaces; rebrand/pipeline work on arbitrary folders and leave no record. Proposed: a rebrand history (source, output, sheet, counts, duration, kit, toggle states) written on completion to DocRefinePro_Data and surfaced as a panel. Matters more now that toggles change output. Awaiting the user's call on placement (fourth tab vs section under the job list).

**Imperial brochure SWAPPED IN (2026-08-05)** — the v145 defect is fully closed in the shipped Batch 4 output. The corrected file was rebranded on its own into `Documents\DUMP_rebranded` and copied over `…_rebranded\imperial-street-specification-sheet-budget-mailboxes.pdf`, **keeping the shipped filename** (drop-in; the 874-unique-names QA stays valid). Verified after the swap: cover 612x792 with its title, doc page 612x850, all 3,256 chars present. The broken original is backed up in this session's scratchpad. Note the DUMP sheet used asset_type `spec-sheet`, so its own delivery name differed — that naming difference was deliberately NOT carried into the deliverable.

**v144 FULL RUN QA'd (2026-08-05, attribution OFF)** — the deliverable is sound: 2,174/2,174 present, 874 rebranded + 1,300 copied, 0 covers with an attribution line (60 sampled), 874 unique names all ≤60 chars, 0 over the 50MB cap (largest 28.2MB), Author/page-count/text/leave-byte-identity all clean, no .part debris, 3.25GB. Clean re-run confirmed: all 874 rebranded files written in one 58-min block; the alarming 34-day mtime spread is just `copystat` preserving source mtimes on the 1,300 copies (1300/1300 match) — NOT stale files. Remember this before panicking next time.

**v145 (2026-08-05) — one real defect the QA caught.** `Imperial-Street-Sign-Brochure.pdf` (1 of 2,174) stores its page box top-down `[0, 792, 612, 0]` — legal PDF (a rect is any two opposite corners) but pypdf reports height **-792**. The engine took it literally: treated it as landscape, scaled the cover by a negative factor, and pushed the document's own content off the page → a branded blank file, **all 3,127 chars of content lost**, silently. Fix: `rebrand.page_size(page)` returns `abs()` of both dims, used by `_dominant_orientation` and `rebrand_pdf`. Regression test `verify_pagebox.py` builds a top-down page synthetically AND asserts against the real file. 245/245 suite.
- **Only that one file is wrong in the shipped output** — advised rebranding it alone and swapping it in rather than re-running the hour.

METHOD NOTE worth repeating: every substantive defect in v140-v145 (numbered filenames, useless triage column, wrong manufacturer attributions, the negative page box) was found by **analysing the real batch output**, never by the unit tests, which were green throughout. Always QA the actual deliverable after a run.

**v146 (2026-08-05) — "Keep the original filenames" toggle** (`CFG.rebrand_keep_original_names`, default OFF = brief's pattern). Threads through run_rebrand_apply/run_pipeline/_assign_output_names/_folder_rebrand; checkbox in both dialogs. When on, branded files keep the source name EXACTLY (no brand suffix, no marker) and the contested-name fallback is skipped (originals are already unique). 266/266 suite.
- NOT done deliberately: normalising asset_type slugs in filenames — checked on the real batch, affects **1 file of 874** (213 use `spec-sheet`, one used `specification-sheet`). Cover titles already unify synonyms. Not worth the churn.

**RENAMING vs LIVE URLS — the decisive finding (2026-08-05).** `Batch 4\_unique-to-rebrand\_unique-index.csv` has columns `filename, pdf_url, md5, bytes, products_using` — **keyed on FILENAME**, carrying each file's live URL on media.budgetmailboxes.com. Every per-product `<Brand>\<SKU>\_manifest.csv` is also a single `filename` column. Numbers: 2,224 unique files, **16,374 product-page references**; the most-referenced file is used by 599 products. Renaming the 874 rebranded files means **874 301-redirects protecting 11,638 product-page references**, and every per-product manifest goes stale. Keeping original names = drop-in replacement, zero redirects, manifests stay valid.
- The Batch 4 brief contains BOTH "FILENAME: [product]-[asset-type]-budget-mailboxes.pdf" AND "don't change live URLs (301 redirects handled later by Ramez)" — they pull against each other.
- **Batch 2 task 86ewmngk3** (Kunchana Godahewa → Jason, status complete, Shehara signed off): scope "178 product PDFs… according to the approved template" — NO filename pattern, NO attribution line, NO version/disclaimer stamps. Jason's comment: "All **2277** PDFs have been rebranded" (178 unique → 2,277 SKU copies) and he hit a bug where "the app isn't **redistributing the rebranded files back to the original locations**". That redistribute-to-original-locations workflow only works if names are unchanged — which is WHY Batch 1/2 kept original names.
- **Recommended to the user: re-run with keep-original-names ON.** Awaiting their call with Adam.

BRIEF-vs-PRECEDENT is now a PATTERN, worth raising with Adam as ONE question, not three: (1) attribution line, (2) filenames, (3) the three missing stamps ("Trusted by the Nation", Version 1.0 + Last Updated, disclaimer). The Batch 4 brief describes a target the approved template shell never implemented, and Batch 1/2 shipped and were signed off without any of it.

**CORRECTION to the brief-vs-precedent read (2026-08-05).** I initially judged the Batch 4 extras a drafting slip because Kunchana's Canva guide (ClickUp doc `8cr937k-390278`, Nov 2026 — cover/header/footer/watermark/back-cover, export PDF Print + RGB, NO filename or stamp requirements) omits them. **Wrong.** The MBW brief (`86ey9rdk5`, authored by Kunchana) cites a formal SOP — `team/kunchana/designer-sidekick/knowledge-sources/vp_proc_sitecontent_rebranding-asset-management.md` — which is the CURRENT standard. The Canva guide is the superseded pilot process. So attribution line, tagline, Version 1.0 + Last Updated, disclaimer and the naming pattern ARE the real standard; Batch 1/2 are the outliers. Renaming is sanctioned and "Ramez owns live re-upload naming/301s at deploy".
- **Batch 1 origin** (`86evrpk8a`, parent `86evnee38`): pilot, "established Canva process", "Add Budget Mailboxes branding (**co-branded or Budget Mailboxes only**)" — so the attribution toggle maps onto a documented sanctioned choice; Batch 1 delivered "BM only". Batch 2 (`86ewmngk3`, Kunchana→Jason, Shehara signed off): "178 product PDFs… according to the approved template"; Jason's comment says **2,277** PDFs rebranded (178 unique → SKU copies) and he hit a bug where the app "isn't redistributing the rebranded files back to the original locations" — that redistribute step is WHY names were preserved.
- **CONCRETE SOP MISMATCH to raise with Kunchana**: the SOP puts manufacturer attribution in a **small FOOTER on every page**; we render it as a COVER subtitle. Same text, wrong place.

**NEXT JOB — MailboxWorks** (`86eyaantb`, full brief `86ey9rdk5`, design lead Kunchana, TA lead Sean): 1,410 SKU folders, **9,087 PDFs** + a `_manifest.csv` per folder, ~10,497 files. Templates at Drive `_rebrand-templates` = Portrait/5 + Landscape/5 (Header · Footer · Cover · **Back-Cover** hyphenated · Watermark PNG) + `_psd-masters/` + `_preview.html` — **exactly the BrandKit layout, works today**; a `brand.json` gives name "MailboxWorks" / slug "mailboxworks" so covers, metadata and filenames all follow. Naming `[product-name]-[asset-type]-mailboxworks.pdf`. 83 Whitehall decor SKUs out of scope. QA: first N per brand full review → spot-check → Catalog → Shehara → Adam.
- READY today: brand kit, brand slug, naming (either convention), classification. NOT BUILT: tagline, Version 1.0 + Last Updated, standard disclaimer (per-page text stamps = real work), and footer-placement for the attribution.

**UNIQUE-COUNT DISCREPANCY explained (2026-08-05).** Workspace `Batch 4_20260709_121017` (36,996 files ingested 2026-07-09) reported **5,165 "Unique Masters"** vs the supplied worklist's 2,224 — the user asked why. Answer: the headline counts EVERY supported type — **2,622 .pdf + 2,533 .csv + 10 .png**. The CSVs are the 6,340 per-SKU `_manifest.csv` files being treated as documents (ingest has no file-type filter). Comparable figure is **2,622 unique PDFs vs 2,224** (~18% over). That residual gap is our dedup being conservative: **50% of masters have no extractable text → byte-hash fallback**, so the same drawing re-exported with different metadata counts twice. Sampling 400 PDF masters, 161 shared a content signature (text + page count + artwork) with another. Erring toward MORE uniques is the safe direction (nothing lost, just extra work) but **our unique count is NOT directly comparable to a supplied worklist** — matters for MBW where we do the dedup ourselves.
- That job is also still stuck on status PROCESSING from 2026-07-09 (the stale-status bug, fixed in v140; the job predates the fix).

**v147 (2026-08-05) — the SOP page stamps, each its own toggle.** New `docrefine/stamps.py` + `runs.py`. Closes two of the three brief gaps.
- **Stamps**: manufacturer attribution *in the page footer* (the SOP mismatch — was a cover subtitle), tagline, `Version 1.0 · Last Updated [Month Year]`, disclaimer. Four independent checkboxes (`CFG.rebrand_footer_attribution` / `_stamp_tagline` / `_stamp_version` / `_stamp_disclaimer`), **all default OFF**; off = byte-identical to v146 output. Shared `StampOptions` group box in both dialogs.
- **Mechanism**: a light band added between content and the footer art — page extension, never an overlay, so the document is never covered/cropped. `_draw_overlay` now returns `(hh, fh, sh)` and content is lifted by `fh + sh`. Ink colour is read from the kit's own art (`_art_ink`), so it's on-brand for any kit. Sizes scale with page width (works on the 44x34in drawings).
- **Wording is DATA, never code** — `brand.json` in the kit holds tagline/disclaimer/version_label/last_updated/attribution/manufacturer_aliases/stamp_ink/stamp_bg. Neither brief gives the disclaimer or tagline text (it's in the SOP), so nothing is hardcoded; a blank field prints nothing and the run log names the gap. Template shipped at `docrefine/assets/brand.example.json`. This also stops BM's tagline appearing on MailboxWorks docs.
- **Manufacturer cleaning** (the blocker v143 flagged): a value that is a website, the seller, or the brand itself yields NO attribution rather than a wrong one. Measured on the real 2,174-row sheet: of 874 branded, **442 would print, 324 blank, 108 unusable** (84 `Florencemailboxes.com`, 10 `WebLife Stores LLC`, 7 `Budget Mailboxes`…). Counts logged before the run.
- **Run history**: `runs.py` → `DocRefinePro_Data/rebrand_runs.jsonl` (JSONL, capped 500, `ts` in ms = identity). Dashboard section **"Rebrand & Processing Runs"** under the job list, with the job list now labelled **"Ingest Jobs"**. Rows = `<parent>\<folder>` + result + when; details show kit, sheet, duration and **the settings that produced it** (which is what distinguishes two runs of the same folder). Buttons: Output / Sheet / Forget.
- Released: commit ea7daa1, tag v147, CI green both platforms, both assets, ClickUp subtask `86eyhpkn5` created automatically.

**v148 (2026-08-05)** — found by running v147's attribution over the REAL sheet (the method note again): **9 manufacturers are spelled more than one way** — `Venia Products LLC` in 4 forms, Whitehall in 5, `SALSBURY INDUSTRIES`/`Salsbury Industries`. Same company credited differently across one delivery set. `stamps.spelling_variants()` groups values identical after case/punctuation stripping and the log names `manufacturer_aliases` as the fix — it reports, never rewrites (which spelling is right is the brand's call). Also fixed `runs.label()` showing a Windows separator on macOS. NOTE `Salsbury` (122) vs `Salsbury Industries` (31) do NOT group — genuinely different strings — so aliases still need a human eye. Released: commit 0902384, tag v148, CI green both platforms, both assets, ClickUp subtask `86eyhpp2j` auto-created.

**NEXT STEP the user still owes Kunchana/Adam** (they said "keeping our options open depending on Kunchana's response"): the app is now ready either way, but the **tagline and disclaimer wording must come from the SOP** — until a `brand.json` with that text exists in `Batch 4\Template`, those two toggles print nothing by design. The three-part question to raise remains: (1) attribution — now available in the SOP's footer position, (2) filenames, (3) the tagline / Version+Updated / disclaimer wording. Footer-vs-cover placement is no longer a build question, only a decision.

Verify suite is now **399 checks / 17 scripts**, consolidated into ONE place: scratchpad `…\1d0f73d0-cf96-4933-bcf2-9e0a3ced0edb\scratchpad\` (they had been scattered across 4 session folders). Run `python verify_X.py <repo> "<…>\BrandKit\Sample Files"`; `verify_phase0.py` needs a 3rd arg (an output dir). `verify_boot.py` pins the version and now also asserts README + CHANGELOG agree with `CURRENT_VERSION` — bump all three.

Real brand kit in use: `C:\Users\WORK\Documents\Batch 4\Template` (no brand.json yet — add one to use the stamps). Brand navy = RGB(32,14,101). The footer art ALREADY carries the URL + phone + USPS/Made-in-USA marks, so the brief's URL/phone items were never missing — only the four stamp items were.

═══════════════════════════════════════════════════════════════════
**STATE AS OF 2026-08-12 (end of session) — READ THIS FIRST**
═══════════════════════════════════════════════════════════════════

**Released this session: v147 → v154, all CI green both platforms, all ClickUp subtasks auto-created.**
- v147 SOP page stamps (4 toggles, off by default) + rebrand run history. v148 manufacturer spelling-variant reporting. v149 `Last Updated:` colon + centred disclaimer. v150 footer keeps its shape with no manufacturer (343 files). v151 **/Rotate — landscape pages were squeezed into portrait and CROPPED (217 files)**. v152 bookmarks (109) + AcroForm (379) now survive. v153 drawings left alone on a confident call + OS junk excluded. v154 classifier reads page SHAPE (density<10 + landscape), catching 29 drawings the filename missed.
- Verify suite **443 checks / 18 scripts**, consolidated in scratchpad `…\1d0f73d0-cf96-4933-bcf2-9e0a3ced0edb\scratchpad\`. New: verify_nav, verify_vision, verify_vision_live, pixel_truth.py, deep_scan.py.

**KUNCHANA'S DECISIONS — all settled and applied (ClickUp 86ey7j4wx thread):** attribution ON in the page **footer**; **keep original filenames** (no renaming, no 301s needed, Ramez not involved); tagline "Trusted by the Nation"; Version 1.0 on every file; "Last Updated: [Month Year]" (colon); disclaimer verbatim *"This guide is provided for reference. Always consult manufacturer specifications for complete details."*; stamps done by the tool, not the artwork. **All 19 canonical manufacturer names supplied and written into `Batch 4\Template\brand.json`** (47 aliases). Sample approved. Watermark confirmed as the established Batch 1 look (verified against the hand-made set on G:) — no hold.

**RUN SETTINGS for Batch 4:** footer attribution ON · tagline ON · version ON · disclaimer ON · cover attribution OFF · keep original filenames ON · **complete set OFF** (source holds `desktop.ini` + `_unique-index.csv`; both are sitting in the current shipped output and must not ship). Delete `_unique-to-rebrand_rebranded` first — Apply skips existing files.

**NEXT STEPS (in order):**
1. **Re-Analyze from scratch with the visual pass ON** (user's stated intent) → produces the corrected sheet in one go, including the drawing reclassification. Expect ~244 of 874 to flip to `leave` (→ ~630 rebrand). Attribution will print on ~531 files after aliases (was 442).
2. Packaged-build check (3 files) — the frozen path has still never exercised stamps.
3. Full run → QA with `pixel_truth.py` (threshold ~6%) + the trial's 100 checks.
4. Reply on ClickUp + release notes. Deadline: **Thursday 2026-08-13 EOD**.

**OPEN FEATURE REQUEST (user's, not yet built): hibernate/sleep on completion**, like qBittorrent — a toggle so a long run puts the machine to sleep when it finishes. Windows `shutdown /h` or `SetSuspendState`; macOS `pmset sleepnow`. Should be per-run (dialog checkbox), remembered in CFG, and must not fire if the run was stopped or failed.

**DECIDED AGAINST: hooking the app to the Claude API for classification.** Costed properly — 1,512 files ≈ 184MB payload, ~2,200 input tokens/file → **$2 (Haiku 4.5 batch) to $11 (Opus 5 batch)**, one Batch API job under the 256MB cap. Rejected because (a) the local `qwen2.5vl:7b` already scored **14/14** on hand-labelled ground truth, so there is no accuracy headroom to buy — the only gain is wall-clock; and (b) the README promises *"runs 100% locally — no cloud uploads"*, which a cloud classifier would break. If revisited, it must be opt-in with local as the default; the real argument is machines that can't host a 7B model at all, not this laptop.

STILL UNVALIDATED: the frozen/packaged rebrand path (bundled reportlab + Poppins under PyInstaller) has only ever been checked by build success + boot smoke test. User's packaged Batch 4 run (Export → Rebrand Unique Masters → Analyze → review CSV → Apply) is the first real test — advised trying a handful of files before the full ~3,314.

FEATURE COMPLETE (all 5 phases + audit hardening + Ollama setup UX). This 15.6GB machine → 2 rebrand workers. For the real 3,314-file job: Analyze folder → review `_rebrand_plan.csv` in Excel → Apply.

═══════════════════════════════════════════════════════════════════
**STATE AS OF 2026-08-12 LATE (session 2) — READ AFTER THE BLOCK ABOVE**
═══════════════════════════════════════════════════════════════════

**v155 RELEASED** — the vision pass (see [the vision pass](vision-pass.md)) is shipped; it is no longer uncommitted. Suite 468/468 / 19 scripts.

**v156 BUILT ON A BRANCH, DELIBERATELY NOT MERGED YET** — branch `feature/sleep-on-completion`, commit b1366bd, worktree under this session's scratchpad. Held back so the Batch 4 delivery runs on clean released v155 code; merge + release *after* the delivery. Contains two things:
- **Sleep/hibernate on completion** (the user's open request): `docrefine/power.py`, `SleepWhenDone` + `SleepCountdown` in dialogs, checkbox on BOTH rebrand dialogs (it applies to whichever step you start). `CFG.sleep_when_done` + `sleep_when_done_action` ("sleep"|"hibernate"). The substance is the refusals — never sleeps after Stop, after an `EventType.ERROR`, or if the countdown is cancelled, and each says why in the run log. `verify_sleep.py` = 38 checks with `power.suspend` stubbed so the suite can never sleep the machine.
- **`main.py --self-test-rebrand <workdir> [src] [kit]`** (`docrefine/selftest.py`): lets a PACKAGED build prove it can brand a PDF, stamps and all. Generates its own source PDFs *and* brand kit (no fixtures), so CI can run it. Writes `_self_test_report.txt` and signals via exit code, because a windowed build has no console. 42/42 from source.

**GOTCHA worth remembering: `brand.json`'s `last_updated` overrides the WHOLE suffix.** Leave it `""` (as the real Batch 4 kit does) and the engine composes `"Last Updated: <Month Year>"` from the run date. Put a bare `"August 2026"` in it and you silently get `Version 1.0 · August 2026` with **no "Last Updated:" label** — the v149 fix appears reverted when it isn't. My self-test fixture hit this; the assertion now checks for the colon.

**`verify_boot.py` no longer hard-codes the version** — it takes the expected value from CHANGELOG's newest `## [vNNN]` heading and asserts config + README agree. Same guarantee, one less thing to hand-edit each release.

**v157 (2026-08-14) — resumable classification + PDF rescue + one failure rule.** Suite **563/563 / 23 scripts** (new: `verify_checkpoint` 25, `verify_repair` 23).
- **`docrefine/checkpoint.py`** — append-only JSONL beside the sheet, flushed per file, tolerates a truncated last line, reuses a record only while size+mtime match. **NOTE: the module and the worker wiring were already in the tree, not written by me** — I reviewed it and fixed the real bug: `clear()` was never called, so after a SUCCESSFUL run the progress file survived and the next analyze would skip every file and return a stale sheet while looking like it worked. Also closes the writer on all exits.
- **A stopped analyze still writes NO sheet** (deliberate: an incomplete classification must not replace a reviewed one). The v155 changelog line "You can Stop mid-pass; the text results are kept" was **false when written** — the code discarded everything. Now true and precisely worded.
- **`processing.repair_unreadable_pdf()`** — rewrites a pypdf-unreadable PDF via bundled `pdftocairo -pdf`, accepted only if it opens, matches pdfinfo's page count and keeps ≥90% of pdftotext's characters. Fixes the two malformed-104-bit-RC4 files. `processing._poppler()` decodes subprocess output as bytes→utf-8/replace, because poppler emits whatever a PDF carries and text mode raises UnicodeDecodeError on ordinary files.
- **`Worker._fallback_deliver(pdf, dst, rel, exc, rebrand=None)`** is now the single failure rule for BOTH rebrand paths; returns "repaired" | "unbranded" | None.
- **CropBox: deliberately NOT fixed.** Measured 1 page in ~1,800 (0 of MBW's 1,122 distinct docs, 1 of Batch 4's 652), no content lost. `rebrand_pdf` returns `trimmed_pages` and the log names the file. A real fix must expand the crop by the strip heights or it clips our own branding.
- **STILL OPEN (latent divergence):** the two paths return different status strings for the same outcome — `"failed"`/`"fail"`, `"rebranded"`/`"done"` — and each computes its own counts. Left alone because changing them risks miscounting a delivery.

**MBW DECISIONS ALL SETTLED — Kunchana on `86eyaantb` (comment `90180247128578`, 2026-08-18). Clear to start.**
- **Shape: A** — 970 branded documents + the manifests, as BM Batch 4. NOT the 9,087-copy tree.
- **Filenames: keep the originals.** Renaming and 301s happen at site deploy with Ramez, not in our output.
- **Tagline: OFF — MailboxWorks has no tagline.** Disclaimer verbatim, same as BM. `Version 1.0` + `Last Updated: [rebrand month]` ON. Attribution `Manufactured by [X] | Sold by MailboxWorks` in the page footer, cover attribution OFF.
- **No SKU exclusions.** The 83 out-of-scope Whitehall decor items were never uploaded, so all 100 Whitehall folders in the tree are in scope.
- Complete set OFF (source holds `desktop.ini` + 1,410 `_manifest.csv`).
- He still needs the **manufacturer spellings table after the analysis** — same round-trip as Batch 4, and historically the real gate.
- **`brand.json` written to `MBW Downloadable re-branding\_rebrand-templates\` and backed up to `DocRefinePro_Data\Brand Kits\mailboxworks-brand.json`.** Verified: brand_name resolves to MailboxWorks (no longer the Budget Mailboxes default), attribution/version/disclaimer all render, `last_updated` left blank so the engine composes the `Last Updated:` label.

**GOTCHA — `BrandKit` looks for the literal filename `brand.json`.** Anything else (e.g. `budget-mailboxes-brand.json`) is ignored and the kit **silently falls back to defaults**: zero aliases, no tagline, no disclaimer, and the brand name defaults to *Budget Mailboxes*. It looks like it works and prints raw, inconsistent manufacturer names. Hit this on 2026-08-18 when the Batch 4 kit was restored under the backup's filename. **Batch 4 kit is now correct again** (`Batch 4\Template\brand.json`, 56 aliases, both orientations verified).

**MBW SOURCE SURVEYED 2026-08-14** at `C:\Users\WORK\Documents\MBW Downloadable re-branding` (10,518 files / 11.31GB): 9,087 PDFs, 1,410 `_manifest.csv` (single `filename` column, same as BM), `_rebrand-templates` (the kit), 10 PNG + 9 PSD, and a stray `desktop.ini`.
- **The real work is 970 documents, not 9,087** — 1,122 distinct filenames, **970 distinct by content hash**, **0.46GB of 11.31GB**. ~9x duplicated; `4C-Surface-Mount-Collar-Installation-Manual.pdf` appears **321 times**. So MBW is **less than half of Batch 4 (2,174)** and ~5h of vision, NOT the 28h I first estimated from the raw count. Don't extrapolate runtime from raw file counts.
- **0 cases of one filename carrying different content** → dedup is safe here. But **35 documents exist under multiple filenames, one under 43** — that decides deliverable shape (970 files vs 9,087 copies) and interacts with the brief's rename pattern. Asked Kunchana on `86eyaantb` (comment `90180246790629`).
- **The MBW kit has NO `brand.json`**, so `BrandKit` silently falls back to defaults naming **Budget Mailboxes** — it would put the wrong company on every MailboxWorks cover. Must author one (name "MailboxWorks", slug "mailboxworks") before any run. Art itself is complete: Portrait + Landscape, all five assets each.
- 100 SKU folders reference "whitehall" vs the brief's 83 out-of-scope decor SKUs — needs a list or a rule.
- **Kunchana is he/him.**

**THE BATCH 4 BRAND KIT NO LONGER EXISTS ON DISK (noticed 2026-08-14).** `C:\Users\WORK\Documents\Batch 4\Template` and the `_unique-to-rebrand_rebranded_previous_*` rollback copy were both deleted after the upload. The artwork is re-downloadable from Drive; the **manufacturer decisions were not**, so the reconstructed `brand.json` (Kunchana's 19 canonical names + the 8 case-variant groups he approved, 56 aliases, plus tagline/disclaimer/version wording) now lives at **`Documents\DocRefinePro_Data\Brand Kits\budget-mailboxes-brand.json`** — verified parsing through `BrandKit` + `clean_manufacturer`. Copy it back beside the artwork when the kit is restored.
- Consequence: the QA scripts that hard-code `Batch 4\Template` (`pixel_truth.py`, `trial_run.py`, `trial_qa.py`, `deep_scan.py`, `verify_attrib`-adjacent helpers) will fail until the kit is restored or the path is parameterised.
- Lesson: a brand kit holds *decisions*, not just images. Keep `brand.json` somewhere durable, not only inside a working folder that gets tidied up.

**BATCH 4 SIGNED OFF AND UPLOADED 2026-08-13.** Kunchana approved on ClickUp `86ey7j4wx`: the 652 count ("the 62 mislabeled installation guides coming IN is a real catch — those were missed in every previous batch"), all 8 case-variant spellings including `bobi`, and the two RC4 files + CropBox note. **One change he asked for: downsample the oversize file rather than ship it.** `frank_lloyd_wright_collection.pdf` 58.2MB → 26.0MB downsampled → **27.1MB branded**, done by resampling only its 20 over-target images to a **200 DPI floor** (scratchpad `downsample_flw.py`), then rebranding it alone and swapping it over the shipped file keeping the filename (`swap_flw.py`, same pattern as the earlier Imperial swap). 32 pages and all 9,596 text chars intact. Final set: **2,174 files, 0 over 50MB, 0 non-PDFs, largest 28.2MB**. Uploaded from `Batch 4\_unique-to-rebrand_rebranded`.
- **DPI-floor trick worth reusing:** compute an image's resolution as if it spanned the full page width — that is the LOWEST it could display at, so resampling to a 200 estimate never puts the true figure below 200. Conservative by construction.
- Rollback: previous delivery kept as the sibling folder `_unique-to-rebrand_rebranded_previous_20260813_100008`. **Do not upload the parent dir** or both sets ship.
- **v157 candidate:** a "reduce to fit the size cap" option in the rebrand dialog, so this isn't a hand-written script each time — MBW's 9,087 files will have their own oversize cases.

**BATCH 4 RUN DETAIL 2026-08-13 (vision-pass run).** Analyze took 10.35h (1,515 vision files at ~24s each). Sheet: **652 rebrand / 1,522 leave** (was 874/1,300) — the vision pass pulled **284 drawings OUT** and **62 genuine guides IN** (e.g. `4001_line-drawing.pdf`–`4005`, named like drawings but actually step-by-step guides the filename rule had skipped). 1,513 of 2,174 decided by seeing the page. Apply took 21.3 min → 2,174 files, 2.69GB, 0 non-PDFs (complete set OFF finally kept `desktop.ini`/`_unique-index.csv` out).

**THREE REAL DEFECTS FOUND BY RUNNING IT — the method note again:**
1. **`cryptography` was missing everywhere** (requirements.txt, venv, frozen bundle). pypdf needs it for AES; 9 rebrand rows raised `DependencyError` and would have shipped **silently unbranded**. Added to requirements + spec, and an AES-encrypted PDF is now generated inside `selftest.py` so CI catches it forever (self-test 43→53 checks).
2. **`run_rebrand_apply` DROPPED files that failed to brand.** Its `except` logged and returned "fail" with **no fallback copy** — unlike `_folder_rebrand`, which does copy. Two files were simply absent from a 2,174-file set. Fixed: it now copies the original through and says so. *Don't confuse the two failure paths again — `_folder_rebrand` is the pipeline, `run_rebrand_apply` is the delivery path.*
3. **An emoji in a log string masked a real 50MB breach.** `⚠️` cannot encode under cp1252 on redirected stdout, so `self.log(...)` raised, was caught upstream as a generic file failure, and the oversize warning vanished — `frank_lloyd_wright_collection.pdf` shipped at **59.4MB** unreported. All emoji removed from worker log strings, and the `"oversize"` status no longer depends on the log sink succeeding. **General rule: a logging failure must never be counted as a work failure.**

**Also known, shipped and flagged to Kunchana:** that 59.4MB file is 58.2MB *at source*, so it cannot meet the 50MB limit branded OR left — only recompression would, which alters content (user chose: ship branded, flag). Two files (`instruction-installation-receptacles.pdf`, `instructioninstall2255.pdf`) have malformed RC4 that pypdf rejects even with `cryptography` — delivered unchanged; poppler CAN read them, so a decrypt-via-poppler pre-step would fix it.

**Manufacturer aliases: Kunchana's 19 groups + 8 more case-only groups** found by running `spelling_variants` over the *cleaned* names (`ABC STEEL CO.`, `Architectural Mailboxes LLC`, `Bobi`, `FLORENCE CORPORATION`, `GOT IT WHOLESALE`, `IMI MAILBOX SYSTEMS`, `MAYNE`, `QUALARC`/`Qualarc`). Now 0 variants, 65 manufacturers, 495/652 carry an attribution line. **Gotcha: `worker`'s spelling warning computes over RAW sheet values with no aliases applied, so it cries wolf forever once aliases exist** — fix it to report only what aliases don't resolve.

**Open, not blocking:** `install_keystone_post_cuff.pdf` (1 of 652) has a CropBox 36pt inside its MediaBox; branding resets the CropBox to the MediaBox, so a trimmed margin becomes visible (5.9% pixel diff). No content lost. A fix must EXPAND the original crop by the strip heights — naively preserving it would crop our own branding out of view.

**PACKAGED-BUILD CHECK (was: never validated).** `verify_frozen_assets.py` (this session's scratchpad) points `sys._MEIPASS` at the extracted real release zip and asserts what the resolution code actually finds: **18/18**. The risk it was hunting is silent — `rebrand._draw_overlay` does `sh = stamps.height(...) if (stamps and _ensure_font()) else 0.0`, so a bundled Poppins that fails to resolve drops **every page stamp with no error at all**. Confirmed resolving, plus brand.example.json, report.html, poppler and tesseract; a stamp band renders with all four stamps and Poppins embedded. Frozen exe boots (`--dry-run`, exit 0). Note pure-Python deps (reportlab/openpyxl) are in the PYZ, so they show 0 loose files in the zip — that is normal, not a packaging fault. `--self-test-rebrand` above is what closes the remaining hidden-import gap.

---

**MBW RUN STATE 2026-08-18 — analysis done, waiting on Kunchana's answers.**

Ingested via the **v157 GUI in Lightning mode** (Jason drove it; my push to script it was wrong —
the GUI is tested code and the headless driver would have been new code). Lightning = pure byte
hash, chosen deliberately: it reproduces the **970** figure Kunchana was quoted and cannot merge
two genuinely different documents, whereas Standard's smart-text hash on a corpus where half the
PDFs have no text gives a mixed methodology comparable to nothing (the Batch 4 discrepancy).
- Reconciled **exactly**: 10,507 of 10,518 scanned (11 skipped = `desktop.ini`, 9 `.psd`,
  `_preview.html` — no documents), **970 PDF masters**, 630 unique CSVs of 1,410, 10 kit PNGs,
  **0 quarantined, 0 collision suffixes**, no name drift either direction. Ingest 99.9s.
  `duplicates_report.csv` = 8,897 rows = 8,117 PDF dupes + 780 CSV dupes. **Caveat: Lightning
  never opens a PDF, so "0 quarantined" is not "0 unreadable" — pass 1 is the first real read.**
- Workspace `MBW Downloadable re-branding_20260818_155929`; masters staged (verified byte-identical,
  970/970) to **`C:\Users\WORK\Documents\MBW rebrand work\_unique-to-rebrand`** — a SIBLING of the
  download, never inside it, or a re-ingest would walk it and double-count.
- Analyze with vision ON: **44 min**, 0 errors. Sheet **474 rebrand / 496 leave**, 1,416 pages,
  418 carrying attribution. 454 decided by seeing the page, 514 by text, 2 by fallback.
- **Comment `90180247866426` posted to `86eyaantb`**, assigned to Kunchana, with 8 numbered asks +
  3 attachments (the template render, `MBW_no-attribution-line.csv`, the full markdown).

**Open, waiting on him:** the 14 printable templates (item 5), the 15 product-line credits (item 4),
`IMI`/`AMCO`/`Mail Boss`/`BOBI ROUND` spellings, the surviving-filename rule, the manifest rewrite,
and the 13 inferred manufacturers. Proposed aliases (25, from his own Batch 4 decisions, 59→36
manufacturers, 0 residual variants) are in this session's scratchpad `mbw_aliases_proposed.json`.

**14 FILES ARE PRINTABLE TO-SCALE TEMPLATES** (13 `PP-*` + `Standing_Tall_Planter.pdf`), each
saying "Product Template and installation aid"; `PP-Large-Plaque.pdf` says "This template is
printed to scale... 21" from center of one mounting screw to the other". The sheet would brand **3
and leave 11** — one class, two treatments. **I first claimed branding breaks the scale; that was
wrong and Jason caught it.** Measured: page **width is identical to 3 decimals**, content is never
rescaled (it is page extension), so at 100% print the template still measures true. The real
effects are that 1 page becomes **3** (template is now page 2, behind a cover) and the **watermark
lands across both mounting-screw callouts**. Presentation call, not a defect — his to make.

**GOTCHA: U+FB00 `ﬀ` HAS NO GLYPH IN Poppins-Bold.** `TedStuﬀ` renders as `TedStu\x00` with **no
exception** — `stringWidth` happily returns the notdef advance. Audited all 59 MBW manufacturer
names: 2 affected, both TedStuff, and Kunchana's Batch 4 alias already maps them to plain `TedStuff`.
Worth re-running that audit per batch (`scratchpad/glyph_audit.py`: render each name through the
real stamp path, look for `\x00` in the extracted text).

**THE TEXT PASS LEAVES FILES ON A FILENAME RULE ITS OWN doc_type CONTRADICTS.** 165 MBW `leave`
rows were typed `installation-guide` by the text model yet actioned `leave` by the filename/
page-shape heuristic — because they scored >=0.9 on text, `needs_a_look` never queued them, so
vision never arbitrated. The sheet's own note even says "Flip to rebrand if this one is really a
guide". Ran vision over all 165 (10.3 min): **151 confirmed leave, 14 flips** — but only **3 of the
14 were genuine** (2 Florence foundation-prep manual pages, 1 Bobi brochure); the other 11 were the
templates above. So the review is worth doing and its output needs judgement, not auto-application.
Applied only the 3 via `verify\apply_row_decisions.py` (explicit reasoned table, sheet backed up to
`.previous.xlsx`, every row read back and asserted).

**`verify\batch_paths.py` NOW CENTRALISES PER-BATCH PATHS** — `DRP_BATCH` env var (`mbw` default,
`batch4`), plus `plan()`, `wording()` (tagline/disclaimer from the KIT), `require_kit()` (refuses a
kit with no `brand.json`), `FOCUS`. Six copies of the same hard-coded triple are gone and the
helpers no longer depend on cwd. New helpers: `apply_batch.py`, `manufacturer_table.py`,
`rewrite_manifests.py`, `review_leave_guides.py`, `apply_row_decisions.py`, `no_attribution_list.py`.

**THE LESSON THAT COST THE MOST TIME — brand literals in QA cut both ways:**
- hard-coded in an **expectation** (`Author == "Budget Mailboxes"`, `f"... Sold by Budget
  Mailboxes"`) → **correct output FAILS**. Cost me a false "wrong brand on every document" alarm.
- hard-coded in a **search** (`all(TAG in page)` with TAG blanked) → **broken output PASSES**,
  because `"" in page` is true everywhere. This is the dangerous direction.
- The fix that actually disambiguated it: a **wrong-kit sentinel** — assert no OTHER brand's name
  appears anywhere in the output. It searched for "Budget Mailboxes", found none, and proved the
  documents were right and my expectations wrong. Keep that check in every batch's QA.

**MBW BUILT AND QA'D 2026-08-19, NOT YET DELIVERED.** Kunchana answered all eight asks on
`86eyaantb` comment **`90180247980961`**; everything applied and verified.
- Kit: `brand.json` **40 aliases / 14 canonical names**, 3 deliberately blank (`Keystone`,
  `Triple Mount`, `Town And Country` — "a wrong credit is worse than none"). An alias whose value
  is `""` makes `clean_manufacturer` return `""`, case-insensitively — verified, and it survives the
  JSON round-trip. Rebuildable from `verify\write_mbw_brand_json.py`, backed up to `Brand Kits\`.
- Sheet: **471 rebrand / 499 leave**, 426 carrying attribution. Reconciles as 418 baseline − 4
  blanked + 13 approved fills − 1 (`Standing_Tall_Planter` left scope).
- Templates: all **14 left unbranded** per his item 5, including the 3 that had been in scope.
- Filenames: **9 survivor renames** applied to folder AND sheet together; 272 redirects avoided.
- Apply: **13.1 min → 970 PDFs, 0.89 GB, 471 branded / 499 left, 0 failed, 0 non-PDFs.**
- `deliver_qa` over the real tree: **ALL CHECKS PASSED**, one expected warn (36 filenames over 60
  chars — inherent to his "keep the originals" override of the brief's naming pattern).
- Manifests: **1,410 rewritten + `filename-aliases.csv` (152 rows)**, 913 rows changed across 405
  sheets, 0 unresolved, 0 rows naming an undelivered file. At `MBW rebrand work\_delivery-manifests`.
- FLW: 58.2 → 26.0 MB source at a 200 DPI floor (32 pages, 9,596 chars intact — an exact Batch 4
  reproduction), branded **26.6 MB**, the largest file in the delivery. `downsample_flw.py --install`
  now swaps it into the staged folder so Apply produces a compliant file, no post-hoc swap.

**ONE OPEN ITEM BEFORE DELIVERY — asked on comment `90180248874809`:** item 2 folded bare `IMI`
into Imperial Mailbox Systems, but item 1 (confirmed as-is) keeps `IMI Mailbox Systems` as its own
canonical name, so 20 documents credit "Imperial Mailbox Systems" and **6** credit "IMI Mailbox
Systems". The 6 are `barcelona-pier-mount-ai-1.pdf`, `Imperial-Powder-Coating{,-1,-2,-3}.pdf`,
`Imperial-Series-Assembly-Instructions.pdf` — all Imperial-named, and `Imperial-Powder-Coating.pdf`
is credited differently from its own `_Revised` sibling. If he says fold, it is a one-line
`write_mbw_brand_json.py --imi-consistent --write` plus deleting those 6 outputs and re-running
Apply (which skips existing files, so it only redoes the gaps).

**`verify\survivors.py` IS THE SINGLE SOURCE OF "WHICH FILE DID WE SHIP".** The rename step and
`rewrite_manifests` both call `resolve()`, and `rewrite_manifests` hard-fails if the staged folder
disagrees with the resolved set. Two implementations of that question is how Batch 4 lost two files.
Tie-break: keep the CURRENT master when it is already among the most-referenced, so an exact tie
never churns a live URL for nothing (2 of the 35 groups).

**SIZE GROWTH, MEASURED AND UNEXPLAINED — v158 investigation, not a defect.** Branded text/vector
documents come out **~2.9x their source size** (`imperial-pdf.pdf` 6.4 → 26.3 MB over 19 pages),
while image-heavy ones barely grow (FLW 26.0 → 26.6 MB over 32). Measured, so do not re-guess:
the branding art IS shared correctly (3 XObjects reused across all 19 pages, 0.47 MB of images
total), and the growth is in page content streams. Two hypotheses tested and BOTH WRONG —
per-page art embedding (no: art is shared) and uncompressed content streams (no:
`compress_content_streams()` made it *bigger*, 26.3 → 32.7 MB). Nothing breaches the cap (max
26.6 MB, 0 over 40 MB) and Batch 4 shipped the same way, so it is efficiency, not correctness.

**MBW FINAL — IMI FOLDED, QA GREEN, AWAITING UPLOAD (2026-08-21).** Kunchana answered the IMI
question on `86eyaantb` comment **`90180249036574`**: "yes, fold those 6 as well: one credit,
Imperial Mailbox Systems, across all 26 documents. Same company — item 1's canonical list shouldn't
preserve the split item 2 removed."
- Rebuilt the kit with `write_mbw_brand_json.py --imi-consistent --write` → **41 aliases, 13
  canonical names**. Measured the delta FIRST: exactly **6** files changed, nothing else.
- Deleted those 6 outputs and re-ran `apply_batch.py --resume` → 6 rebranded, 964 skipped, 5
  seconds. `--resume` is what stops Apply moving the whole previous tree aside.
- Verified by reading the PDFs back: all 6 print "Manufactured by Imperial Mailbox Systems | Sold by
  MailboxWorks", and a sweep of **all 970** files found **0** still printing the retired credit.
- **`deliver_qa` ALL CHECKS PASSED** (253 documents inspected page by page, 228 attributions printed
  + 25 correctly blank). Only warn: 36 filenames over 60 chars, inherent to his keep-the-originals
  override of the brief's naming pattern.

**FINAL DELIVERY NUMBERS:** 970 files (471 branded / 499 byte-identical), 1,416 pages branded,
0.96 GB, largest 26.6 MB, **0 over 40 MB**. **426 of 471 carry an attribution line**; **22 distinct
manufacturers from 59 raw spellings, 0 variants left**. Manifests: 1,410 rewritten +
`filename-aliases.csv` (152 rows), 913 rows repointed across 405 sheets, 0 unresolved.
- `Imperial Mailbox Systems` ends up on **37** documents, not the 26 Kunchana quoted — his own item 4
  (USPS/Barcelona/Quad/Twin/Norris) and item 8 (3 Newspaper Holder sheets) added the rest. Say so
  in the delivery note so the number is not a surprise.
- Trees: `MBW rebrand work\_unique-to-rebrand_rebranded` (970 PDFs) and
  `MBW rebrand work\_delivery-manifests` (1,410 sheets + aliases CSV).

**REMAINING: Jason uploads to Drive, THEN the delivery comment goes out with the link, THEN the task
moves to status `done / review`.** That status string exists on this list (others: `pending
deployment`, `complete`, `revision`, `waiting reply`). Delivery draft is written and approved in
substance — only the Drive link is missing. Nothing gets posted without Jason's go-ahead.

---

**THE CONFIDENT-WRONG CLASSIFICATION GAP (found 2026-08-21 by Kunchana querying two delivered
Batch 4 files). The most important defect class in this project so far.**

`classify.needs_a_look` queues a file for the visual pass only when it has no extractable text or
the text model's confidence is < 0.9. **A file the text model reads CONFIDENTLY AND WRONGLY is
therefore never looked at.** Measured: **413 of 652 branded Batch 4 files (63%)** and **328 of 471
MBW branded files (70%)** were decided by the text model alone, all at 0.9-1.0 confidence.
- Symptom the client saw: product catalogues, brochures and `E*-spec.pdf` spec sheets carrying a
  cover title of **"INSTALLATION MANUAL"**. The cover title is the largest text on page one, so
  this is the most visible thing we can get wrong.
- **`verify\review_branded_titles.py`** asks vision about every branded row the text model decided
  alone. Report-only, and it separates TITLE disagreements from ACTION disagreements because the
  latter change counts the client approved.
- **Vision disagrees on ~74% of titles, and almost all of it is noise.** `verify\apply_title_fixes.py`
  filters with a document-FAMILY rule (catalog / warranty / care / spec / sustainability /
  regulatory / finish / install) and applies a change only when the family changes:
  MBW 244 diffs -> 227 wording ('Product Warranty'->'Warranty') + 10 product-name + **7 real**.
- **Second filter that matters: reject a new title whose family is "other".** Vision often answers
  with the PRODUCT name ('Installation Manual' -> 'Mail Slots', -> 'Modern Standoff Address Signs').
  The SOP wants the cover to say what the document IS, so those make covers worse.
- **Vision's ACTION answers are mostly unusable — test its self-consistency before believing it.**
  It examined 73 MBW warranty documents and said brand on 60, leave on 13. Same class, 82/18 = noise.
  3 more contradicted its own calls on sibling documents. Of 18 action flags only **1** was
  self-consistent (`grande-s-round.pdf`: said "dimensioned technical drawing" AND "leave" — confirmed
  by eye, line art with measurement callouts). Applied that one; MBW went 471 -> **470 branded**.

**WE BRANDED A CORRUPT SOURCE AND SHIPPED IT.** `2B-Global-Product-Catalog.pdf` (Batch 4): damaged
xref (`xref num 485 not found`), font resources that are not dictionaries, bad colour spaces.
Declares `/Count=32`, only **5 pages reachable**, all render blank; Acrobat throws "An error exists
on this page". **The source renders blank too — we did not damage it, but branding it turned an
obviously broken file into something that looks like a finished deliverable.** Recovery tested and
failed: `pdftocairo -pdf` reconstructs all 32 pages and recovers **0 characters**. Needs a clean
copy from the manufacturer; do not ship it branded.

**`verify\scan_empty_pages.py`** checks EVERY file for (a) declared `/Count` vs reachable pages,
(b) orphaned page objects, (c) pages with no text, no image XObjects and a content stream < 60 bytes.
Results: **Batch 4** 1 file with unreachable content (the above) + the 2 known RC4 files; **MBW** 0.
- **Orphaned page objects ALONE are benign** — 7 Batch 4 and 3 MBW files have them with
  `/Count == reachable`, i.e. nothing lost. The discriminating signal is `/Count != reachable`.
  Do not fail a delivery on orphans.
- **KNOWN LIMITATION, do not forget: the corrupt catalogue EXTRACTS 1,909 characters and still
  renders blank**, because the fonts are broken. So a text-based blank-page check cannot see this
  class at all — only the `/Count` mismatch caught it. A file with an intact page tree but broken
  fonts would pass every check we have. Closing that properly needs rendering page one of every
  file (~5-10 min per batch); a cheaper first pass is to flag extracted text that is mostly
  replacement/control characters and render only the suspects. NOT YET BUILT.

**WHY OUR QA MISSED IT:** `deliver_qa` does compare source text to output page-for-page, but only
across a sample plus a focus list. A blank file outside the sample passes silently. Fold the two
structural signals into `deliver_qa` as ALL-FILE checks — that is the durable fix.

**GOTCHA: `review_branded_titles.py` / `apply_title_fixes.py` originally hard-coded an `MBW_`
prefix on their JSON hand-off file**, so running them for batch4 would have read MBW's results.
Now `f"{batch_paths.BATCH}_vision-title-review.json"`. Same class as the brand-literal bug —
anything batch-specific belongs in `batch_paths`, never in a filename literal.

**BATCH 4 DISCLOSURE POSTED 2026-08-21** — `86ey7j4wx` comment **`90180249202161`**, tagged and
assigned to Kunchana. Owns both defects, gives the 63% cause, states the all-2,174 scope, and offers
the fix. **Nothing applied to Batch 4 yet — awaiting his word**, plus a clean `2B-Global-Product-
Catalog.pdf` from 2B Global. Timing point that made this cheap: Catalog's QA leg is closed and the
set sits with **Akram Khan for review + upload**, so it is NOT live — a correction, not a rollback.

**THE FILENAME-CORROBORATION FILTER IS THE ONE THAT SAVES YOU.** Vision proposed **344** title
changes on Batch 4; only **25** survived three filters (254 rewording, 62 product-name, **3
contradicted by their own filename**). Without the third filter we would have retitled
`Venia-Products-LLC-Street-Light-Pole-Limited-Warranty-2024.pdf` as "Installation Manual" —
introducing the exact defect we were fixing. `apply_title_fixes.filename_family()` treats the
filename as independent evidence written by whoever produced the document: where it names a document
class and vision disagrees, **the filename wins**. Derived independently of my own hand-review and
rejected precisely the 3 cases I had flagged by eye — a good sign the rule is real and not fitted.
MBW re-checked under the same filter: still 6, 0 contradicted, so nothing applied there needed undoing.

**VISION HALLUCINATES CONFIDENTLY ON UNREADABLE PAGES.** It returned "Dimensioned Technical Drawing"
at **confidence 1.0** for `2B-Global-Product-Catalog.pdf`, whose pages are blank. A vision answer on
a page it cannot actually see is worthless and does not look it — never trust the confidence alone,
and always keep an independent signal (filename, SKU folder, page geometry) in the loop.

**MBW FINAL AFTER THE TITLE FIXES: 970 files, 470 branded / 500 left, 425 carrying attribution,
`deliver_qa` ALL CHECKS PASSED** (252 documents inspected). Two changes to counts Kunchana had
approved, both to be disclosed in the delivery note: 6 cover titles corrected, and
`grande-s-round.pdf` moved to `leave` (471 -> 470). Still awaiting Jason's Drive upload before the
delivery comment goes out and the task moves to `done / review`.

**MBW DELIVERED 2026-08-21.** Comment **`90180249203557`** on `86eyaantb`, tagged/assigned to
Kunchana, and the task moved to status **`done / review`**. Uploaded by Jason to the shared drive at
`H:\Shared drives\[VP] Venia Products KMS\...\MBW Downloadable re-branding\`:
- `_unique-to-rebrand_rebranded` — 970 PDFs, 912.64 MB
  https://drive.google.com/drive/folders/19KhAn2v4WHDVDaBo8BQyuucnlYtxHUsM
- `_delivery-manifests` — 1,411 CSVs + README.txt, 0.33 MB
  https://drive.google.com/drive/folders/1t-NLljWzppPkM0z_H5COVYfNmNCRjyJG
- **Upload verified against local:** 970/970 and 1,412/1,412 matching by name and size, 0 extra,
  0 missing, 0 size differences; SHA-256 identical on a 20-file sample (largest files, every file
  corrected that day, both renamed survivors, 11 random). `_unique-to-rebrand` deliberately NOT
  uploaded — it holds the 58.2 MB `.original` FLW backup and its FLW is the downsampled copy.

**GOTCHA FOR NEXT TIME — the deliverables now live INSIDE the source folder on the shared drive**,
alongside the 1,410 SKU directories. Pointing a fresh ingest at
`...\MBW Downloadable re-branding` would now walk the branded output as if it were source and
double-count. Flagged in the delivery comment. Locally I avoided this by staging as a SIBLING
(`MBW rebrand work\`), which is the pattern to keep.

**The delivery comment disclosed both post-sign-off changes** (6 cover titles, and
`grande-s-round.pdf` 471->470) rather than shipping altered counts silently, and corrected the
Imperial credit figure to **37 documents** — Kunchana had said 26, but his own item 4 and item 8
rulings added 11 more. Volunteering that arithmetic beats him finding it.

**`_delivery-manifests\README.txt`** explains the SKU mapping, the 152 aliases as Ramez's redirect
map, why filenames were kept (and hence the 36 over-60-char names), and the FLW downsample — written
for someone opening the folder from Drive with none of the ClickUp thread. Every figure in it was
verified against the artifacts on disk before shipping (`scratchpad/verify_readme.py`, 10/10).

**POWERSHELL GOTCHA that produced a FALSE PASS:** I used `$h`/`$l` as hashtables while `$H`/`$L`
held the compare roots. PowerShell variables are case-INSENSITIVE, so `$h = @{}` clobbered `$H` and
the second folder's comparison silently ran against nothing and printed "0 / 0 / 0" — which reads
exactly like a clean result. Re-ran with distinct names (`$driveRoot`/`$localRoot`) to get the real
answer. Never reuse a letter-case variant of a variable already in play in PowerShell.
