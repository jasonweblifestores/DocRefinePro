# DocRefine Pro - Changelog

## [v156] - 2026-08-12
### Added — A long run can put the computer to sleep when it finishes
Analyzing a few thousand files with the visual pass on runs for hours, so the machine gets left on overnight for work that finished at 2am. There is now a checkbox on both rebranding dialogs, the same offer a torrent client makes: finish the work, then sleep.

**The interesting part is when it refuses.** Sleeping is only ever offered after a run that finished on its own:

* **Stop was pressed** → stays awake. Someone is at the keyboard.
* **The run reported an error** → stays awake. Whoever comes back to it needs to read what went wrong.
* **The countdown was cancelled** → stays awake, and says so in the log.

Every one of those says in the run log why the machine was left on, so an overnight run that *didn't* sleep explains itself rather than looking broken. Closing the countdown window, pressing Escape and "Stay awake" all mean stay awake — every exit that isn't a deliberate choice leaves the machine running, because that is the recoverable direction. The countdown is a full minute and "Stay awake" is the default button.

The setting is remembered, it is per-run, and it is reset at the start of every run so a decision can never leak into the next one. Downloading a model doesn't count as the run you asked to sleep after — that one ends by asking you to click Analyze again.

Sleep is the default action; `sleep_when_done_action` can be set to `hibernate` instead. On Windows it goes through `SetSuspendState` and falls back to the shell if that is refused; macOS uses `pmset sleepnow`, Linux `systemctl suspend`. Where the machine offers no way to sleep on command the checkbox is disabled rather than failing at the end of a long run, and macOS is not promised a hibernate it doesn't have.

## [v155] - 2026-08-12
### Added — The classifier can look at a page when there is no text to read
**Roughly half of a typical batch has no extractable text at all** — 1,043 of 2,174 files on the current one. To a text-only model a scanned installation guide, a dimensioned drawing and a UL certificate are indistinguishable, because all three supply nothing to read. Those files were classified from their filename alone, and the previous two releases were both attempts to squeeze more out of that filename.

A local vision model now *looks* at the page instead. Where it is consulted its answer wins over the filename and page-shape rules, which only ever existed to compensate for not being able to see.

**It is measured against hand-labelled ground truth, not assumed: 14 of 14 correct.** That includes all seven Florence cut sheets, the three drawings the filename rule missed, and `1590-T1V-Spec-Sheet.pdf` — a file *named* like a spec sheet that is actually a drawing, which no filename rule can ever get right. Its stated reasoning was sound every time: "measurement arrows and dimension labels" for a drawing against "numbered steps and exploded diagram" for a guide.

**The two models never run at the same time.** The text model reads everything first and queues what it cannot settle; it is then unloaded, the vision model works through the queue, and it is unloaded at the end. This is not tidiness — it was measured with `ollama ps`:

| | seconds per file |
|---|---|
| both models resident (8.5 GB wanted on an 8 GB card, vision pushed onto the CPU) | 14.3 |
| text model unloaded first | 3.5 |

Four times faster at no cost to accuracy. Ollama also used to hold 6 GB for its keep-alive window after a run had finished and the user had moved on; the models are now released explicitly.

**Speed depends entirely on the machine, so the app measures rather than promises.** The run log reads the model's actual GPU residency and says what *that* computer is doing — "only 73% of `qwen2.5vl:7b` fits in this machine's GPU memory; the rest runs on the CPU" — suggests a smaller model when it does not fit, and reports a measured rate and estimate taken from *after* the first file, which also pays for loading 6 GB of weights and made early estimates several times too pessimistic.

### Fixed — Large-format drawings were the one thing the vision pass could not see
Rendered at a fixed resolution, a 44x34in drawing became a 4840x3740 image and the model **refused it outright** — so the very files this feature exists to identify were the ones it failed on. The render resolution is now derived from the page size to hit the target directly, which also took a large-format page from **38 seconds to 0.5**.

Off by default, and Analyze offers to download the model if it is missing. You can Stop mid-pass; the text results are kept.

## [v154] - 2026-08-12
### Changed — The classifier can now tell a drawing by its shape, not just its name
Chasing why 215 CAD drawings were labelled "installation guide" with full confidence turned up the real cause: **a drawing is a visual thing, and the model only ever sees extracted text.** On a Florence cut sheet that text is a few dimension labels plus a note reading *"Designed to mount … For use with front loading modules only"* — a fair description of an installation guide. The model answered the evidence it was given; the evidence was the problem.

Two measurements supply what was missing, taken from the real batch:

| | drawings | documents |
|---|---|---|
| text per square inch of page | 4.6 | 24.4 |
| landscape | 99% | 6% |

A landscape page carrying almost no text is a drawing. Together with the filename rule this separates 244 drawings out of 874, and **29 of them the filename never caught** — including the `1570-*-BM` cut sheets, which were doubly wrong to brand because they already carry Budget Mailboxes branding.

It costs nothing to run: no rendering, no model call, and the review sheet can explain its reasoning in a sentence. A filename that names itself as instructions still wins, so a sparse landscape mounting template stays a guide.

## [v153] - 2026-08-12
### Changed — Technical drawings are left alone even when the model is certain
The brief says to leave CAD drawings as-is, and filenames like `tech-54368` or `4c06d-02-sm_cutsheet` are how manufacturers label them. That rule previously only applied when the local model was *unsure* — a confident call still won.

On the current batch that let **215 dimensioned CAD drawings** through as things to brand, because the model labelled them "installation guide" with full confidence. A sample of six turned out to be six drawings and no guides. Confidence measures how sure the model is, not whether it is right.

The filename now decides outright, and **every row it overrides is flagged for review**, so the decision is visible rather than silent. Leaving is also the reversible direction: the original file ships untouched, and one cell in the sheet sends it back to be branded.

### Fixed — Operating-system clutter never reaches a delivery set
`desktop.ini`, `Thumbs.db` and `.DS_Store` are skipped when copying a complete set. A `desktop.ini` was found sitting in the shipped Batch 4 output.

## [v152] - 2026-08-12
### Fixed — Bookmarks and form fields now survive rebranding
A rebranded file kept its text, its links and its layout, but quietly lost two things the reader had paid for:

* **Bookmarks.** A long installation manual arrived with an empty navigation pane. On the current batch that is **109 files**. Bookmarks are now carried across, with every destination shifted by one page to account for the front cover.
* **Form fields.** The interactive fields themselves came across, but the catalog entry that registers them as a form did not — so they were orphaned and stopped working. **379 files** carry that entry, 85 of them with real interactive fields. It is now preserved.

Neither loss was visible in the output's text, its page count or its file size, which is why nothing had flagged it.

## [v151] - 2026-08-12
### Fixed — Landscape pages that are stored rotated are no longer squeezed into portrait
A PDF page can be stored as a portrait box with a `/Rotate` instruction telling the reader to turn it. Everyone who opens the file sees a landscape page; the stored box still says portrait. The rebranding engine read the box and believed it — so those documents were given **portrait covers, portrait header and footer strips, and their own content was compressed into a portrait frame and cropped at the edge**.

On the current batch that is **217 of the 874 files being branded** — a quarter of the set, mostly the Florence cut-sheet drawings. The title block on the right-hand side of those drawings was being cut off entirely.

Page dimensions are now read the way a reader sees them, so orientation, cover sizing and the page strips all agree with what the document actually looks like. Those files now get the landscape template and keep their full width.

**This affects output produced by earlier versions.** Any set rebranded before v151 should be re-run if it contains rotated pages.

Worth noting how it hid: the text of a cropped page still extracts perfectly, so every automated check passed. It was only visible on a rendered page — the same blind spot that hid the v145 defect.

## [v150] - 2026-08-12
### Fixed — The page footer keeps its shape when a document has no manufacturer
Found in a trial run over real files, not in testing: when a document had no usable manufacturer, the version and last-updated line slid across to the **left** of the footer, because it was the only thing on that row. Every other file had it on the right, beside the attribution.

On the current batch that is 343 of 874 files — so a set delivered this way would have carried two different footer layouts, which reads as a mistake rather than a rule. The attribution now always sits left and the version always sits right, whether or not the other one is there.

## [v149] - 2026-08-09
### Changed — The last-updated line now matches the SOP exactly
The wording for the page stamps came back confirmed against the SOP, and the version line carries a colon that ours did not: it reads **`Last Updated: August 2026`**, not `Last Updated August 2026`. One character, but it is the difference between matching the standard and approximating it.

The **disclaimer is now centred**, sitting square under the tagline. The attribution and version line pair off to the left and right edges because they belong together as a pair; the two lines that speak for the brand read better as one centred block.

Everything else in that line is unchanged — `Version 1.0` still resets on every file, and the month is still the month you run the rebrand, shared by every file in the batch.

## [v148] - 2026-08-05
### Added — The app now tells you when one manufacturer is credited two ways
Both findings came from running v147's new attribution stamp over the current batch's real review sheet rather than test data.

* **Inconsistent spellings are flagged before the run starts.** On the current sheet, nine manufacturers appear under more than one spelling — `Venia Products LLC` turns up four ways (`VENIA PRODUCTS LLC`, `Venia Products LLC.`, …) and Whitehall five. With the footer attribution on, the same company would be credited differently from one file to the next in a single delivery set. The run log now names the groups and points at `manufacturer_aliases` in `brand.json`.
* Nothing is rewritten automatically. Which spelling is correct is the brand's decision, not the app's — so it reports and lets you settle it.
* For context, on that sheet the attribution line prints on 442 of the 874 files being branded: 324 have no manufacturer recorded at all and 108 hold a value that names a website, the seller or Budget Mailboxes itself. All of that is now stated in the log up front rather than discovered afterwards.

### Fixed
* Rebrand history rows showed a Windows-style path on macOS. They now use the platform's own separator.

## [v147] - 2026-08-05
### Added — The page stamps the rebranding SOP asks for
The rebranding standard asks for four things on every page that the template artwork does not provide: the manufacturer attribution, a tagline, a version and last-updated line, and a standard disclaimer. None of them were in the output, because none of them are in the template. They are now available, **each as its own checkbox and all off by default** — so nothing changes unless you ask for it, and the previously signed-off sets remain the default behaviour.

* **Nothing is ever printed over your document.** The stamps get their own light band, added between the page content and the footer bar — the same way the header and footer strips already extend the page. Page size grows by the height of the band; the document itself is never overlapped, scaled or cropped.
* **Manufacturer attribution now goes in the page footer**, which is where the SOP puts it. The existing cover-subtitle option is untouched, so either placement is a tick.
* **The wording lives with the brand, not in the app.** A brand kit can now carry a `brand.json` holding its tagline, disclaimer, version label and attribution format. Copy `docrefine/assets/brand.example.json` into your kit as a starting point. Anything left blank is not printed, and the run log names what is missing — so no approximate wording can reach a document, and one brand's tagline can never appear on another brand's PDFs.
* The stamp colour is read from your kit's own artwork, so it is on-brand for whichever kit is loaded; `stamp_ink` overrides it if you need to.
* Stamps scale with the page, so they read correctly on both a Letter guide and a 44x34in drawing.

**A manufacturer name that can't be trusted is now left off rather than printed.** On the current batch the review sheet's manufacturer column contained websites (`Florencemailboxes.com`), the seller (`WebLife Stores LLC`) and the brand itself — the last of which reads as "Manufactured by Budget Mailboxes | Sold by Budget Mailboxes". Any such value means that document simply gets no attribution line, and the run log counts them up front so you know the size of the gap before an hour-long run rather than after it. `manufacturer_aliases` in `brand.json` is how you correct one.

### Added — Rebranding runs are no longer invisible
Rebranding and the pipeline work on ordinary folders and never created a job, so once a run finished the app forgot it happened. Since the output now depends on which toggles were set, "what produced this folder" had become a question with no answer.

* The dashboard now lists **Rebrand & Processing Runs** under the ingest jobs, each showing the folder, what it produced, and when.
* Selecting a run shows the brand kit, the review sheet, how long it took, and **the settings that shaped the output** — which is what tells two runs of the same folder apart.
* Buttons open the output folder or the review sheet; **Forget** drops a run from the list and touches no files.
* Each row names the folder together with its parent, because every deduplicated job's masters folder is called `01_Master_Files` and that name alone identifies nothing.
* Both dashboard lists are now labelled (**Ingest Jobs** / **Rebrand & Processing Runs**) so neither is mistaken for the other.

## [v146] - 2026-08-05
### Added — "Keep the original filenames"
Rebranded files are renamed to the delivery pattern the Batch 4 brief asks for, `product-asset-type-budget-mailboxes.pdf`. But the Batch 1 and 2 sets — both already signed off — kept every original filename. This is now a checkbox, so either convention is a tick rather than a rebuild.

* **Off (default):** rename to the brief's pattern.
* **On:** every branded file keeps the exact name it came in with — original casing, spacing and punctuation, no brand suffix added. Files marked *leave* keep their original name either way.
* Available in **Rebrand a Folder**, **Rebrand Unique Masters** and **Process a Folder**, and your choice is remembered.

Keeping the original name also avoids a side effect of the delivery pattern: when a product name and a long asset type together exceed the 60-character limit, the product is trimmed to fit — `Imperial Street Signs` became `imperial-street`, losing "signs".

Note the previous sets appended a marker to the name (`_REBRANDED`, `-RE-BRANDED`, ` REBRANDED` — inconsistently). This option does not reproduce that: the name is left exactly as it was.

## [v145] - 2026-08-05
### Fixed — A page stored "upside down" no longer loses its content
A PDF page box is defined by any two opposite corners, so writing it top-down is perfectly legal — but the page then reports a *negative* height. The rebranding engine took that at face value: the document was treated as landscape, the cover was scaled by a negative factor, and **the document's own content was pushed off the page entirely**, leaving a branded but effectively blank file.

Found during QA of the Batch 4 run: one file in 2,174 (`Imperial-Street-Sign-Brochure.pdf`) was affected, and all 3,127 characters of its content had been lost. Page dimensions are now read the way a PDF viewer reads them, and that file rebrands correctly with every word intact.

## [v144] - 2026-08-05
### Fixed — Deduplication no longer merges documents that only *read* the same
Smart deduplication compared the text it could extract from a PDF. Two documents with identical wording but different drawings — the same spec sheet for a different model, say — therefore hashed the same, and one was filed away as a duplicate of the other. On a batch you didn't deduplicate yourself, that's content quietly disappearing before you ever see it.

* The fingerprint now includes the **artwork on the page** — the dimensions and stored size of every embedded image — alongside the text. Documents that read alike but look different stay separate. Nothing is decoded, so ingest speed is unchanged.
* Genuine duplicates still collapse exactly as before: measured against 65 known duplicate pairs in the current batch, the change kept all 63 that previously merged and separated none of them wrongly.
* **When smart hashing can't run** on a file, the app now says so in the log and falls back to a plain byte comparison. It used to fall back silently, so a batch could quietly deduplicate far less accurately than you'd expect with no indication.

## [v143] - 2026-08-05
### Changed — The cover attribution line is now optional
* **New checkbox: "Manufacturer line on covers".** The Batch 1 and 2 sets, both already signed off, carry the document title alone; the task brief additionally asks for a `Manufactured by [X] | Sold by Budget Mailboxes` line. This setting decides which standard applies, and your choice is remembered.
* **It now defaults to off**, matching the previously approved sets. Tick it to follow the brief.
* Nothing else about the output changes either way — filenames, page counts, PDF metadata and the document text are identical.

Context: on the current batch the attribution line was wrong on 109 of 874 covers, because the model's `manufacturer` column contained things like `Florencemailboxes.com` (a website), `WebLife Stores LLC` (the seller) and `Budget Mailboxes` itself — which read as "Manufactured by Budget Mailboxes | Sold by Budget Mailboxes". None of those rows were flagged for review, since the classification was confident and only the manufacturer was wrong. Turning the line off removes that whole class of error; if it is ever turned back on, those values need normalising first.

Re-running Apply with the setting changed does **not** require re-running Analyze — the review sheet is unaffected. Delete the existing output folder first, though, since Apply skips files that already exist.

## [v142] - 2026-08-04
### Fixed — Delivery filenames that actually identify the document
Output names are built as `product-asset-type-budget-mailboxes.pdf`, but the local model leaves the product blank on a large share of documents (40% of the current batch). Every one of those collapsed onto the same name and got a numbered suffix — **255 files would have shipped as `installation-guide-budget-mailboxes.pdf` through `-255.pdf`**. Nothing was ever lost, but the names told a customer nothing and didn't match the delivery pattern.

* **A blank product now falls back to the source filename**, which is where the model or part number actually lives: `1570-12-BM.pdf` becomes `1570-12-bm-installation-guide-budget-mailboxes.pdf`.
* **When several documents genuinely share a product and asset type, all of them use their source name** — previously whichever row happened to come first in the sheet kept the clean name and the rest were numbered, so the outcome depended on sheet order.
* Long names are shortened on a word boundary instead of mid-word (no more `wall-mount-mailbo-…`), and the product is trimmed before the asset type so the name always still says what the document is.
* A numbered suffix, now only needed as a last resort, stays inside the 60-character limit — it previously pushed names to 61.

On the current batch this takes colliding files from **481 down to 53**, with all 874 names unique and within the length cap.

## [v141] - 2026-08-04
### Changed — A review sheet you can actually triage
* **The "review?" column now flags decisions, not just unreadable files.** On the current batch it was marking 70% of rows, which made it useless as a filter — most of those were technical drawings correctly left as-is, needing no sign-off at all. A file we couldn't read now only gets flagged when we're proposing to *brand* it. **Flagged rows drop from 1,517 to 528 (70% → 24%)**, and what remains is genuinely uncertain.
* **Technical drawings are no longer branded on a coin-flip.** When the local model is unsure (confidence below 0.9) and the filename labels the document as a drawing — `tech-…`, `…-cut-sheet`, `…-bolt-pattern`, `…-elevation` — it now leaves the file as-is and says why in the sheet's notes. Leaving a document alone is the reversible choice. A confident read of the document text still wins, so genuine spec sheets and guides are unaffected. On the current batch this moves 209 low-confidence rows out of the rebrand list while leaving all 215 confident ones untouched.

## [v140] - 2026-08-04
### Changed — Simple, customer-facing cover titles
* **Covers now say what the document is**, in the same plain style as the hand-made Batch 1 set: "INSTALLATION MANUAL", "SPECIFICATION SHEET", "PRODUCT WARRANTY". Titles are derived from the document's asset type, so they're consistent across the whole batch and there's nothing to hand-edit row by row.
* **Long titles wrap onto a second line and shrink to fit** instead of running off both edges of the page.
* Model and part numbers keep their own casing (`4C11D`, `3635RL`) rather than being mangled into `4C11d`.
* Documents with no readable text now get a clean generic title instead of a title made out of their filename.

### Added
* **Excel review sheets.** Analyze now produces an `.xlsx`: a "review?" column that flags every row worth a second look, filenames that open the actual PDF when clicked, dropdowns for action and asset type, frozen headers and filters. Existing `.csv` sheets are still read and can still be used.
* **Scanned installation guides are no longer skipped in silence.** A PDF with no text layer whose filename names it as instructions is now flagged for rebranding instead of being left as-is with everything else unreadable. (In the current batch that's 55 genuine guides that would otherwise have been missed.)
* **Brand kits can carry their own name.** Drop a `brand.json` (`{"name": ..., "slug": ...}`) beside the artwork and the cover attribution, PDF author and output filenames all follow it. Kits without one behave exactly as before.

### Fixed
* **A run that is interrupted can no longer leave a corrupt file behind.** Output is written to a temporary file and swapped into place only once complete — previously a half-written PDF was treated as finished by the resume check and would have shipped as the deliverable.
* **Brand kits are validated properly.** A kit missing its cover or back-cover artwork is now rejected up front with a message naming what's absent, instead of appearing valid and then failing on every single document.
* A duplicate internal function meant the local-AI auto-start added in v138 was being silently bypassed in one code path.
* Document text is now read from the first few pages rather than just two, so a file that opens with a full-page image is no longer mistaken for having no text at all.

### Fixed — Core engine (found in a full audit of the main application)
* **Stop no longer locks up the app.** Pressing Stop partway through Ingest, Unique Export, Distribution or CSV Export left the buttons permanently disabled, with a restart the only way out. Every one of those paths now finishes cleanly and says what happened.
* **Exporting a CSV from a job with no manifest** used to do nothing at all — no message, no error, and the app locked. It now explains the problem.
* **`manifest.json` is written atomically.** It's the single source of truth for Unique Export, Distribution and CSV Export; a 37k-file manifest takes about half a second to write, and an interruption inside that window previously corrupted it and made the whole ingest unreadable. The same protection now covers `stats.json` and `status.json`.
* **Refined files can no longer be left half-written.** Flatten, OCR, resize, image-to-PDF and Office sanitize all write to a temporary file and swap it in when complete — a run that's interrupted used to leave a truncated file that every later run treated as finished work.
* **Office sanitize no longer inflates documents.** Files were being re-zipped without compression, which made sanitized `.docx`/`.xlsx` files dramatically larger (over 200x on text-heavy documents in testing).
* **Office metadata with accents or non-Latin characters** is now read and written as UTF-8. Previously the machine's regional encoding was used, which could make sanitizing silently skip a file.
* **The "Max Pixels" setting now works.** It was shown in Settings and saved, but the underlying limit was hardcoded and ignored it.

### UI
* The Rebrand dialog remembers the last source folder, so you don't re-browse to it between Analyze and Apply.
* An "Open folder" button next to the review sheet field.
* Apply stays disabled until a review sheet is actually available, with a tooltip saying why.

## [v139] - 2026-08-03
### Added — Complete output sets
* **"Complete set" option:** Rebranding can now copy every non-PDF file (images, Office docs, spreadsheets, anything else) through to the output folder unchanged, so the `_rebranded` tree is the whole upload set rather than PDFs only. Available as a checkbox in **Rebrand a Folder**, **Rebrand Unique Masters**, and **Process a Folder**; your choice is remembered. Turn it off for the previous PDFs-only behaviour.
* Copies are byte-identical, mirror the source folder structure, and re-runs skip what's already there.
* PDFs found in the source that aren't in the review sheet are now called out in the log instead of quietly missing from the output.

### Changed — Review sheets have a home
* **The review sheet is no longer written into the folder you analyzed**, where it was easy to lose among thousands of subfolders. Sheets now go to **Documents\DocRefinePro_Data\Rebrand Reviews**, named after the job and folder they describe (e.g. `BM-Batch4__01_Master_Files_rebrand_plan.csv`).
* **Apply finds the sheet by itself** — the Rebrand dialog pre-fills it as soon as you pick the source folder. Sheets written by earlier versions (inside the source folder) are still found automatically.
* Re-analyzing a folder keeps your reviewed sheet as `…_rebrand_plan.previous.csv` instead of overwriting it outright.
* Review sheets are never copied into the output, even with "Complete set" on.
* The output folder no longer gets an internal `stats.json` timing file — the delivery tree now contains only your documents. Timings are in the log.

## [v138] - 2026-07-30
### Fixed / Improved — Local AI (Ollama) detection & setup
* **Auto-start:** If Ollama is installed but not running, the app now starts it automatically before analyzing — instead of reporting "not detected" and silently falling back to filename-based titles.
* **Guided setup:** When Ollama isn't installed or the model isn't downloaded, the Analyze step now offers to open the Ollama download page, or to download the model (~2 GB) with a progress bar — rather than quietly giving up.
* **Connection fix:** Connect via `127.0.0.1` instead of `localhost` to avoid a Windows IPv6 resolution issue that could cause a false "connection refused".
* Clearer status messages in the log about whether the local AI is being used.

## [v137] - 2026-07-30
### Added
* **Rebrand Unique Masters:** Rebrand a deduplicated job's unique master files directly from the Export tab (Option D) — no need to hand-navigate to the workspace folder. Point-and-go from any ingested batch.
### Fixed — Rebranding robustness at scale
* **No more lost files on name clashes:** Two documents that resolve to the same output filename are now kept as separate, numbered files instead of one silently overwriting the other. Re-runs stay stable (resume-safe).
* **Parallel image safety:** Branding images are now built per worker, removing a rare risk of corrupted output when rebranding many files at once.
* **Scanned documents default to "leave as-is":** PDFs with no readable text are no longer auto-rebranded (safer for scanned certificates/drawings); they're flagged for review instead.
* Internal cleanup: removed dead code and hardened cross-platform path handling.

## [v136] - 2026-07-30
### Added — Document Rebranding
* **Rebrand a Folder:** Apply branding to a whole folder of PDFs while preserving the original, searchable content. Each document gets a titled cover + back cover, header/footer strips added as page *extensions* (the original page is never overlapped or cropped), and a watermark. Output keeps its text layer, stays well under 50 MB, embeds fonts, and sets PDF metadata (Author = Budget Mailboxes).
* **Smart classification (runs locally):** A local AI model reads each document to decide what to rebrand vs. leave as-is (skipping CAD/technical drawings and certifications) and drafts the product, asset type, manufacturer, and cover title into a review spreadsheet you approve before anything runs. No cloud, no data leaves the machine.
* **Process a Folder (pipeline):** Run your chosen steps — Flatten, Rebrand, OCR — over a folder in a safe order, with OCR always last so scanned files still end up searchable.
### Performance
* Rebranding and the pipeline run multi-threaded, scaled to available memory.

## [v135] - 2026-07-09
### Performance (Large-Batch Hardening)
* **Big PDFs:** Flattening now processes and merges one page at a time instead of holding an entire document in memory, so very large multi-page PDFs no longer risk exhausting RAM during a batch.
* **Inspector:** The file list now filters using batched inserts with sorting suspended, and search input is debounced. Filtering a manifest of tens of thousands of files stays responsive (40k rows rebuild in ~0.5s).

## [v134] - 2026-07-09
### Fixed
* **Audit Certificate:** Fixed a broken template tag that caused report generation to silently fail on every job. The report template is now also bundled into the packaged app so receipts render in installed builds.
* **OCR / PDF Engines:** The app now locates the bundled Tesseract and poppler engines in their actual folders, and the Windows build packages them, so Flatten/OCR/Preview work in the released `.exe` without a separate install.
* **Quarantine Visibility:** Files that fail ingestion are now recorded in the manifest, so they appear in the Inspector and the exported CSV instead of silently disappearing.
### Added
* **Mark Unique:** The Forensic viewer's "Mark Unique" button now works — it promotes a duplicate into its own standalone master file with a new ID.
### Maintenance
* Added missing package files; refreshed README version references; removed a stray test document from the repository.

## [v133.1] - 2026-04-17
### Hotfix
* **CI/CD Fix:** Resolved `ModuleNotFoundError` during build phase by syncing the GitHub pipeline installer process with `requirements.txt`.

## [v133] - 2026-04-17
### Infrastructure & CI/CD
* **Zero-Touch Deployments:** Fully automated GitHub Releases pipeline, permanently replacing the manual Google Drive deployment process.
* **Mac App Optimization:** Successfully removed the dangerous post-build Python stripping script. PyInstaller now natively filters out `QtWebEngine`, `QtQuick`, and `Qt3D` at the binary analysis level, retaining the massive file size reduction without corrupting Apple code signatures.
* **Mac Smoke Testing:** Programmed a `--dry-run` flag into the core engine to enable fully automated, headless sanity verification. GitHub Actions now tests the compiled MacOS binary to ensure it boots without crashing before releasing it.
* **Architectural Fixes:** Re-wrote Python internal configuration and validation using `pydantic`. Integrated `Jinja2` templating for all HTML reporting engines. Eliminated closure-loop memory leaks in PySide6 rendering.

## [v132] - 2026-03-12
### Added
* **Chained Workflow Bypass:** Added a new UI toggle in the Refine tab (`Source from Flattened Cache`) allowing users to execute sequential refinement actions (e.g. OCR) directly on files residing in the output cache rather than forcing a pull from the strict master hub. This fully supports intermediate manual editing steps like third-party visual rebranding.

## [v131] - 2026-03-12
### Fixed
* **Reconstruction Safety:** Re-wrote `get_best_source` in the Worker module to support intelligent ID prefix matching for processed files. The engine now ignores minor filename variations or unexpected double extensions (e.g. `.pdf.pdf`) applied by external rebranding tools.
* **Crash Protection:** Added strict validation checks to prevent `[Errno 2]` fatal failures if a fallback master file has been manually moved or deleted from the workspace.

## [v130] - 2026-03-12
### Fixed
* **Reconstruction Pipeline:** Wired up the "Override Source" checkbox to properly prompt for an external directory. This allows users to seamlessly inject externally modified files (like rebranded PDFs) into the redistribution engine, ignoring exact file extension mismatches as long as the base Unique ID matches.

## [v129] - 2026-01-19
### Maintenance
* **Legacy Cleanup:** Permanently removed the deprecated Tkinter UI module (`docrefine/gui/app.py`).
* **Architecture:** The application is now exclusively Qt/PySide6.

## [v128.5] - 2026-01-16
### UX & Stability
* **Inspector First:** Changed default tab order to prioritize Forensic Inspection.
* **Progress Fix:** Resolved issue where progress bar would hang at 99%.
* **Pause Safety:** Fixed critical bug where pausing the refinement process caused files to be copied without processing.
* **Smart Skip:** The engine now detects existing output files and skips them to save time.
* **Log Hygiene:** App logs now reset on startup to prevent bloat.
* **Timing:** Added detailed breakdown of Ingest vs. Refine times in the dashboard stats.

## [v128.4] - 2026-01-16
### Architecture & Optimization
* **Spec-First Build System:** Migrated from CLI overrides to a pure Python Spec file architecture for consistent cross-platform building.
* **Mac Diet (Surgical):** Reduced macOS app bundle size from **1.3GB to ~200MB** via:
    * **Pre-Build:** Filtering `hiddenimports` to prevent PyInstaller hooks from loading unwanted Qt frameworks (WebEngine, Quick, 3D).
    * **Post-Build:** "Nuclear" stripping script that physically removes any surviving bloat frameworks and cleans up broken symlinks to prevent installer crashes.
* **Stability:** Fixed Windows build crash caused by tuple unpacking errors in the Spec file.
* **Maintenance:** Added `tools/inventory.py` for project auditing.

## [v128] - 2026-01-16
### Architecture
* **Build System Overhaul:** Switched to a "Spec-First" build architecture.
* **Mac Diet:** Implemented aggressive binary filtering at the PyInstaller Spec level to block `QtWebEngine`, `QtQuick`, and `Qt3D` *before* bundling. This targets the 1.3GB bloat issue directly.
* **Inventory Control:** Added `tools/inventory.py` for project auditing.
* **Cleanup:** Removed deprecated CLI overrides from build scripts.

## [v127] - 2026-01-15
### Fixed
- **Mac Build:** Resolved `OSError` in stripping script by handling symlinks correctly.
- **Optimization:** Refined `strip_mac.py` to differentiate between directories (`rmtree`) and symbolic links (`unlink`) during framework cleanup.

## [v126] - 2026-01-15
### Infrastructure
- **Mac Optimization:** Implemented `strip_mac.py` to programmatically remove unused Qt Frameworks (`QtQuick`, `QtQml`, `QtWebEngine`) post-build.
- **CI/CD:** Replaced fragile bash commands with Python scripting for reliable path resolution during the build process.
- **Size Reduction:** Forced removal of PyInstaller-protected frameworks to reduce DMG size from ~1.2GB to target (~400MB).

## [v125] - 2026-01-15
### Infrastructure
- **Mac Optimization:** Implemented manual framework stripping in CI/CD to reduce DMG size.
- **Cleanup:** Removed unused Qt translations and debug symbols from the macOS binary.

## [v124] - 2026-01-15
### Fixed
- **Build Fix**: Resolved `TypeError` in Spec file caused by deprecated `include_pycache` argument in PyInstaller 6.18.

## [v123] - 2026-01-15
### Infrastructure
- **Size Optimization:** Reduced DMG/EXE footprint by ~50% via targeted PySide6 stripping.
- **Compression:** Integrated UPX compression and enabled binary symbol stripping.
- **Architecture:** Moved from `collect_all` to manual dependency mapping to prevent "Universal" binary bloat.

## [v122] - 2026-01-15
### Infrastructure
- **Build Fix:** Added `BUNDLE` block to Spec file for correct macOS `.app` generation.
- **Asset Safety:** Added fallback logic for missing application icons during build.

## [v121] - 2026-01-15
### Maintenance
- **Release:** Fixed Git tag synchronization for CI/CD pipeline.

## [v120] - 2026-01-15
### Infrastructure
- **Spec-Based Build:** Switched from CLI commands to `DocRefinePro.spec` for release builds.
- **Cross-Platform Fix:** Added `collect_all('PySide6')` to the build spec to ensure Mac/Win DLLs are bundled correctly.

## v119 (The Great Refactor)
* [Architecture] **UI Migration:** Complete rewrite of the UI layer from Tkinter to **PySide6 (Qt)**.
    * Modern "Fusion" theme with Dark Mode support.
    * Non-blocking, thread-safe architecture using Signals & Slots.
* [New] **Forensic Comparator 2.0:**
    * Synchronized Zoom & Pan.
    * Dark background for high-contrast inspection.
    * Smooth page scrolling.
* [New] **Active Worker Visualizer:** Real-time grid showing multi-threaded status.
* [Improved] **Timer Logic:** Job timer now respects "Pause" state.
* [Fixed] **Windows Explorer:** Fixed "Reveal in Folder" failing due to subprocess restrictions.

## v118-patch1
* [Fixed] **Context Menu Bug:** Fixed "Reveal in Folder" failure by using unique ID lookup instead of fragile filename matching.

## v118 (Modular Foundation)
* [Architecture] **Modular Restructure:** Application logic split into `gui`, `worker`, `processing`, and `config` modules for improved stability.
* [Improvement] **Update Signal:** Hardcoded verification of v118.
* [Docs] Updated bundled documentation paths.

## v117 (Hotfixes)
* [Fixed] **Ingest Crash:** Resolved regression in `run_inventory` arguments.
* [Fixed] **Debug Export:** Fixed threading violation in export tool.
* [Refactor] **Thread Safety:** Hardened UI/Worker separation.

## v116
* [Added] In-App Documentation Viewer.
* [Improved] Intelligent Resource Loader.

## v115
* [Fixed] Mac UI freeze (queue throttling).
* [Fixed] Windows "Ghost Windows" patch.