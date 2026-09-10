# Known issues — bundled-binary quirks, quarantine model, template bundling

> **Provenance.** This began as one of Jason Diaz's local Claude Code memory files
> (`~/.claude/projects/.../memory/docrefine-pro-known-issues.md`) and was moved into the repo on 2026-09-10 as part of
> his handoff. It is preserved close to verbatim — the numbers in here were measured, often
> expensively, and paraphrasing them would lose the point. Dates are absolute. Treat it as a
> record of what was true when written, not a guarantee about today's code: **verify any file,
> function or flag it names still exists before acting on it.**

---

Findings from the DocRefine Pro repo scan on 2026-07-09 (all fixed in commit 7fc7858 except where noted):

- The bundled `Tesseract-OCR/` (134 files) and `poppler/` (455 files) folders are committed to the repo (~144MB of the total) and are Windows binaries. `config.find_binary()` now searches `Tesseract-OCR/` and `poppler/Library/bin/`; the PyInstaller spec bundles them on Windows only (Mac uses Homebrew). tessdata ships `eng` + `osd` only.
- The `.spec` did NOT bundle `docrefine/templates/` — now added, else the Jinja Audit Certificate report can't render in frozen builds.
- Quarantine data model: `run_inventory` now writes `status: "QUARANTINE"` entries into `manifest.json` (id `Q0001`…). They have no `uid`; GUI open/compare handlers guard against missing `uid`.
Large-batch hardening (v135, released 2026-07-09) for a 58.6GB / 37k-file ingest:
- Flatten (`PdfProcessor.flatten_or_ocr`) now writes one single-page PDF per page and merges with pypdf incrementally — peak memory bounded to one page (was: loaded all pages into RAM for the Pillow multipage save; OOM risk on a 15.6GB machine with 2 workers).
- Inspector filter (`main_window.filter_inspector`) uses batched `addTopLevelItems` with sorting/updates suspended + a 250ms debounce on the search box (40k rows rebuild ~0.5s).
- Disk note: every pipeline stage keeps a full copy under `Documents\DocRefinePro_Data` on C:; full ingest→refine→organize→distribute can need ~4x the source size. User freed space to >500GB on 2026-07-09.

- Not addressed: README download filenames are illustrative; the app relies on `sys._MEIPASS` layout when frozen. Latest released version: v135.

**BEFORE ANY LONG UNATTENDED RUN ON THIS MACHINE (learned the hard way 2026-08-13):** check for a queued Windows restart *first* — `HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending` and `…\WindowsUpdate\Auto Update\RebootRequired`. A 12-hour vision analyze was launched with **both flags already set** and active hours of 08:00–02:00, so Windows was free to force a restart at 02:00 — mid-run. Fix is to pause updates (Settings → Windows Update → Pause), not to reboot, which would have cost the run. Also: `powercfg /change standby-timeout-ac 0` does NOT cover **lid close** (the setting is hidden on this machine) and only covers **AC** — the DC sleep timeout is still 240 min, so unplugging reintroduces the problem. Display/monitor timeout is harmless, leave it alone.
- Compounding this: `worker.run_rebrand_analyze` writes its sheet **only after both passes finish**, and a stop returns without writing anything — so any interruption discards the entire run, including the completed text pass. Checkpoint/resume is planned as v157. Until then, an interrupted long analyze means falling back to the previous sheet (kept as `*.previous.xlsx`).
- Note the timezone trap when reading GitHub API timestamps here: they are UTC and this machine is **+05:30**. Misreading one produced an ETA five and a half hours wrong.

See [project setup](project-setup.md).
