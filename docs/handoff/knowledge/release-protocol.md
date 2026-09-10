# Release protocol — the ordered checklist, and the automatic ClickUp subtask

> **Provenance.** This began as one of Jason Diaz's local Claude Code memory files
> (`~/.claude/projects/.../memory/docrefine-release-protocol.md`) and was moved into the repo on 2026-09-10 as part of
> his handoff. It is preserved close to verbatim — the numbers in here were measured, often
> expensively, and paraphrasing them would lose the point. Dates are absolute. Treat it as a
> record of what was true when written, not a guarantee about today's code: **verify any file,
> function or flag it names still exists before acting on it.**

---

Every DocRefine Pro release must also update the **ClickUp project task "DocRefine Pro"** — task id `86ex00r23`, list `901815817512` ("Project Managment", folder "TA - Operations"). User asked for this on 2026-08-04 after the tracker had drifted four versions behind (last entry was v136 while v140 shipped). See [project setup](project-setup.md).

**Why:** the task is how this project is reported upward; a release that isn't on it effectively didn't happen.

**How to apply — release steps, in order:**
1. Bump `SystemUtils.CURRENT_VERSION` in `docrefine/config.py`, update `CHANGELOG.md` and the `vNNN` strings in `README.md`.
2. Run the full verify suite from the session scratchpad (see [the rebrand playbook](rebrand-playbook.md) for the list) — must be all green, including the 18/18 regression.
3. Commit, `git tag -a vNNN`, push main **and** the tag.
4. Confirm CI (both Windows and macOS jobs) and that the release has both assets.
5. **ClickUp is automatic — verified working as of v147 (2026-08-05).** The `notify-clickup` job in `.github/workflows/build.yml` creates a **subtask** under `86ex00r23` once both builds pass. The `CLICKUP_API_TOKEN` repo secret **is set** (added 2026-08-04). Do not post manually — you will duplicate it.
   - It names the subtask from the changelog's first `###` heading: `<outcome> (vNNN)`. So **the changelog heading you write becomes the ClickUp task name** — write it as an outcome, not "Added — ...".
   - It only ever creates, never overwrites, and skips if a subtask already ends in `(vNNN)`.
   - **To check whether it worked, look at SUBTASKS, not comments.** It posts no comments; `clickup_get_task_comments` on `86ex00r23` returns 0 and always will. (An earlier version of this workflow posted comments — commit `f081c65` changed it to subtasks.)
   - The job reports its outcome to the run summary, so a green job that did nothing is visible rather than silent.

**Suite now lives at `Documents\DocRefinePro_Data\verify\`** — moved there 2026-08-14 because it had been sitting in `AppData\Local\Temp`, which Windows Disk Cleanup and Storage Sense delete without asking. `run_suite.py` sits beside the scripts, finds them by `Path(__file__).parent`, tallies every `verify_*.py` and takes `--only=name,name` to filter. Run it as:

    .venv\Scripts\python.exe Documents\DocRefinePro_Data\verify\run_suite.py <repo> "<...>\BrandKit\Sample Files" Documents\DocRefinePro_Data\verify\_p0out

**563/563 across 23 scripts as of v157**, verified from the new location. The per-batch helper scripts (deliver_qa, pixel_truth, trial_run/trial_qa, deep_scan, analyze_vision, apply_batch4, compare_sheets, downsample_flw, swap_flw, restore_brand_json) are alongside it. Two of them — `verify_phase2` and `verify_vision_live` — drive **real Ollama**, so never run the full suite while a vision analyze is in flight: both models resident on an 8GB card drops throughput from 3.5s to 14.3s a file and adds hours to the run. Run them after.

**`verify_boot.py` no longer hard-codes the version** (as of 2026-08-13). It reads the expected value from CHANGELOG's newest `## [vNNN]` heading and asserts `CURRENT_VERSION` and README agree — same guarantee that all three are bumped, minus the literal that had to be hand-edited each release.
