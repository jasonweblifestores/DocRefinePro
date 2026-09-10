# The vision pass — local vision classification, hardware limits, accuracy

> **Provenance.** This began as one of Jason Diaz's local Claude Code memory files
> (`~/.claude/projects/.../memory/docrefine-vision-pass.md`) and was moved into the repo on 2026-09-10 as part of
> his handoff. It is preserved close to verbatim — the numbers in here were measured, often
> expensively, and paraphrasing them would lose the point. Dates are absolute. Treat it as a
> record of what was true when written, not a guarantee about today's code: **verify any file,
> function or flag it names still exists before acting on it.**

---

Built 2026-08-12 (v147-v154 era). Roughly **half of a typical corpus has no extractable text** — on Batch 4, 1,043 of 2,174 files — so a text-only model classifies them from the filename alone. `qwen2.5vl:7b` via Ollama now *looks* at those pages. See [the rebrand playbook](rebrand-playbook.md).

**Architecture — two sequential passes, never both models at once.** Pass 1 runs the text model over every file and queues the ones it can't settle (`classify.needs_a_look`: no text, or confidence < 0.9). Then the text model is **unloaded**, the vision model runs the queue, and it is unloaded at the end. Where vision is consulted **its answer wins** over the filename and page-shape heuristics — those exist only to compensate for not being able to see.

**Why sequential, measured with `ollama ps` on an 8GB RTX 5060:**
- Both models resident want 8.5GB on an 8GB card → the vision model is pushed onto the CPU → **14.3s/file** (spikes to 25s).
- Text model unloaded first → **3.5s/file**. Four times faster, no accuracy cost.
- Sustained rate on this machine settles around **~22-28s/file** regardless of `num_ctx` (2048/4096/8192 all measured the same) — an early 3.5s reading was an outlier, don't quote it as an estimate.
- **RE-MEASURED 2026-08-12 on the 14 ground-truth files: 28.2s/file** (first file 38.7s for the model load, warm files 25-31s). So budget **~12 hours for 1,512 files, not 9.** Quote the range, and note the app reports its own measured ETA anyway.

**Two real bugs found and fixed while building it:**
- Large-format drawings (44x34in) rendered to 4840x3740 at fixed DPI — the model **refused them outright**, so the very files this exists to identify were the ones it couldn't see. Fixed by computing DPI from page size to hit `VISION_MAX_PX` directly: render time **38s → 0.5s**.
- Nothing was released after a run: Ollama held 6GB for its keep-alive window while the user had moved on. `classify.unload_model()` now releases explicitly (it's **asynchronous** — don't assert `loaded_models() == []` immediately, that races).

**Hardware reporting is measured, never assumed** — nothing is tuned to this laptop. `classify.model_placement()` reads Ollama's `size` vs `size_vram` and the run log says what *that* machine is doing ("only 73% of qwen2.5vl:7b fits in this machine's GPU memory… the rest runs on the CPU"), suggests `granite3.2-vision:2b` when it doesn't fit, and reports a **measured** rate + ETA from *after* the first file (the first pays the model load and made the estimate several times too pessimistic).

**Accuracy: 14/14** on hand-labelled ground truth (`verify_vision_live.py`) — all seven CAD cut sheets, the three drawings the filename rule missed, and `1590-T1V-Spec-Sheet.pdf`, which is *named* like a spec sheet but is a drawing. Its stated evidence was right every time ("measurement arrows and dimension labels" vs "numbered steps and exploded diagram").

Opt-in: `CFG.rebrand_vision_pass` (default off) + `rebrand_vision_model`; checkbox in RebrandDialog with a speed warning; Analyze offers to download the model if missing.

**RELEASED as v155 on 2026-08-12** (commit 968bb21, tag v155, CI green Win+Mac, both assets, ClickUp subtask `86eym3evz` auto-created). It had been sitting uncommitted in the working tree. Verify suite at release: **468/468 across 19 scripts** (443 + verify_vision's 25), plus verify_vision_live 14/14 live against the ground truth.

**GPU PLACEMENT IS THE WHOLE BALLGAME — measured 2026-08-18 on MBW.** Jason switched the
laptop from Optimus (hybrid iGPU+dGPU) to **dGPU-only**, and `qwen2.5vl:7b` now runs
**100% on the GPU** (`ollama ps`: 5.37GB of 5.37GB in VRAM; run log says "qwen2.5vl:7b is
running entirely on the GPU"). Vision throughput: **~3.6s/file sustained** over 456 files,
against Batch 4's **24s/file** full-run average when only 73% fitted and the rest ran on CPU.
The whole MBW analyze — 970 files, 456 of them seen — took **44 minutes**, not the 4-5h
budgeted. Batch 4's 10.35h analyze would be ~1.5h on this configuration.
- **Check placement BEFORE any long run.** The log line `"running entirely on the GPU"` versus
  `"only X% of ... fits"` predicts whether you are in for 45 minutes or 10 hours. This is the
  single highest-leverage pre-flight check; the 12-hour estimates in this file assumed a CPU spill.
- **The mechanism is NOT more free VRAM** — counter-intuitively dGPU-only gives *less*, because
  the whole Windows shell moved onto the dGPU (`explorer.exe`, SearchHost, Brave, ClickUp, Claude,
  Roam all hold VRAM; `display_active: Enabled`). Measured **8151 MiB total, 7497 used, 403 free**.
  Likely Ollama's layer-fit heuristic reads hybrid-mode driver figures differently, but that is
  UNVERIFIED — it needs a hybrid-mode A/B with `ollama ps` on both sides.
- **403 MiB spare is tight.** If a browser grabs a few hundred MB mid-run, layers spill back to
  CPU and the rate silently collapses to 24s/file. Close heavy GPU consumers before long runs.
- Page count is a secondary factor at most: `VISION_MAX_PAGES = 2` and `VISION_MAX_PX = 1200`
  cap the spread at ~2x, so it cannot explain a 6x gap.

**BUG: `VISION_NUM_CTX = 2048` SILENTLY DROPS FILES (found 2026-08-18, NOT yet fixed).** Some
renders exceed the context, Ollama returns **HTTP 400 `exceed_context_size_error`**,
`classify_visually`'s bare `except Exception: return None` swallows it, and the file falls back to
the **filename heuristic at confidence 0** — the visual pass quietly not applying to a file it
exists for. Symptom: a ~0.7s "answer" (far too fast for inference) and a `source=fallback` row.
- Hit **2 of 456** MBW files; for one the filename guess was WRONG (`mailbox_post_matrix-0520-3.pdf`
  marked `leave`, but at `num_ctx=4096` vision reads it as a spec sheet to rebrand at 0.95). It
  would have shipped unbranded. Batch 4 also had exactly 2 fallback rows — same cause, unexamined.
- **v158 fix:** retry once with a context sized from `n_prompt_tokens` (or just raise to 4096 —
  this file already records 2048/4096/8192 measuring the same speed), and LOG the rejection
  instead of swallowing it. A silent failure here defeats the whole feature.
