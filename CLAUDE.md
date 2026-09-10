# CLAUDE.md — orientation for Claude Code

Read this first, then **[docs/handoff/README.md](docs/handoff/README.md)**.

## Context you need before doing anything

DocRefine Pro was built and operated by one person, Jason Diaz, who left WebLife in October 2026. The person
running you now is probably **not** the person who built this. Do not assume they know the history — the
history is written down in `docs/handoff/`, and it is your job to read it rather than to ask them.

`AGENTS.md` in this directory predates the handover. Its instructions ("the user is the Product Manager…has
ZERO coding knowledge") described Jason specifically. **Treat its architecture rules as current** — PySide6,
Signals/Slots for thread communication, no local builds — and treat its assumptions about who the user is as
possibly stale. Ask rather than guess.

## Where things are

| Path | What |
|---|---|
| `docrefine/` | The engine. `worker.py`, `processing.py`, `rebrand.py`, `classify.py`, `stamps.py`, `reviews.py`, `runs.py`, `config.py`, plus `gui/`. |
| `docs/handoff/` | **The handoff pack.** Start at its README. |
| `verify/` | 59 verification scripts. 563/563 across 23 `verify_*.py` as of v157, plus per-batch QA helpers. |
| `brandkits/` | The two delivered brand kits' `brand.json` decisions. Irreplaceable — see below. |
| `poppler/`, `Tesseract-OCR/` | Bundled binaries. `pdftotext` is at `poppler/Library/bin/pdftotext.exe`. |
| `CHANGELOG.md` | Version history. The newest `## [vNNN]` heading is authoritative — `verify_boot.py` asserts `config.py` and `README.md` agree with it. |

Runtime data lives **outside** the repo in `Documents\DocRefinePro_Data\` (workspaces, review sheets, logs,
run history). That is by design, but it means it is not backed up by git — which is how it nearly got lost.

## Rules that will save you

1. **QA the real output, always.** Every substantive defect in this project's history was found by inspecting
   an actual delivery, never by the unit tests, which were green throughout. A file that silently lost all
   3,127 characters of its content. 255 files that would have shipped with numbered junk filenames. A corrupt
   source we branded and shipped. Run the suite, then look at the deliverable.
2. **The brand kit loader wants the literal filename `brand.json`.** Any other name is ignored and the kit
   falls back to defaults naming *Budget Mailboxes* — no aliases, no tagline, no disclaimer — while appearing
   to work. If a run prints raw, inconsistent manufacturer names, check this first.
3. **Wording is data.** Taglines, disclaimers, version labels and manufacturer aliases belong in `brand.json`,
   never in code. A blank field prints nothing and the run log names the gap.
4. **Releases follow a fixed order** and ClickUp updates itself. See
   [docs/handoff/knowledge/release-protocol.md](docs/handoff/knowledge/release-protocol.md). Never post a
   release to ClickUp by hand — CI creates a subtask under `86ex00r23`, and posting manually duplicates it.
   To check whether it worked, look at **subtasks**, not comments.
5. **Never run the full verify suite while a vision analyze is in flight.** Two scripts drive real Ollama;
   two models resident on an 8GB card drops throughput from 3.5s to 14.3s per file and adds hours to a run.
6. **The app does not edit vendor page content.** It extends pages and overlays our branding. Any request to
   change what a document *says* is a different operation — see
   [docs/handoff/02-sustainable-hearth.md](docs/handoff/02-sustainable-hearth.md) for why that distinction
   matters and what it cost to establish.

## Running it

Python 3.10.11. The `python` and `py` commands on PATH may be Microsoft Store stubs — do not use them; call
the interpreter by full path. Set up a venv, install `requirements.txt`, then:

```bash
QT_QPA_PLATFORM=offscreen .venv/Scripts/python.exe main.py --dry-run
```

Full setup detail in [docs/handoff/knowledge/project-setup.md](docs/handoff/knowledge/project-setup.md).

## Known landmines

`docs/handoff/knowledge/known-issues.md` and `clickup-api-quirks.md` exist specifically because these things
look like bugs and are not. Read them before "fixing" something that looks broken — particularly before
reposting a ClickUp comment that appears mangled, which double-notifies real people to repair a defect that
does not exist.
