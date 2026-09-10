# Project setup — what the app is, where things live, how to run and verify

> **Provenance.** This began as one of Jason Diaz's local Claude Code memory files
> (`~/.claude/projects/.../memory/docrefine-pro-setup.md`) and was moved into the repo on 2026-09-10 as part of
> his handoff. It is preserved close to verbatim — the numbers in here were measured, often
> expensively, and paraphrasing them would lose the point. Dates are absolute. Treat it as a
> record of what was true when written, not a guarantee about today's code: **verify any file,
> function or flag it names still exists before acting on it.**

---

DocRefine Pro is a local document-processing desktop app (Python + PySide6/Qt): dedup, PDF flatten/OCR, Office sanitize, export. Engine in `docrefine/` (worker/processing/reporting/config + gui/). Version tracked in `docrefine/config.py` (`CURRENT_VERSION`) and `CHANGELOG.md`.

- GitHub repo: https://github.com/jasonweblifestores/DocRefinePro.git (public; `main`). Committed with local git identity Jason Diaz <jason@weblifestores.com>.
- The clone was restored after a local data loss on 2026-07-09 to `C:\Users\WORK\Documents\WebLife Labs\PROJECTS\DocRefine Pro\DocRefinePro` (shallow, `--depth 1`, no full history/tags).
- No system Python originally. Installed Python 3.10.11 via winget to `C:\Users\WORK\AppData\Local\Programs\Python\Python310\python.exe` (the `python`/`py` commands on PATH are Microsoft Store stubs — do not use them).
- Verification venv at `<project>\.venv` (git-ignored) with `requirements.txt` installed. Headless GUI smoke test: `.venv\Scripts\python.exe main.py --dry-run` with `QT_QPA_PLATFORM=offscreen`.
- Builds run in GitHub Actions on `v*` tag push (Win + Mac); Mac job also runs the `--dry-run` smoke test. See [known issues](known-issues.md).
