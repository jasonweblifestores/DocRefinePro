# SAVE AS: docrefine/runs.py
"""
A record of every rebrand, analyze and pipeline run.

The dashboard's job list is built from ingest workspaces on disk. Rebranding and
the pipeline don't create a workspace — they read an arbitrary folder and write a
sibling output folder — so until now that work left no trace in the app at all:
you could not tell from the dashboard whether a folder had been rebranded, when,
with which brand kit, or with which settings.

That last part is what makes this matter rather than merely tidy. Since v143 the
output depends on toggles (attribution, filenames, page stamps), so "which
settings produced this folder" is a question with a real answer that was
previously unrecoverable. Every run now appends a line here, and the dashboard
shows them under the job list.

The file is JSON Lines — append a line per run, so a crash mid-write costs at
most the newest entry and never the history.
"""
import json
import os
import threading
from datetime import datetime
from pathlib import Path

from .config import USER_DIR, SystemUtils

RUNS_PATH = USER_DIR / "rebrand_runs.jsonl"
MAX_RECORDS = 500          # trimmed oldest-first; this is a history, not an archive

_LOCK = threading.Lock()

# How each setting reads in the dashboard. Only states worth naming appear — the
# defaults stay quiet so the ones that changed the output stand out.
_SETTING_LABELS = (
    ("keep_original_names", True, "original names"),
    ("complete_set", True, "complete set"),
    ("show_attribution", True, "cover attribution"),
    ("footer_attribution", True, "footer attribution"),
    ("stamp_tagline", True, "tagline"),
    ("stamp_version", True, "version/updated"),
    ("stamp_disclaimer", True, "disclaimer"),
)


def record(kind, source, output=None, sheet=None, kit=None, brand=None,
           counts=None, seconds=None, settings=None):
    """Append one run to the history. Never raises — a failed write must not fail a run."""
    entry = {
        # Milliseconds, because the timestamp is also this run's identity — "Forget"
        # looks it up by ts, and two runs recorded in the same second would then be
        # indistinguishable and both removed.
        "ts": datetime.now().isoformat(timespec="milliseconds"),
        "version": SystemUtils.CURRENT_VERSION,
        "kind": kind,
        "source": str(source or ""),
        "output": str(output or ""),
        "sheet": str(sheet or ""),
        "kit": str(kit or ""),
        "brand": str(brand or ""),
        "counts": dict(counts or {}),
        "seconds": round(float(seconds), 1) if seconds is not None else None,
        "settings": dict(settings or {}),
    }
    try:
        with _LOCK:
            RUNS_PATH.parent.mkdir(parents=True, exist_ok=True)
            with open(RUNS_PATH, "a", encoding="utf-8") as f:
                f.write(json.dumps(entry) + "\n")
            _trim()
    except Exception:
        pass
    return entry


def _trim():
    """Keep the newest MAX_RECORDS lines. Called with the lock held."""
    try:
        lines = RUNS_PATH.read_text(encoding="utf-8").splitlines()
    except OSError:
        return
    if len(lines) <= MAX_RECORDS:
        return
    keep = lines[-MAX_RECORDS:]
    tmp = RUNS_PATH.with_suffix(".tmp")
    tmp.write_text("\n".join(keep) + "\n", encoding="utf-8")
    tmp.replace(RUNS_PATH)


def load(limit=None):
    """Every recorded run, newest first. A corrupt line is skipped, not fatal."""
    out = []
    try:
        with open(RUNS_PATH, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                try:
                    rec = json.loads(line)
                except ValueError:
                    continue
                if isinstance(rec, dict):
                    out.append(rec)
    except OSError:
        return []
    out.reverse()
    return out[:limit] if limit else out


def remove(ts):
    """Drop the run with this timestamp from the history — one run, never several."""
    try:
        with _LOCK:
            lines = RUNS_PATH.read_text(encoding="utf-8").splitlines()
            kept, dropped = [], False
            for line in lines:
                try:
                    if not dropped and json.loads(line).get("ts") == ts:
                        dropped = True
                        continue
                except ValueError:
                    pass
                kept.append(line)
            RUNS_PATH.write_text(("\n".join(kept) + "\n") if kept else "", encoding="utf-8")
        return True
    except OSError:
        return False


# --------------------------------------------------------------------------
#  Presentation helpers (used by the dashboard, kept here with the data)
# --------------------------------------------------------------------------

def label(rec):
    """What the run was pointed at, named so two runs can be told apart.

    Includes the parent folder because every deduplicated job's masters folder is
    called `01_Master_Files` — on its own that name identifies nothing.
    """
    raw = str(rec.get("source") or "")
    src = Path(raw.replace("\\", "/")) if "\\" in raw else Path(raw)
    if not src.name:
        return "(unknown folder)"
    parent = src.parent.name
    return f"{parent}{os.sep}{src.name}" if parent else src.name


def result_text(rec):
    """The outcome in a few words, e.g. "874 branded · 1,300 copied"."""
    c = rec.get("counts") or {}
    order = (("rebranded", "branded"), ("left", "left as-is"), ("copied", "copied"),
             ("analyzed", "analyzed"), ("to_rebrand", "to rebrand"),
             ("processed", "processed"), ("skipped", "skipped"), ("failed", "failed"))
    parts = [f"{c[key]:,} {word}" for key, word in order if c.get(key)]
    if not parts:
        return "no files"
    return " · ".join(parts)


def settings_text(rec):
    """The toggles that shaped this output, as a short readable list."""
    s = rec.get("settings") or {}
    on = [name for key, want, name in _SETTING_LABELS if s.get(key) == want]
    return ", ".join(on) if on else "defaults"


def when_text(rec):
    """The run date, formatted for the dashboard column."""
    try:
        return datetime.fromisoformat(rec["ts"]).strftime("%Y-%m-%d %H:%M")
    except (KeyError, ValueError):
        return str(rec.get("ts") or "")


def duration_text(rec):
    secs = rec.get("seconds")
    if not secs:
        return "-"
    if secs < 90:
        return f"{secs:.0f} sec"
    return f"{secs / 60:.1f} min"
