"""Progress that survives an interrupted classification run.

A visual-pass analyze over a few thousand files runs for hours — the Batch 4 set
took 10.35 — and until now the review sheet was only written once *both* passes
had finished. A stop, a crash, or a Windows update restart at hour nine threw
away everything, including the completed text pass. On a corpus four times the
size that is not a risk worth carrying.

The file is **append-only JSONL**, written as each file is classified, so a hard
power loss can cost at most the last line rather than the run. Reading tolerates
a truncated final record, because that is exactly what a power loss leaves
behind. Later records for the same file win, which is how the visual pass
supersedes what the text pass decided.

A record is only reused when the source file is unchanged (same size, same
mtime) — an edited or replaced document is classified again rather than
resurrected from a stale answer.
"""
import json
import os
from pathlib import Path

SUFFIX = ".progress.jsonl"


def path_for(plan_path):
    """The progress file that belongs beside a given review sheet."""
    p = Path(plan_path)
    return p.with_name(p.stem + SUFFIX)


def signature(pdf):
    """Cheap identity for a source file: size and modification time.

    Deliberately not a hash — hashing thousands of PDFs to decide whether to
    skip them would cost a slice of the time this is meant to save.
    """
    try:
        st = os.stat(pdf)
        return int(st.st_size), int(st.st_mtime)
    except OSError:
        return None, None


def load(path):
    """Every usable record, keyed by relative path, last one winning.

    Never raises: a missing, empty, partially-written or corrupt file simply
    means there is no progress to resume from.
    """
    out = {}
    p = Path(path)
    if not p.is_file():
        return out
    try:
        with open(p, "r", encoding="utf-8") as fh:
            for line in fh:
                line = line.strip()
                if not line:
                    continue
                try:
                    rec = json.loads(line)
                except ValueError:
                    continue          # truncated last line after a power loss
                f = rec.get("file")
                if isinstance(f, str) and isinstance(rec.get("row"), dict):
                    out[f] = rec
    except OSError:
        return {}
    return out


def usable(rec, pdf):
    """Is this record still true of the file on disk?"""
    if not rec:
        return False
    size, mtime = signature(pdf)
    return size is not None and rec.get("size") == size and rec.get("mtime") == mtime


class Writer:
    """Appends records as they are produced. Opened lazily, closed politely."""

    def __init__(self, path):
        self.path = Path(path)
        self._fh = None

    def _open(self):
        if self._fh is None:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            self._fh = open(self.path, "a", encoding="utf-8")
        return self._fh

    def add(self, rel, pdf, row, needs_look=False, vision_done=False):
        """Record one classified file. Flushed immediately — an unflushed
        buffer is exactly what a power cut discards."""
        size, mtime = signature(pdf)
        rec = {"file": rel, "size": size, "mtime": mtime,
               "needs_look": bool(needs_look), "vision_done": bool(vision_done),
               "row": row}
        try:
            fh = self._open()
            fh.write(json.dumps(rec, ensure_ascii=False) + "\n")
            fh.flush()
        except Exception:
            pass          # progress-keeping must never break the run itself

    def close(self):
        try:
            if self._fh is not None:
                self._fh.close()
        except Exception:
            pass
        finally:
            self._fh = None


def clear(path):
    """Drop the progress file once its sheet has been written."""
    try:
        Path(path).unlink()
    except OSError:
        pass
