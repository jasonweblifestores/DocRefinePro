# SAVE AS: docrefine/reviews.py
"""
Where rebrand review sheets live.

Up to v138 the sheet was written as `_rebrand_plan.csv` *inside* the analyzed
source folder, where it was easy to lose among thousands of subfolders — and
where a "complete set" copy would sweep it into the upload tree.

From v139 every sheet goes to one fixed, obvious place:

    Documents/DocRefinePro_Data/Rebrand Reviews/

named after the source it describes, so Apply can find it again on its own.
The name carries the parent folder too, because deduplicated jobs all call their
masters folder `01_Master_Files` — the job name is what tells them apart.

Sheets written by older versions are still found (see `plan_candidates`).
"""
import re
from pathlib import Path

from .config import REVIEWS_ROOT

PLAN_SUFFIX = "_rebrand_plan.csv"
LEGACY_PLAN_NAME = "_rebrand_plan.csv"


def is_plan_file(name):
    """True for any review sheet, wherever it was written."""
    n = str(name).lower()
    return n == LEGACY_PLAN_NAME or n.endswith(PLAN_SUFFIX)


def _safe(part):
    """Make one path component safe (and short enough) for a filename."""
    cleaned = re.sub(r'[<>:"/\\|?*]+', "-", (part or "").strip())
    return cleaned.strip(". ")[:60] or "folder"


def plan_name_for(src):
    """Stable sheet name for a source folder: `<parent>__<folder>_rebrand_plan.csv`."""
    try:
        p = Path(src).resolve()
    except OSError:
        p = Path(src)
    parent = p.parent.name
    stem = f"{_safe(parent)}__{_safe(p.name)}" if parent else _safe(p.name)
    return stem + PLAN_SUFFIX


def plan_path_for(src):
    """Where to WRITE the review sheet for a source folder (creates the folder)."""
    REVIEWS_ROOT.mkdir(parents=True, exist_ok=True)
    return REVIEWS_ROOT / plan_name_for(src)


def plan_candidates(src):
    """Every place a sheet for this source may live, best first."""
    p = Path(src)
    name = plan_name_for(src)
    return [
        REVIEWS_ROOT / name,                    # v139 default
        p.parent / name,                        # hand-placed beside the source
        p.parent / f"{p.name}{PLAN_SUFFIX}",    # hand-placed, short name
        p / LEGACY_PLAN_NAME,                   # pre-v139: inside the source
    ]


def find_plan(src):
    """Locate an existing review sheet for a source folder, or None."""
    for c in plan_candidates(src):
        try:
            if c.is_file():
                return c
        except OSError:
            continue
    return None
