"""Which batch the per-batch helpers act on — one place instead of six copies.

Every helper used to carry its own `SRC = .../Batch 4/...` triple, so a new
batch meant hunting the same three literals through seven files, and the Batch 4
kit going missing from disk broke all of them at once. Select a batch with the
DRP_BATCH environment variable:

    set DRP_BATCH=batch4     (or mbw, the default)

Wording (tagline/disclaimer) is read from the KIT, never hard-coded, because a
literal from one brand silently passes on another: `all("" in b)` is True for
every page, so a blanked tagline check verifies nothing while looking green.
Use `TAGLINE_EXPECTED` to branch — it says whether this kit stamps one at all.
"""
import os
import sys
from pathlib import Path

DOCS = Path(r"C:\Users\WORK\Documents")
REPO = DOCS / "WebLife Labs" / "PROJECTS" / "DocRefine Pro" / "DocRefinePro"

# Importing this module is enough to reach `docrefine`. The helpers used to do
# `sys.path.insert(0, ".")`, which only worked when run from the repo directory.
if str(REPO) not in sys.path:
    sys.path.insert(0, str(REPO))

BATCHES = {
    "batch4": {
        "root": DOCS / "Batch 4",
        "src":  DOCS / "Batch 4" / "_unique-to-rebrand",
        "kit":  DOCS / "Batch 4" / "Template",
        "out":  DOCS / "Batch 4" / "_unique-to-rebrand_rebranded",
        # Files that caused real defects on this batch — always inspected in full.
        "focus": [
            "Imperial-Street-Sign-Brochure.pdf",             # v145 negative page box
            "metropolis-style-wall-mount-instructions.pdf",   # origin offset
            "3001-Anchor-Bolt-Specifications.pdf",            # origin offset
            "install_keystone_post_cuff.pdf",                 # CropBox < MediaBox
            "30815-lock-change.pdf",                          # encrypted
            "4c11d-10-sm_cutsheet_pdp_.pdf",                  # mixed size + rotation
            "1570-cbu-installation-manual.pdf",               # AcroForm + long manual
            "1570-12-BM.pdf",                                 # bookmarks
            "1570-12v-cut-sheet.pdf",                         # links/annotations
        ],
    },
    "mbw": {
        "root": DOCS / "MBW rebrand work",
        "src":  DOCS / "MBW rebrand work" / "_unique-to-rebrand",
        # The kit stays with the source download; only the working tree is ours.
        "kit":  DOCS / "MBW Downloadable re-branding" / "_rebrand-templates",
        "out":  DOCS / "MBW rebrand work" / "_unique-to-rebrand_rebranded",
        # These are the POST-rename names. Both changed when the most-referenced
        # survivor rule was applied (Kunchana item 6); the pre-rename names are no
        # longer in the sheet or the staged folder, so a stale entry here silently
        # drops the file from the trial's known-awkward picks.
        "focus": [
            # The document reached under 43 different filenames — the widest
            # alias collapse in the batch, so the one most worth eyeballing.
            "5076-Hinged-Locking-Rear-Horizontal-All-Florence-CutSheet-1.pdf",
            # The most duplicated by content (420 copies across 4 names).
            "florence-powdercoat-referencechart-brochure-color-options.pdf",
            # Re-classified after the num_ctx bug dropped it to a filename guess;
            # kept its name because its group was an exact reference tie.
            "mailbox_post_matrix-0520-3.pdf",
        ],
    },
}

BATCH = (os.environ.get("DRP_BATCH") or "mbw").strip().lower()
if BATCH not in BATCHES:
    sys.exit(f"DRP_BATCH={BATCH!r} is not one of {sorted(BATCHES)}")

_B = BATCHES[BATCH]
ROOT, SRC, KIT, OUT = _B["root"], _B["src"], _B["kit"], _B["out"]
FOCUS = list(_B["focus"])


def plan(ext=".xlsx"):
    """The review sheet for this batch, via the app's own naming rule."""
    from docrefine import reviews
    return reviews.plan_path_for(SRC, ext)


def wording(kit_dir=None):
    """(tagline, tagline_expected, disclaimer) for this batch, from the kit itself.

    tagline_expected is False when the kit supplies no tagline — MailboxWorks has
    none — and callers must then assert the tagline is ABSENT rather than search
    for an empty string, which matches every page.
    """
    from docrefine.rebrand import BrandKit
    k = BrandKit(kit_dir or KIT)
    tag = str(k.brand.get("tagline") or "").strip()
    disc = str(k.brand.get("disclaimer") or "").strip()
    return tag, bool(tag), disc


def require_kit():
    """Fail loudly if the kit is not on disk — it has gone missing before."""
    if not KIT.is_dir():
        sys.exit(f"brand kit missing: {KIT}\n"
                 f"Restore it (artwork + brand.json) before running this.")
    if not (KIT / "brand.json").is_file():
        sys.exit(f"{KIT} has no brand.json — BrandKit would silently fall back to "
                 f"the Budget Mailboxes defaults (no aliases, wrong brand name).")
    return KIT


def summary():
    return (f"batch={BATCH}\n  src {SRC}\n  kit {KIT}\n  out {OUT}")


if __name__ == "__main__":
    print(summary())
    print(f"  plan {plan()}")
    print(f"  kit ok: {require_kit()}")
    t, exp, d = wording()
    print(f"  tagline={t!r} expected={exp}")
    print(f"  disclaimer={d[:50]!r}...")
    print(f"  focus={len(FOCUS)} file(s)")
