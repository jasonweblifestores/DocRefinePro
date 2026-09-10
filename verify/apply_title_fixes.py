"""Apply only the cover-title corrections that change what the document IS.

review_branded_titles.py asks the vision model for a title on every branded file the
text model decided alone. Most disagreements are wording, not error:

    'Product Warranty'      -> 'Warranty'                  same document, reworded
    'Mounting Instructions' -> 'Installation Instructions' same document, reworded
    'Installation Manual'   -> 'Maintenance Manual'        DIFFERENT DOCUMENT
    'Installation Manual'   -> 'Product Catalog'           DIFFERENT DOCUMENT

Rebuilding a file to reword its cover is churn: it changes bytes the client has
already reviewed, for no reader-visible benefit. Getting the document class wrong is
what Kunchana actually found in Batch 4, and that is worth rebuilding for.

So a change is applied only when the two titles fall into DIFFERENT families below.
Everything else is reported and left alone.

Usage: [set DRP_BATCH=mbw] python apply_title_fixes.py [--write]
"""
import json
import shutil
import sys

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
except Exception:
    pass

import batch_paths                      # also puts the repo on sys.path

from docrefine import reviews
from docrefine.rebrand import ASSET_TYPE_TITLES
from docrefine.worker import Worker

WRITE = "--write" in sys.argv

# Ordered: the first family whose keyword appears wins, so 'installation cut sheet'
# reads as a cut sheet rather than an installation guide.
FAMILIES = [
    ("catalog",        ("catalog", "catalogue", "brochure", "lookbook", "product line")),
    ("warranty",       ("warranty", "guarantee")),
    ("care",           ("maintenance", "care", "cleaning", "upkeep")),
    ("spec",           ("spec", "cut sheet", "cutsheet", "dimension", "data sheet",
                        "compatibility", "reference chart", "material")),
    ("sustainability", ("sustainab", "leed", "environment")),
    ("regulatory",     ("postal regulation", "usps regulation", "compliance", "certification")),
    ("finish",         ("powder", "finish", "color option", "colour option", "paint")),
    ("install",        ("install", "assembly", "assemble", "mounting", "mount",
                        "setup", "set up", "template", "procedure")),
]


def family(title):
    t = " ".join(str(title or "").lower().split())
    for name, kws in FAMILIES:
        if any(k in t for k in kws):
            return name
    return "other"


def filename_family(name):
    """What the FILENAME claims the document is, or 'other' if it says nothing.

    The filename is independent evidence, written by whoever produced the document.
    Where it disagrees with vision, vision is the one guessing.
    """
    return family(str(name).replace("_", " ").replace("-", " ").rsplit(".", 1)[0])


def main():
    plan = batch_paths.plan()
    review = plan.with_name(f"{batch_paths.BATCH}_vision-title-review.json")
    if not review.is_file():
        sys.exit(f"no review file at {review} — run review_branded_titles.py first")
    data = json.loads(review.read_text(encoding="utf-8"))
    changes = data.get("title_changes", [])

    rows = reviews.read_plan(plan)
    by_file = {r["file"]: r for r in rows}

    substantive, churn, product_name, contradicted, missing = [], [], [], [], []
    for c in changes:
        r = by_file.get(c["file"])
        if r is None:
            missing.append(c["file"]); continue
        fo, fn = family(c["old"]), family(c["new"])
        ff = filename_family(c["file"])
        if fo == fn:
            churn.append((c, fo, fn))
        elif fn == "other":
            # e.g. 'Installation Manual' -> 'Mail Slots'. Vision has named the
            # PRODUCT, not the document. The SOP wants the cover to say what the
            # document IS, so this would make the cover less useful, not more.
            product_name.append((c, fo, fn))
        elif ff != "other" and ff != fn:
            # The filename says one thing and vision says another. Applying this
            # would introduce the very defect we are fixing — it is what would
            # retitle 'Venia-...-Limited-Warranty-2024.pdf' as an installation
            # manual, and what would accept a confident title read off the blank
            # pages of a corrupt file. The filename wins.
            contradicted.append((c, fo, fn, ff))
        else:
            substantive.append((c, fo, fn))

    print(f"sheet        : {plan}")
    print(f"title diffs  : {len(changes)}")
    print(f"  substantive: {len(substantive)}   (document class changes — will apply)")
    print(f"  wording     : {len(churn)}   (same class — left alone)")
    print(f"  product name: {len(product_name)}   (names the product, not the document "
          f"— left alone)")
    print(f"  contradicted: {len(contradicted)}   (filename disagrees with vision "
          f"— left alone)")
    if missing:
        print(f"  NOT IN SHEET: {len(missing)} {missing[:3]}")

    print("\nSUBSTANTIVE:")
    for c, fo, fn in substantive:
        print(f"  [{fo} -> {fn}] conf={c['confidence']}")
        print(f"     {c['file']}")
        print(f"     {c['old']!r}  ->  {c['new']!r}")
    if not substantive:
        print("  none")

    print(f"\nFILENAME CONTRADICTS VISION (not applied), {len(contradicted)}:")
    for c, fo, fn, ff in contradicted:
        print(f"  filename says {ff}, vision says {fn}: {c['old']!r} -> {c['new']!r}")
        print(f"        {c['file']}")

    print(f"\nPRODUCT NAME RATHER THAN DOCUMENT TYPE (not applied), {len(product_name)}:")
    for c, fo, _ in product_name:
        print(f"  [{fo}] {c['old']!r} -> {c['new']!r}\n        {c['file']}")

    print(f"\nWORDING ONLY (not applied), first 12 of {len(churn)}:")
    for c, fo, _ in churn[:12]:
        print(f"  [{fo}] {c['old']!r} -> {c['new']!r}   {c['file'][:46]}")

    if not substantive:
        print("\nnothing to apply")
        return 0
    if not WRITE:
        print("\nDRY RUN — sheet untouched. Re-run with --write")
        return 0

    bak = plan.with_name(plan.stem + ".previous" + plan.suffix)
    if not bak.exists():
        shutil.copy2(plan, bak)
        print(f"\noriginal kept as {bak.name}")
    for c, fo, fn in substantive:
        r = by_file[c["file"]]
        r["title"] = c["new"]
        if c.get("asset_type"):
            r["asset_type"] = c["asset_type"]
        r["source"] = "vision"
        r["confidence"] = str(c["confidence"])
        r["notes"] = (f"cover title corrected after the visual pass: the text model read this "
                      f"as a {fo} document, the page shows a {fn} document")
    reviews.write_plan(plan, rows, Worker.REBRAND_PLAN_COLUMNS,
                       src_root=str(batch_paths.SRC),
                       asset_types=sorted(ASSET_TYPE_TITLES))

    again = {r["file"]: r for r in reviews.read_plan(plan)}
    assert len(again) == len(rows), "row count changed"
    for c, _, _ in substantive:
        got = again[c["file"]]["title"].strip()
        assert got == c["new"].strip(), f"{c['file']}: read back {got!r}"
    print(f"\nwrote {len(substantive)} corrected title(s); all read back as written")
    print("\nnext: delete these outputs and re-run apply_batch.py --resume")
    for c, _, _ in substantive:
        print(f"   {c['file']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
