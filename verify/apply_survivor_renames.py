"""Rename the staged masters to their chosen surviving filenames (item 6).

Kunchana approved the most-referenced rule, so 9 of the 35 collapsed documents ship
under a different name than `os.walk` happened to pick. Two things must move together
or the delivery breaks:

  * the file in the staged source folder
  * the `file` column of the review sheet that drives Apply

A sheet row whose file is missing is silently skipped by Apply, which is how a delivery
loses documents. So this does both in one step and verifies both afterwards, and the
choice itself comes from survivors.resolve() rather than being restated here.

Run BEFORE Apply. Safe to re-run: renames already done are detected and skipped.

Usage: [set DRP_BATCH=mbw] python apply_survivor_renames.py [--write]
"""
import shutil
import sys
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
except Exception:
    pass

import batch_paths                      # also puts the repo on sys.path
import survivors

from docrefine import reviews
from docrefine.rebrand import ASSET_TYPE_TITLES
from docrefine.worker import Worker

WRITE = "--write" in sys.argv
DOWNLOAD = Path(r"C:\Users\WORK\Documents\MBW Downloadable re-branding")


def main():
    SRC = batch_paths.SRC
    plan = batch_paths.plan()
    rows = reviews.read_plan(plan)
    by_file = {r["file"]: r for r in rows}

    delivered, alias_of, renames = survivors.resolve(DOWNLOAD)
    refs = survivors.reference_counts(DOWNLOAD)

    print(f"source : {SRC}")
    print(f"sheet  : {plan}")
    print(f"rows   : {len(rows)}   renames to apply: {len(renames)}\n")

    todo, done, problems = [], [], []
    for old, new in sorted(renames.items()):
        src_old, src_new = SRC / old, SRC / new
        row = by_file.get(old)
        if src_new.exists() and not src_old.exists() and new in by_file:
            done.append((old, new))
            continue
        if not src_old.exists():
            problems.append(f"{old}: not in the staged folder")
            continue
        if row is None:
            problems.append(f"{old}: not in the sheet")
            continue
        if src_new.exists():
            problems.append(f"{new}: target already exists in the staged folder")
            continue
        if new in by_file:
            problems.append(f"{new}: target name is already a sheet row")
            continue
        todo.append((old, new, row))

    for old, new in done:
        print(f"    already renamed: {old} -> {new}")
    for p in problems:
        print(f"    PROBLEM: {p}")
    if problems:
        sys.exit("\nrefusing to proceed with unresolved problems above")

    for old, new, row in todo:
        print(f"    {refs[old]:4d} -> {refs[new]:4d} refs   {old}")
        print(f"                      -> {new}")

    if not todo:
        print("\nnothing to do")
        return 0
    if not WRITE:
        print(f"\nDRY RUN — {len(todo)} rename(s) not applied. Re-run with --write")
        return 0

    bak = plan.with_name(plan.stem + ".previous" + plan.suffix)
    if not bak.exists():
        shutil.copy2(plan, bak)
        print(f"\noriginal sheet kept as {bak.name}")

    for old, new, row in todo:
        (SRC / old).rename(SRC / new)
        row["file"] = new

    reviews.write_plan(plan, rows, Worker.REBRAND_PLAN_COLUMNS,
                       src_root=str(SRC), asset_types=sorted(ASSET_TYPE_TITLES))

    # Verify both halves agree, and that nothing was lost or duplicated.
    again = reviews.read_plan(plan)
    names = [r["file"] for r in again]
    assert len(again) == len(rows), f"row count changed: {len(rows)} -> {len(again)}"
    assert len(set(names)) == len(names), "duplicate file names in the sheet"
    on_disk = {p.name for p in SRC.glob("*.pdf")}
    for old, new, _ in todo:
        assert new in names, f"{new} missing from the sheet"
        assert old not in names, f"{old} still in the sheet"
        assert new in on_disk, f"{new} missing from the staged folder"
        assert old not in on_disk, f"{old} still in the staged folder"
    missing = [n for n in names if n not in on_disk]
    orphan = [n for n in on_disk if n not in set(names)]
    print(f"\nrenamed {len(todo)} file(s) and their sheet rows")
    print(f"staged PDFs {len(on_disk)}   sheet rows {len(again)}")
    print(f"sheet rows with no file : {len(missing)} {missing[:5]}")
    print(f"files with no sheet row : {len(orphan)} {orphan[:5]}")
    assert not missing and not orphan, "sheet and staged folder disagree"
    print("sheet and staged folder agree exactly")

    # The delivered set must now match what survivors.resolve() promised.
    assert on_disk == delivered, (
        f"staged set != resolved delivered set "
        f"(+{len(on_disk - delivered)} / -{len(delivered - on_disk)})")
    print(f"staged set matches the resolved survivor set ({len(delivered)} files)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
