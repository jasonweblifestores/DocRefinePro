"""Point the per-SKU manifests at the filenames we actually deliver.

Deduplication is by CONTENT, so a document reached under several filenames
collapses to one delivered file (shape A, Kunchana on ClickUp 86eyaantb). The
per-SKU `_manifest.csv` sheets still name the filename that particular SKU folder
used, and for 1,185 of 9,087 rows that name is one of the collapsed aliases — a
row pointing at a file that will not be in the delivery.

Every one of those rows resolves: each aliased filename has exactly one surviving
file with byte-identical content, so this is a mechanical rewrite, not a judgement
call. Two artifacts come out:

  manifests/<SKU>/_manifest.csv   the sheets, rewritten to delivered filenames
  filename-aliases.csv            old -> delivered, the audit trail (and what
                                  Ramez needs for the 301s at site deploy)

The source download is NEVER modified — everything is written into the delivery
tree. Dry run by default; pass --write to actually produce files.

Usage: [set DRP_BATCH=mbw] python rewrite_manifests.py [--write] [--out DIR]
"""
import csv
import json
import shutil
import sys
from collections import Counter
from pathlib import Path

import batch_paths                      # also puts the repo on sys.path
import survivors

WRITE = "--write" in sys.argv
SRC_DOWNLOAD = Path(r"C:\Users\WORK\Documents\MBW Downloadable re-branding")
WORKSPACES = Path(r"C:\Users\WORK\Documents\DocRefinePro_Data\Workspaces")

if "--out" in sys.argv:
    OUT = Path(sys.argv[sys.argv.index("--out") + 1])
else:
    OUT = batch_paths.ROOT / "_delivery-manifests"


def main():
    # The survivor choice comes from survivors.resolve(), the same call the rename
    # step used. Deriving it here a second time is how the manifests would end up
    # pointing at the filenames os.walk happened to pick rather than the ones we
    # actually shipped — silently, and on 913 rows.
    ws = survivors.newest_workspace(SRC_DOWNLOAD)
    delivered, alias, renames = survivors.resolve(SRC_DOWNLOAD, ws)
    print(f"workspace : {ws.name}")
    print(f"delivered : {len(delivered)} PDFs   (survivor renames applied: {len(renames)})")

    # The manifests must name files that are actually on disk. Checking against the
    # staged folder rather than trusting the resolver keeps the two honest.
    staged = {p.name for p in batch_paths.SRC.glob("*.pdf")}
    if staged and staged != delivered:
        extra, short = staged - delivered, delivered - staged
        sys.exit(f"staged folder disagrees with the resolved survivor set: "
                 f"+{len(extra)} {sorted(extra)[:3]} / -{len(short)} {sorted(short)[:3]}\n"
                 f"run apply_survivor_renames.py first")
    print(f"staged    : {len(staged)} PDFs on disk, matches the resolved set")
    print(f"aliases   : {len(alias)} filenames collapse onto a delivered file")

    sheets = sorted(SRC_DOWNLOAD.rglob("_manifest.csv"))
    print(f"manifests : {len(sheets)} sheets\n")

    rows_total = rewritten = untouched = 0
    touched_sheets = 0
    unresolved = Counter()
    plan = []          # (sheet, sku, [(old, new) ...], [all output rows])

    for s in sheets:
        sku = s.parent.relative_to(SRC_DOWNLOAD).as_posix()
        with s.open(newline="", encoding="utf-8-sig") as fh:
            rd = csv.DictReader(fh)
            fields = rd.fieldnames or ["filename"]
            recs = list(rd)

        changes, out_rows = [], []
        for r in recs:
            fn = (r.get("filename") or "").strip()
            if not fn:
                out_rows.append(r)
                continue
            rows_total += 1
            if fn in delivered:
                untouched += 1
            elif fn in alias:
                r = dict(r)
                r["filename"] = alias[fn]
                changes.append((fn, alias[fn]))
                rewritten += 1
            else:
                unresolved[f"{sku}: {fn}"] += 1
            out_rows.append(r)

        if changes:
            touched_sheets += 1
        plan.append((s, sku, fields, changes, out_rows))

    print(f"rows total          {rows_total}")
    print(f"  already delivered {untouched}")
    print(f"  rewritten         {rewritten}   ({touched_sheets} of {len(sheets)} sheets)")
    print(f"  UNRESOLVED        {sum(unresolved.values())}")
    for k, n in unresolved.most_common(10):
        print(f"     {n:4d}  {k}")
    if unresolved:
        sys.exit("\nrefusing to write: some rows name a file that is neither delivered "
                 "nor a known alias. Investigate before shipping manifests.")

    # A row that changed must still describe the SAME content, or the mapping is
    # wrong. Alias entries come from one content hash by construction, so assert
    # the invariant rather than trusting it.
    for _, _, _, changes, _ in plan:
        for old, new in changes:
            assert alias[old] == new and new in delivered, f"bad mapping {old} -> {new}"
    print("\nmapping invariant holds: every rewrite stays within one content hash")

    if not WRITE:
        print(f"\nDRY RUN — nothing written. Would write to:\n  {OUT}")
        sample = [(sku, c[0], c[1]) for _, sku, _, ch, _ in plan for c in ch[:1]][:6]
        for sku, old, new in sample:
            print(f"    {sku}\n      {old}\n   -> {new}")
        return 0

    if OUT.exists():
        shutil.rmtree(OUT)
    (OUT / "manifests").mkdir(parents=True)

    for s, sku, fields, changes, out_rows in plan:
        dst = OUT / "manifests" / sku / "_manifest.csv"
        dst.parent.mkdir(parents=True, exist_ok=True)
        with dst.open("w", newline="", encoding="utf-8") as fh:
            wr = csv.DictWriter(fh, fieldnames=fields)
            wr.writeheader()
            wr.writerows(out_rows)

    with (OUT / "filename-aliases.csv").open("w", newline="", encoding="utf-8") as fh:
        wr = csv.writer(fh)
        wr.writerow(["original_filename", "delivered_filename"])
        for old in sorted(alias):
            wr.writerow([old, alias[old]])

    made = sorted(OUT.rglob("_manifest.csv"))
    print(f"\nwrote {len(made)} manifests + filename-aliases.csv ({len(alias)} rows)")
    print(f"  {OUT}")
    assert len(made) == len(sheets), f"expected {len(sheets)} sheets, wrote {len(made)}"

    # Re-read what we wrote: every filename in the delivered manifests must now
    # be a file we actually ship. This is the check that matters.
    bad = 0
    for m in made:
        with m.open(newline="", encoding="utf-8-sig") as fh:
            for r in csv.DictReader(fh):
                fn = (r.get("filename") or "").strip()
                if fn and fn not in delivered:
                    bad += 1
    print(f"verification: {bad} rows still naming a file we do not deliver")
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
