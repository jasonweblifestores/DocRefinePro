"""Which filename survives a content collapse — computed once, used everywhere.

Deduplication is by content, so a document reached under several filenames ships as
one file. Ingest keeps whichever copy `os.walk` reached first, which is arbitrary; the
surviving name becomes a live URL and every other becomes a redirect for Ramez.

Kunchana chose the **most-referenced** name (ClickUp 86eyaantb comment 90180247980961,
item 6): "yes, use the most-referenced name. 272 fewer redirects is the right trade."

The tie-break keeps the CURRENT master when it is already among the most-referenced.
Two of the 35 groups are exact ties where a rename would buy nothing, and a rename that
buys nothing is still a live URL changing — so it is not made.

This module is the single source of that decision. The rename step and the manifest
rewrite both call it, because two implementations of "which file did we ship" is exactly
how Batch 4 lost two files from a 2,174-file delivery.
"""
import csv
import json
from collections import Counter
from pathlib import Path

WORKSPACES = Path(r"C:\Users\WORK\Documents\DocRefinePro_Data\Workspaces")


def newest_workspace(download_root):
    """The most recent ingest workspace for a source tree."""
    prefix = Path(download_root).name
    cands = sorted((d for d in WORKSPACES.iterdir()
                    if d.is_dir() and d.name.startswith(prefix)), key=lambda d: d.name)
    if not cands:
        raise SystemExit(f"no ingest workspace under {WORKSPACES} for {prefix}")
    return cands[-1]


def reference_counts(download_root):
    """{filename: how many per-SKU manifest rows name it}."""
    counts = Counter()
    for sheet in Path(download_root).rglob("_manifest.csv"):
        with sheet.open(newline="", encoding="utf-8-sig") as fh:
            for row in csv.DictReader(fh):
                fn = (row.get("filename") or "").strip()
                if fn:
                    counts[fn] += 1
    return counts


def resolve(download_root, workspace=None):
    """Work out the delivered name for every content group.

    Returns (delivered, alias_of, renames):
      delivered  set of filenames the delivery will contain
      alias_of   {every other filename: its delivered name} - the redirect map
      renames    {current master name: new name} - only where the choice moves
    """
    ws = Path(workspace) if workspace else newest_workspace(download_root)
    refs = reference_counts(download_root)
    man = json.loads((ws / "manifest.json").read_text(encoding="utf-8"))

    delivered, alias_of, renames = set(), {}, {}
    for data in man.values():
        if data.get("status") == "QUARANTINE":
            continue
        if not data["name"].lower().endswith(".pdf"):
            continue
        current = data["name"]
        names = {Path(c).name for c in data["copies"]}
        best = max(refs[n] for n in names)
        # Keep the current name when it is already among the most-referenced, so a
        # tie never churns a live URL for nothing.
        keep = current if refs[current] == best else min(
            (n for n in names if refs[n] == best), key=str.lower)
        delivered.add(keep)
        if keep != current:
            renames[current] = keep
        for n in names - {keep}:
            alias_of[n] = keep
    return delivered, alias_of, renames


def summary(download_root, workspace=None):
    delivered, alias_of, renames = resolve(download_root, workspace)
    refs = reference_counts(download_root)
    saved = sum(refs[new] - refs[old] for old, new in renames.items())
    return {
        "delivered": len(delivered),
        "aliases": len(alias_of),
        "renames": len(renames),
        "redirects_avoided": saved,
        "rows_needing_redirect": sum(refs[n] for n in alias_of),
    }


if __name__ == "__main__":
    import sys
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
    except Exception:
        pass
    root = Path(r"C:\Users\WORK\Documents\MBW Downloadable re-branding")
    delivered, alias_of, renames = resolve(root)
    refs = reference_counts(root)
    print(f"delivered files          : {len(delivered)}")
    print(f"aliased filenames        : {len(alias_of)}")
    print(f"renames from walk order  : {len(renames)}")
    print(f"manifest rows to rewrite : {sum(refs[n] for n in alias_of)}")
    print(f"redirects avoided        : "
          f"{sum(refs[new] - refs[old] for old, new in renames.items())}\n")
    for old, new in sorted(renames.items(), key=lambda kv: refs[kv[0]] - refs[kv[1]]):
        print(f"  {refs[old]:4d} -> {refs[new]:4d} refs")
        print(f"       {old}")
        print(f"    -> {new}")
