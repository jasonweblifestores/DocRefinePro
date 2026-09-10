"""Find source PDFs that are structurally damaged or render blank.

Found on 2026-08-21 from a Batch 4 file the client queried:
`2B-Global-Product-Catalog.pdf` declares /Count 32 but only 5 pages are reachable
through the page tree, every reachable page is empty, and its object references
dangle. We branded those 5 blank pages and shipped them. Nothing we did caused it
— the source renders blank too — but nothing we did DETECTED it either.

`deliver_qa` does compare source text to output text, but only on a sample plus a
focus list, so a blank file outside the sample passes silently. This checks EVERY
file, cheaply, on two independent signals:

  tree mismatch  /Count disagrees with the reachable page count, or page objects
                 exist outside the tree. Means pages are unreachable.
  empty pages    a page with no extractable text, no image XObjects and a content
                 stream too small to be drawing anything.

Reports; changes nothing.

Usage: python scan_empty_pages.py <folder> [--limit N]
"""
import logging
import sys
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
except Exception:
    pass

import batch_paths                      # also puts the repo on sys.path

logging.getLogger("pypdf").setLevel(logging.CRITICAL)
from pypdf import PdfReader
from pypdf.generic import IndirectObject

FOLDER = Path(sys.argv[1]) if len(sys.argv) > 1 else batch_paths.SRC
LIMIT = int(sys.argv[sys.argv.index("--limit") + 1]) if "--limit" in sys.argv else None

# A page carrying only a tiny content stream and nothing else is not drawing
# anything a reader would see. Real pages clear this comfortably.
MIN_CONTENT_BYTES = 60


def page_is_empty(pg):
    try:
        if (pg.extract_text() or "").strip():
            return False
    except Exception:
        return False                      # unreadable is a different problem
    try:
        res = pg.get("/Resources") or {}
        xo = res.get("/XObject")
        if xo is not None and len(xo.get_object()) > 0:
            return False
    except Exception:
        pass
    try:
        c = pg.get_contents()
        if c is None:
            return True
        data = c.get_data() if hasattr(c, "get_data") else b""
        return len(data) < MIN_CONTENT_BYTES
    except Exception:
        return False


def tree_mismatch(r):
    """(declared, reachable, orphaned) — orphaned pages are unreachable content."""
    try:
        declared = int(r.trailer["/Root"]["/Pages"].get("/Count", -1))
    except Exception:
        declared = -1
    reachable = len(r.pages)
    tree_ids = set()
    for pg in r.pages:
        ir = pg.indirect_reference
        if ir is not None:
            tree_ids.add((ir.idnum, ir.generation))
    orphaned = 0
    try:
        size = int(r.trailer.get("/Size", 0) or 0)
        for i in range(1, min(size, 5000)):
            try:
                o = IndirectObject(i, 0, r).get_object()
            except Exception:
                continue
            if isinstance(o, dict) and o.get("/Type") == "/Page":
                if (i, 0) not in tree_ids:
                    orphaned += 1
    except Exception:
        pass
    return declared, reachable, orphaned


def main():
    pdfs = sorted(FOLDER.glob("*.pdf"))
    if LIMIT:
        pdfs = pdfs[:LIMIT]
    print(f"scanning {len(pdfs)} PDFs in {FOLDER}\n")

    broken_tree, all_empty, some_empty, unreadable = [], [], [], []
    for i, p in enumerate(pdfs, 1):
        if i % 200 == 0:
            print(f"  ...{i}/{len(pdfs)}", flush=True)
        try:
            r = PdfReader(str(p))
            n = len(r.pages)
        except Exception as e:
            unreadable.append((p.name, f"{type(e).__name__}: {str(e)[:60]}"))
            continue
        declared, reachable, orphaned = tree_mismatch(r)
        if orphaned or (declared >= 0 and declared != reachable):
            broken_tree.append((p.name, declared, reachable, orphaned))
        empties = 0
        for pg in r.pages:
            try:
                if page_is_empty(pg):
                    empties += 1
            except Exception:
                pass
        if empties and empties == n:
            all_empty.append((p.name, n))
        elif empties:
            some_empty.append((p.name, empties, n))

    print(f"\n{'='*66}")
    print(f"UNREADABLE                     {len(unreadable)}")
    for n, e in unreadable[:10]:
        print(f"    {n}  ({e})")
    print(f"\nBROKEN PAGE TREE               {len(broken_tree)}")
    print("    (declared /Count vs reachable pages, plus orphaned page objects)")
    for n, d, reach, orph in broken_tree[:20]:
        print(f"    {n[:52]:54} /Count={d:4} reachable={reach:4} orphaned={orph:4}")
    print(f"\nEVERY PAGE BLANK               {len(all_empty)}")
    for n, c in all_empty[:20]:
        print(f"    {n[:52]:54} {c} pages, all empty")
    print(f"\nSOME PAGES BLANK               {len(some_empty)}")
    for n, e, c in some_empty[:20]:
        print(f"    {n[:52]:54} {e} of {c} pages empty")
    print("=" * 66)
    return 0


if __name__ == "__main__":
    sys.exit(main())
