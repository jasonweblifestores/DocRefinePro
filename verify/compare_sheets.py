"""What did looking at the pages actually change?

Compares the new sheet against the one it replaced, so the reclassification can
be reported as a number of specific files rather than a delta in a total.
"""
import logging
import sys
from collections import Counter
from pathlib import Path

import batch_paths                      # also puts the repo on sys.path
logging.getLogger("pypdf").setLevel(logging.CRITICAL)

from docrefine import reviews, stamps as S
from docrefine.rebrand import BrandKit

NEW = batch_paths.plan()
OLD = NEW.with_name(NEW.stem + ".previous" + NEW.suffix)

kit = BrandKit(batch_paths.require_kit())
AL = kit.brand.get("manufacturer_aliases") or {}
BRAND = kit.brand_name


def act(r):
    return (r.get("action") or "").strip().lower()


new = {r["file"]: r for r in reviews.read_plan(NEW)}
old = {r["file"]: r for r in reviews.read_plan(OLD)}

print(f"new sheet: {len(new)} rows    previous: {len(old)} rows\n")

flips = Counter()
to_leave, to_rebrand = [], []
for f, r in new.items():
    o = old.get(f)
    if not o:
        flips["new file (not in previous sheet)"] += 1
        continue
    a, b = act(o), act(r)
    if a == b:
        flips[f"unchanged: {b}"] += 1
    else:
        flips[f"{a} -> {b}"] += 1
        (to_leave if b == "leave" else to_rebrand).append(r)

for k, v in sorted(flips.items(), key=lambda x: -x[1]):
    print(f"  {k:34} {v}")

print(f"\n--- {len(to_leave)} files the vision pass pulled OUT of branding ---")
for r in to_leave[:15]:
    print(f"  {r['file'][:56]:56} {(r.get('doc_type') or '')[:26]:26} conf={r.get('confidence')}")
    if r.get("notes"):
        print(f"      {str(r['notes'])[:96]}")
if len(to_leave) > 15:
    print(f"  … and {len(to_leave) - 15} more")

if to_rebrand:
    print(f"\n--- {len(to_rebrand)} files it pulled INTO branding ---")
    for r in to_rebrand[:15]:
        print(f"  {r['file'][:56]:56} {(r.get('doc_type') or '')[:26]:26} conf={r.get('confidence')}")
        if r.get("notes"):
            print(f"      {str(r['notes'])[:96]}")
    if len(to_rebrand) > 15:
        print(f"  … and {len(to_rebrand) - 15} more")

# --- what will actually print on the covers/footers -------------------------
reb = [r for r in new.values() if act(r) == "rebrand"]
printed = [r for r in reb if S.clean_manufacturer(r.get("manufacturer"), BRAND, AL)]
print(f"\n--- attribution projection over {len(reb)} branded files ---")
print(f"  will print an attribution line : {len(printed)}")
print(f"  will print none                : {len(reb) - len(printed)}")

raw = Counter((r.get("manufacturer") or "").strip() for r in reb)
unusable = Counter()
for r in reb:
    m = (r.get("manufacturer") or "").strip()
    if m and not S.clean_manufacturer(m, BRAND, AL):
        unusable[m] += 1
print(f"\n  top manufacturers as they will print:")
seen = Counter(S.clean_manufacturer(r.get("manufacturer"), BRAND, AL) for r in printed)
for m, c in seen.most_common(12):
    print(f"     {c:4}  {m}")
if unusable:
    print(f"\n  values deliberately yielding NO attribution (seller/website/self):")
    for m, c in unusable.most_common(8):
        print(f"     {c:4}  {m!r}")
blank = raw.get("", 0)
print(f"\n  blank manufacturer: {blank}")
