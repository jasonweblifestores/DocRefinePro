"""The manufacturer spellings table to send the client after an analyze.

On Batch 4 this human round-trip — not any processing — was the real gate, so it
gets its own script rather than being retyped each batch.

What it reports, in the order that matters:

  1. Names that differ only in case or punctuation, AFTER the kit's existing
     aliases are applied. Order matters: the worker's own warning computes over
     the RAW sheet values, so once aliases exist it cries wolf forever. Applying
     first means every group listed here is a decision still outstanding.
  2. The counts table — one row per cleaned name, with how many documents would
     carry it, so the client can see the weight behind each spelling.
  3. What `clean_manufacturer` DROPPED and why. A dropped name means a document
     silently ships with no attribution line, so it must be visible rather than
     inferred from a total that doesn't add up.

Deliberately reports and never rewrites: which spelling is house style is the
brand's call, and `manufacturer_aliases` in brand.json is where it belongs.

Usage: [set DRP_BATCH=batch4] python manufacturer_table.py [--md]
"""
import sys
from collections import Counter, defaultdict

# Manufacturer names come out of real PDFs and carry whatever the source did —
# this batch has a U+FB00 'ff' ligature. On Windows a redirected stdout is cp1252
# and printing that raises, which is how a genuine 50 MB breach once went
# unreported on Batch 4. Never let writing the report be what fails.
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
except Exception:
    pass

import batch_paths                      # also puts the repo on sys.path

from docrefine import reviews, stamps as S
from docrefine.rebrand import BrandKit

MD = "--md" in sys.argv

KIT = batch_paths.require_kit()
kit = BrandKit(KIT)
AL = kit.brand.get("manufacturer_aliases") or {}
BRAND = kit.brand_name

plan = batch_paths.plan()
rows = reviews.read_plan(plan)
reb = [r for r in rows if (r.get("action") or "").strip().lower() == "rebrand"]

print(f"batch  : {batch_paths.BATCH}   brand {BRAND!r}   aliases in kit: {len(AL)}")
print(f"sheet  : {plan}")
print(f"rows   : {len(rows)}   to rebrand: {len(reb)}\n")

# ---------------------------------------------------------------- clean + count
cleaned = Counter()          # printable name -> documents that would carry it
dropped = defaultdict(list)  # reason -> raw values
blank = 0
raw_for_variants = []

for r in reb:
    raw = (r.get("manufacturer") or "").strip()
    if not raw:
        blank += 1
        continue
    name = S.clean_manufacturer(raw, BRAND, AL)
    if name:
        cleaned[name] += 1
        raw_for_variants.append(name)
    else:
        # Say WHY, using the same rules clean_manufacturer applies, so a silent
        # omission becomes a line the client can argue with.
        low = raw.lower()
        if S._DOMAIN.search(raw):
            why = "looks like a website"
        elif BRAND.lower() in low:
            why = f"is {BRAND} itself"
        elif any(h in low for h in S.SELLER_HINTS):
            why = "reads as a seller, not a manufacturer"
        elif len(raw) < 3:
            why = "too short to be a name"
        else:
            why = "rejected"
        dropped[why].append(raw)

# --------------------------------------------------- 1. variants still open
variants = S.spelling_variants(raw_for_variants)
print("=" * 70)
print(f"1. SPELLINGS THAT DIFFER ONLY IN CASE OR PUNCTUATION: {len(variants)} group(s)")
print("   (computed AFTER the kit's aliases, so each one is still undecided)")
print("=" * 70)
if not variants:
    print("   none — every manufacturer name is already consistent")
for g in variants:
    total = sum(cleaned[n] for n in g)
    print(f"\n   {total} documents:")
    for n in sorted(g, key=lambda x: -cleaned[x]):
        print(f"     {cleaned[n]:5d}  {n}")

# --------------------------------------------------- 2. the counts table
print("\n" + "=" * 70)
print(f"2. COUNTS TABLE: {len(cleaned)} manufacturers over "
      f"{sum(cleaned.values())} documents")
print("=" * 70)
if MD:
    print("\n| Manufacturer (as it would print) | Documents |")
    print("|---|---:|")
    for name, n in sorted(cleaned.items(), key=lambda kv: (-kv[1], kv[0].lower())):
        print(f"| {name} | {n} |")
else:
    for name, n in sorted(cleaned.items(), key=lambda kv: (-kv[1], kv[0].lower())):
        print(f"   {n:5d}  {name}")

# --------------------------------------------------- 3. what got dropped
n_dropped = sum(len(v) for v in dropped.values())
print("\n" + "=" * 70)
print(f"3. NO ATTRIBUTION LINE: {blank + n_dropped} of {len(reb)} documents")
print("=" * 70)
print(f"   {blank:5d}  the sheet has no manufacturer at all")
for why, vals in sorted(dropped.items(), key=lambda kv: -len(kv[1])):
    uniq = Counter(vals)
    print(f"   {len(vals):5d}  {why}")
    for v, n in uniq.most_common(8):
        print(f"            {n:4d}x  {v!r}")
    if len(uniq) > 8:
        print(f"            ... and {len(uniq) - 8} more distinct value(s)")

carry = sum(cleaned.values())
print(f"\n   => {carry} of {len(reb)} branded documents would carry an attribution line")
assert carry + blank + n_dropped == len(reb), "counts do not add up"
print("   counts reconcile against the sheet")
