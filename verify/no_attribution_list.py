"""The documents that will carry NO attribution line, and why — verifiably.

A count the client has to take on trust is not evidence. This names every file,
the raw value the document actually carried, the rule that suppressed it, and a
SKU folder it lives in, so any row can be opened and checked.

Writes a CSV alongside the review sheet for attaching to the task.

Usage: [set DRP_BATCH=mbw] python no_attribution_list.py
"""
import csv
import sys
from collections import defaultdict
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
except Exception:
    pass

import batch_paths                      # also puts the repo on sys.path

from docrefine import reviews, stamps as S
from docrefine.rebrand import BrandKit

DOWNLOAD = Path(r"C:\Users\WORK\Documents\MBW Downloadable re-branding")

# filename -> SKU folders referencing it
skus = defaultdict(list)
for s in DOWNLOAD.rglob("_manifest.csv"):
    with s.open(newline="", encoding="utf-8-sig") as fh:
        for r in csv.DictReader(fh):
            fn = (r.get("filename") or "").strip()
            if fn:
                skus[fn].append(s.parent.name)


def sku_of(name):
    ss = sorted(set(skus.get(name, [])))
    if not ss:
        return "(alias of another file)"
    return ss[0] + (f" +{len(ss)-1}" if len(ss) > 1 else "")


kit = BrandKit(batch_paths.require_kit())
BRAND = kit.brand_name
AL = kit.brand.get("manufacturer_aliases") or {}

rows = reviews.read_plan(batch_paths.plan())
reb = [r for r in rows if (r.get("action") or "").strip().lower() == "rebrand"]

# Order matters: it mirrors clean_manufacturer's own order of checks.
def why(raw):
    if not raw:
        return "no manufacturer named in the document"
    if S._DOMAIN.search(raw):
        return "a website, not a company"
    low = raw.lower()
    if BRAND.lower() in low:
        return "names MailboxWorks itself"
    if any(h in low for h in S.SELLER_HINTS):
        return "reads as the seller, not the manufacturer"
    if len(raw) < 3:
        return "too short to be a company name"
    return "rejected"


out = []
for r in reb:
    raw = (r.get("manufacturer") or "").strip()
    if S.clean_manufacturer(raw, BRAND, AL):
        continue
    out.append({"file": r["file"], "example_sku": sku_of(r["file"]),
                "value_in_document": raw, "why_no_line": why(raw),
                "title": (r.get("title") or "").strip()})

buckets = defaultdict(list)
for o in out:
    buckets[o["why_no_line"]].append(o)

print(f"{len(reb)} to rebrand — {len(reb) - len(out)} carry a line, "
      f"{len(out)} do not\n")
for w, items in sorted(buckets.items(), key=lambda kv: -len(kv[1])):
    vals = sorted({i["value_in_document"] for i in items})
    print(f"{len(items):4d}  {w}")
    if vals and vals != [""]:
        print(f"        value(s): {', '.join(repr(v) for v in vals)}")
    for i in items[:3]:
        print(f"        e.g. {i['file']}   [SKU {i['example_sku']}]")
    if len(items) > 3:
        print(f"        ... and {len(items) - 3} more (all in the CSV)")
    print()

dst = batch_paths.plan().with_name("MBW_no-attribution-line.csv")
with dst.open("w", newline="", encoding="utf-8-sig") as fh:
    w = csv.DictWriter(fh, fieldnames=["file", "example_sku", "value_in_document",
                                       "why_no_line", "title"])
    w.writeheader()
    w.writerows(sorted(out, key=lambda o: (o["why_no_line"], o["file"])))
print(f"wrote {len(out)} rows -> {dst}")
assert len(out) + (len(reb) - len(out)) == len(reb)
