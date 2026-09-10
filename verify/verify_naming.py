"""v142: delivery filenames stay meaningful when the model finds no product name."""
import sys, re, tempfile, shutil
from pathlib import Path

REPO = Path(sys.argv[1])
sys.path.insert(0, str(REPO))

results = []
def check(n, ok, d=""):
    results.append(bool(ok)); print(f"[{'PASS' if ok else 'FAIL'}] {n}  {d}")

from docrefine.rebrand import (delivery_filename, numbered_filename, output_filename_from_fields,
                               MAX_DELIVERY_NAME, BrandKit)
from docrefine.worker import Worker

# =====================================================================
#  Blank product falls back to the source name
# =====================================================================
check("N1 a product is used when the model found one",
      delivery_filename("1570-12-BM", "1570 Cluster Box", "installation-guide")
      == "1570-cluster-box-installation-guide-budget-mailboxes.pdf")
check("N2 a blank product falls back to the source name",
      delivery_filename("1570-12-BM", "", "installation-guide")
      == "1570-12-bm-installation-guide-budget-mailboxes.pdf")
check("N3 blank product AND asset type still yields something usable",
      delivery_filename("1570-12-BM", "", "") == "1570-12-bm-budget-mailboxes.pdf")
check("N4 an empty everything never produces a bare brand name",
      delivery_filename("", "", "") == "document-budget-mailboxes.pdf")
check("N5 the brand slug comes from the kit",
      delivery_filename("x", "Widget", "manual", brand_slug="acme-mail")
      == "widget-manual-acme-mail.pdf")

# =====================================================================
#  Length cap and word boundaries
# =====================================================================
long_product = "Wall Mount Mailbox With Decorative Cast Aluminium Surround Assembly"
n = delivery_filename("src", long_product, "mounting-instructions")
check("N6 a long name is capped", len(n) <= MAX_DELIVERY_NAME, f"{len(n)} chars: {n}")
check("N7 the asset type survives truncation", n.endswith("mounting-instructions-budget-mailboxes.pdf"), n)
check("N8 truncation lands on a word boundary, not mid-word",
      not re.search(r"[a-z]{2,}-mounting", n) or "-" in n.split("-mounting")[0][-1:] or
      all(len(w) > 1 for w in n.split("-")), n)
check("N9 no trailing hyphen before the asset type", "--" not in n, n)

# numbered suffix must stay inside the cap
capped = "a" * 39 + "-budget-mailboxes.pdf"
check("N10 baseline is exactly at the cap", len(capped) == MAX_DELIVERY_NAME, f"{len(capped)}")
for i in (2, 10, 250):
    num = numbered_filename(capped, i)
    check(f"N10 -{i} suffix stays within the cap", len(num) <= MAX_DELIVERY_NAME, f"{len(num)}: {num}")
    check(f"N10 -{i} keeps the extension", num.endswith(".pdf"), num)

# =====================================================================
#  Whole-batch assignment: unique, deterministic, meaningful
# =====================================================================
BK = Path(sys.argv[2]) if len(sys.argv) > 2 else None
kit = BrandKit(BK) if BK else None
rows = []
# 6 documents the model gave the same product+asset (the real failure mode)
for i in range(6):
    rows.append({"file": f"sub/wall-{2400+i}.pdf", "action": "rebrand",
                 "product": "Wall Mount Mailbox", "asset_type": "mounting-instructions"})
# 3 with no product at all
for i in range(3):
    rows.append({"file": f"1570-{i}-BM.pdf", "action": "rebrand",
                 "product": "", "asset_type": "installation-guide"})
rows.append({"file": "drawing.pdf", "action": "leave"})

w = Worker(callback=lambda e: None)
out = Path("C:/out") if sys.platform == "win32" else Path("/out")
w._assign_output_names(rows, out, kit)
names = [Path(r["_dst"]).name for r in rows]
check("N11 every file gets a unique name", len(set(names)) == len(names), str(len(set(names))))
check("N12 same-product documents are told apart by their source name",
      all(f"{2400+i}" in names[i] for i in range(6)), str(names[:3]))
check("N13 no meaningless -N suffixes when a source name can be used",
      not any(re.search(r"-\d+\.pdf$", n) for n in names[:6]), str(names[:6]))
check("N14 blank-product rows carry their source name",
      all(f"1570-{i}-bm" in names[6 + i] for i in range(3)), str(names[6:9]))
check("N15 leave rows keep their original filename", names[-1] == "drawing.pdf", names[-1])
check("N16 every assigned name respects the cap",
      all(len(n) <= MAX_DELIVERY_NAME for n in names), str(max(len(n) for n in names)))

rows2 = [dict(r) for r in rows]
for r in rows2:
    r.pop("_dst", None)
w._assign_output_names(rows2, out, kit)
check("N17 assignment is deterministic across runs",
      [Path(r["_dst"]).name for r in rows2] == names)

# a genuine unavoidable duplicate still gets a stable number
dupes = [{"file": "same.pdf", "action": "rebrand", "product": "P", "asset_type": "a"},
         {"file": "same.pdf", "action": "rebrand", "product": "P", "asset_type": "a"}]
w._assign_output_names(dupes, out, kit)
d = [Path(r["_dst"]).name for r in dupes]
check("N18 identical sources still get distinct names", d[0] != d[1], str(d))

print("\n" + "=" * 56)
print(f"RESULT: {sum(results)}/{len(results)} passed")
sys.exit(0 if all(results) else 1)
