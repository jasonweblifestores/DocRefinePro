"""Trial run: production settings, real files, deliberately awkward selection.

Which batch comes from batch_paths (DRP_BATCH env var, default mbw).
"""
import shutil, sys, time
from pathlib import Path
import batch_paths                      # also puts the repo on sys.path
from docrefine import reviews, stamps as S
from docrefine.rebrand import BrandKit
from docrefine.worker import Worker

SP   = Path(sys.argv[1])
SRC  = batch_paths.SRC
KIT  = batch_paths.require_kit()
PLAN = batch_paths.plan()
OUT  = SP / "TRIAL_output"
if OUT.exists(): shutil.rmtree(OUT)

kit = BrandKit(KIT)
al = kit.brand.get("manufacturer_aliases") or {}
rows = reviews.read_plan(PLAN)
by_file = {r["file"]: r for r in rows}

def pick(pred, n=1, exclude=()):
    out = []
    for r in rows:
        if r["file"] in exclude or r["file"] in [x["file"] for x in out]: continue
        if (SRC / r["file"]).exists() and pred(r):
            out.append(r)
            if len(out) == n: break
    return out

def man(r): return (r.get("manufacturer") or "").strip()
def act(r): return (r.get("action") or "").strip().lower()
def pages(r):
    try: return int(r.get("pages") or 0)
    except ValueError: return 0

chosen, why = [], []
def add(rs, label):
    for r in rs:
        if r["file"] not in [c["file"] for c in chosen]:
            chosen.append(r); why.append((r["file"], label))

# Deliberately awkward cases, not a random sample. Selected by INTENT rather than
# by literal manufacturer strings — the old version named Batch 4's values, which
# either KeyError or silently match nothing on another batch, so the trial would
# quietly shrink instead of failing.
BRAND = kit.brand_name
size = lambda r: (SRC / r["file"]).stat().st_size
cleaned = lambda r: S.clean_manufacturer(man(r), BRAND, al)

# Files this batch has been bitten by before, if they are in the sheet at all.
add([by_file[n] for n in batch_paths.FOCUS if n in by_file], "known-awkward for this batch")

# An alias must actually change the printed name. With an empty alias map (before
# the client returns canonical names) there is nothing to test, and saying so is
# better than a selection that looks complete.
add(pick(lambda r: act(r) == "rebrand" and man(r) and cleaned(r) and cleaned(r) != man(r)),
    "alias applies: raw name is rewritten")
add(pick(lambda r: act(r) == "rebrand" and man(r) and not cleaned(r)
         and S._DOMAIN.search(man(r))), "website value -> NO attribution")
add(pick(lambda r: act(r) == "rebrand" and man(r) and not cleaned(r)
         and BRAND.lower() in man(r).lower()), "our own brand -> NO attribution")
add(pick(lambda r: act(r) == "rebrand" and man(r) and not cleaned(r)
         and any(h in man(r).lower() for h in S.SELLER_HINTS)),
    "seller value -> NO attribution")
add(pick(lambda r: act(r) == "rebrand" and not man(r)), "blank manufacturer -> NO attribution")
add(pick(lambda r: act(r) == "rebrand" and cleaned(r)), "ordinary attribution prints")
add(pick(lambda r: act(r) == "rebrand" and pages(r) >= 12), "multi-page: stamps on every page")
add(pick(lambda r: act(r) == "rebrand" and pages(r) == 1), "single page: cover + body + back")
add(pick(lambda r: act(r) == "leave"), "leave row -> must stay byte-identical")
add(pick(lambda r: act(r) == "leave" and pages(r) >= 2, 1,
         exclude={c["file"] for c in chosen}), "second leave row")
add(pick(lambda r: act(r) == "rebrand" and size(r) > 3e6), "large file (size cap)")
add(pick(lambda r: act(r) == "rebrand" and any(c in r["file"] for c in "'& ")),
    "awkward filename")
# Largest and widest in the batch: the size cap and the large-format geometry.
reb_rows = [r for r in rows if act(r) == "rebrand" and (SRC / r["file"]).exists()]
if reb_rows:
    add([max(reb_rows, key=size)], "largest file in the batch")
    add([max(reb_rows, key=pages)], "most pages in the batch")

print(f"batch {batch_paths.BATCH}   brand {BRAND!r}   aliases in kit: {len(al)}")
if not al:
    print("  NOTE: the kit has no manufacturer aliases yet, so the alias case cannot")
    print("        be exercised. Re-run this trial once the canonical names land.")
print(f"\nTRIAL SELECTION ({len(chosen)} files)")
for f, lbl in why:
    r = by_file[f]
    print(f"  [{act(r):7}] {f[:52]:52s}  {lbl}")


trial_plan = SP / "trial_plan.xlsx"
reviews.write_plan(trial_plan, chosen, Worker.REBRAND_PLAN_COLUMNS, src_root=SRC)

log = []
w = Worker(callback=lambda e: None)
w.log = lambda m, err=False: log.append(("ERR " if err else "    ") + str(m))
t0 = time.time()
w.run_rebrand_apply(str(SRC), str(KIT), str(trial_plan), out_dir=str(OUT),
                    complete_set=False,          # keeps desktop.ini / _unique-index.csv out
                    show_attribution=False,      # footer placement, not cover
                    keep_original_names=True,
                    stamp_opts={"footer_attribution": True,
                                "stamp_tagline": batch_paths.wording(KIT)[1],
                                "stamp_version": True, "stamp_disclaimer": True})
print(f"\nRUN LOG ({time.time()-t0:.1f}s)")
for l in log: print("  " + l)
