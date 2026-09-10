"""QA a whole delivered tree against the sheet that produced it.

Every substantive defect in v140-v157 was found by looking at the real output,
never by the unit tests, which were green throughout. So this checks the actual
deliverable, and it checks the cheap things on EVERY file rather than a sample:
presence, naming, size cap, page count, Author, and byte-identity of the files
we promised to leave alone.

The expensive per-page text and stamp checks run on a sample plus a fixed focus
list of files that have bitten us before. Pass a bigger sample as argv[2] to
widen that.

Which batch it acts on comes from batch_paths (DRP_BATCH env var, default mbw).

Usage: [set DRP_BATCH=batch4] python deliver_qa.py <scratch> [sample=150]
"""
import logging
import random
import sys
from collections import Counter
from pathlib import Path

import batch_paths                      # also puts the repo on sys.path
logging.getLogger("pypdf").setLevel(logging.CRITICAL)
from pypdf import PdfReader

from docrefine import reviews, stamps as S
from docrefine.rebrand import BrandKit, page_size

SP = Path(sys.argv[1])
SAMPLE = int(sys.argv[2]) if len(sys.argv) > 2 else 150

SRC, OUT, KIT = batch_paths.SRC, batch_paths.OUT, batch_paths.require_kit()

# Wording comes from the kit, never a literal: MailboxWorks has no tagline, and
# a blanked TAG would make `all(TAG in b for b in body)` true for every page —
# the check would look green while verifying nothing. TAG_EXPECTED says whether
# to assert the tagline is present or assert it is absent.
TAG, TAG_EXPECTED, DISC = batch_paths.wording(KIT)

# Files that have caused real defects before — always inspected in full.
FOCUS = batch_paths.FOCUS

kit = BrandKit(KIT)
AL = kit.brand.get("manufacturer_aliases") or {}
BRAND = kit.brand_name

fails, warns = [], []
stats = Counter()


def chk(ok, msg, hard=True):
    if ok:
        return True
    (fails if hard else warns).append(msg)
    return False


def norm(t):
    return "".join((t or "").split())


plan = reviews.plan_path_for(SRC)
rows = reviews.read_plan(plan)
by_file = {r["file"]: r for r in rows}
reb = [r for r in rows if (r.get("action") or "").strip().lower() == "rebrand"]
leave = [r for r in rows if (r.get("action") or "").strip().lower() != "rebrand"]

print(f"batch:  {batch_paths.BATCH}   brand {BRAND!r}   "
      f"tagline {'ON: ' + repr(TAG) if TAG_EXPECTED else 'OFF (assert absent)'}")
print(f"sheet:  {plan}")
print(f"rows:   {len(rows)}   rebrand {len(reb)}   leave {len(leave)}")
print(f"output: {OUT}\n")

# ---------------------------------------------------------------- inventory
on_disk = [f for f in OUT.rglob("*") if f.is_file()]
pdfs = [f for f in on_disk if f.suffix.lower() == ".pdf"]
others = [f for f in on_disk if f.suffix.lower() != ".pdf"]

chk(len(pdfs) == len(rows), f"every sheet row produced a PDF: {len(pdfs)} of {len(rows)}")
chk(not others, f"no non-PDF files in the delivery (complete set OFF): "
                f"{[f.name for f in others[:8]]}")
debris = [f.name for f in on_disk if f.suffix.lower() in (".part", ".tmp") or f.name.startswith("~")]
chk(not debris, f"no temp/partial debris: {debris[:8]}")

names = [f.name for f in pdfs]
dupes = [n for n, c in Counter(names).items() if c > 1]
chk(not dupes, f"all output filenames unique: {len(set(names))} of {len(names)}  dupes={dupes[:5]}")
toolong = [n for n in names if len(n) > 60]
chk(not toolong, f"filenames within 60 chars: {len(toolong)} over  {toolong[:3]}", hard=False)

missing = [r["file"] for r in rows if not (OUT / r["file"]).is_file()]
chk(not missing, f"nothing missing from the output: {len(missing)}  {missing[:5]}")
extra = sorted({f.relative_to(OUT).as_posix() for f in pdfs} - set(by_file))
chk(not extra, f"nothing in the output that isn't on the sheet: {len(extra)}  {extra[:5]}")

# keep-original-names was ON: every delivered name must equal its source name
renamed = [r["file"] for r in rows
           if (OUT / r["file"]).is_file() and (OUT / r["file"]).name != (SRC / r["file"]).name]
chk(not renamed, f"every file keeps its original name: {len(renamed)} renamed  {renamed[:5]}")

# ------------------------------------------------------- cheap checks, all files
over_cap, bad_pages, bad_author, unreadable, not_identical = [], [], [], [], []
sizes = []
for r in rows:
    rel = r["file"]
    f = OUT / rel
    s = SRC / rel
    if not f.is_file():
        continue
    mb = f.stat().st_size / 1e6
    sizes.append((mb, rel))
    if mb >= 50:
        over_cap.append(f"{rel} {mb:.1f}MB")

    if (r.get("action") or "").strip().lower() != "rebrand":
        # promised untouched — prove it byte for byte
        try:
            if f.read_bytes() != s.read_bytes():
                not_identical.append(rel)
        except Exception as e:
            unreadable.append(f"{rel}: {e}")
        continue

    try:
        sr, dr = PdfReader(str(s)), PdfReader(str(f))
        if len(dr.pages) != len(sr.pages) + 2:
            bad_pages.append(f"{rel} {len(sr.pages)}->{len(dr.pages)}")
        if (dr.metadata or {}).get("/Author") != BRAND:
            bad_author.append(f"{rel} {(dr.metadata or {}).get('/Author')!r}")
    except Exception as e:
        unreadable.append(f"{rel}: {type(e).__name__}: {e}")

chk(not over_cap, f"every file under the 50 MB cap: {len(over_cap)} over  {over_cap[:5]}")
chk(not bad_pages, f"every branded file is source+2 pages: {len(bad_pages)} wrong  {bad_pages[:5]}")
chk(not bad_author, f"Author is '{BRAND}' throughout: {len(bad_author)} wrong  {bad_author[:5]}")
chk(not not_identical, f"every 'leave' file is byte-identical: {len(not_identical)} differ  {not_identical[:5]}")
chk(not unreadable, f"every output file opens: {len(unreadable)} failed  {unreadable[:5]}")

sizes.sort(reverse=True)
print(f"\nlargest files:")
for mb, rel in sizes[:5]:
    print(f"   {mb:7.1f} MB  {rel}")
print(f"total delivered: {sum(m for m, _ in sizes) / 1000:.2f} GB\n")

# ------------------------------------------- expensive checks, focus + sample
pool = [r["file"] for r in reb if (OUT / r["file"]).is_file()]
random.seed(11)
picks = list(dict.fromkeys([n for n in FOCUS if n in by_file]
                           + random.sample(pool, min(SAMPLE, len(pool)))))
picks = [p for p in picks if (OUT / p).is_file()][:SAMPLE + len(FOCUS)]

print(f"inspecting {len(picks)} documents page by page "
      f"({len([n for n in FOCUS if n in by_file])} known-awkward + sample)\n")

attrib_printed = attrib_absent = 0
lost_text, no_tag, no_ver, no_disc, stamped_cover, geom, wrong_attrib = [], [], [], [], [], [], []
stray_tag, foreign_brand = [], []

# Brand wording that belongs to a DIFFERENT kit. If brand.json fails to resolve,
# BrandKit silently falls back to the Budget Mailboxes defaults and the run looks
# successful while printing the wrong company — so look for the neighbours by
# name rather than trusting that our own name showed up.
OTHER_BRANDS = [b for b in ("Budget Mailboxes", "MailboxWorks") if b != BRAND]
OTHER_TAGS = [t for t in ("Trusted by the Nation",) if t != TAG]

for rel in picks:
    r = by_file[rel]
    if (r.get("action") or "").strip().lower() != "rebrand":
        continue
    try:
        sr = PdfReader(str(SRC / rel))
        dr = PdfReader(str(OUT / rel))
    except Exception as e:
        unreadable.append(f"{rel}: {e}")
        continue
    stats["inspected"] += 1

    body = [" ".join((dr.pages[i].extract_text() or "").split())
            for i in range(1, len(dr.pages) - 1)]

    # source text must survive, page for page (offset by the front cover)
    bad = [i + 1 for i in range(len(sr.pages))
           if norm(sr.pages[i].extract_text())
           and i + 1 < len(dr.pages) - 1
           and norm(sr.pages[i].extract_text()) not in norm(dr.pages[i + 1].extract_text())]
    if bad:
        lost_text.append(f"{rel} pages {bad[:6]}")

    if body:
        if TAG_EXPECTED:
            if not all(TAG in b for b in body):
                no_tag.append(rel)
        else:
            # This kit stamps no tagline. There is no string to search for, so
            # assert the absence of every tagline we know of instead — a blank
            # `TAG in b` would be true on every page and prove nothing.
            hits = [t for t in OTHER_TAGS if any(t in b for b in body)]
            if hits:
                stray_tag.append(f"{rel}: {hits}")
        if not all("Version 1.0" in b and "Last Updated:" in b for b in body):
            no_ver.append(rel)
        if not all(DISC in b for b in body):
            no_disc.append(rel)

    covers = " ".join(((dr.pages[0].extract_text() or "")
                       + (dr.pages[-1].extract_text() or "")).split())
    if TAG_EXPECTED and TAG in covers:
        stamped_cover.append(rel)

    # Wrong-kit sentinel: another brand's name must never appear anywhere.
    whole = covers + " " + " ".join(body)
    bad_brands = [b for b in OTHER_BRANDS if b in whole]
    if bad_brands:
        foreign_brand.append(f"{rel}: {bad_brands}")

    # attribution: what the kit's rules say should print, and whether it did
    expect = S.clean_manufacturer(r.get("manufacturer"), BRAND, AL)
    if expect:
        line = f"Manufactured by {expect} | Sold by {BRAND}"
        if body and all(line in b for b in body):
            attrib_printed += 1
        else:
            wrong_attrib.append(f"{rel}: expected {line!r}")
    else:
        attrib_absent += 1
        if body and any("Manufactured by" in b for b in body):
            wrong_attrib.append(f"{rel}: expected NO attribution "
                                f"(manufacturer {r.get('manufacturer')!r})")

    sw, sh = page_size(sr.pages[0])
    dw, dh = page_size(dr.pages[1])
    if abs(dw - sw) > 1 or dh <= sh:
        geom.append(f"{rel} src {sw:.0f}x{sh:.0f} -> body {dw:.0f}x{dh:.0f}")

chk(not lost_text, f"source text preserved page-for-page: {len(lost_text)} lost  {lost_text[:4]}")
if TAG_EXPECTED:
    chk(not no_tag, f"tagline on every content page: {len(no_tag)} missing  {no_tag[:4]}")
else:
    chk(not stray_tag, f"no tagline stamped (this kit has none): "
                       f"{len(stray_tag)} carry one  {stray_tag[:4]}")
chk(not foreign_brand, f"no other brand's name anywhere (wrong-kit sentinel): "
                       f"{len(foreign_brand)}  {foreign_brand[:4]}")
chk(not no_ver, f"version + 'Last Updated:' on every content page: {len(no_ver)} missing  {no_ver[:4]}")
chk(not no_disc, f"disclaimer on every content page: {len(no_disc)} missing  {no_disc[:4]}")
chk(not stamped_cover, f"covers carry no page stamps: {len(stamped_cover)}  {stamped_cover[:4]}")
chk(not geom, f"pages extended, never overlaid or resized: {len(geom)}  {geom[:4]}")
chk(not wrong_attrib, f"attribution matches the kit's rules: {len(wrong_attrib)} wrong  {wrong_attrib[:4]}")

print(f"\nattribution in the inspected sample: {attrib_printed} printed, "
      f"{attrib_absent} correctly blank")

# whole-set projection of how many files will carry an attribution line
proj = sum(1 for r in reb if S.clean_manufacturer(r.get("manufacturer"), BRAND, AL))
print(f"across all {len(reb)} branded files, {proj} should carry an attribution line "
      f"and {len(reb) - proj} none")

print("\n" + "=" * 68)
print(f"DELIVERY QA: {'ALL CHECKS PASSED' if not fails else f'{len(fails)} FAILURES'}")
for m in fails:
    print("  FAIL  " + m)
for m in warns:
    print("  warn  " + m)
print("=" * 68)
sys.exit(1 if fails else 0)
