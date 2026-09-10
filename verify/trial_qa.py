import sys
from pathlib import Path
import batch_paths                      # also puts the repo on sys.path
from pypdf import PdfReader
from docrefine import reviews, stamps as S
from docrefine.rebrand import BrandKit, page_size

SP = Path(sys.argv[1]); OUT = SP / "TRIAL_output"
SRC = batch_paths.SRC
kit = BrandKit(batch_paths.require_kit())
al = kit.brand.get("manufacturer_aliases") or {}
rows = {r["file"]: r for r in reviews.read_plan(SP / "trial_plan.xlsx")}
# From the kit, not a literal — TAG_EXPECTED is False for a kit with no tagline
# (MailboxWorks), where `TAG in text` would be vacuously true on every page.
TAG, TAG_EXPECTED, DISC = batch_paths.wording()

fails = []
def chk(ok, msg):
    if not ok: fails.append(msg)
    print(("  ok   " if ok else "  FAIL ") + msg)

print(f"files produced: {len(list(OUT.rglob('*')))}")
for f in sorted(OUT.rglob("*")):
    if not f.is_file(): continue
    rel = f.relative_to(OUT).as_posix()
    row = rows.get(rel)
    src = SRC / rel
    action = (row.get("action") or "").lower()
    print(f"\n--- {rel}   [{action}]")
    chk(f.name == src.name, "filename unchanged from source")

    if action == "leave":
        chk(f.read_bytes() == src.read_bytes(), "byte-identical to source")
        continue

    sr, dr = PdfReader(str(src)), PdfReader(str(f))
    chk(len(dr.pages) == len(sr.pages) + 2, f"page count = source+2 ({len(sr.pages)}->{len(dr.pages)})")
    chk((dr.metadata or {}).get("/Author") == kit.brand_name,
        f"Author = {kit.brand_name} (found {(dr.metadata or {}).get('/Author')!r})")
    mb = f.stat().st_size / 1e6
    chk(mb < 50, f"under 50 MB ({mb:.1f})")

    body = ["".join((dr.pages[i].extract_text() or "")) for i in range(1, len(dr.pages) - 1)]
    joined = " ".join(" ".join(b.split()) for b in body)
    # Compare page against page. Comparing a slice of the whole document fails
    # spuriously as soon as it crosses a page boundary, because the stamps sit
    # between one page's text and the next.
    norm = lambda t: "".join((t or "").split())
    lost = [i + 1 for i in range(len(sr.pages))
            if norm(sr.pages[i].extract_text()) and
               norm(sr.pages[i].extract_text()) not in norm(dr.pages[i + 1].extract_text())]
    chk(not lost, f"every source page's text preserved in place ({len(sr.pages)}pp)"
                  + (f" — LOST ON {lost}" if lost else ""))
    # stamps on EVERY content page
    if TAG_EXPECTED:
        chk(all(TAG in " ".join(b.split()) for b in body), f"tagline on all {len(body)} content pages")
    else:
        chk(not any("Trusted by the Nation" in " ".join(b.split()) for b in body),
            "no tagline stamped (this kit has none)")
    chk(all("Version 1.0" in " ".join(b.split()) and "Last Updated:" in " ".join(b.split()) for b in body),
        "version + 'Last Updated:' on all content pages")
    chk(all(DISC in " ".join(b.split()) for b in body), "disclaimer on all content pages")
    covers = " ".join(((dr.pages[0].extract_text() or "")
                       + (dr.pages[-1].extract_text() or "")).split())
    # The disclaimer is a per-page stamp on every kit, so it works as the
    # cover-leak sentinel whether or not this brand has a tagline.
    chk(DISC not in covers and (not TAG_EXPECTED or TAG not in covers),
        "stamps NOT on covers")
    other = [b for b in ("Budget Mailboxes", "MailboxWorks") if b != kit.brand_name]
    chk(not [b for b in other if b in covers + " " + joined],
        f"no other brand's name present (expected {kit.brand_name!r})")

    expect = S.clean_manufacturer(row.get("manufacturer"), kit.brand_name, al)
    has = "Manufactured by" in joined
    if expect:
        line = f"Manufactured by {expect} | Sold by {kit.brand_name}"
        chk(all(line in " ".join(b.split()) for b in body), f"attribution: {line}")
    else:
        chk(not has, f"NO attribution (manufacturer {row.get('manufacturer')!r} unusable)")
    # geometry: content page must be taller than source, same width
    sw, sh = page_size(sr.pages[0]); dw, dh = page_size(dr.pages[1])
    chk(abs(dw - sw) < 1, f"width unchanged ({sw:.0f})")
    chk(dh > sh, f"page extended, not overlaid ({sh:.0f} -> {dh:.0f})")

print("\n" + "=" * 60)
print("TRIAL QA: " + ("ALL CHECKS PASSED" if not fails else f"{len(fails)} FAILURES"))
for m in fails: print("  - " + m)
sys.exit(1 if fails else 0)
