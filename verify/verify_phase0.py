"""Phase 0 verification: production rebrand engine — searchable, <50MB, metadata set."""
import sys, time
from pathlib import Path

REPO, BK, OUT = Path(sys.argv[1]), Path(sys.argv[2]), Path(sys.argv[3])
sys.path.insert(0, str(REPO))
from pypdf import PdfReader
from docrefine.rebrand import BrandKit, rebrand_pdf

OUT.mkdir(parents=True, exist_ok=True)
kit = BrandKit(BK)   # BK contains Portrait/ and Landscaape/ subfolders

results = []
def check(name, ok, detail=""):
    results.append(ok); print(f"[{'PASS' if ok else 'FAIL'}] {name}  {detail}")

check("brand kit: portrait assets found", kit.has("portrait"))
check("brand kit: landscape assets found", kit.has("landscape"))

jobs = [
    (BK / "1570_FCBU_Installation_Instructions.pdf", "Assembly Instructions", 98),
    (BK / "1570-4T5-BM.pdf", "Cluster Box Unit", 200),
]
for src, title, min_text in jobs:
    out = OUT / f"P0_{src.stem}.pdf"
    t0 = time.time()
    stats = rebrand_pdf(src, out, kit, title, author="Budget Mailboxes")
    dt = time.time() - t0
    r = PdfReader(str(out))
    # searchable? check a content page (index 1, right after the cover)
    txt = (r.pages[1].extract_text() or "").strip()
    author = (r.metadata or {}).get("/Author")
    raw = out.read_bytes()
    embedded = b"FontFile" in raw
    print(f"\n-- {src.name} --  {stats}  {dt:.1f}s")
    check(f"{src.stem}: output created", out.exists())
    check(f"{src.stem}: UNDER 50 MB", stats["size_mb"] < 50, f"{stats['size_mb']} MB")
    check(f"{src.stem}: text still searchable", len(txt) >= min_text, f"{len(txt)} chars")
    check(f"{src.stem}: metadata Author=Budget Mailboxes", author == "Budget Mailboxes", f"{author!r}")
    check(f"{src.stem}: fonts embedded", embedded)

print("\n" + "=" * 52)
print(f"RESULT: {sum(results)}/{len(results)} passed")
sys.exit(0 if all(results) else 1)
