"""F4 fix: two source PDFs that map to the same output name must both survive (no data loss)."""
import sys, csv, shutil, tempfile
from pathlib import Path
REPO, BK = Path(sys.argv[1]), Path(sys.argv[2])
sys.path.insert(0, str(REPO))
from pypdf import PdfReader
from reportlab.pdfgen import canvas as rlc
from docrefine.worker import Worker

results = []
def check(n, ok, d=""):
    results.append(ok); print(f"[{'PASS' if ok else 'FAIL'}] {n}  {d}")

def text_pdf(path, text):
    c = rlc.Canvas(str(path), pagesize=(612, 792)); c.setFont("Helvetica", 24)
    c.drawString(90, 700, text); c.showPage(); c.save()

work = Path(tempfile.mkdtemp(prefix="drp_col_")); src = work / "batch"; src.mkdir()
text_pdf(src / "one.pdf", "DOCUMENT ONE ALPHA")
text_pdf(src / "two.pdf", "DOCUMENT TWO BRAVO")   # different content, SAME product/asset -> same base name

plan = src / "_rebrand_plan.csv"
cols = Worker.REBRAND_PLAN_COLUMNS
with open(plan, "w", newline="", encoding="utf-8-sig") as f:
    w = csv.DictWriter(f, fieldnames=cols); w.writeheader()
    for fn in ("one.pdf", "two.pdf"):
        w.writerow({"file": fn, "action": "rebrand", "product": "Widget", "asset_type": "manual",
                    "manufacturer": "Acme", "title": "Widget Manual"})

out = work / "out"
Worker(callback=lambda e: None).run_rebrand_apply(str(src), str(BK), str(plan), out_dir=str(out))

produced = sorted(p.name for p in out.rglob("*.pdf"))
check("both colliding docs produced a file (no overwrite)", len(produced) == 2, str(produced))
# v142: a contested product/asset name makes BOTH files fall back to their source
# name, so neither ends up as a meaningless "-2" and neither wins by sheet order.
check("names disambiguated by source name, not a number",
      produced == ["one-manual-budget-mailboxes.pdf",
                   "two-manual-budget-mailboxes.pdf"], str(produced))
if len(produced) == 2:
    texts = set()
    for p in out.rglob("*.pdf"):
        texts.add(" ".join((pg.extract_text() or "") for pg in PdfReader(str(p)).pages).replace(" ", ""))
    check("both distinct source contents preserved", ("DOCUMENTONEALPHA" in "".join(texts)) and ("DOCUMENTTWOBRAVO" in "".join(texts)))

# resume: re-run must skip (deterministic names), not create -3/-4 duplicates
Worker(callback=lambda e: None).run_rebrand_apply(str(src), str(BK), str(plan), out_dir=str(out))
check("re-run is stable (still exactly 2 files)", len(list(out.rglob('*.pdf'))) == 2)

shutil.rmtree(work, ignore_errors=True)
print("\n" + "=" * 50); print(f"RESULT: {sum(results)}/{len(results)} passed")
sys.exit(0 if all(results) else 1)
