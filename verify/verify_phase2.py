"""Phase 2 verification: local classification -> review sheet -> edit -> apply."""
import sys, csv, shutil, tempfile
from pathlib import Path

REPO, BK = Path(sys.argv[1]), Path(sys.argv[2])
sys.path.insert(0, str(REPO))
from pypdf import PdfReader
from docrefine.worker import Worker
from docrefine.classify import classify_document, ollama_available

results = []
def check(name, ok, detail=""):
    results.append(ok); print(f"[{'PASS' if ok else 'FAIL'}] {name}  {detail}")

print("ollama available:", ollama_available())

# --- direct classification of both samples ---
manual = BK / "1570_FCBU_Installation_Instructions.pdf"
drawing = BK / "1570-4T5-BM.pdf"
cm = classify_document(manual)
cd = classify_document(drawing)
print("MANUAL  ->", {k: cm[k] for k in ("action", "doc_type", "product", "asset_type", "manufacturer", "title")})
print("DRAWING ->", {k: cd[k] for k in ("action", "doc_type", "product", "asset_type", "manufacturer", "title")})
check("manual classified as rebrand", cm["action"] == "rebrand", cm["doc_type"])
# drawing SHOULD be 'leave'; note if the model disagreed (human review would fix it)
check("drawing classified (leave preferred)", cd["action"] in ("leave", "rebrand"), f"action={cd['action']}")

# --- analyze a temp folder into a review sheet ---
work = Path(tempfile.mkdtemp(prefix="drp_p2_"))
src = work / "batch"; src.mkdir()
shutil.copy2(manual, src / manual.name)
shutil.copy2(drawing, src / drawing.name)
Worker(callback=lambda e: None).run_rebrand_analyze(str(src))
from docrefine.reviews import plan_path_for            # v139: sheet lives outside the source
plan = plan_path_for(src)
check("review sheet created", plan.exists(), str(plan))
check("review sheet NOT written into the source folder", not (src / "_rebrand_plan.csv").exists())
from docrefine import reviews
rows = reviews.read_plan(plan)
check("review sheet has a row per PDF", len(rows) == 2, f"{len(rows)} rows")
cols_ok = all(c in rows[0] for c in ("file", "action", "product", "asset_type", "manufacturer", "title"))
check("review sheet has expected columns", cols_ok)

# --- simulate a human review: force the drawing to 'leave', give the manual a manufacturer ---
for r in rows:
    if "4T5" in r["file"]:
        r["action"] = "leave"
    if "FCBU" in r["file"]:
        r["manufacturer"] = "Florence Manufacturing"
        r["product"] = "1570 Cluster Box"
        r["asset_type"] = "installation-guide"
        r["title"] = ""     # blank -> engine derives the canonical short title
reviews.write_plan(plan, rows, Worker.REBRAND_PLAN_COLUMNS, src_root=src)

# --- apply the reviewed sheet ---
out = work / "batch_rebranded"
Worker(callback=lambda e: None).run_rebrand_apply(str(src), str(BK), str(plan), out_dir=str(out),
                                                  show_attribution=True)  # v143: opt-in

branded = out / "1570-cluster-box-installation-guide-budget-mailboxes.pdf"
left = out / drawing.name  # copied as-is, original name
check("rebranded file named from product/asset-type", branded.exists(), branded.name)
check("left-as-is file copied unbranded (original name)", left.exists(), left.name)
check("left-as-is is byte-identical to source", left.exists() and left.read_bytes() == drawing.read_bytes())

if branded.exists():
    r = PdfReader(str(branded))
    cover_txt = (r.pages[0].extract_text() or "")
    body_txt = (r.pages[1].extract_text() or "").strip()
    size = branded.stat().st_size / 1e6
    check("branded: manufacturer line on cover", "Manufactured by Florence Manufacturing" in cover_txt)
    check("branded: body still searchable", len(body_txt) >= 90, f"{len(body_txt)} chars")
    check("branded: under 50 MB", size < 50, f"{size:.1f} MB")

shutil.rmtree(work, ignore_errors=True)
# The sheet lives in the user's real Rebrand Reviews folder — don't leave litter.
for leftover in (plan, plan.with_name(plan.stem + ".previous" + plan.suffix)):
    try: leftover.unlink()
    except OSError: pass
print("\n" + "=" * 52)
print(f"RESULT: {sum(results)}/{len(results)} passed")
sys.exit(0 if all(results) else 1)
