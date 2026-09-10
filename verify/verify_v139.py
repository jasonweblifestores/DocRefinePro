"""v139 verification: (1) Complete-set output option, (2) review sheet relocated + auto-found."""
import sys, csv, shutil, tempfile
from pathlib import Path

REPO, BK = Path(sys.argv[1]), Path(sys.argv[2])
sys.path.insert(0, str(REPO))
from reportlab.pdfgen import canvas as rlc
from docrefine.worker import Worker
from docrefine.config import REVIEWS_ROOT
from docrefine import reviews

results = []
def check(n, ok, d=""):
    results.append(bool(ok)); print(f"[{'PASS' if ok else 'FAIL'}] {n}  {d}")

def text_pdf(path, text):
    path.parent.mkdir(parents=True, exist_ok=True)
    c = rlc.Canvas(str(path), pagesize=(612, 792)); c.setFont("Helvetica", 24)
    c.drawString(90, 700, text); c.showPage(); c.save()

def W():
    return Worker(callback=lambda e: None)

work = Path(tempfile.mkdtemp(prefix="drp_v139_"))
written_sheets = []

# =====================================================================
#  A. reviews module — naming, location, discovery order
# =====================================================================
job = work / "BM-Batch4"; masters = job / "01_Master_Files"; masters.mkdir(parents=True)
p = reviews.plan_path_for(masters)
check("A1 sheet goes to the Rebrand Reviews folder", p.parent == REVIEWS_ROOT, str(p.parent))
check("A2 sheet name carries job + folder", p.name == "BM-Batch4__01_Master_Files_rebrand_plan.xlsx", p.name)

job2 = work / "BM-Batch5"; masters2 = job2 / "01_Master_Files"; masters2.mkdir(parents=True)
check("A3 same-named masters folders don't collide",
      reviews.plan_path_for(masters2).name != p.name, reviews.plan_path_for(masters2).name)

check("A4 no sheet found before analyze", reviews.find_plan(masters) is None)
legacy = masters / "_rebrand_plan.csv"; legacy.write_text("file\n", encoding="utf-8")
check("A5 pre-v139 in-source sheet still found", reviews.find_plan(masters) == legacy)
sibling = job / "01_Master_Files_rebrand_plan.csv"; sibling.write_text("file\n", encoding="utf-8")
check("A6 sheet beside the source wins over the legacy one", reviews.find_plan(masters) == sibling)
p.write_text("file\n", encoding="utf-8"); written_sheets.append(p)
check("A7 Reviews-folder sheet wins over both", reviews.find_plan(masters) == p)
check("A8 is_plan_file spots sheets by either name/format",
      reviews.is_plan_file("_rebrand_plan.csv") and reviews.is_plan_file("X__Y_rebrand_plan.csv")
      and reviews.is_plan_file("X__Y_rebrand_plan.xlsx") and not reviews.is_plan_file("data.csv"))
p.unlink(); sibling.unlink(); legacy.unlink()

# =====================================================================
#  B. Analyze writes to the new home (and keeps a reviewed sheet safe)
# =====================================================================
src = work / "corpus"
text_pdf(src / "manual.pdf", "INSTALLATION AND ASSEMBLY GUIDE ALPHA")
text_pdf(src / "sub" / "spec.pdf", "PRODUCT SPECIFICATION SHEET BRAVO")
(src / "notes.txt").write_text("loose note", encoding="utf-8")
(src / "sub" / "photo.jpg").write_bytes(b"\xff\xd8\xff\xe0not-a-real-jpeg")
(src / "sub" / "pricing.xlsx").write_bytes(b"PK\x03\x04not-a-real-xlsx")

W().run_rebrand_analyze(str(src))
plan = reviews.plan_path_for(src); written_sheets.append(plan)
check("B1 analyze wrote the sheet to the Reviews folder", plan.exists(), str(plan))
check("B2 analyze left nothing in the source folder", not (src / "_rebrand_plan.csv").exists())
check("B3 apply finds the sheet on its own", reviews.find_plan(src) == plan)
rows = reviews.read_plan(plan)
check("B4 one row per PDF, non-PDFs ignored", len(rows) == 2, f"{len(rows)} rows")

COLS = Worker.REBRAND_PLAN_COLUMNS
def save_reviewed():
    reviews.write_plan(plan, rows, COLS, src_root=src)

for r in rows:                      # simulate the human review
    r["action"] = "rebrand"; r["manufacturer"] = "Acme"
    r["product"] = Path(r["file"]).stem; r["asset_type"] = "guide"
save_reviewed()

W().run_rebrand_analyze(str(src))   # re-analyze must not silently destroy the review
bak = plan.with_name(plan.stem + ".previous" + plan.suffix); written_sheets.append(bak)
check("B5 re-analyze backs up the reviewed sheet", bak.exists(), bak.name)
check("B6 backup holds the reviewed edits",
      any((r.get("manufacturer") or "") == "Acme" for r in reviews.read_plan(bak)))

save_reviewed()                     # restore the reviewed version

# =====================================================================
#  C. Apply — PDFs only (default) vs Complete set
# =====================================================================
out_pdf = work / "out_pdfs_only"
W().run_rebrand_apply(str(src), str(BK), None, out_dir=str(out_pdf))   # plan auto-found
names = sorted(q.name for q in out_pdf.rglob("*") if q.is_file())
check("C1 apply ran with no sheet passed in", len(names) > 0, str(names))
check("C2 PDFs only: 2 branded PDFs, nothing else", names == [
    "manual-guide-budget-mailboxes.pdf", "spec-guide-budget-mailboxes.pdf"], str(names))

out_all = work / "out_complete"
W().run_rebrand_apply(str(src), str(BK), None, out_dir=str(out_all), complete_set=True)
rel = sorted(str(q.relative_to(out_all)).replace("\\", "/") for q in out_all.rglob("*") if q.is_file())
check("C3 complete set: branded PDFs + every non-PDF", rel == [
    "manual-guide-budget-mailboxes.pdf", "notes.txt",
    "sub/photo.jpg", "sub/pricing.xlsx", "sub/spec-guide-budget-mailboxes.pdf"], str(rel))
check("C4 copied non-PDFs are byte-identical",
      (out_all / "sub" / "photo.jpg").read_bytes() == (src / "sub" / "photo.jpg").read_bytes()
      and (out_all / "notes.txt").read_text(encoding="utf-8") == "loose note")
check("C5 nested folder structure mirrored", (out_all / "sub").is_dir())
check("C5b no app artifacts (stats.json) in the delivery tree",
      not list(out_all.rglob("stats.json")) and not list(out_pdf.rglob("stats.json")))

before = sorted(q.name for q in out_all.rglob("*") if q.is_file())
W().run_rebrand_apply(str(src), str(BK), None, out_dir=str(out_all), complete_set=True)
check("C6 re-run is stable (resume-safe, no duplicates)",
      sorted(q.name for q in out_all.rglob("*") if q.is_file()) == before)

# a stray legacy sheet in the source must never land in the upload set
(src / "_rebrand_plan.csv").write_text("file\n", encoding="utf-8")
out_np = work / "out_noplan"
W().run_rebrand_apply(str(src), str(BK), str(plan), out_dir=str(out_np), complete_set=True)
check("C7 review sheets are never copied into the output",
      not any(reviews.is_plan_file(q.name) for q in out_np.rglob("*")),
      str(sorted(q.name for q in out_np.rglob('*') if q.is_file())))
(src / "_rebrand_plan.csv").unlink()

# a PDF added after analysis is reported, not silently dropped
logs = []
w2 = Worker(callback=lambda e: None); w2.log = lambda m, err=False: logs.append(m)
text_pdf(src / "late.pdf", "ADDED AFTER ANALYSIS")
w2.run_rebrand_apply(str(src), str(BK), str(plan), out_dir=str(work / "out_late"), complete_set=True)
check("C8 un-analyzed PDFs are flagged in the log",
      any("not in the review sheet" in m for m in logs), "")
(src / "late.pdf").unlink()

# =====================================================================
#  D. Pipeline honours the same option
# =====================================================================
psrc = work / "pipe"
text_pdf(psrc / "doc.pdf", "PIPELINE DOCUMENT CHARLIE")
(psrc / "readme.txt").write_text("keep me", encoding="utf-8")

pout = work / "pipe_complete"
W().run_pipeline(str(psrc), do_flatten=False, do_rebrand=True, do_ocr=False,
                 kit_dir=str(BK), out_dir=str(pout), complete_set=True)
pnames = sorted(q.name for q in pout.rglob("*") if q.is_file())
check("D1 pipeline complete set keeps non-PDFs", "readme.txt" in pnames, str(pnames))

pout2 = work / "pipe_pdfs"
W().run_pipeline(str(psrc), do_flatten=False, do_rebrand=True, do_ocr=False,
                 kit_dir=str(BK), out_dir=str(pout2), complete_set=False)
pnames2 = sorted(q.name for q in pout2.rglob("*") if q.is_file())
check("D2 pipeline PDFs-only drops non-PDFs", "readme.txt" not in pnames2 and len(pnames2) == 1, str(pnames2))

# the pipeline's rebrand step reads the relocated sheet too
psrc2 = work / "pipe2"
text_pdf(psrc2 / "leaveme.pdf", "TECHNICAL DRAWING DELTA")
p2 = reviews.plan_path_for(psrc2); written_sheets.append(p2)
reviews.write_plan(p2, [{"file": "leaveme.pdf", "action": "leave"}],
                   Worker.REBRAND_PLAN_COLUMNS, src_root=psrc2)
pout3 = work / "pipe2_out"
W().run_pipeline(str(psrc2), do_flatten=False, do_rebrand=True, do_ocr=False,
                 kit_dir=str(BK), out_dir=str(pout3), complete_set=True)
check("D3 pipeline uses the relocated sheet ('leave' honoured)",
      (pout3 / "leaveme.pdf").exists()
      and (pout3 / "leaveme.pdf").read_bytes() == (psrc2 / "leaveme.pdf").read_bytes(),
      str([q.name for q in pout3.rglob('*.pdf')]))

# =====================================================================
#  E. Config default
# =====================================================================
from docrefine.config import ConfigData
check("E1 complete set defaults on", ConfigData().rebrand_complete_set is True)

for s in written_sheets:            # leave the user's Reviews folder as we found it
    try: s.unlink()
    except OSError: pass
shutil.rmtree(work, ignore_errors=True)
print("\n" + "=" * 56)
print(f"RESULT: {sum(results)}/{len(results)} passed")
sys.exit(0 if all(results) else 1)
