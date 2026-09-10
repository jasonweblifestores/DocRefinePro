"""Phase 3 verification: composable pipeline (Flatten -> Rebrand -> OCR, OCR last)."""
import sys, shutil, tempfile
from pathlib import Path

REPO, BK = Path(sys.argv[1]), Path(sys.argv[2])
sys.path.insert(0, str(REPO))
from pypdf import PdfReader
from reportlab.pdfgen import canvas as rlc
from docrefine.worker import Worker
from docrefine import processing

results = []
def check(name, ok, detail=""):
    results.append(ok); print(f"[{'PASS' if ok else 'FAIL'}] {name}  {detail}")

print("HAS_TESSERACT:", processing.HAS_TESSERACT, "| POPPLER:", bool(processing.POPPLER_BIN))

def make_text_pdf(path):
    c = rlc.Canvas(str(path), pagesize=(612, 792))
    c.setFont("Helvetica", 26); c.drawString(90, 700, "HELLO PIPELINE TEST 12345")
    c.setFont("Helvetica", 18); c.drawString(90, 650, "Installation and assembly guide")
    c.showPage(); c.save()

# --- Full chain: Flatten -> Rebrand -> OCR ---
work = Path(tempfile.mkdtemp(prefix="drp_p3_"))
src = work / "batch"; src.mkdir()
make_text_pdf(src / "sample_doc.pdf")

Worker(callback=lambda e: None).run_pipeline(str(src), do_flatten=True, do_rebrand=True, do_ocr=True,
                                             kit_dir=str(BK), dpi=200, out_dir=str(work / "out_full"))
out_full = work / "out_full"
files = list(out_full.rglob("*.pdf"))
check("full pipeline produced output", len(files) == 1, str([f.name for f in files]))
if files:
    f = files[0]
    r = PdfReader(str(f))
    all_txt = " ".join((pg.extract_text() or "") for pg in r.pages)
    size = f.stat().st_size / 1e6
    check("output named as branded delivery", f.name.endswith("-budget-mailboxes.pdf"), f.name)
    check("flattened text RESTORED by OCR-last (searchable)", len(all_txt.strip()) >= 8, f"{len(all_txt.strip())} chars")
    check("full pipeline output under 50 MB", size < 50, f"{size:.1f} MB")

# --- Rebrand-only: text must be preserved directly (no OCR) ---
src2 = work / "batch2"; src2.mkdir()
make_text_pdf(src2 / "guide.pdf")
Worker(callback=lambda e: None).run_pipeline(str(src2), do_flatten=False, do_rebrand=True, do_ocr=False,
                                             kit_dir=str(BK), dpi=200, out_dir=str(work / "out_reb"))
reb = list((work / "out_reb").rglob("*.pdf"))
check("rebrand-only produced output", len(reb) == 1)
if reb:
    txt = " ".join((pg.extract_text() or "") for pg in PdfReader(str(reb[0])).pages)
    check("rebrand-only keeps original text (no OCR needed)", "PIPELINE TEST" in txt.replace("  ", " "), "")

shutil.rmtree(work, ignore_errors=True)
print("\n" + "=" * 50)
print(f"RESULT: {sum(results)}/{len(results)} passed")
sys.exit(0 if all(results) else 1)
