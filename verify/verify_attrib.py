"""v143: the cover attribution line is a toggle, off by default (Batch 1/2 house style)."""
import sys, inspect, tempfile, shutil
from pathlib import Path

REPO, BK = Path(sys.argv[1]), Path(sys.argv[2])
sys.path.insert(0, str(REPO))

results = []
def check(n, ok, d=""):
    results.append(bool(ok)); print(f"[{'PASS' if ok else 'FAIL'}] {n}  {d}")

from PySide6.QtWidgets import QApplication
app = QApplication.instance() or QApplication([])

from pypdf import PdfReader
from reportlab.pdfgen import canvas as rlc
from docrefine import reviews
from docrefine.worker import Worker
from docrefine.rebrand import BrandKit
from docrefine.config import CFG, ConfigData
from docrefine.gui.dialogs import RebrandDialog, PipelineDialog

kit = BrandKit(BK)
work = Path(tempfile.mkdtemp(prefix="drp_attr_"))

def text_pdf(path, body="INSTALLATION AND ASSEMBLY GUIDE FOR THE UNIT"):
    path.parent.mkdir(parents=True, exist_ok=True)
    c = rlc.Canvas(str(path), pagesize=(612, 792)); c.setFont("Helvetica", 20)
    c.drawString(70, 700, body); c.showPage(); c.save()

# =====================================================================
#  Default is OFF — matching the signed-off Batch 1 and 2 covers
# =====================================================================
check("A1 config default is off", ConfigData().rebrand_show_attribution is False)
sig = inspect.signature(Worker.run_rebrand_apply).parameters
check("A2 run_rebrand_apply exposes the flag, default off",
      "show_attribution" in sig and sig["show_attribution"].default is False)
check("A3 run_pipeline exposes the flag, default off",
      inspect.signature(Worker.run_pipeline).parameters["show_attribution"].default is False)

# =====================================================================
#  What lands on the cover
# =====================================================================
src = work / "src"
text_pdf(src / "guide.pdf")
rows = [{"file": "guide.pdf", "action": "rebrand", "product": "1570 CBU",
         "asset_type": "installation-guide", "manufacturer": "Florence Corporation",
         "title": "", "source": "llm", "confidence": "0.95"}]
plan = work / "plan.xlsx"
reviews.write_plan(plan, rows, Worker.REBRAND_PLAN_COLUMNS, src_root=src)

def cover_text(out_dir, show):
    Worker(callback=lambda e: None).run_rebrand_apply(
        str(src), str(BK), str(plan), out_dir=str(out_dir), show_attribution=show)
    made = list(Path(out_dir).rglob("*.pdf"))
    assert len(made) == 1, made
    return " ".join((PdfReader(str(made[0])).pages[0].extract_text() or "").split()).replace(" ", "")

off = cover_text(work / "out_off", False)
check("A4 attribution absent when off", "Manufactured" not in off and "Soldby" not in off, off[:70])
check("A5 the title still renders when off", "INSTALLATIONMANUAL" in off, off[:70])

on = cover_text(work / "out_on", True)
check("A6 attribution present when on", "ManufacturedbyFlorenceCorporation" in on, on[:80])
check("A7 and names the seller", "SoldbyBudgetMailboxes" in on)
check("A8 the title is unchanged by the toggle", "INSTALLATIONMANUAL" in on)

# =====================================================================
#  The toggle must not disturb anything else about the deliverable
# =====================================================================
a = list((work / "out_off").rglob("*.pdf"))[0]
b = list((work / "out_on").rglob("*.pdf"))[0]
check("A9 filenames are identical either way", a.name == b.name, a.name)
ra, rb = PdfReader(str(a)), PdfReader(str(b))
check("A10 page counts match", len(ra.pages) == len(rb.pages) == 3)
check("A11 Author metadata unaffected",
      (ra.metadata or {}).get("/Author") == (rb.metadata or {}).get("/Author") == "Budget Mailboxes")
check("A12 body text preserved either way",
      "GUIDE" in (ra.pages[1].extract_text() or "").upper()
      and "GUIDE" in (rb.pages[1].extract_text() or "").upper())

# =====================================================================
#  Wired into both dialogs
# =====================================================================
prev = CFG.get("rebrand_show_attribution")
try:
    CFG._data.rebrand_show_attribution = False
    d = RebrandDialog(None, default_kit=str(BK), default_source=str(src))
    check("A13 RebrandDialog has the checkbox, unticked by default",
          hasattr(d, "chk_attrib") and not d.chk_attrib.isChecked())
    # give it a sheet so on_apply() validates instead of popping a blocking warning
    d.plan_path = str(plan); d.txt_plan.setText(str(plan)); d._sync_apply()
    d.chk_attrib.setChecked(True); d.on_apply()
    check("A14 ticking it is carried out of the dialog", d.show_attribution is True)
    check("A14 and the mode is apply", d.mode == "apply", str(d.mode))
    d.deleteLater()
    p = PipelineDialog(None, default_kit=str(BK))
    check("A15 PipelineDialog has it too, unticked", hasattr(p, "chk_attrib")
          and not p.chk_attrib.isChecked())
    p.deleteLater()
finally:
    CFG._data.rebrand_show_attribution = prev

shutil.rmtree(work, ignore_errors=True)
print("\n" + "=" * 56)
print(f"RESULT: {sum(results)}/{len(results)} passed")
sys.exit(0 if all(results) else 1)
