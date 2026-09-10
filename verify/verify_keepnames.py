"""v146: 'Keep the original filenames' — the Batch 1/2 convention, as an option."""
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
work = Path(tempfile.mkdtemp(prefix="drp_keep_"))
w = Worker(callback=lambda e: None)

def text_pdf(path, body="INSTALLATION AND ASSEMBLY GUIDE"):
    path.parent.mkdir(parents=True, exist_ok=True)
    c = rlc.Canvas(str(path), pagesize=(612, 792)); c.setFont("Helvetica", 20)
    c.drawString(70, 700, body); c.showPage(); c.save()

# =====================================================================
#  Defaults: the brief's pattern, since that is what Batch 4 asks for
# =====================================================================
check("K1 config default keeps renaming on", ConfigData().rebrand_keep_original_names is False)
for fn in (Worker.run_rebrand_apply, Worker.run_pipeline):
    p = inspect.signature(fn).parameters
    check(f"K2 {fn.__name__} exposes the flag, default off",
          "keep_original_names" in p and p["keep_original_names"].default is False)

# =====================================================================
#  Name assignment both ways
# =====================================================================
ROWS = [
    {"file": "Imperial-Street-Sign-Brochure.pdf", "action": "rebrand",
     "product": "Imperial Street Signs", "asset_type": "specification-sheet"},
    {"file": "sub/1570-12-BM.pdf", "action": "rebrand",
     "product": "", "asset_type": "installation-guide"},
    {"file": "tech-3625rl.pdf", "action": "leave"},
]
out = Path("C:/out") if sys.platform == "win32" else Path("/out")

renamed = [dict(r) for r in ROWS]
w._assign_output_names(renamed, out, kit, keep_original_names=False)
kept = [dict(r) for r in ROWS]
w._assign_output_names(kept, out, kit, keep_original_names=True)

rn = [Path(r["_dst"]).name for r in renamed]
kn = [Path(r["_dst"]).name for r in kept]
check("K3 renaming still follows the brief", rn[0].endswith("-budget-mailboxes.pdf"), rn[0])
check("K4 keeping names returns the source name exactly",
      kn[0] == "Imperial-Street-Sign-Brochure.pdf", kn[0])
check("K5 original case and punctuation preserved",
      kn[0] == ROWS[0]["file"], f"{kn[0]!r} vs {ROWS[0]['file']!r}")
check("K6 no brand suffix is appended when keeping names",
      "budget-mailboxes" not in kn[0].lower() and "rebrand" not in kn[0].lower(), kn[0])
check("K7 nested paths keep their folder", Path(kept[1]["_dst"]).parent.name == "sub"
      and Path(kept[1]["_dst"]).name == "1570-12-BM.pdf", str(Path(kept[1]["_dst"]).relative_to(out)))
check("K8 leave rows are unaffected either way",
      rn[2] == kn[2] == "tech-3625rl.pdf", f"{rn[2]} / {kn[2]}")
check("K9 every path is still unique", len(set(kn)) == len(kn))

# a name that would have been truncated by the delivery pattern survives intact
check("K10 no truncation when keeping names",
      "Imperial-Street-Sign-Brochure" in kn[0] and "imperial-street-specification" in rn[0],
      f"kept={kn[0]} | renamed={rn[0]}")

# =====================================================================
#  End to end: the file on disk really carries the original name
# =====================================================================
src = work / "src"
text_pdf(src / "Imperial-Street-Sign-Brochure.pdf")
text_pdf(src / "sub" / "1570-12-BM.pdf")
rows = [dict(r) for r in ROWS[:2]]
plan = work / "plan.xlsx"
reviews.write_plan(plan, rows, Worker.REBRAND_PLAN_COLUMNS, src_root=src)

o_keep = work / "out_keep"
Worker(callback=lambda e: None).run_rebrand_apply(
    str(src), str(BK), str(plan), out_dir=str(o_keep), keep_original_names=True)
made = sorted(p.relative_to(o_keep).as_posix() for p in o_keep.rglob("*.pdf"))
check("K11 output files carry the original names",
      made == ["Imperial-Street-Sign-Brochure.pdf", "sub/1570-12-BM.pdf"], str(made))

f = o_keep / "Imperial-Street-Sign-Brochure.pdf"
r = PdfReader(str(f))
check("K12 they are genuinely rebranded, not copied", len(r.pages) == 3, f"{len(r.pages)} pages")
cover = " ".join((r.pages[0].extract_text() or "").split()).upper()
check("K13 cover still titled from the asset type", "SPECIFICATION SHEET" in cover, cover[:40])
check("K14 author metadata still set", (r.metadata or {}).get("/Author") == "Budget Mailboxes")
check("K15 body text preserved",
      "GUIDE" in (r.pages[1].extract_text() or "").upper())
check("K16 the output is NOT byte-identical to the source (it was branded)",
      f.read_bytes() != (src / "Imperial-Street-Sign-Brochure.pdf").read_bytes())

o_ren = work / "out_ren"
Worker(callback=lambda e: None).run_rebrand_apply(
    str(src), str(BK), str(plan), out_dir=str(o_ren), keep_original_names=False)
ren = sorted(p.name for p in o_ren.rglob("*.pdf"))
check("K17 renaming mode still produces delivery names",
      all(n.endswith("-budget-mailboxes.pdf") for n in ren), str(ren))

# =====================================================================
#  Wired into both dialogs
# =====================================================================
prev = CFG.get("rebrand_keep_original_names")
try:
    CFG._data.rebrand_keep_original_names = False
    d = RebrandDialog(None, default_kit=str(BK), default_source=str(src))
    check("K18 RebrandDialog has the checkbox, unticked by default",
          hasattr(d, "chk_keepnames") and not d.chk_keepnames.isChecked())
    d.plan_path = str(plan); d.txt_plan.setText(str(plan)); d._sync_apply()
    d.chk_keepnames.setChecked(True); d.on_apply()
    check("K19 ticking it is carried out of the dialog", d.keep_original_names is True)
    d.deleteLater()
    p = PipelineDialog(None, default_kit=str(BK))
    check("K20 PipelineDialog has it too", hasattr(p, "chk_keepnames")
          and not p.chk_keepnames.isChecked())
    p.deleteLater()
finally:
    CFG._data.rebrand_keep_original_names = prev

shutil.rmtree(work, ignore_errors=True)
print("\n" + "=" * 56)
print(f"RESULT: {sum(results)}/{len(results)} passed")
sys.exit(0 if all(results) else 1)
