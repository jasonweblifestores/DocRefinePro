"""Regression check: original engine flows still work after the rebranding feature."""
import sys, json, shutil, tempfile, importlib
from pathlib import Path
REPO = Path(sys.argv[1]); sys.path.insert(0, str(REPO))

results = []
def check(n, ok, d=""):
    results.append(ok); print(f"[{'PASS' if ok else 'FAIL'}] {n}  {d}")

# every module imports (incl GUI + new rebrand modules)
for m in ["docrefine.config", "docrefine.core.events", "docrefine.processing",
          "docrefine.reporting", "docrefine.worker", "docrefine.rebrand", "docrefine.classify",
          "docrefine.gui.qt_adapter", "docrefine.gui.forensic", "docrefine.gui.dialogs",
          "docrefine.gui.main_window", "docrefine.gui.app_qt"]:
    try:
        importlib.import_module(m); check(f"import {m}", True)
    except Exception as e:
        check(f"import {m}", False, repr(e))

from docrefine.worker import Worker, promote_duplicate_to_master
from docrefine.reporting import generate_job_report
from docrefine.config import WORKSPACES_ROOT
from PIL import Image

work = Path(tempfile.mkdtemp(prefix="drp_reg_")); src = work / "source"; src.mkdir()
Image.new("RGB", (200, 200), "white").save(src / "a.pdf", "PDF")
shutil.copy2(src / "a.pdf", src / "b.pdf")               # duplicate
Image.new("RGB", (120, 120), "blue").save(src / "pic.jpg", "JPEG")
(src / "broken.pdf").write_bytes(b"")                     # -> quarantine

Worker(callback=lambda e: None).run_inventory(str(src), "Standard")
ws = sorted(WORKSPACES_ROOT.glob("source_*"), key=lambda p: p.stat().st_mtime)[-1]
manifest = json.loads((ws / "manifest.json").read_text(encoding="utf-8"))
stats = json.loads((ws / "stats.json").read_text(encoding="utf-8"))
masters = {k: v for k, v in manifest.items() if v.get("status") != "QUARANTINE"}
quar = {k: v for k, v in manifest.items() if v.get("status") == "QUARANTINE"}
check("ingest: 2 masters", stats.get("masters") == 2, f"masters={stats.get('masters')}")
check("ingest: 1 quarantined recorded", len(quar) == 1)
pdf_master = next((v for v in masters.values() if v["name"] == "a.pdf"), None)
check("dedup grouped a.pdf/b.pdf", pdf_master and len(pdf_master["copies"]) == 2)

rpt = generate_job_report(ws, "Regression Batch", file_results=[
    {"file": "a.pdf", "orig_size": 1000, "new_size": 500, "ok": True},
    {"file": "bad.pdf", "orig_size": 100, "new_size": 0, "ok": False, "error": "boom"}])
check("report generated", rpt and Path(rpt).exists())
check("report errors table rendered", rpt and "boom" in Path(rpt).read_text(encoding="utf-8"))

root = Path(pdf_master["root"])
dup_rel = next(c for c in pdf_master["copies"] if c != pdf_master["master"])
new_id = promote_duplicate_to_master(ws, pdf_master["uid"], root / dup_rel)
m2 = {k: v for k, v in json.loads((ws / "manifest.json").read_text(encoding="utf-8")).items() if v.get("status") != "QUARANTINE"}
check("promote adds new master", new_id == "[0003]" and len(m2) == 3, f"id={new_id} n={len(m2)}")

shutil.rmtree(work, ignore_errors=True); shutil.rmtree(ws, ignore_errors=True)
print("\n" + "=" * 50); print(f"RESULT: {sum(results)}/{len(results)} passed")
sys.exit(0 if all(results) else 1)
