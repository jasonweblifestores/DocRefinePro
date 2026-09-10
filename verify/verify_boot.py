"""Boot + GUI wiring: dialogs construct offscreen and expose the new complete-set option."""
import sys, inspect, tempfile
from pathlib import Path

REPO = Path(sys.argv[1])
sys.path.insert(0, str(REPO))

results = []
def check(n, ok, d=""):
    results.append(bool(ok)); print(f"[{'PASS' if ok else 'FAIL'}] {n}  {d}")

from PySide6.QtWidgets import QApplication
app = QApplication.instance() or QApplication([])

from docrefine.gui.dialogs import RebrandDialog, PipelineDialog
from docrefine.gui.app_qt import AppController
from docrefine.worker import Worker
from docrefine.config import CFG, REVIEWS_ROOT, SystemUtils
from docrefine import reviews

# The three places a version lives must agree — a release that only bumps one of
# them ships an app that misreports itself. Taking the expected value from the
# CHANGELOG's newest entry rather than a literal keeps this honest on any branch
# and stops the pin itself becoming a thing to remember to edit.
import re as _re
_readme = (REPO / "README.md").read_text(encoding="utf-8")
_chlog = (REPO / "CHANGELOG.md").read_text(encoding="utf-8")
_newest = _re.search(r"^## \[(v\d+)\]", _chlog, _re.M)
check("CHANGELOG's newest entry is a version", _newest is not None)
_want = _newest.group(1) if _newest else ""
check(f"code version is the newest changelog entry ({_want})",
      SystemUtils.CURRENT_VERSION == _want, SystemUtils.CURRENT_VERSION)
check("README matches the code version", f"# DocRefine Pro {SystemUtils.CURRENT_VERSION}" in _readme)
check("CHANGELOG has an entry for it", f"## [{SystemUtils.CURRENT_VERSION}]" in _chlog)
check("Reviews folder exists", REVIEWS_ROOT.is_dir(), str(REVIEWS_ROOT))
check("config knows the option", isinstance(CFG.get("rebrand_complete_set"), bool))

sig = inspect.signature(Worker.run_rebrand_apply).parameters
check("run_rebrand_apply takes complete_set", "complete_set" in sig)
check("run_rebrand_apply plan is optional", sig["plan_csv"].default is None)
check("run_pipeline takes complete_set", "complete_set" in inspect.signature(Worker.run_pipeline).parameters)

# a source with a sheet already present -> dialog should pre-fill it
src = Path(tempfile.mkdtemp(prefix="drp_boot_"))
sheet = reviews.plan_path_for(src); sheet.write_text("file\n", encoding="utf-8")

d = RebrandDialog(None, default_kit="", default_source=str(src))
check("RebrandDialog has the complete-set checkbox", hasattr(d, "chk_complete"))
check("RebrandDialog exposes complete_set", isinstance(d.complete_set, bool))
check("RebrandDialog auto-finds the review sheet", d.plan_path == str(sheet), str(d.plan_path))
d.deleteLater(); sheet.unlink()

p = PipelineDialog(None, default_kit="")
check("PipelineDialog has the complete-set checkbox", hasattr(p, "chk_complete"))
check("PipelineDialog exposes complete_set", isinstance(p.complete_set, bool))
p.deleteLater()

check("AppController importable", AppController is not None)
src.rmdir()

print("\n" + "=" * 50)
print(f"RESULT: {sum(results)}/{len(results)} passed")
sys.exit(0 if all(results) else 1)
