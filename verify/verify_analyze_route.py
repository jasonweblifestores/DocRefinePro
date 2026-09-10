"""Does clicking Analyze in the GUI actually ask for a visual pass?

The worker was tested directly and the dialog was tested for construction, but
nothing covered the wiring between them: RebrandDialog's checkbox -> CFG ->
AppController._start_analyze -> the arguments the worker is finally called with.

That gap matters because it is silent in the worst direction. If the gating in
_start_analyze were wrong, a user who ticks the box gets a sheet built from
filenames alone and nothing says so — the run looks identical, just faster.

Everything here is stubbed: no Ollama, no worker thread, no window.

Run: python verify_analyze_route.py <repo>
"""
import sys
import types
from pathlib import Path

REPO = Path(sys.argv[1])
sys.path.insert(0, str(REPO))

results = []


def check(n, ok, d=""):
    results.append(bool(ok))
    print(f"[{'PASS' if ok else 'FAIL'}] {n}  {d}")


from PySide6.QtWidgets import QApplication, QMessageBox
app = QApplication.instance() or QApplication([])

from docrefine import classify
from docrefine.gui import app_qt as AQ

SRC = "C:/some/folder"


class Ctl(AQ.AppController):
    """Just the analyze-routing decision, with no window and no worker thread."""

    def __init__(self):
        self.window = None
        self.worker = types.SimpleNamespace(run_rebrand_analyze="ANALYZE",
                                            run_pull_model="PULL")
        self.calls = []
        self._sleep_when_done = False
        self._run_had_error = False

    def start_process(self, target, args, multi_threaded=False, **kw):
        self.calls.append((target, args))


_real = {k: getattr(classify, k) for k in
         ("ollama_status", "start_server", "has_model", "has_vision_model")}


def stub_classify(state="ready", has_vision=True):
    classify.ollama_status = lambda: {"state": state}
    classify.start_server = lambda: state == "ready"
    classify.has_model = lambda *a, **k: True
    classify.has_vision_model = lambda *a, **k: has_vision


class FakeBox:
    """Stands in for the QMessageBox, answering with a chosen button role."""
    answer = "skip"          # "get" | "skip" | "cancel"

    def __init__(self, *a, **k):
        self._buttons = {}

    def setWindowTitle(self, *a): pass
    def setIcon(self, *a): pass
    def setText(self, *a): pass
    def exec(self): pass

    def addButton(self, *a):
        # (text, role) for custom buttons; a bare StandardButton for Cancel
        if len(a) == 2:
            role = "get" if a[1] == QMessageBox.AcceptRole else "skip"
        else:
            role = "cancel"
        b = object()
        self._buttons[role] = b
        return b

    def clickedButton(self):
        return self._buttons.get(FakeBox.answer, object())


_realbox = AQ.QMessageBox
AQ.QMessageBox = FakeBox
for attr in ("Question", "AcceptRole", "DestructiveRole", "Cancel", "Information"):
    setattr(FakeBox, attr, getattr(_realbox, attr, None))

# ---------------------------------------------------------------------------
# The case that actually produced the Batch 4 sheet: model present, box ticked.
# ---------------------------------------------------------------------------
stub_classify("ready", has_vision=True)
c = Ctl()
c._start_analyze(SRC, True)
check("A1 ticking the box reaches the worker with vision on",
      c.calls and c.calls[0][1][:3] == (SRC, None, True), str(c.calls))
check("A2 and it is the analyze worker that was started",
      c.calls and c.calls[0][0] == "ANALYZE")

# unticked must stay off — not merely default to config
c = Ctl()
c._start_analyze(SRC, False)
check("A3 leaving it unticked runs without the visual pass",
      c.calls and c.calls[0][1][:3] == (SRC, None, False), str(c.calls))

# ---------------------------------------------------------------------------
# Model missing: the user is asked, and the answer is honoured.
# ---------------------------------------------------------------------------
stub_classify("ready", has_vision=False)

FakeBox.answer = "skip"
c = Ctl()
c._start_analyze(SRC, True)
check("B1 'analyze without it' proceeds with the pass OFF",
      c.calls and c.calls[0][1][:3] == (SRC, None, False), str(c.calls))

FakeBox.answer = "get"
c = Ctl()
c._start_analyze(SRC, True)
check("B2 'download model' pulls instead of analyzing",
      c.calls and c.calls[0][0] == "PULL", str(c.calls))
check("B2b and it does NOT start an analyze as well", len(c.calls) == 1, str(c.calls))

FakeBox.answer = "cancel"
c = Ctl()
c._start_analyze(SRC, True)
check("B3 cancelling starts nothing at all", c.calls == [], str(c.calls))

# ---------------------------------------------------------------------------
# Ollama unusable: the filename fallback must not claim a visual pass.
# ---------------------------------------------------------------------------
stub_classify("not_installed", has_vision=False)
FakeBox.answer = "skip"          # "Use filenames"
c = Ctl()
c._start_analyze(SRC, True)
check("C1 the filename fallback never asks for a visual pass",
      c.calls and c.calls[0][1][:3] == (SRC, None, False), str(c.calls))

for k, v in _real.items():
    setattr(classify, k, v)
AQ.QMessageBox = _realbox

print("\n" + "=" * 50)
print(f"RESULT: {sum(results)}/{len(results)} passed")
sys.exit(0 if all(results) else 1)
