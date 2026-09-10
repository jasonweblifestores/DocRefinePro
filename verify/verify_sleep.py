"""v156: sleep the machine when a long run finishes.

The interesting part is not the sleeping, it is the *refusing* — a run that was
stopped by hand or that failed must leave the machine awake. Those rules are
asserted here against the real controller, with the actual suspend call stubbed
so running this test cannot put the machine to sleep.

Run: python verify_sleep.py <repo> "<...>\BrandKit\Sample Files"
"""
import inspect
import sys
import tempfile
from pathlib import Path

REPO = Path(sys.argv[1])
sys.path.insert(0, str(REPO))

results = []


def check(n, ok, d=""):
    results.append(bool(ok))
    print(f"[{'PASS' if ok else 'FAIL'}] {n}  {d}")


from PySide6.QtWidgets import QApplication
app = QApplication.instance() or QApplication([])

from docrefine import power
from docrefine.config import CFG, ConfigData
from docrefine.gui import app_qt as AQ
from docrefine.gui.dialogs import RebrandDialog, PipelineDialog, SleepCountdown, SleepWhenDone

# =====================================================================
#  A. Off by default, remembered, and honest about the platform
# =====================================================================
cfg = ConfigData()
check("A1 sleep-when-done is off by default", cfg.sleep_when_done is False)
check("A2 the action defaults to plain sleep", cfg.sleep_when_done_action == "sleep")
check("A3 both actions are named", set(power.ACTIONS) == {"sleep", "hibernate"})
check("A4 there is a countdown, and it is not instant", power.COUNTDOWN_SECONDS >= 10)
check("A5 describe() speaks the platform's language",
      power.describe("sleep") == "sleep" and power.describe("hibernate") in ("hibernate", "sleep"))
check("A6 available() answers without raising", isinstance(power.available(), bool))

# =====================================================================
#  B. The checkbox is in both dialogs, and exposes a per-run answer
# =====================================================================
src = Path(tempfile.mkdtemp(prefix="drp_sleep_"))
d = RebrandDialog(None, default_kit="", default_source=str(src))
check("B1 RebrandDialog has the checkbox", isinstance(d.chk_sleep, SleepWhenDone))
check("B2 RebrandDialog exposes sleep_when_done", isinstance(d.sleep_when_done, bool))
check("B3 it starts false before a button is pressed", d.sleep_when_done is False)
d.chk_sleep.setChecked(True)
d.source_path = str(src)
d.on_analyze()
check("B4 Analyze carries the answer through", d.sleep_when_done is True)
check("B5 Analyze still sets its own mode", d.mode == "analyze")
d.deleteLater()

p = PipelineDialog(None, default_kit="")
check("B6 PipelineDialog has the checkbox", isinstance(p.chk_sleep, SleepWhenDone))
check("B7 PipelineDialog exposes sleep_when_done", isinstance(p.sleep_when_done, bool))
p.deleteLater()

# a checkbox on a machine that cannot sleep is disabled, not a false promise
if not power.available():
    c = SleepWhenDone()
    check("B8 disabled where sleeping is impossible", not c.isEnabled() and not c.isChecked())
else:
    check("B8 platform supports sleeping, so the box is live", SleepWhenDone().isEnabled())

# =====================================================================
#  C. THE SAFETY RULES — when the controller must NOT sleep
# =====================================================================
suspended = []
countdowns = []


class FakeWorker:
    def __init__(self):
        self.stop_sig = False
        self.logs = []

    def log(self, m, err=False):
        self.logs.append(str(m))


class Ctl(AQ.AppController):
    """Only the completion logic, with no Qt window or worker thread."""
    def __init__(self):
        self.worker = FakeWorker()
        self.window = None
        self._sleep_when_done = False
        self._run_had_error = False


def stub(countdown_accepts=True):
    suspended.clear(); countdowns.clear()
    power.suspend = lambda action="sleep": (suspended.append(action), (True, "ok"))[1]

    class FakeCountdown:
        def __init__(self, *a, **k):
            countdowns.append(k.get("action", "sleep"))

        def exec(self):
            return 1 if countdown_accepts else 0

    AQ.SleepCountdown = FakeCountdown


_real_suspend, _real_countdown = power.suspend, AQ.SleepCountdown

# C1: the ordinary happy path — ticked, finished cleanly
stub()
c = Ctl(); c._sleep_when_done = True
c._maybe_sleep()
check("C1 a clean finish sleeps", suspended == ["sleep"], f"{suspended}")

# C2: not ticked — nothing happens at all, not even a countdown
stub()
c = Ctl(); c._sleep_when_done = False
c._maybe_sleep()
check("C2 unticked never sleeps and never prompts", suspended == [] and countdowns == [])

# C3: STOP was pressed
stub()
c = Ctl(); c._sleep_when_done = True; c.worker.stop_sig = True
c._maybe_sleep()
check("C3 a stopped run stays awake", suspended == [] and countdowns == [])
check("C3b and it says why", any("stopped" in m.lower() for m in c.worker.logs), c.worker.logs)

# C4: the run reported an error
stub()
c = Ctl(); c._sleep_when_done = True; c._run_had_error = True
c._maybe_sleep()
check("C4 a failed run stays awake", suspended == [] and countdowns == [])
check("C4b and it says why", any("problem" in m.lower() for m in c.worker.logs), c.worker.logs)

# C5: the user cancels the countdown
stub(countdown_accepts=False)
c = Ctl(); c._sleep_when_done = True
c._maybe_sleep()
check("C5 cancelling the countdown stays awake", suspended == [] and countdowns == ["sleep"])
check("C5b and it says so", any("cancel" in m.lower() for m in c.worker.logs), c.worker.logs)

# C6: one run, one offer — a second DONE cannot sleep again
stub()
c = Ctl(); c._sleep_when_done = True
c._maybe_sleep(); c._maybe_sleep()
check("C6 a repeated DONE only ever sleeps once", suspended == ["sleep"], f"{suspended}")

# C7: the hibernate setting is passed through rather than ignored
stub()
CFG.set("sleep_when_done_action", "hibernate")
c = Ctl(); c._sleep_when_done = True
c._maybe_sleep()
check("C7 the hibernate choice reaches power.suspend", suspended == ["hibernate"], f"{suspended}")
CFG.set("sleep_when_done_action", "sleep")

power.suspend, AQ.SleepCountdown = _real_suspend, _real_countdown

# =====================================================================
#  D. Every run resets the state — no decision leaks into the next run
# =====================================================================
sig = inspect.signature(AQ.AppController.start_process).parameters
check("D1 start_process takes sleep_when_done", "sleep_when_done" in sig)
check("D2 and it defaults to off", sig["sleep_when_done"].default is False)
srcs = inspect.getsource(AQ.AppController.start_process)
check("D3 start_process clears the error flag", "_run_had_error = False" in srcs)
check("D4 on_done consults the sleep logic", "_maybe_sleep" in inspect.getsource(AQ.AppController.on_done))
check("D5 errors are recorded", "_run_had_error = True" in inspect.getsource(AQ.AppController.on_run_error))
check("D6 a model download is not treated as the sleepable run",
      "sleep_when_done=sleep_when_done" in inspect.getsource(AQ.AppController._start_analyze))
check("D7 _start_analyze accepts the flag",
      "sleep_when_done" in inspect.signature(AQ.AppController._start_analyze).parameters)

# =====================================================================
#  E. The countdown dialog: every accidental exit means stay awake
# =====================================================================
cd = SleepCountdown(None, action="sleep", seconds=3)
check("E1 it counts from the number given", cd.remaining == 3)
check("E2 Stay awake is the default button", cd.btn_stay.isDefault())
check("E3 the label names the action and the time",
      "sleep" in cd.lbl.text() and "3" in cd.lbl.text(), cd.lbl.text().replace("\n", " "))
cd._tick()
check("E4 a tick counts down", cd.remaining == 2)
cd.reject()
check("E5 rejecting stops the timer", not cd._timer.isActive())
cd.deleteLater()

# reaching zero accepts (and therefore sleeps) exactly once
cd = SleepCountdown(None, action="sleep", seconds=1)
done = []
cd.accept = lambda: done.append(1)
cd._tick()
check("E6 reaching zero accepts", done == [1])
check("E7 and stops ticking", not cd._timer.isActive())
cd.deleteLater()

src.rmdir()
print("\n" + "=" * 50)
print(f"RESULT: {sum(results)}/{len(results)} passed")
sys.exit(0 if all(results) else 1)
