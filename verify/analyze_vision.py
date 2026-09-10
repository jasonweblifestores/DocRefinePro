"""Re-Analyze Batch 4 from scratch with the visual pass ON.

Drives exactly the code path the GUI's Analyze button drives
(Worker.run_rebrand_analyze with vision_pass=True), so the sheet it writes is
the same artifact the app would produce. Logs progress with timestamps to
stdout, flushed, so a 12-hour run can be checked in on.
"""
import logging
import sys
import time
from pathlib import Path

REPO = Path(r"C:\Users\WORK\Documents\WebLife Labs\PROJECTS\DocRefine Pro\DocRefinePro")
sys.path.insert(0, str(REPO))
logging.getLogger("pypdf").setLevel(logging.CRITICAL)

from docrefine.worker import Worker
from docrefine import reviews

SRC = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(
    r"C:\Users\WORK\Documents\Batch 4\_unique-to-rebrand")

T0 = time.time()


def stamp():
    e = time.time() - T0
    return f"[{int(e // 3600):02d}:{int(e % 3600 // 60):02d}:{int(e % 60):02d}]"


last_main = [""]


def on_event(ev):
    """Echo coarse progress; the worker emits one per file, so throttle it.

    Payload for PROGRESS_MAIN is {"percent", "text"}; text reads
    "Reading 12/2174" in pass 1 and "Looking at 5/1512" in pass 2.
    """
    if getattr(ev.type, "name", "") != "PROGRESS_MAIN":
        return
    text = str((ev.payload or {}).get("text") or "")
    if not text:
        return
    # print on every phase change, then every 25th file
    head = text.rsplit(" ", 1)[0]
    tail = text.rsplit(" ", 1)[-1]
    cur = tail.split("/")[0] if "/" in tail else ""
    if head != last_main[0] or (cur.isdigit() and int(cur) % 25 == 0):
        last_main[0] = head
        print(f"{stamp()} {text}", flush=True)


w = Worker(callback=on_event)
w.log = lambda m, err=False: print(f"{stamp()} {'ERROR: ' if err else ''}{m}", flush=True)

print(f"{stamp()} source: {SRC}")
print(f"{stamp()} sheet:  {reviews.plan_path_for(SRC)}")
print(f"{stamp()} vision pass: ON", flush=True)

w.run_rebrand_analyze(str(SRC), None, True)

print(f"{stamp()} analyze finished in {(time.time() - T0) / 3600:.2f} hours", flush=True)

# --- summarise the sheet we just wrote -------------------------------------
plan = reviews.plan_path_for(SRC)
rows = reviews.read_plan(plan)
n_reb = sum(1 for r in rows if (r.get("action") or "").strip().lower() == "rebrand")
n_leave = len(rows) - n_reb
by_source = {}
for r in rows:
    by_source[r.get("source") or "?"] = by_source.get(r.get("source") or "?", 0) + 1
print(f"\nSHEET: {plan}")
print(f"  rows          {len(rows)}")
print(f"  rebrand       {n_reb}")
print(f"  leave         {n_leave}")
print("  decided by:")
for k, v in sorted(by_source.items(), key=lambda x: -x[1]):
    print(f"    {k or '(blank)':<24} {v}")
