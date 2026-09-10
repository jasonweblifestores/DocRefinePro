"""v157: progress that survives an interrupted classification run.

The valuable assertions here are the awkward ones:

* a stop must SAVE progress but must NOT write a review sheet, and must not
  overwrite a sheet someone already reviewed;
* a completed run must DELETE its progress file — left behind, the next analyze
  would find every file "already classified", skip both passes and hand back a
  stale sheet while looking like it worked;
* a file edited since it was classified must be classified again, not
  resurrected from the saved answer;
* a truncated final line (what a power cut leaves) must cost one record, not the
  file.

Ollama is stubbed, so this runs fast and offline.

Run: python verify_checkpoint.py <repo> "<...>\BrandKit\Sample Files"
"""
import json
import logging
import os
import sys
import tempfile
import time
from pathlib import Path

REPO = Path(sys.argv[1])
sys.path.insert(0, str(REPO))
logging.getLogger("pypdf").setLevel(logging.CRITICAL)

results = []


def check(n, ok, d=""):
    results.append(bool(ok))
    print(f"[{'PASS' if ok else 'FAIL'}] {n}  {d}")


from reportlab.pdfgen import canvas as rlc

from docrefine import checkpoint as cp
from docrefine import classify, reviews
from docrefine.worker import Worker

work = Path(tempfile.mkdtemp(prefix="drp_cp_"))


def pdf(d, name, body="HELLO"):
    p = d / name
    c = rlc.Canvas(str(p), pagesize=(612, 792))
    c.setFont("Helvetica", 14)
    c.drawString(40, 700, body)
    c.showPage()
    c.save()
    return p


# =====================================================================
#  A. The progress file itself
# =====================================================================
src = work / "src"
src.mkdir()
files = [pdf(src, f"doc{i}.pdf", f"CONTENT {i}") for i in range(5)]

plan = work / "sheet.xlsx"
cpath = cp.path_for(plan)
check("A1 the progress file sits beside its sheet",
      cpath.parent == plan.parent and cpath.name.startswith("sheet"), cpath.name)
check("A2 loading a missing file yields nothing, rather than raising", cp.load(cpath) == {})

w = cp.Writer(cpath)
w.add("doc0.pdf", files[0], {"file": "doc0.pdf", "action": "rebrand"}, needs_look=True)
w.add("doc1.pdf", files[1], {"file": "doc1.pdf", "action": "leave"})
w.close()
recs = cp.load(cpath)
check("A3 records come back keyed by file", set(recs) == {"doc0.pdf", "doc1.pdf"})
check("A4 the row survives the round trip", recs["doc1.pdf"]["row"]["action"] == "leave")
check("A5 needs_look is remembered", recs["doc0.pdf"]["needs_look"] is True)
check("A6 vision_done defaults false", recs["doc0.pdf"]["vision_done"] is False)

# later record wins — that is how the visual pass supersedes the text pass
w = cp.Writer(cpath)
w.add("doc0.pdf", files[0], {"file": "doc0.pdf", "action": "leave"},
      needs_look=True, vision_done=True)
w.close()
recs = cp.load(cpath)
check("A7 a later record supersedes an earlier one",
      recs["doc0.pdf"]["row"]["action"] == "leave" and recs["doc0.pdf"]["vision_done"] is True)

# a truncated final line is what a power cut leaves behind
with open(cpath, "a", encoding="utf-8") as fh:
    fh.write('{"file": "doc2.pdf", "size": 1, "mtime"')
recs = cp.load(cpath)
check("A8 a truncated last line costs one record, not the file",
      set(recs) == {"doc0.pdf", "doc1.pdf"}, f"{sorted(recs)}")

check("A9 a record is usable while the file is unchanged", cp.usable(recs["doc1.pdf"], files[1]))
time.sleep(0.01)
os.utime(files[1], (time.time() + 10, time.time() + 10))
check("A10 a file touched since classification is NOT reused",
      not cp.usable(recs["doc1.pdf"], files[1]))
files[1].write_bytes(files[1].read_bytes() + b"\n% grown")
check("A11 a file whose size changed is NOT reused", not cp.usable(recs["doc1.pdf"], files[1]))
check("A12 clear() removes it", (cp.clear(cpath), not cpath.exists())[1])
check("A13 clearing an absent file is harmless", cp.clear(cpath) is None)

# =====================================================================
#  B. The worker: stop saves progress and writes no sheet
# =====================================================================
_real = classify.classify_document
calls = {"n": 0}


def fake_classify(p, *a, **k):
    calls["n"] += 1
    return {"action": "leave", "doc_type": "drawing", "product": "", "asset_type": "",
            "manufacturer": "", "title": "", "confidence": 1.0, "source": "stub", "notes": ""}


classify.classify_document = fake_classify

src2 = work / "src2"
src2.mkdir()
for i in range(6):
    pdf(src2, f"f{i}.pdf", f"BODY {i}")

plan2 = work / "sheet2.xlsx"
cpath2 = cp.path_for(plan2)


def run(worker, stop_after=None):
    """Run analyze, optionally stopping after N files via the progress callback."""
    seen = {"n": 0}

    def cb(ev):
        if getattr(ev.type, "name", "") == "PROGRESS_MAIN":
            t = str((ev.payload or {}).get("text") or "")
            if t.startswith("Reading"):
                seen["n"] += 1
                if stop_after and seen["n"] >= stop_after:
                    worker.stop_sig = True

    worker.callback = cb
    worker.run_rebrand_analyze(str(src2), str(plan2), False)


w1 = Worker(callback=lambda e: None)
w1.log = lambda m, err=False: None
calls["n"] = 0
run(w1, stop_after=3)
check("B1 a stopped run writes NO review sheet", not plan2.exists())
check("B2 but it does save progress", cpath2.is_file())
saved = cp.load(cpath2)
check("B3 progress covers the files it got through", 0 < len(saved) <= 6, f"{len(saved)} records")
first_calls = calls["n"]

# resume: the classifier must not be asked about files already done
w2 = Worker(callback=lambda e: None)
w2.log = lambda m, err=False: None
calls["n"] = 0
run(w2)
check("B4 the resumed run completes and writes the sheet", plan2.exists())
check("B5 it re-classified only what was left",
      calls["n"] == 6 - len(saved), f"{calls['n']} calls, {len(saved)} were already saved")
rows = reviews.read_plan(plan2)
check("B6 the sheet holds every file", len(rows) == 6, f"{len(rows)} rows")

# =====================================================================
#  C. THE TRAP: a completed run must not leave progress behind
# =====================================================================
check("C1 a completed run deletes its progress file", not cpath2.exists(),
      "left behind, the next analyze would skip everything and return a stale sheet")

calls["n"] = 0
w3 = Worker(callback=lambda e: None)
w3.log = lambda m, err=False: None
run(w3)
check("C2 so a fresh analyze really re-classifies", calls["n"] == 6, f"{calls['n']} calls")

# =====================================================================
#  D. Wiring
# =====================================================================
import inspect  # noqa: E402
srcs = inspect.getsource(Worker.run_rebrand_analyze)
check("D1 progress is written during pass 1", "cp_writer.add" in srcs)
check("D2 progress is cleared once the sheet exists", "cp.clear" in srcs)
check("D3 the stop path keeps progress rather than clearing it",
      srcs.index("cp.clear") > srcs.index("Analysis stopped by user"))
check("D4 the stop message no longer promises a sheet",
      "no review sheet" in srcs.lower())

classify.classify_document = _real
print("\n" + "=" * 50)
print(f"RESULT: {sum(results)}/{len(results)} passed")
sys.exit(0 if all(results) else 1)
