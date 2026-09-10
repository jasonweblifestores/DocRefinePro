"""The Batch 4 delivery run, with the settings Kunchana signed off.

    footer attribution ON · tagline ON · version ON · disclaimer ON
    cover attribution  OFF   (footer placement, not the cover)
    keep original filenames ON   (live URLs unchanged, no 301s)
    complete set OFF   (source holds desktop.ini and _unique-index.csv;
                        both are in the currently shipped output and must not
                        ship again)

Apply skips outputs that already exist, so the previous delivery tree has to be
out of the way. It is MOVED ASIDE, not deleted — it is the current deliverable,
a rename costs nothing, and if anything about this run looks wrong the old set
is still there to compare against.
"""
import logging
import shutil
import sys
import time
from pathlib import Path

REPO = Path(r"C:\Users\WORK\Documents\WebLife Labs\PROJECTS\DocRefine Pro\DocRefinePro")
sys.path.insert(0, str(REPO))
logging.getLogger("pypdf").setLevel(logging.CRITICAL)

from docrefine import reviews
from docrefine.worker import Worker

SRC = Path(r"C:\Users\WORK\Documents\Batch 4\_unique-to-rebrand")
KIT = Path(r"C:\Users\WORK\Documents\Batch 4\Template")
OUT = Path(r"C:\Users\WORK\Documents\Batch 4\_unique-to-rebrand_rebranded")

T0 = time.time()


def stamp():
    e = time.time() - T0
    return f"[{int(e // 3600):02d}:{int(e % 3600 // 60):02d}:{int(e % 60):02d}]"


def main():
    plan = reviews.plan_path_for(SRC)
    if not Path(plan).is_file():
        print(f"no review sheet at {plan} — run the analyze first")
        return 2

    rows = reviews.read_plan(plan)
    n_reb = sum(1 for r in rows if (r.get("action") or "").strip().lower() == "rebrand")
    print(f"{stamp()} sheet:   {plan}")
    print(f"{stamp()} rows:    {len(rows)}  ({n_reb} to rebrand, {len(rows) - n_reb} to leave)")

    # Move the previous delivery aside rather than destroying it.
    # Skipped on a repair pass: Apply is resume-safe (it skips outputs that
    # already exist), so re-running fills only the gaps.
    if OUT.exists() and "--resume" not in sys.argv:
        keep = OUT.with_name(OUT.name + "_previous_" + time.strftime("%Y%m%d_%H%M%S"))
        print(f"{stamp()} moving the previous output aside:")
        print(f"{stamp()}   {OUT}")
        print(f"{stamp()}   -> {keep}")
        OUT.rename(keep)

    log = []

    def w_log(m, err=False):
        line = f"{stamp()} {'ERROR: ' if err else ''}{m}"
        log.append(line)
        # Never let the act of logging raise. A cp1252 stdout choking on one
        # character used to propagate out of the worker and be counted as a
        # failed FILE, which is how a 59.4 MB breach of the size cap went
        # unreported.
        try:
            print(line, flush=True)
        except Exception:
            print(line.encode("ascii", "replace").decode("ascii"), flush=True)

    last = [""]

    def on_event(ev):
        if getattr(ev.type, "name", "") != "PROGRESS_MAIN":
            return
        text = str((ev.payload or {}).get("text") or "")
        if not text:
            return
        head = text.rsplit(" ", 1)[0]
        tail = text.rsplit(" ", 1)[-1]
        cur = tail.split("/")[0] if "/" in tail else ""
        if head != last[0] or (cur.isdigit() and int(cur) % 50 == 0):
            last[0] = head
            print(f"{stamp()} {text}", flush=True)

    w = Worker(callback=on_event)
    w.log = w_log

    w.run_rebrand_apply(
        str(SRC), str(KIT), str(plan), out_dir=str(OUT),
        complete_set=False,           # keeps desktop.ini / _unique-index.csv out
        show_attribution=False,       # footer placement, not the cover
        keep_original_names=True,     # live URLs unchanged
        stamp_opts={"footer_attribution": True, "stamp_tagline": True,
                    "stamp_version": True, "stamp_disclaimer": True},
    )

    print(f"\n{stamp()} apply finished in {(time.time() - T0) / 60:.1f} min")

    files = sorted(OUT.rglob("*"))
    pdfs = [f for f in files if f.is_file() and f.suffix.lower() == ".pdf"]
    others = [f for f in files if f.is_file() and f.suffix.lower() != ".pdf"]
    print(f"  output PDFs      {len(pdfs)}")
    print(f"  output non-PDFs  {len(others)}   {[f.name for f in others[:5]]}")
    print(f"  total size       {sum(f.stat().st_size for f in pdfs) / 1e9:.2f} GB")
    return 0


if __name__ == "__main__":
    sys.exit(main())
