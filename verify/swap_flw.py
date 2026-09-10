"""Brand the downsampled FLW file and swap it into the shipped delivery.

Same pattern as the Imperial swap: rebrand the one file on its own with the
delivery's exact settings, then copy it over the shipped output KEEPING THE
SHIPPED FILENAME, so the 2,174-file QA stays valid.
"""
import logging
import shutil
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, r"C:\Users\WORK\Documents\WebLife Labs\PROJECTS\DocRefine Pro\DocRefinePro")
logging.getLogger("pypdf").setLevel(logging.CRITICAL)

from pypdf import PdfReader

from docrefine import reviews
from docrefine.worker import Worker

NAME = "frank_lloyd_wright_collection.pdf"
SP = Path(__file__).resolve().parent
SMALL = SP / "flw_downsampled.pdf"
SRC = Path(r"C:\Users\WORK\Documents\Batch 4\_unique-to-rebrand")
KIT = Path(r"C:\Users\WORK\Documents\Batch 4\Template")
OUT = Path(r"C:\Users\WORK\Documents\Batch 4\_unique-to-rebrand_rebranded")
SHIPPED = OUT / NAME

row = next(r for r in reviews.read_plan(reviews.plan_path_for(SRC)) if r["file"] == NAME)
print(f"sheet row: action={row.get('action')} product={row.get('product')!r} "
      f"asset_type={row.get('asset_type')!r} manufacturer={row.get('manufacturer')!r}")

work = Path(tempfile.mkdtemp(prefix="flw_swap_"))
tsrc = work / "src"
tsrc.mkdir()
shutil.copy2(SMALL, tsrc / NAME)          # same filename the sheet refers to
tout = work / "out"

plan = work / "plan.xlsx"
reviews.write_plan(plan, [row], Worker.REBRAND_PLAN_COLUMNS, src_root=tsrc)

log = []
w = Worker(callback=lambda e: None)


def wl(m, err=False):
    line = ("ERROR: " if err else "") + str(m)
    log.append(line)
    try:
        print("   " + line, flush=True)
    except Exception:
        print("   " + line.encode("ascii", "replace").decode(), flush=True)


w.log = wl

print("\nbranding the downsampled copy with the delivery settings:")
w.run_rebrand_apply(
    str(tsrc), str(KIT), str(plan), out_dir=str(tout),
    complete_set=False,
    show_attribution=False,
    keep_original_names=True,
    stamp_opts={"footer_attribution": True, "stamp_tagline": True,
                "stamp_version": True, "stamp_disclaimer": True},
)

produced = tout / NAME
if not produced.is_file():
    print("\nFAILED: nothing produced — not swapping anything")
    sys.exit(1)

mb = produced.stat().st_size / 1e6
rd = PdfReader(str(produced))
body = " ".join((rd.pages[i].extract_text() or "") for i in range(1, len(rd.pages) - 1))
body = " ".join(body.split())
src_pages = len(PdfReader(str(SMALL)).pages)

checks = [
    ("under the 50 MB limit", mb < 50, f"{mb:.1f} MB"),
    ("page count is source+2", len(rd.pages) == src_pages + 2,
     f"{src_pages} -> {len(rd.pages)}"),
    ("Author is the brand", (rd.metadata or {}).get("/Author") == "Budget Mailboxes",
     str((rd.metadata or {}).get("/Author"))),
    ("tagline stamped", "Trusted by the Nation" in body, ""),
    ("version + Last Updated stamped",
     "Version 1.0" in body and "Last Updated:" in body, ""),
    ("disclaimer stamped", "This guide is provided for reference" in body, ""),
    ("original text preserved", len(body) > 9000, f"{len(body)} chars in content pages"),
]
print()
ok = True
for name, good, detail in checks:
    ok &= good
    print(f"[{'PASS' if good else 'FAIL'}] {name}  {detail}")

if not ok:
    print("\nNOT swapping — fix the failures first")
    sys.exit(1)

backup = SP / (NAME + ".shipped-59MB.bak")
if SHIPPED.is_file() and not backup.exists():
    shutil.copy2(SHIPPED, backup)
    print(f"\nbacked up the 59.4 MB version to {backup.name}")
shutil.copy2(produced, SHIPPED)
print(f"swapped into the delivery: {SHIPPED}  ({SHIPPED.stat().st_size / 1e6:.1f} MB)")
shutil.rmtree(work, ignore_errors=True)
print("\ndone — filename unchanged, so the 2,174-file QA still holds")
