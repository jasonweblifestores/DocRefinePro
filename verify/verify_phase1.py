"""Phase 1 verification: rebrand a nested folder with no review sheet -> mirrored branded tree.

(v137 removed the standalone `run_rebrand`; the surviving path with the same
contract — filename-derived names, no plan — is the pipeline's Rebrand step.)
"""
import sys, shutil, tempfile
from pathlib import Path

REPO, BK = Path(sys.argv[1]), Path(sys.argv[2])
sys.path.insert(0, str(REPO))
from pypdf import PdfReader
from docrefine.worker import Worker
from docrefine.core.events import EventType

results = []
def check(name, ok, detail=""):
    results.append(bool(ok)); print(f"[{'PASS' if ok else 'FAIL'}] {name}  {detail}")

# Build a nested source tree
work = Path(tempfile.mkdtemp(prefix="drp_p1_"))
src = work / "batch"; (src / "brandA").mkdir(parents=True)
shutil.copy2(BK / "1570_FCBU_Installation_Instructions.pdf", src / "brandA" / "1570_FCBU_Installation_Instructions.pdf")
shutil.copy2(BK / "1570-4T5-BM.pdf", src / "1570-4T5-BM.pdf")
out = work / "out"

events = []
w = Worker(callback=lambda e: events.append(e.type))
w.run_pipeline(str(src), do_flatten=False, do_rebrand=True, do_ocr=False,
               kit_dir=str(BK), out_dir=str(out))

# progress + completion signalled?
check("emitted progress", EventType.PROGRESS_MAIN in events)
check("emitted DONE", EventType.DONE in events)
check("emitted completion notification", EventType.NOTIFICATION in events)

# mirrored tree + naming
nested = out / "brandA" / "1570-fcbu-installation-instructions-budget-mailboxes.pdf"
root = out / "1570-4t5-bm-budget-mailboxes.pdf"
check("nested output mirrored + named", nested.exists(), str(nested.relative_to(out)) if nested.exists() else "missing")
check("root output mirrored + named", root.exists(), str(root.relative_to(out)) if root.exists() else "missing")

for f in [nested, root]:
    if not f.exists():
        continue
    r = PdfReader(str(f))
    txt = (r.pages[1].extract_text() or "").strip()
    size = f.stat().st_size / 1e6
    author = (r.metadata or {}).get("/Author")
    check(f"{f.name[:28]}…: <=60 char name", len(f.name) <= 60, f"{len(f.name)}")
    check(f"{f.name[:28]}…: searchable", len(txt) >= 90, f"{len(txt)} chars")
    check(f"{f.name[:28]}…: <50MB", size < 50, f"{size:.1f} MB")
    check(f"{f.name[:28]}…: author set", author == "Budget Mailboxes")

# re-run produces the same set (deterministic names, no -2/-3 duplicates)
before = sorted(p.relative_to(out).as_posix() for p in out.rglob("*.pdf"))
Worker(callback=lambda e: None).run_pipeline(str(src), do_flatten=False, do_rebrand=True, do_ocr=False,
                                             kit_dir=str(BK), out_dir=str(out))
after = sorted(p.relative_to(out).as_posix() for p in out.rglob("*.pdf"))
check("re-run is deterministic (same file set)", before == after, str(after))

shutil.rmtree(work, ignore_errors=True)
print("\n" + "=" * 50)
print(f"RESULT: {sum(results)}/{len(results)} passed")
sys.exit(0 if all(results) else 1)
