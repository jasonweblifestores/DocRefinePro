"""Score the live vision model against files whose type was confirmed by eye.

The point is to decide whether to trust it on 1,512 files, so it is scored on
ground truth rather than on its own confidence. Labels below come from rendering
each page and looking at it during the v151-v154 investigation.

Run: python verify_vision_live.py <repo>
"""
import sys, time, logging
from pathlib import Path

REPO = Path(sys.argv[1])
sys.path.insert(0, str(REPO))
logging.getLogger("pypdf").setLevel(logging.CRITICAL)
from docrefine import classify

SRC = Path(r"C:\Users\WORK\Documents\Batch 4\_unique-to-rebrand")

# label, file, expected action
TRUTH = [
    # dimensioned CAD drawings — confirmed by eye, title blocks and dimension arrows
    ("drawing", "4c06d-02-sm_cutsheet_pdp_.pdf", "leave"),
    ("drawing", "4c12d-11-sm_cutsheet_pdp_.pdf", "leave"),
    ("drawing", "4c09d-06-sm_cutsheet_pdp_.pdf", "leave"),
    ("drawing", "4c15d-09-sm_cutsheet_pdp_.pdf", "leave"),
    ("drawing", "tech-54368.pdf", "leave"),
    ("drawing", "tech-18-52168.pdf", "leave"),
    ("drawing", "tech-64168.pdf", "leave"),
    # drawings the FILENAME rule missed — already carry BM branding
    ("drawing", "1570-13-bm_1_1.pdf", "leave"),
    ("drawing", "1570-8-bm_1.pdf", "leave"),
    ("drawing", "1590-T1V-Spec-Sheet.pdf", "leave"),
    # genuine documents — confirmed by eye
    ("guide", "1570-cbu-installation-manual.pdf", "rebrand"),
    ("guide", "inst-lockerbenches-wood-77000-series.pdf", "rebrand"),
    ("guide", "inst-standardopenaccesslockers-70000-series.pdf", "rebrand"),
    ("brochure", "Imperial-Street-Sign-Brochure.pdf", "rebrand"),
    ("spec sheet", "4c06d-02-sm-installation-guide-budget-mailboxes.pdf", None),  # skipped if absent
]

model = classify.DEFAULT_VISION_MODEL
if not classify.has_vision_model(model):
    print(f"[skip] {model} is not downloaded yet — nothing to score")
    sys.exit(0)
print(f"scoring {model} against {len([t for t in TRUTH if t[2]])} hand-labelled files\n")

ok = bad = skipped = 0
wrong = []
t0 = time.time()
for label, name, expect in TRUTH:
    if expect is None:
        continue
    f = SRC / name
    if not f.is_file():
        print(f"  [skip] {name} not present"); skipped += 1; continue
    t1 = time.time()
    got = classify.classify_visually(f, model=model)
    dt = time.time() - t1
    if not got:
        print(f"  [FAIL] {name}: model gave no answer ({dt:.1f}s)"); bad += 1
        wrong.append((name, label, expect, "no answer")); continue
    hit = got["action"] == expect
    ok += hit; bad += (not hit)
    if not hit:
        wrong.append((name, label, expect, f"{got['action']} / {got['doc_type']}"))
    print(f"  [{'ok  ' if hit else 'WRONG'}] {name[:46]:46} "
          f"expected={expect:8} got={got['action']:8} "
          f"({got['doc_type'][:22]:22}) conf={got['confidence']:.2f} {dt:5.1f}s")
    if got.get("notes"):
        print(f"           saw: {got['notes'][:90]}")

n = ok + bad
print("\n" + "=" * 60)
print(f"correct {ok}/{n}   ({100*ok/n if n else 0:.0f}%)   skipped {skipped}   "
      f"total {time.time()-t0:.0f}s   ~{(time.time()-t0)/max(1,n):.1f}s per file")
if wrong:
    print("\ngot these wrong:")
    for name, label, expect, got in wrong:
        print(f"   {name}  (it is a {label}; expected {expect}, said {got})")
est = (time.time() - t0) / max(1, n) * 1512
print(f"\nat this rate 1,512 files would take about {est/3600:.1f} hours")
