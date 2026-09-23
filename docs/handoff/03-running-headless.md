# Running deliveries without the GUI — the recommended way

**Jason's recommendation on leaving (2026-09-23): drive this from Claude Code, not the desktop UI, and do not
invest in GUI development for rebranding work.** The app started as a general document tool with a Qt
interface, but it has niched hard toward PDF rebranding, and that work is better done by script.

This is not a theory. **Both shipped deliveries already ran this way.** The Budget Mailboxes Batch 4 and
MailboxWorks applies were executed by scripts that call the engine directly — no dialogs were involved in
producing either deliverable.

## The entry point

Everything the rebrand pipeline does hangs off `docrefine.worker.Worker`, which takes a callback for progress
events and needs nothing from Qt:

```python
from docrefine.worker import Worker
from docrefine import reviews

w = Worker(callback=lambda e: None)          # or print(e) to watch it
w.run_rebrand_analyze(src, kit, ...)          # classify → writes the review sheet
w.run_rebrand_apply(src, kit, plan, out_dir=..., ...)   # reads the (edited) sheet → produces the delivery
```

Two stages, with a human in the middle. `run_rebrand_analyze` classifies each PDF and writes a review sheet;
somebody reads and corrects that sheet; `run_rebrand_apply` turns the approved sheet into the deliverable.
**The review step is the point of the whole design and must not be automated away** — the playbook is a long
list of things the classifier got confidently wrong that a human caught in the sheet.

## Worked examples already in the repo

Read these before writing anything new. They are the real scripts that produced the real deliveries:

| Script | What it does |
|---|---|
| `verify/trial_run.py` | Builds a small trial plan and applies it — the right shape to copy for a first run. |
| `verify/apply_batch4.py` | The Batch 4 apply. |
| `verify/apply_mbw_rulings.py` | Applies a set of decisions to the MBW sheet. |
| `verify/deliver_qa.py` | Page-by-page QA of a finished delivery. Run it every time. |
| `verify/batch_paths.py` | Per-batch source/kit/output paths, selected with the `DRP_BATCH` environment variable (`batch4` or `mbw`). Import this instead of hard-coding paths — it exists because the same three literals were once copied through seven files and all broke at once. |

## Why this suits Claude Code

The work is not "click a button". A delivery is: survey a corpus, decide what is in scope, classify, review a
few hundred rows of judgement calls, apply, QA the output page by page, then write up what shipped and why.
Most of that is reading, measuring and explaining — and the parts that are mechanical are a handful of Worker
calls. An agent with the repo, the playbook and the verify scripts can do the whole loop. A GUI mostly gets in
the way of the reviewing and the QA, which is where the real effort is.

## What the GUI still gives you, honestly

Not nothing — just nothing you need for a delivery. It offers the folder and brand-kit pickers, the stamp
toggles, live progress and the "Rebrand & Processing Runs" history panel. All of those map to arguments and
config fields you can set directly, and the run history is a JSONL file
(`DocRefinePro_Data/rebrand_runs.jsonl`) you can read yourself.

## The caveat, if you act on this recommendation

**Not developing the GUI is not the same as removing it.** DocRefine Pro is still a released desktop app: CI
builds Windows and macOS binaries on every `v*` tag, and the non-rebrand features — dedup, flatten, OCR,
export — are used through that interface. Deciding to *drop* the GUI is a bigger call than deciding not to add
to it, and it would mean changing what the release actually is. Talk to whoever owns the product before going
that far.

Also note: if you do touch GUI code, seven verification scripts exercise its wiring (`verify_boot.py`,
`verify_regression.py`, `verify_analyze_route.py`, `verify_attrib.py`, `verify_keepnames.py`,
`verify_sleep.py`, `verify_v147.py`). They must still pass. Letting them rot is how you would discover at
release time that the app no longer starts.
