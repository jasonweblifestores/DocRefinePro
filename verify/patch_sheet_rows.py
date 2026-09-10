"""Replace specific sheet rows with a vision result the run couldn't obtain.

Why this exists: `VISION_NUM_CTX = 2048` is too small for some renders. Ollama
answers HTTP 400 `exceed_context_size_error`, `classify_visually` catches it and
returns None, and the file silently drops to the filename heuristic at confidence
0 — the visual pass quietly not applying to a file it exists for. On MBW that hit
2 of 456 files, and for one of them the filename guess was WRONG (a compatibility
chart marked "leave" that vision reads as a spec sheet to rebrand at 0.95).

Re-asking with a context big enough for the images is what the run would have done
had the ceiling been high enough, so this writes that answer into the sheet rather
than hand-editing a decision. The original sheet is kept as `.previous.xlsx` so
`compare_sheets.py` can show exactly what moved.

The repo is deliberately NOT patched here: MBW ships on clean released v157, the
same reason v156 was held off Batch 4's delivery. The context ceiling is a v158 fix.

Usage: [set DRP_BATCH=mbw] python patch_sheet_rows.py [--write]
"""
import json
import shutil
import sys
import time
import urllib.request
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
except Exception:
    pass

import batch_paths                      # also puts the repo on sys.path

from docrefine import classify as C, reviews
from docrefine.rebrand import ASSET_TYPE_TITLES
from docrefine.worker import Worker

WRITE = "--write" in sys.argv
NUM_CTX = 4096          # measured on Batch 4 as no slower than 2048
TARGETS = ["mailbox_post_matrix-0520-3.pdf",
           "whitehall-quad-post-installation-instructions.pdf"]


def ask(pdf, num_ctx=NUM_CTX):
    """The same request classify_visually makes, with a context that fits."""
    images = C.render_pages_b64(str(pdf))
    if not images:
        return None, 0.0
    payload = json.dumps({
        "model": C.DEFAULT_VISION_MODEL, "prompt": C._VISION_PROMPT, "images": images,
        "stream": False, "format": "json", "keep_alive": "5m",
        "options": {"temperature": 0.1, "num_predict": 220, "num_ctx": num_ctx},
    }).encode()
    req = urllib.request.Request(C.OLLAMA_URL + "/api/generate", data=payload,
                                 headers={"Content-Type": "application/json"})
    t = time.time()
    with urllib.request.urlopen(req, timeout=C.VISION_TIMEOUT) as resp:
        raw = json.loads(resp.read())["response"]
    return json.loads(raw), time.time() - t


def main():
    plan = batch_paths.plan()
    rows = reviews.read_plan(plan)
    by_file = {r["file"]: r for r in rows}
    print(f"sheet : {plan}")
    print(f"rows  : {len(rows)}\n")

    changes = []
    for name in TARGETS:
        row = by_file.get(name)
        if row is None:
            sys.exit(f"{name} is not in the sheet")
        if (row.get("source") or "").strip().lower() != "fallback":
            print(f"  SKIP {name}: source is {row.get('source')!r}, not 'fallback' "
                  f"— already decided by a model, leaving it alone")
            continue

        d, dt = ask(batch_paths.SRC / name)
        if not d:
            print(f"  {name}: still no answer — leaving the fallback in place")
            continue
        action = str(d.get("action", "")).strip().lower()
        if action not in ("rebrand", "leave"):
            print(f"  {name}: unusable action {action!r} — leaving the fallback")
            continue

        before = dict(row)
        row["action"] = action
        row["source"] = "vision"
        for k in ("doc_type", "asset_type", "title", "manufacturer", "product"):
            v = str(d.get(k, "") or "").strip()
            if v:
                row[k] = v
        try:
            row["confidence"] = str(round(float(d.get("confidence", 0) or 0), 2))
        except (TypeError, ValueError):
            row["confidence"] = "0"
        row["notes"] = (f"re-read with num_ctx={NUM_CTX}; the run's 2048 ceiling "
                        f"rejected the images (HTTP 400) and it fell back to the filename")

        flipped = (before.get("action") or "").strip().lower() != action
        changes.append((name, before, dict(row), flipped, dt))
        print(f"  {name}  ({dt:.1f}s){'  ** DECISION FLIPPED **' if flipped else ''}")
        print(f"      action     {before.get('action')!r} -> {row['action']!r}")
        print(f"      confidence {before.get('confidence')!r} -> {row['confidence']!r}")
        print(f"      title      {before.get('title')!r} -> {row['title']!r}")
        print(f"      manufacturer {before.get('manufacturer')!r} -> {row['manufacturer']!r}")

    try:
        C.unload_model(C.DEFAULT_VISION_MODEL)
    except Exception:
        pass

    if not changes:
        print("\nnothing to change")
        return 0

    reb = sum(1 for r in rows if (r.get("action") or "").strip().lower() == "rebrand")
    print(f"\nsheet would become: {reb} rebrand / {len(rows) - reb} leave")
    print(f"decision flips    : {sum(1 for c in changes if c[3])}")

    if not WRITE:
        print("\nDRY RUN — sheet not written. Re-run with --write")
        return 0

    bak = plan.with_name(plan.stem + ".previous" + plan.suffix)
    shutil.copy2(plan, bak)
    print(f"\noriginal kept as {bak.name}")
    reviews.write_plan(plan, rows, Worker.REBRAND_PLAN_COLUMNS,
                       src_root=str(batch_paths.SRC),
                       asset_types=sorted(ASSET_TYPE_TITLES))

    # Read it back: the sheet on disk is what Apply will use, not our in-memory list.
    again = reviews.read_plan(plan)
    got = {r["file"]: r for r in again}
    reb2 = sum(1 for r in again if (r.get("action") or "").strip().lower() == "rebrand")
    print(f"written and re-read: {len(again)} rows, {reb2} rebrand / {len(again)-reb2} leave")
    assert len(again) == len(rows), f"row count changed: {len(rows)} -> {len(again)}"
    for name, _, after, _, _ in changes:
        landed = got[name]
        print(f"  {name}: action={landed['action']!r} source={landed['source']!r} "
              f"conf={landed['confidence']!r}")
        assert landed["action"].strip().lower() == after["action"], \
            f"{name}: wrote {after['action']!r} but read back {landed['action']!r}"
    print("every patched row reads back as written")
    return 0


if __name__ == "__main__":
    sys.exit(main())
