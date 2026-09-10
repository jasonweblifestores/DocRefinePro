"""Let vision look at the rows a heuristic left behind without seeing them.

The text model reads these files confidently enough that `needs_a_look` never
queued them, so their action was settled by filename and page shape alone — and on
164 MBW rows the model's own doc_type says "installation guide" while the action
says "leave". The sheet's own note asks for exactly this check ("Flip to rebrand if
this one is really a guide").

Batch 4 is the precedent: where vision was allowed to look, it pulled 62 genuine
guides INTO branding that the filename rule had skipped, and Kunchana called that
the real catch of the batch. So this asks the same question of the same class of
file, using num_ctx large enough that the images are not rejected outright.

Reports only. Writing the sheet is a separate, explicit step.

Usage: [set DRP_BATCH=mbw] python review_leave_guides.py [--write] [--limit N]
"""
import json
import shutil
import sys
import time
import urllib.request
from collections import Counter
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
LIMIT = int(sys.argv[sys.argv.index("--limit") + 1]) if "--limit" in sys.argv else None
NUM_CTX = 4096          # 2048 rejects some renders outright with HTTP 400
T0 = time.time()


def stamp():
    e = time.time() - T0
    return f"[{int(e // 60):02d}:{int(e % 60):02d}]"


def ask(pdf, num_ctx=NUM_CTX):
    images = C.render_pages_b64(str(pdf))
    if not images:
        return None, "no render"
    payload = json.dumps({
        "model": C.DEFAULT_VISION_MODEL, "prompt": C._VISION_PROMPT, "images": images,
        "stream": False, "format": "json", "keep_alive": "15m",
        "options": {"temperature": 0.1, "num_predict": 220, "num_ctx": num_ctx},
    }).encode()
    req = urllib.request.Request(C.OLLAMA_URL + "/api/generate", data=payload,
                                 headers={"Content-Type": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=C.VISION_TIMEOUT) as resp:
            raw = json.loads(resp.read())["response"]
        return json.loads(raw), ""
    except Exception as e:
        return None, f"{type(e).__name__}: {str(e)[:120]}"


def candidates(rows):
    """Leave rows the text model itself typed as a guide, never seen by vision."""
    out = []
    for r in rows:
        if (r.get("action") or "").strip().lower() == "rebrand":
            continue
        if (r.get("source") or "").strip().lower() == "vision":
            continue          # vision already looked and said leave — respect it
        blob = f"{r.get('doc_type','')} {r.get('asset_type','')}".lower()
        if "guide" in blob or "manual" in blob or "instruction" in blob:
            out.append(r)
    return out


def main():
    plan = batch_paths.plan()
    rows = reviews.read_plan(plan)
    todo = candidates(rows)
    if LIMIT:
        todo = todo[:LIMIT]
    print(f"sheet : {plan}")
    print(f"rows  : {len(rows)}")
    print(f"asking vision about {len(todo)} 'leave' rows the text model typed as a guide\n",
          flush=True)

    flips, confirms, failed = [], [], []
    for i, r in enumerate(todo, 1):
        name = r["file"]
        d, err = ask(batch_paths.SRC / name)
        if d is None:
            failed.append((name, err))
            print(f"{stamp()} {i}/{len(todo)}  NO ANSWER  {name}  ({err})", flush=True)
            continue
        action = str(d.get("action", "")).strip().lower()
        if action not in ("rebrand", "leave"):
            failed.append((name, f"action={action!r}"))
            continue
        try:
            conf = round(float(d.get("confidence", 0) or 0), 2)
        except (TypeError, ValueError):
            conf = 0.0
        if action == "rebrand":
            flips.append((r, d, conf))
            print(f"{stamp()} {i}/{len(todo)}  FLIP -> rebrand  conf={conf}  {name}", flush=True)
            print(f"          {str(d.get('evidence') or d.get('notes') or '')[:150]}", flush=True)
        else:
            confirms.append((r, d, conf))
            if i % 25 == 0:
                print(f"{stamp()} {i}/{len(todo)}  (confirmed leave x{len(confirms)})", flush=True)

    try:
        C.unload_model(C.DEFAULT_VISION_MODEL)
    except Exception:
        pass

    print(f"\n{stamp()} done in {(time.time()-T0)/60:.1f} min")
    print(f"  vision confirms LEAVE : {len(confirms)}")
    print(f"  vision says REBRAND   : {len(flips)}")
    print(f"  no usable answer      : {len(failed)}")
    for n, e in failed[:10]:
        print(f"      {n}  ({e})")

    if flips:
        print(f"\n  the {len(flips)} it would pull INTO branding:")
        for r, d, conf in flips:
            print(f"    conf={conf:<5} {r['file']}")
            print(f"        title={str(d.get('title') or '')[:60]!r} "
                  f"manufacturer={str(d.get('manufacturer') or '')!r}")
        print("\n  titles they would carry:")
        for t, n in Counter(str(d.get("title") or "") for _, d, _ in flips).most_common(12):
            print(f"    {n:4d}  {t!r}")

    if not WRITE:
        print("\nREPORT ONLY — sheet untouched. Re-run with --write to apply.")
        return 0

    for r, d, conf in flips:
        r["action"] = "rebrand"
        r["source"] = "vision"
        r["confidence"] = str(conf)
        for k in ("doc_type", "asset_type", "title", "manufacturer", "product"):
            v = str(d.get(k, "") or "").strip()
            if v:
                r[k] = v
        r["notes"] = ("vision looked at the page after the text pass had left it on a "
                      "filename/page-shape rule; its answer wins")
    for r, d, conf in confirms:
        r["source"] = "vision"
        r["confidence"] = str(conf)
        r["notes"] = "vision confirmed this is a drawing, not a guide"

    bak = plan.with_name(plan.stem + ".previous" + plan.suffix)
    if not bak.exists():
        shutil.copy2(plan, bak)
        print(f"\noriginal kept as {bak.name}")
    reviews.write_plan(plan, rows, Worker.REBRAND_PLAN_COLUMNS,
                       src_root=str(batch_paths.SRC),
                       asset_types=sorted(ASSET_TYPE_TITLES))
    again = reviews.read_plan(plan)
    reb = sum(1 for r in again if (r.get("action") or "").strip().lower() == "rebrand")
    print(f"written and re-read: {len(again)} rows, {reb} rebrand / {len(again)-reb} leave")
    assert len(again) == len(rows), "row count changed"
    return 0


if __name__ == "__main__":
    sys.exit(main())
