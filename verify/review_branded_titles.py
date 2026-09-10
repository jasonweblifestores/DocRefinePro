"""Let vision look at the branded documents the text model decided alone.

Why: on MBW, 328 of the 471 branded files were classified by the text model at
0.9-1.0 confidence and therefore never queued for the visual pass — `needs_a_look`
only queues what is unreadable or uncertain. A confident wrong answer is invisible
to that rule. Kunchana found the consequence in the shipped Batch 4 set: product
catalogues and brochures carrying a cover title of "INSTALLATION MANUAL".

The cover title is the largest text on page one, so a wrong one is the most visible
defect we can ship. This asks the model that can actually see the page.

Deliberately REPORT-ONLY, and it separates two very different kinds of disagreement:

  TITLE      vision reads a different document type. Cheap to correct, and not
             something the client signed off line by line.
  ACTION     vision would brand something we are leaving, or vice versa. This
             CHANGES COUNTS THE CLIENT APPROVED, so it is listed for a human, never
             applied. The MBW template case is why: vision was right that 11 files
             were "installation aids" and still wrong that they should be branded.

Usage: [set DRP_BATCH=mbw] python review_branded_titles.py [--limit N] [--all]
  --all  also examine rows vision has already seen (default: only text-decided)
"""
import json
import sys
import time
import urllib.request
from collections import Counter

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
except Exception:
    pass

import batch_paths                      # also puts the repo on sys.path

from docrefine import classify as C, reviews

LIMIT = int(sys.argv[sys.argv.index("--limit") + 1]) if "--limit" in sys.argv else None
INCLUDE_SEEN = "--all" in sys.argv
NUM_CTX = 4096          # 2048 rejects some renders outright with HTTP 400
T0 = time.time()


def stamp():
    e = time.time() - T0
    return f"[{int(e // 60):02d}:{int(e % 60):02d}]"


def ask(pdf):
    images = C.render_pages_b64(str(pdf))
    if not images:
        return None, "no render"
    payload = json.dumps({
        "model": C.DEFAULT_VISION_MODEL, "prompt": C._VISION_PROMPT, "images": images,
        "stream": False, "format": "json", "keep_alive": "15m",
        "options": {"temperature": 0.1, "num_predict": 220, "num_ctx": NUM_CTX},
    }).encode()
    req = urllib.request.Request(C.OLLAMA_URL + "/api/generate", data=payload,
                                 headers={"Content-Type": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=C.VISION_TIMEOUT) as resp:
            return json.loads(json.loads(resp.read())["response"]), ""
    except Exception as e:
        return None, f"{type(e).__name__}: {str(e)[:100]}"


def norm(s):
    return " ".join(str(s or "").lower().split())


def main():
    plan = batch_paths.plan()
    rows = reviews.read_plan(plan)
    todo = [r for r in rows
            if (r.get("action") or "").strip().lower() == "rebrand"
            and (INCLUDE_SEEN or (r.get("source") or "").strip().lower() != "vision")]
    if LIMIT:
        todo = todo[:LIMIT]
    print(f"sheet : {plan}")
    print(f"asking vision about {len(todo)} branded rows "
          f"({'all' if INCLUDE_SEEN else 'text-decided only'})\n", flush=True)

    title_diff, action_diff, failed, agreed = [], [], [], 0
    for i, r in enumerate(todo, 1):
        d, err = ask(batch_paths.SRC / r["file"])
        if d is None:
            failed.append((r["file"], err))
            print(f"{stamp()} {i}/{len(todo)}  NO ANSWER  {r['file']}  ({err})", flush=True)
            continue
        act = str(d.get("action", "")).strip().lower()
        new_t = str(d.get("title") or "").strip()
        old_t = (r.get("title") or "").strip()
        try:
            conf = round(float(d.get("confidence", 0) or 0), 2)
        except (TypeError, ValueError):
            conf = 0.0

        if act == "leave":
            action_diff.append((r, d, conf))
            print(f"{stamp()} {i}/{len(todo)}  ACTION  vision says LEAVE  {r['file']}", flush=True)
        if new_t and norm(new_t) != norm(old_t):
            title_diff.append((r, old_t, new_t, conf, str(d.get("asset_type") or "")))
            print(f"{stamp()} {i}/{len(todo)}  TITLE   {old_t!r} -> {new_t!r}   {r['file']}",
                  flush=True)
        elif act != "leave":
            agreed += 1
        if i % 50 == 0:
            print(f"{stamp()} ...{i}/{len(todo)}  "
                  f"titles differing {len(title_diff)}, action differing {len(action_diff)}",
                  flush=True)

    try:
        C.unload_model(C.DEFAULT_VISION_MODEL)
    except Exception:
        pass

    print(f"\n{stamp()} done in {(time.time()-T0)/60:.1f} min")
    print(f"  agreed outright        {agreed}")
    print(f"  different TITLE        {len(title_diff)}")
    print(f"  different ACTION       {len(action_diff)}   (for a human, not applied)")
    print(f"  no usable answer       {len(failed)}")

    if title_diff:
        print(f"\n  most common replacement titles:")
        for t, n in Counter(t[2] for t in title_diff).most_common(15):
            print(f"    {n:4d}  {t}")
    if action_diff:
        print(f"\n  vision would LEAVE these {len(action_diff)} currently-branded files:")
        for r, d, conf in action_diff[:25]:
            print(f"    conf={conf:<5} {r['file'][:56]}")

    out = plan.with_name(f"{batch_paths.BATCH}_vision-title-review.json")
    out.write_text(json.dumps({
        "title_changes": [{"file": r["file"], "old": o, "new": nt,
                           "confidence": c, "asset_type": at}
                          for r, o, nt, c, at in title_diff],
        "action_changes": [{"file": r["file"], "confidence": c,
                            "vision": {k: d.get(k) for k in
                                       ("action", "doc_type", "asset_type", "title")}}
                           for r, d, c in action_diff],
        "failed": [{"file": f, "error": e} for f, e in failed],
    }, indent=2), encoding="utf-8")
    print(f"\nwrote {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
