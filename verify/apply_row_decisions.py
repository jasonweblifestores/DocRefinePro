"""Apply named, reasoned row decisions to a review sheet — one table, auditable.

Deliberately explicit rather than a rule: each change names the file, the new
action and why, so the sheet's history is readable months later. Anything that is
a matter of presentation rather than correctness is NOT listed here — it goes to
the client instead.

Usage: [set DRP_BATCH=mbw] python apply_row_decisions.py [--write]
"""
import shutil
import sys

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
except Exception:
    pass

import batch_paths                      # also puts the repo on sys.path

from docrefine import reviews
from docrefine.rebrand import ASSET_TYPE_TITLES
from docrefine.worker import Worker

WRITE = "--write" in sys.argv

# (file, new action, fields to set, reason)
DECISIONS = [
    ("1570-cbu-foundation-plan-pad-spec_2.pdf", "rebrand",
     {"doc_type": "installation guide", "asset_type": "installation-guide",
      "title": "Concrete Foundation Preparation", "manufacturer": "Florence Corporation",
      "confidence": "0.95", "source": "vision"},
     "pages 6-7 of a 16-page Florence installation manual - prose instructions "
     "('consult local building codes', 'refer to Table 1'), not a dimensioned drawing"),

    ("1570-cbu-foundation-plan-pad-spec_2-1-1.pdf", "rebrand",
     {"doc_type": "installation guide", "asset_type": "installation-guide",
      "title": "Concrete Foundation Preparation", "manufacturer": "Florence Corporation",
      "confidence": "0.9", "source": "vision"},
     "same document under a second filename"),

    ("Timeless-Design-Durable-Mail-Solutions-compressed.pdf", "rebrand",
     {"doc_type": "product catalog", "asset_type": "product-catalog",
      "title": "Bobi Mailbox", "manufacturer": "bobi",
      "confidence": "0.95", "source": "vision"},
     "9-page Bobi marketing brochure ('a most wanted mailbox in Finland for over "
     "30 years'), left as-is only because the filename gave nothing away"),

    # Added after the visual review of the 328 branded files the text model decided
    # alone. Approved by Jason 2026-08-21; takes the branded count 471 -> 470, so it
    # is disclosed to Kunchana in the delivery note rather than moved quietly.
    ("grande-s-round.pdf", "leave",
     {"doc_type": "dimensioned-technical-drawing", "asset_type": "dimensioned-technical-drawing",
      "title": "Dimensioned Technical Drawing", "confidence": "0.95", "source": "vision"},
     "line art with measurement callouts (165, ca 60, o 3.8, ca. 20-30 cm), a scale "
     "icon and no prose - 'BOBI GRANDE S + BOBI ROUND, cm'. The brief leaves "
     "dimensioned drawings as-is, and 499 others in this batch already are. Confirmed "
     "by eye, not only by the model"),
]


def main():
    plan = batch_paths.plan()
    rows = reviews.read_plan(plan)
    by_file = {r["file"]: r for r in rows}

    before_reb = sum(1 for r in rows if (r.get("action") or "").strip().lower() == "rebrand")
    print(f"sheet : {plan}")
    print(f"before: {before_reb} rebrand / {len(rows) - before_reb} leave\n")

    applied = 0
    for name, action, fields, reason in DECISIONS:
        r = by_file.get(name)
        if r is None:
            sys.exit(f"{name} is not in the sheet")
        was = (r.get("action") or "").strip().lower()
        if was == action:
            print(f"  already {action}: {name}")
            continue
        r["action"] = action
        r.update(fields)
        r["notes"] = reason
        applied += 1
        print(f"  {was} -> {action}  {name}")
        print(f"      {reason}")

    after_reb = sum(1 for r in rows if (r.get("action") or "").strip().lower() == "rebrand")
    print(f"\nafter : {after_reb} rebrand / {len(rows) - after_reb} leave   ({applied} changed)")

    if not WRITE:
        print("\nDRY RUN — sheet untouched. Re-run with --write")
        return 0

    bak = plan.with_name(plan.stem + ".previous" + plan.suffix)
    if not bak.exists():
        shutil.copy2(plan, bak)
        print(f"original kept as {bak.name}")
    reviews.write_plan(plan, rows, Worker.REBRAND_PLAN_COLUMNS,
                       src_root=str(batch_paths.SRC),
                       asset_types=sorted(ASSET_TYPE_TITLES))

    again = reviews.read_plan(plan)
    got = {r["file"]: r for r in again}
    assert len(again) == len(rows), f"row count changed: {len(rows)} -> {len(again)}"
    for name, action, _, _ in DECISIONS:
        assert got[name]["action"].strip().lower() == action, \
            f"{name}: wrote {action!r}, read back {got[name]['action']!r}"
    reb = sum(1 for r in again if (r.get("action") or "").strip().lower() == "rebrand")
    print(f"written and re-read: {len(again)} rows, {reb} rebrand / {len(again)-reb} leave")
    print("every decision reads back as written")
    return 0


if __name__ == "__main__":
    sys.exit(main())
