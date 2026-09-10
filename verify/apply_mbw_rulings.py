"""Apply Kunchana's item 5 and item 8 rulings to the MBW review sheet.

Source: ClickUp 86eyaantb comment 90180247980961 (2026-08-19).

Item 5 — "leave all 14 unbranded ... So don't brand the 3 currently in scope either".
  11 of the 14 printable templates were already `leave`; these are the 3 that were not.

Item 8 — "confirmed, all 13". These documents name no manufacturer anywhere in their
  text, so the value is an inference from the filename and the SKU prefix that he has
  explicitly approved. Written into the sheet as the canonical name directly, since an
  alias cannot rescue a blank value.
  Two amendments he made in the same comment:
    streetscape-maintenance.pdf -> QualArc  (not Streetscape, per his item 4)
    PP-Curb-Planter.pdf         -> no fill; it is one of the 14 templates, so it leaves

Items 1-4 are handled by the kit's manufacturer_aliases, not here — see
write_mbw_brand_json.py. Item 6 is the filename survivors, item 7 the manifests.

Usage: [set DRP_BATCH=mbw] python apply_mbw_rulings.py [--write]
"""
import shutil
import sys

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
except Exception:
    pass

import batch_paths                      # also puts the repo on sys.path

from docrefine import reviews, stamps as S
from docrefine.rebrand import ASSET_TYPE_TITLES, BrandKit
from docrefine.worker import Worker

WRITE = "--write" in sys.argv

# item 5 — the 3 templates still marked rebrand. The other 11 are already `leave`.
TO_LEAVE = [
    "PP-Curb-Planter.pdf",
    "PP-Large-Plaque.pdf",
    "Standing_Tall_Planter.pdf",
]
LEAVE_REASON = ("printable to-scale template - Kunchana item 5: leave all 14 unbranded, "
                "branding puts the watermark over the mounting-screw measurements")

# item 8 — file -> approved manufacturer
FILLS = {
    "dvault-frontaccess-installation3.pdf": "dVault",
    "gaines-classic-installation1.pdf": "Gaines Manufacturing, Inc.",
    "salsbury-vertical-mailbox-installation-instructions.pdf": "Salsbury Industries",
    "grande-s-round.pdf": "bobi",
    "mailbox-deluxe-post-assembly-instructions.pdf": "Whitehall Products, LLC",
    "mailbox-standard-post-assembly-instructions.pdf": "Whitehall Products, LLC",
    "whitehall-quad-post-installation-instructions.pdf": "Whitehall Products, LLC",
    "Newspaper-Holder-6-Assembly-Instructions_Revised-2.pdf": "Imperial Mailbox Systems",
    "Newspaper-Holder-6-Assembly-Instructions_Revised-4.pdf": "Imperial Mailbox Systems",
    "Newspaper-Holder-6-Assembly-Instructions_Revised-5.pdf": "Imperial Mailbox Systems",
    "streetscape-maintenance.pdf": "QualArc",
    "ecco-7-installation1.pdf": "Ecco",
    "E4-Installation.pdf": "Ecco",
}
FILL_REASON = ("no manufacturer in the document text; filename + SKU prefix agree and "
               "Kunchana approved the inference (item 8)")


def main():
    plan = batch_paths.plan()
    rows = reviews.read_plan(plan)
    by_file = {r["file"]: r for r in rows}
    kit = BrandKit(batch_paths.require_kit())
    AL = kit.brand.get("manufacturer_aliases") or {}
    BRAND = kit.brand_name

    def counts():
        reb = [r for r in rows if (r.get("action") or "").strip().lower() == "rebrand"]
        attrib = sum(1 for r in reb
                     if S.clean_manufacturer(r.get("manufacturer"), BRAND, AL))
        return len(reb), len(rows) - len(reb), attrib

    r0, l0, a0 = counts()
    print(f"sheet  : {plan}")
    print(f"kit    : {len(AL)} aliases, brand {BRAND!r}")
    print(f"before : {r0} rebrand / {l0} leave, {a0} carrying an attribution line\n")

    print("item 5 — templates out of scope:")
    moved = 0
    for name in TO_LEAVE:
        r = by_file.get(name)
        if r is None:
            sys.exit(f"{name} is not in the sheet")
        was = (r.get("action") or "").strip().lower()
        if was == "leave":
            print(f"    already leave: {name}")
            continue
        r["action"] = "leave"
        r["notes"] = LEAVE_REASON
        moved += 1
        print(f"    rebrand -> leave  {name}")

    print("\nitem 8 — approved manufacturer fills:")
    filled = 0
    for name, who in FILLS.items():
        r = by_file.get(name)
        if r is None:
            sys.exit(f"{name} is not in the sheet")
        if (r.get("action") or "").strip().lower() != "rebrand":
            print(f"    SKIP {name}: action is {r.get('action')!r}, would carry no line")
            continue
        had = (r.get("manufacturer") or "").strip()
        if had:
            print(f"    SKIP {name}: already has {had!r}")
            continue
        r["manufacturer"] = who
        r["notes"] = FILL_REASON
        filled += 1
        # It must survive clean_manufacturer, or the fill prints nothing.
        got = S.clean_manufacturer(who, BRAND, AL)
        assert got == who, f"{name}: {who!r} would print as {got!r}"
        print(f"    {who:32} <- {name}")

    r1, l1, a1 = counts()
    print(f"\nafter  : {r1} rebrand / {l1} leave, {a1} carrying an attribution line")
    print(f"changed: {moved} moved to leave, {filled} manufacturers filled")
    assert len(rows) == r1 + l1

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
    for name in TO_LEAVE:
        assert got[name]["action"].strip().lower() == "leave", name
    for name, who in FILLS.items():
        if (got[name].get("action") or "").strip().lower() == "rebrand":
            assert got[name]["manufacturer"].strip() == who, \
                f"{name}: read back {got[name]['manufacturer']!r}"
    reb = sum(1 for r in again if (r.get("action") or "").strip().lower() == "rebrand")
    print(f"written and re-read: {len(again)} rows, {reb} rebrand / {len(again)-reb} leave")
    print("every ruling reads back as written")
    return 0


if __name__ == "__main__":
    sys.exit(main())
