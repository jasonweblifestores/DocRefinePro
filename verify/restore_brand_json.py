"""Rebuild the Batch 4 brand.json and put it somewhere that survives cleanup.

The kit folder it lived in has been deleted, and with it the only copy of the
manufacturer decisions that gated the whole delivery: Kunchana's 19 canonical
names, plus the 8 case-variant groups he approved afterwards. The artwork can be
re-downloaded from Drive; these decisions cannot be re-derived, only re-asked.

Writes to the app's own data directory, which persists, rather than a scratchpad.
"""
import json
import shutil
from pathlib import Path

SP = Path(__file__).resolve().parent
BACKUP = SP / "brand.json.backup-pre-alias-normalise"
DEST_DIR = Path.home() / "Documents" / "DocRefinePro_Data" / "Brand Kits"
DEST = DEST_DIR / "budget-mailboxes-brand.json"

# The 8 groups added after Kunchana's list, approved by him on 2026-08-13.
# Three apply a canonical name he had already supplied; the other five are
# case-only variants he confirmed, including `bobi` as the brand's own styling.
ADDED = {
    "ABC STEEL CO.": "ABC Steel Co.",
    "Architectural Mailboxes LLC": "Architectural Mailboxes, LLC",
    "Bobi": "bobi",
    "FLORENCE CORPORATION": "Florence Corporation",
    "GOT IT WHOLESALE": "Got It Wholesale",
    "IMI MAILBOX SYSTEMS": "IMI Mailbox Systems",
    "MAYNE": "Mayne",
    "QUALARC": "QualArc",
    "Qualarc": "QualArc",
}

brand = json.loads(BACKUP.read_text(encoding="utf-8"))
al = brand.get("manufacturer_aliases") or {}
before = len(al)
al.update(ADDED)
brand["manufacturer_aliases"] = dict(sorted(al.items(), key=lambda kv: kv[0].lower()))

note = brand.get("_source") or []
if isinstance(note, list):
    note = note + [
        "",
        "RECOVERED 2026-08-14. The kit folder (Batch 4\\Template) was deleted after",
        "the delivery was uploaded; this is the reconstructed wording + alias map.",
        "Artwork is re-downloadable from Drive, these decisions are not.",
        "8 further case-variant groups approved by Kunchana on 2026-08-13 are folded in.",
    ]
    brand["_source"] = note

DEST_DIR.mkdir(parents=True, exist_ok=True)
DEST.write_text(json.dumps(brand, indent=2, ensure_ascii=False), encoding="utf-8")

print(f"aliases: {before} -> {len(brand['manufacturer_aliases'])}")
print(f"tagline    : {brand.get('tagline')!r}")
print(f"disclaimer : {str(brand.get('disclaimer'))[:60]!r}...")
print(f"written to : {DEST}")

# sanity: does it still load through the real code path?
import sys
sys.path.insert(0, r"C:\Users\WORK\Documents\WebLife Labs\PROJECTS\DocRefine Pro\DocRefinePro")
tmp = DEST_DIR / "_verify_kit"
shutil.rmtree(tmp, ignore_errors=True)
(tmp / "Portrait").mkdir(parents=True)
shutil.copy2(DEST, tmp / "brand.json")
from docrefine.rebrand import BrandKit
bk = BrandKit(tmp)
print(f"\nparsed by BrandKit: name={bk.brand.get('name')!r} slug={bk.brand.get('slug')!r} "
      f"aliases={len(bk.brand.get('manufacturer_aliases') or {})}")
from docrefine import stamps as S
for raw in ("MAYNE", "Salsbury", "Florencemailboxes.com", "QUALARC", "Budget Mailboxes"):
    print(f"   {raw!r:26} -> "
          f"{S.clean_manufacturer(raw, bk.brand_name, bk.brand.get('manufacturer_aliases'))!r}")
shutil.rmtree(tmp, ignore_errors=True)
