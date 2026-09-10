"""Write the MBW kit's brand.json from Kunchana's decisions — reproducibly.

Every alias below traces to ClickUp `86eyaantb` comment `90180247980961` (2026-08-19),
which answered the eight-item analysis. Written as a script rather than by hand so the
provenance of each group is in version-controllable text and the file can be regenerated.

A brand kit holds DECISIONS, not just images: the artwork is re-downloadable from Drive,
these names are not. Also backed up to DocRefinePro_Data\Brand Kits\.

Usage: python write_mbw_brand_json.py [--imi-consistent] [--write]
  --imi-consistent  fold 'IMI Mailbox Systems' into 'Imperial Mailbox Systems' too.
                    Item 2 said "credit all of them as Imperial Mailbox Systems" but
                    item 1 (confirmed as-is) keeps 'IMI Mailbox Systems' as its own
                    canonical name. Default is the LITERAL reading of item 1 — the two
                    stay separate — because that is what he explicitly confirmed.
"""
import json
import shutil
import sys
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
except Exception:
    pass

import batch_paths                      # also puts the repo on sys.path

WRITE = "--write" in sys.argv
IMI_CONSISTENT = "--imi-consistent" in sys.argv

LIGATURE_FF = "\ufb00"          # 'ﬀ' — no glyph in Poppins-Bold, renders as \x00

# canonical -> raw values seen in the MBW sheet
GROUPS = {
    # --- item 1: carried over verbatim from Budget Mailboxes Batch 4 -------------
    "Florence Corporation": [
        "Florence Mailboxes", "FlorenceMailboxes", "Florence Manufacturing", "Florence",
        "Auth Florence Manufacturing Company", "Auth-Florence Manufacturing Company",
    ],
    "Architectural Mailboxes, LLC": ["Architectural Mailboxes"],
    "Whitehall Products, LLC": [
        "Whitehall", "Whitehall Products", "Whitehall Products LLC",
        "White Hall Products LLC", "WHITEHALL PRODUCTS",
    ],
    "Epoch Design, LLC": ["Epoch Design", "Epoch Design LLC"],
    "Salsbury Industries": ["Salsbury"],
    "bobi": ["Bobi"],                        # BOBI ROUND stays separate — item 3
    "Gaines Manufacturing, Inc.": ["Gaines Manufacturing"],
    "QualArc": ["Qualarc"],
    "Special Lite Products Company, Inc.": [
        "Special Lite Products", "Special Lite Products LLC",
    ],
    # The source spells this with a typographic ff ligature the stamp font cannot
    # draw; his Batch 4 spelling with plain 'ff' is also the fix.
    "TedStuff": [f"TedStu{LIGATURE_FF}", f"Tedstu{LIGATURE_FF}"],

    # --- item 3: new to MBW -----------------------------------------------------
    "AMCO": ["Amco"],
    "The Mail Boss": ["Mail Boss"],

    # --- items 2 and 4: everything that resolves to Imperial --------------------
    # item 2: bare 'IMI' folds in (SKU folders and document text both read Imperial).
    # item 4: the eight product-line / postal-service values he ruled Imperial.
    "Imperial Mailbox Systems": [
        "IMPERIAL MAILBOX SYSTEMS",          # case variant, item 1
        "IMI",                               # item 2
        "USPS", "US Postal Service",          # item 4 — regulator, not the maker
        "Barcelona System", "Quad System", "Twin System", "Norris",
    ],

    # --- item 4: the rest -------------------------------------------------------
    # Balmoral -> Whitehall is already covered by the Whitehall group below.
    "Salsbury Industries ": [],              # placeholder removed below
}

# item 4 continued, added explicitly so each ruling reads as its own line
GROUPS["Whitehall Products, LLC"].append("Balmoral")
GROUPS["Salsbury Industries"].append("Outdoor Parcel Locker")
GROUPS["QualArc"].append("Streetscape")
GROUPS["Architectural Mailboxes, LLC"].append("RetroBox & Uptown")
GROUPS.pop("Salsbury Industries ", None)

# item 4: "I can't name the maker with confidence, and a wrong credit is worse than
# none." An alias to "" makes clean_manufacturer return "" — verified, case-insensitive.
BLANKED = ["Keystone", "Triple Mount", "Town And Country"]

if IMI_CONSISTENT:
    GROUPS["Imperial Mailbox Systems"] += ["IMI Mailbox Systems", "IMI MAILBOX SYSTEMS"]
else:
    GROUPS["IMI Mailbox Systems"] = ["IMI MAILBOX SYSTEMS"]

aliases = {}
for canon, raws in GROUPS.items():
    for raw in raws:
        if raw in aliases and aliases[raw] != canon:
            sys.exit(f"conflict: {raw!r} mapped to both {aliases[raw]!r} and {canon!r}")
        aliases[raw] = canon
for raw in BLANKED:
    aliases[raw] = ""

brand = {
    "_source": [
        "Manufacturer names confirmed by Kunchana Godahewa on ClickUp task 86eyaantb,",
        "comment 90180247980961 (2026-08-19), answering the eight-item MBW analysis.",
        "Groups marked 'item 1' are carried over verbatim from his Budget Mailboxes",
        "Batch 4 decisions so the same company is credited identically across brands.",
        "",
        "Settings that go with this kit for MailboxWorks:",
        "  Manufacturer attribution in the page footer  ON",
        "  Tagline                                      OFF  (MailboxWorks has no tagline)",
        "  Version and last-updated line                ON",
        "  Standard disclaimer                          ON",
        "  Manufacturer line on covers                  OFF  (footer placement, per SOP)",
        "  Keep the original filenames                  ON",
        "  Complete set                                 OFF",
        "",
        "An alias whose value is an empty string prints NO attribution. Used for the",
        "values he could not attribute with confidence (Keystone, Triple Mount, Town",
        "And Country) - a wrong credit is worse than an absent one.",
        "",
        "last_updated is deliberately EMPTY: the engine composes 'Last Updated: <Month",
        "Year>' from the run date. Filling it overrides the whole suffix and silently",
        "drops the 'Last Updated:' label the SOP requires.",
    ],
    "name": "MailboxWorks",
    "slug": "mailboxworks",
    "tagline": "",
    "disclaimer": "This guide is provided for reference. Always consult manufacturer specifications for complete details.",
    "version_label": "Version 1.0",
    "last_updated": "",
    "attribution": "Manufactured by {manufacturer} | Sold by {brand}",
    "manufacturer_aliases": dict(sorted(aliases.items(), key=lambda kv: kv[0].lower())),
}

dst = batch_paths.KIT / "brand.json"
backup = Path(r"C:\Users\WORK\Documents\DocRefinePro_Data\Brand Kits\mailboxworks-brand.json")

print(f"IMI reading      : {'CONSISTENT (folded)' if IMI_CONSISTENT else 'LITERAL (kept separate)'}")
print(f"aliases          : {len(aliases)}  ({len(BLANKED)} deliberately blank)")
print(f"canonical names  : {len({v for v in aliases.values() if v})}")
print(f"target           : {dst}")

if not WRITE:
    print("\nDRY RUN. Alias map:")
    for k, v in brand["manufacturer_aliases"].items():
        print(f"    {k!r:38} -> {v!r}")
    sys.exit(0)

# The filename MUST be exactly brand.json — anything else is ignored and BrandKit
# silently falls back to defaults naming Budget Mailboxes.
assert dst.name == "brand.json", dst
dst.write_text(json.dumps(brand, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
backup.parent.mkdir(parents=True, exist_ok=True)
shutil.copy2(dst, backup)
print(f"\nwrote {dst}\nbacked up to {backup}")

# Read it back through BrandKit — the only check that matters.
from docrefine.rebrand import BrandKit
from docrefine import stamps as S
k = BrandKit(batch_paths.KIT)
assert k.brand_name == "MailboxWorks", k.brand_name
assert k.brand_slug == "mailboxworks", k.brand_slug
assert not str(k.brand.get("tagline") or "").strip(), "tagline must be empty"
assert not str(k.brand.get("last_updated") or "").strip(), "last_updated must be empty"
got = k.brand.get("manufacturer_aliases") or {}
assert got == brand["manufacturer_aliases"], "aliases did not round-trip"
print(f"BrandKit resolves: {k.brand_name!r}, {len(got)} aliases, "
      f"tagline empty, last_updated empty")
for raw, expect in (("IMI", "Imperial Mailbox Systems"), ("USPS", "Imperial Mailbox Systems"),
                    ("Keystone", ""), ("Amco", "AMCO"), ("Mail Boss", "The Mail Boss"),
                    (f"TedStu{LIGATURE_FF}", "TedStuff"), ("BOBI ROUND", "BOBI ROUND")):
    got1 = S.clean_manufacturer(raw, k.brand_name, got)
    assert got1 == expect, f"{raw!r} -> {got1!r}, expected {expect!r}"
    print(f"    {raw!r:22} -> {got1!r}")
print("spot checks pass")
