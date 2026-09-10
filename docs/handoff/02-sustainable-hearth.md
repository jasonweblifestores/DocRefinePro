# Runbook — "European Home → Sustainable Hearth" PDFs

**ClickUp [86eypep2z](https://app.clickup.com/t/86eypep2z). Status as of 2026-09-10: scope not agreed,
nothing built, waiting on a TA Task Request from Kunchana Godahewa.** If this comes back, start here.

## The one thing to get right

The task was handed over as *"the same rebrand treatment as your BM batches"*. **It is not.** Our pipeline
wraps an **untouched** original page in *our* retailer branding. This job asks to change **the vendor's own
branding printed inside the page** — their logo, their company name, their address. The app has never done
that and was deliberately designed not to.

If you accept it as a like-for-like rebrand you will discover this three days in. The analysis below was done
on 2026-08-28 so that nobody has to discover it again.

## What was measured

**13 documents** across the two shipped deliveries carry "European Home" — 10 in Budget Mailboxes Batch 4,
3 in MailboxWorks. Where the old name actually lives, and whether it can be changed:

| Where it appears | Docs | Changeable? |
|---|---:|---|
| **Our own attribution line** (`Manufactured by European Home | Sold by …`) | 6 | **Yes.** Brand-kit edit + re-run. This is the cheap, safe part. |
| Page text in an ordinary font | 3 | Technically yes — but it is an address block, not a name. See below. |
| Page text in a **subset font with a scrambled encoding** | 2 | **No safe find/replace.** "Log Bridge Road" extracts as `-PH#SJEHF3PBE`. The glyphs a replacement needs may not exist in the subset at all. |
| **Baked into a full-page raster** | 3 | **No.** All three MBW files are a single scanned image, zero text objects. The EH house-mark and wordmark are pixels. (MailboxWorks already swapped the *footer* to their own logo and left the EH logo at the top.) |
| **Text converted to outlines** | 5 | **No.** Zero text objects *and* zero images — the lettering is vector paths. |

Counts overlap; a document can be in more than one row.

### It is not a name swap

The documents carry a full vendor identity block — company, street address, phone, fax and
`www.europeanhome.com` — and across the 13 files there are **three different addresses**:

- Log Bridge Road, Bld/Unit, Middleton, MA
- 376 Washington St., Suite 203, Malden, MA 02148
- 501-B Main Street, Saugus, MA 01906

Changing only the name ships a stale address and a dead domain. Any path that has us setting type needs the
current address, phone, fax, website **and** the new logo artwork, in writing, before a single file is touched.

### Two data problems found in passing

- `european-home-mail-slot-installation-instructions.pdf` (MBW) is credited to **"The Mailbox Works"** in our
  review sheet. That is a misclassification — the classifier read the retailer footer off the scan. It is a
  European Home document.
- The live BM brand page's Download Library entries **163** ("European Home Mailbox Installation
  Instructions") and **168** ("Mail Slot- European Home Dimensions") still show the old name in their titles
  and thumbnail alt text. That is **site metadata, not the PDFs** — Ramez's side, not this job.

### Scope is wider than our data suggests

The live BM Download Library has **13 entries** (MageWorx ids 161–172 and 361). Only **2** name European Home
in their title; the rest are named by product — Torgen, View Point, Mailbox Stand Info Flyer. So filtering on
our `manufacturer` column **under-counts**, and filtering on the title under-counts worse. Get an
id-to-filename mapping from Catalog before agreeing a file count. Kunchana offered to pull the affected file
list (comment `90180250616508`) — take him up on it.

MBW's Sustainable Hearth group page carries no PDFs; MBW documents sit on product pages.

## The seven questions that must be answered first

Posted as comment `90180251134504` — full text at
[deliveries/SH_scoping_comment.md](deliveries/SH_scoping_comment.md). Do not start until these are answered:

1. **Which file list** — the whole 13-entry library, only those printing the old name, or wider?
2. **Is MBW in?** Three delivered MBW documents are affected; the assignment only mentions BM.
3. **Our branding only, or the vendor's artwork too?** The fork that decides everything.
4. **If the artwork: where does the replacement come from?**
5. **If we set the type: the exact wording** — current address, phone, fax, website, logo art.
6. **The 6 leave-as-is technical drawings** — in scope or untouched?
7. **Filenames and delivery destination** — MBW pattern (keep names, renames and 301s at deploy with Ramez),
   or rename in our output?

## Decision tree

### If the answer to (3) is "our branding only" — do this

Genuinely small. One kit edit and a targeted re-run.

1. Edit `brandkits/budget-mailboxes-brand.json` and `brandkits/mailboxworks-brand.json`. In
   `manufacturer_aliases`, the canonical value is currently `European Home` with an alias `EuropeanHome`.
   Point both — and any new spelling — at `Sustainable Hearth`.
2. **Copy the kit back beside the artwork as the literal filename `brand.json`.** Any other filename is
   ignored and the kit silently falls back to defaults naming *Budget Mailboxes*: zero aliases, no tagline, no
   disclaimer. It looks like it worked. This has bitten before (2026-08-18).
3. Fix the MBW misattribution row in the review sheet at the same time.
4. Re-run **Apply only** on the affected files — a re-Analyze is not needed, the sheet is unaffected by
   Apply-time settings. Delete the existing outputs first: Apply skips files that already exist.
5. QA with `verify/deliver_qa.py`, then confirm the attribution line with `verify/verify_attrib.py`.

### If the answer is "the vendor's artwork too" — push back once, then cost it

**The recommendation on record: ask Sustainable Hearth for re-issued source PDFs.** They are mid-rebrand and
will almost certainly have re-cut this collateral themselves. That turns the job back into a normal
ingest-and-rebrand run — which the app does well — instead of rebuilding another company's artwork and owning
the result.

If WebLife declines to ask, say plainly that the remaining path is **manual page reconstruction** for at least
10 of the 13 files, and get it costed **before** starting. It is not a pipeline feature and should not be
estimated as one. Do not attempt a text-layer find/replace on the subset-font files; it will either fail or
produce missing glyphs, and the failure is silent.

## Reproducing the file list

The review sheets for both deliveries are in [deliveries/](deliveries/). To regenerate the affected-file list
from scratch, run this from the repo root with the project venv:

```python
import openpyxl, re, pathlib

SHEETS = [
    ("BATCH4", "Batch 4___unique-to-rebrand_rebrand_plan.xlsx"),
    ("MBW", "MBW rebrand work___unique-to-rebrand_rebrand_plan.xlsx"),
]
for tag, name in SHEETS:
    path = pathlib.Path("docs/handoff/deliveries") / name
    ws = openpyxl.load_workbook(path, read_only=True).active
    rows = ws.iter_rows(values_only=True)
    hdr = list(next(rows))
    for r in rows:
        blob = " ".join(str(x) for x in r if x)
        if re.search(r"european\s*home", blob, re.I):
            row = dict(zip(hdr, r))
            print(tag, row["action"], row["manufacturer"], row["file"])
```

To re-check what is *inside* a PDF rather than what the sheet claims, `pdftotext` ships with the repo at
`poppler/Library/bin/pdftotext.exe`. A file that returns **zero characters** is either a scan or outlined
vector — check the page's image count with pypdf to tell which, because the fix differs and neither is a text
edit.

## What not to do

- **Do not** treat this as a normal rebrand batch because the ClickUp task says so.
- **Do not** find/replace text in the two subset-font installation sheets.
- **Do not** change any filename without telling Ramez Sedra first — the delivered corpus is keyed on
  filename, and on Batch 4 renaming 874 files would have meant 874 redirects protecting 11,638 product-page
  references.
- **Do not** change the vendor's printed name without also changing the address, phone and domain.
