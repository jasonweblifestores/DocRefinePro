**MBW analysis complete — eight things for your call**

970 unique documents, deduplicated from 9,087 files (~9× duplication — one document appears 420 times). **474 to rebrand, 496 to leave** as technical drawings; 1,416 pages to brand. The delivery is **970 files** — the 474 branded plus the 496 left byte-for-byte identical — plus the manifests.

**Why 418 of the 474 carry an attribution line and not all of them.** The other 56 are cases where a line would be wrong, not gaps we've left. Full list attached as `MBW_no-attribution-line.csv` (every file, the value the document actually carried, the reason, and a SKU folder to find it in) so you can check any of them rather than take my word for it:

| Count | Reason | Value in the document | Examples |
|---:|---|---|---|
| 35 | reads as the seller, not the maker | `The Mailbox Works` | `6-Post-Mounting-Instructions_Revised.pdf` (SKU IMZP6), `amco-colonialpedestal-assembly1.pdf` (SKU VM-280), `amco-victorian-pedestal-assembly.pdf` (SKU VM-2) |
| 14 | no manufacturer named in the document | *(blank)* | see item 8 below |
| 4 | names MailboxWorks itself | `MailboxWorks`, `The MailboxWorks` | `janzer-post-installation1.pdf` (SKU JB-JP), `janzer-doublepost-installation1.pdf` (SKU JDP) |
| 3 | a website, not a company | `FlorenceMailboxes.com`, `Florencemailboxes.com`, `postandporch.com` | `4B-Replacement-Mailbox-Information-1-1.pdf` (SKU 140052A), `PP-Large-Plaque.pdf` (SKU PP-Large-Plaque) |

We leave the line off rather than print "Manufactured by MailboxWorks | Sold by MailboxWorks". Tell me if you'd rather those carried something.

---

**1. Confirm your Batch 4 spellings carry over**

The raw data has **59 distinct manufacturer spellings**. Applying the canonical names you approved on Budget Mailboxes Batch 4 brings it to **36, with zero case-or-punctuation variants left**, and credits each company identically across both brands:

| Canonical name | Absorbs | Docs |
|---|---|---:|
| Florence Corporation | Florence Mailboxes, FlorenceMailboxes, Florence Manufacturing, Florence, Auth Florence Manufacturing Company, Auth-Florence Manufacturing Company | 223 |
| Architectural Mailboxes, LLC | Architectural Mailboxes | 56 |
| Whitehall Products, LLC | Whitehall, Whitehall Products, Whitehall Products LLC, White Hall Products LLC, WHITEHALL PRODUCTS | 29 |
| Imperial Mailbox Systems | IMPERIAL MAILBOX SYSTEMS | 9 |
| Epoch Design, LLC | Epoch Design, Epoch Design LLC | 7 |
| Salsbury Industries | Salsbury | 7 |
| IMI Mailbox Systems | IMI MAILBOX SYSTEMS | 6 |
| bobi | Bobi | 4 |
| Gaines Manufacturing, Inc. | Gaines Manufacturing | 4 |
| QualArc | Qualarc | 2 |
| Special Lite Products Company, Inc. | Special Lite Products, Special Lite Products LLC | 2 |
| TedStuff | TedStuﬀ, Tedstuﬀ | 2 |

*On TedStuff — the MBW source spells it with a typographic `ﬀ` ligature our stamp font has no glyph for; it would print as "TedStu" with a gap. Your Batch 4 spelling with plain `ff` fixes it, so nothing needed from you. Mentioning it because it wasn't cosmetic.*

**2. `IMI` (11 documents) — fold into Imperial Mailbox Systems, or keep separate?**

Batch 4 kept `IMI` separate, but the MBW evidence points one way: all 11 sit in SKU folders named `Imperial-Barcelona-System`, `Imperial-Series`, `IMP-QD`, `IMM119`, and the documents themselves read `IMPERIAL MAILBOX SYSTEMS`. I'd fold them.

**3. Three small spellings**

- **`AMCO` / `Amco`** — one document each (`amco-colonialpedestal-installation1.pdf`, SKU `VM-280`; `amco-victorian-pedestal-installation.pdf`, SKU `VM-2`). Which? Note the *assembly* sheets for those same SKUs credit "The Mailbox Works", so they carry no line either way.
- **`Mail Boss` (3 docs)** — Batch 4 used *The Mail Boss*. Match that, keep `Mail Boss`, or fold into *Epoch Design, LLC* (same company)?
- **`BOBI ROUND` (1 doc, `bobi-round-stand.pdf`)** — Batch 4 kept it separate from `bobi`. Still separate?

**4. 15 documents credit a product line or the postal service — needs a ruling**

These would print e.g. *"Manufactured by Quad System"* or *"Manufactured by USPS"*. A wrong credit is worse than none, so I haven't guessed. Where the SKU folder makes the maker obvious I've suggested it:

| Current value | Document | SKU | Suggest |
|---|---|---|---|
| USPS | `Imperial-Series-Assembly-Instructions_Revised.pdf` | IMPP1–3 | Imperial Mailbox Systems |
| USPS | `C1-and-D1-Assembly-Instructions_Revised-1.pdf` | IMZP7 | Imperial Mailbox Systems |
| US Postal Service | `C1-and-D1-Assembly-Instructions_Revised.pdf` | IMZP8 | Imperial Mailbox Systems |
| Barcelona System | `Barcelona-Assembly-Instructions_Revised.pdf` | IMZP2 | Imperial Mailbox Systems |
| Barcelona System | `Barcelona-Assembly-Instructions_Revised-2.pdf` | IMZP3 | Imperial Mailbox Systems |
| Quad System | `Quad-Imperial-Series-Assembly-Instructions_Revised.pdf` | IMP-QD | Imperial Mailbox Systems |
| Twin System | `Twin-Imperial-Series-Assembly-Instructions_Revised.pdf` | IMP-TWP | Imperial Mailbox Systems |
| Norris | `Brass-Knob-and-Flag-Assembly-Instructions_Revised.pdf` | Flag-Housing, Large-Flag | Imperial Mailbox Systems — the same document under its other filename is already credited IMI, and `imperial-pdf-4.pdf` shares its SKU folder |
| Balmoral | `Balmoral-Installation-Instructions.pdf` | WH-BAL-FLW-COONLEY-PKG1 +8 | Whitehall Products, LLC (`WH-` prefix) |
| Triple Mount | `install_keystone_multimounts.pdf` | KD3-KS, KD4-KS +5 | ? |
| Keystone | `Keystone-Deluxe-Post.pdf` | BEAM, EndCap +13 | ? |
| Town And Country | `Bracket-Installation.pdf`, `town-country-pro-series-bracket-installation_1_.pdf` | SCB-1005 +2 | ? |
| RetroBox & Uptown | `RetroBox-Install-Instructions.pdf`, `UptownBox-Installation-Instructions.pdf` | QS-UP, RBAW +7 | ? |
| Outdoor Parcel Locker | `inst-parcellockers-pedestal.pdf` | 3302 | ? |
| Streetscape | `streetscape-mailbox-installation.pdf` | SI-155L | ? is Streetscape the maker or the line? |

For each: the real manufacturer, or leave blank?

**5. 14 printable templates — one treatment for all of them (see attached render)**

14 documents describe themselves as *"Product Template and installation aid"*, printed to scale — 13 `PP-*` files plus `Standing_Tall_Planter.pdf`. As it stands the plan would **brand 3 and leave 11**: the same class of document treated two ways in one delivery.

To be precise about what branding does: it does **not** rescale the artwork — the page width is identical to three decimal places, so printed at actual size the template still measures true. What it does do, in the attached render:

- the template becomes **page 2 of 3**, behind a cover — so "print the template" becomes "print page 2", and printing the whole file gives three 24-inch-wide sheets
- the watermark sits **across both mounting-screw callouts** (*"9 7/16″ from center of one mounting screw to the other"*) — legible, but over the only numbers that matter

Your call and I'll apply it to all 14. If you want them branded, we can suppress the watermark on just this class.

**6. Which filename survives a collapse — recommendation**

35 documents exist under more than one filename (one under 43), collapsing to a single branded file each. The survivor is currently whichever copy the scan reached first — arbitrary. Choosing **the name the most SKU folders already reference** is deterministic and cuts the redirects Ramez has to set up:

| Rule | Rows needing a redirect |
|---|---:|
| Arbitrary (current) | 1,185 |
| **Most-referenced name** | **913** |
| Alphabetically first | 2,175 |

**272 fewer URL changes**, moving 9 of the 35. The biggest:

| Now | Would become | SKU refs |
|---|---|---|
| `florence-powdercoat-referencechart-brochure-032021_2-merged-compressed-1.pdf` | `florence-powdercoat-referencechart-brochure-color-options.pdf` | 28 → 231 |
| `Imperial-Powder-Coating_Revised_4th-Sep.pdf` | `Imperial-Powder-Coating_Revised_4th-Sep-1.pdf` | 6 → 31 |
| `Cluster-Mailbox-Materials.pdf` | `cbu-florence-material-spec_1-1-1.pdf` | 2 → 19 |
| `4c-15h-family-cut-sheet-1.pdf` | `4c-15h-family-cut-sheet.pdf` | 6 → 18 |

The first alone is 203 of the 272. It's chosen purely on reference count, so a couple of the new names are less tidy than the old — `cbu-florence-material-spec_1-1-1.pdf` over `Cluster-Mailbox-Materials.pdf` — but far better linked.

**7. Confirm the manifest rewrite**

Of 9,087 manifest rows, **1,185 across 559 of the 1,410 sheets** name a filename that collapses away. Every one resolves to exactly one surviving byte-identical file — **zero ambiguous cases** — so it's mechanical. I'd deliver:

- `manifests/<SKU>/_manifest.csv` — pointing at the filenames actually delivered
- `filename-aliases.csv` — 152 rows, original → delivered, as the audit trail and Ramez's redirect map

The source download isn't modified. Flagging it because it changes what's inside the manifests you receive.

**8. 14 documents name no manufacturer — 13 we can fill in, if you confirm**

These carry no maker anywhere in the text, but the filename and the SKU prefix agree on who made them. Confirming these lifts attribution coverage from **418 to 431 of 474**. Nine use canonical names you already approved on Batch 4:

| Document | SKU | Proposed | Basis |
|---|---|---|---|
| `dvault-frontaccess-installation3.pdf` | DVCS0023 | **dVault** | filename + `DV` SKU prefix — your Batch 4 name |
| `gaines-classic-installation1.pdf` | CL | **Gaines Manufacturing, Inc.** | filename — your Batch 4 name |
| `salsbury-vertical-mailbox-installation-instructions.pdf` | 3503 +4 | **Salsbury Industries** | filename — your Batch 4 name |
| `grande-s-round.pdf` | BobiGrandeS-Config +35 | **bobi** | SKU — your Batch 4 name |
| `mailbox-deluxe-post-assembly-instructions.pdf` | WH-DXE-FLW-COONLEY-PKG1 +1 | **Whitehall Products, LLC** | `WH-` SKU prefix |
| `mailbox-standard-post-assembly-instructions.pdf` | WH-160xx +4 | **Whitehall Products, LLC** | `WH-` SKU prefix |
| `whitehall-quad-post-installation-instructions.pdf` | WH-WP-Quad | **Whitehall Products, LLC** | filename + `WH-` SKU prefix |
| `Newspaper-Holder-6-Assembly-Instructions_Revised-2.pdf` | IMM119 | **Imperial Mailbox Systems** | `IMM` SKU prefix — `imperial-pdf-3.pdf` sits in the same SKU folder |
| `Newspaper-Holder-6-Assembly-Instructions_Revised-4.pdf` | IMM631-IMM631EST | **Imperial Mailbox Systems** | `IMM` SKU prefix |
| `Newspaper-Holder-6-Assembly-Instructions_Revised-5.pdf` | IMM188 | **Imperial Mailbox Systems** | `IMM` SKU prefix — `imperial-pdf-6.pdf` sits in the same SKU folder |
| `streetscape-maintenance.pdf` | SI-155L | **Streetscape** | same SKU as `streetscape-mailbox-installation.pdf` — depends on your answer to item 4 |
| `ecco-7-installation1.pdf` | E7 | **Ecco** | filename + SKU; new name, not on Batch 4 |
| `E4-Installation.pdf` | E4 | **Ecco** | SKU matches the `E7` pattern above — least certain of the set |
| `PP-Curb-Planter.pdf` | PP-Curb-Planter | **Post & Porch** | SKU; also one of the 14 templates in item 5 |

These are inferences from filename and SKU, not from anything the document says — so I'd rather you confirm than have us print a maker the document doesn't claim.

---

**One FYI, no action needed:** `frank_lloyd_wright_collection.pdf` is in this batch too at 58.2 MB. I'll apply the treatment you approved on Batch 4 — resample its oversize images to a 200 DPI floor so it lands under 50 MB, content unchanged. It's the only file over 40 MB.
