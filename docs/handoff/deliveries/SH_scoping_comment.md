**Sustainable Hearth PDFs — what the TA request needs to specify**

@Kunchana Godahewa — following my note above asking you to put this through the TA Task Request form. That is not process for its own sake: this one is **not the same shape as the BM and MBW batches**, and if it comes in described as "same rebrand treatment" we will build the wrong thing. Below is what I would need the request to answer, plus what I have already pulled so you do not have to.

**Why it is different.** Those batches wrap an *untouched* original page in *our* branding — cover, header/footer strips, watermark, and a `Manufactured by X | Sold by Y` line. We never edit the vendor's own artwork inside the page. This request is about the vendor's own branding printed inside the page. Different operation, different risk, different cost.

**You offered to pull the affected file list — already done, no need.** Across the two delivered sets, **13 documents carry European Home** (10 in BM Batch 4, 3 in MBW). Where the old name actually lives:

| Where it appears | Docs | Can we change it? |
|---|---:|---|
| Our own attribution line (`Manufactured by European Home`) | 6 | **Yes** — brand-kit edit and a re-run, quick |
| Page text in a normal font | 3 | Yes, but see item 4 — it is an address block, not just a name |
| Page text in a **subset font with a scrambled encoding** | 2 | **No safe find/replace.** "Log Bridge Road" extracts as `-PH#SJEHF3PBE`; the glyphs we would need may not exist in the subset |
| **Baked into a full-page raster** (the 3 MBW scans) | 3 | **No.** The EH house-mark and wordmark are pixels. MailboxWorks already swapped the footer to their own logo and left the EH logo at the top |
| **Text converted to outlines** (technical drawings) | 5 | **No.** Zero text objects, zero images |

*(Counts overlap — a document can appear in more than one row.)* I can attach the per-file list to the request whenever you want it.

---

**What the request needs to pin down**

**1. The file list.** The live BM brand-page Download Library has **13 entries** (MageWorx ids 161–172 and 361). Only **2** name European Home in their title; the rest are named by product — Torgen, View Point, Mailbox Stand Info Flyer. So filtering on our manufacturer field under-counts, and filtering on the title under-counts worse. The request should name the set explicitly: every document in that library, only the ones printing the old name, or something wider such as product-page PDFs. An id-to-filename mapping would settle it outright.

**2. Whether MBW is in.** The assignment says BM, but MBW carries three European Home documents too, already delivered. In, later, or out.

**3. What "updated" means per document — our line, or their artwork.** Two very different deliverables:

- **(a) Our branding only.** Every line *we* print reads Sustainable Hearth. Cheap and safe, quick turnaround, and the vendor's own logo and address stay as they are inside the page.
- **(b) The vendor's printed identity too.** Their logo, name, address and website inside the document change as well. Per the table, that is not a text edit for 10 of the 13 — it means re-creating pages.

The request needs to say which one the client expects to see, because (a) and (b) are not the same job.

**4. If (b) — where the replacement artwork comes from.** My recommendation: **ask Sustainable Hearth for re-issued source PDFs.** They are mid-rebrand and will almost certainly have re-cut this collateral themselves. That turns this back into a normal ingest-and-rebrand run instead of us rebuilding another company's artwork and owning the result. If we are not going to ask them, the request should say so explicitly, because the alternative is manual page reconstruction and I want that costed before it starts rather than after.

**5. If we are setting the type ourselves — the exact wording.** It is not a name swap. The documents carry a full identity block: company, street address, phone, fax, and `www.europeanhome.com`. Across the 13 there are **three different addresses**:

- Log Bridge Road, Bld/Unit, Middleton, MA
- 376 Washington St., Suite 203, Malden, MA 02148
- 501-B Main Street, Saugus, MA 01906

Name-only ships a stale address and a dead domain. The request needs the current address, phone, fax and website, plus the new logo art.

**6. The 6 technical drawings we deliberately leave alone.** `Centauri-Dimensions`, `Galaxy-Large-Mail-Slot-Dimensions`, `Galaxy-Small-Mail-Slot-Dimensions`, `mailsloteurohome`, `Curb-Appeal-and-Holder-Dimensions`, `Vega-Dimensions`. These were classified leave-as-is and shipped byte-identical, as agreed. Four have their text converted to outlines, so there is nothing to edit even if we wanted to. In scope because they carry the old name, or unchanged.

**7. Filenames and where the delivery lands.** Four files carry the old name in the filename — `EH-Mail-Slot-Dimensions.pdf`, `European-Home-mailbox-install-inst.pdf`, and the three `european-home-*` on MBW. On MBW we kept original filenames and let renames and 301s happen at site deploy with @Ramez Sedra; the request should say whether that holds here or we rename in our output. It should also say where finished files go — back into the Drive source folder as with MBW, or straight to whoever swaps the Download Library entries.

---

**Two things I will fix regardless, flagging so they are not a surprise:**

- `european-home-mail-slot-installation-instructions.pdf` (MBW) is credited to **"The Mailbox Works"** in our data. That is a misclassification — the classifier read the retailer footer on the scan. It is a European Home document and I will correct it.
- Download Library entries **163** ("European Home Mailbox Installation Instructions") and **168** ("Mail Slot- European Home Dimensions") still show the old name in their titles and thumbnail alt text on the live BM page. That is site metadata rather than the PDFs, so I read it as @Ramez Sedra's side rather than this job — confirming so it does not fall between us.

For visibility on @Laleesha Wijeratne's question above: the PDF leg has not started and is waiting on this scope, so it is not ready to close.

— Sent via Friday OS
