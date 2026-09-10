"""Structural scan of every file we intend to BRAND, looking for the class of
assumption that text-based checks cannot see."""
import sys, logging
from collections import Counter, defaultdict
from pathlib import Path
import batch_paths                      # also puts the repo on sys.path
logging.getLogger("pypdf").setLevel(logging.CRITICAL)
from pypdf import PdfReader
from docrefine import reviews
from docrefine.rebrand import page_rotation

SRC = batch_paths.SRC
PLAN = batch_paths.plan()
rows = [r for r in reviews.read_plan(PLAN) if (r.get("action") or "").lower() == "rebrand"]

findings = defaultdict(list)
stats = Counter()

def num(x):
    try: return float(x)
    except Exception: return 0.0

for r in rows:
    f = SRC / r["file"]
    if not f.exists():
        findings["missing from disk"].append(r["file"]); continue
    try:
        rd = PdfReader(str(f))
    except Exception as e:
        findings["UNREADABLE"].append(f"{r['file']}: {e}"); continue
    stats["files"] += 1
    if getattr(rd, "is_encrypted", False):
        findings["encrypted"].append(r["file"])

    # One unreadable file must not end the scan — an encrypted PDF used to raise
    # here and take the whole report with it, hiding every finding after it.
    try:
        _ = len(rd.pages)
    except Exception as e:
        findings["UNREADABLE (pages)"].append(f"{r['file']}: {type(e).__name__}: {e}")
        continue

    sizes, rots = set(), set()
    for i, p in enumerate(rd.pages):
        stats["pages"] += 1
        mb = p.mediabox
        mw, mh = abs(num(mb.width)), abs(num(mb.height))
        sizes.add((round(mw), round(mh))); rots.add(page_rotation(p))

        # origin not at 0,0 — we translate content assuming it is
        if abs(num(mb.left)) > 1 or abs(num(mb.bottom)) > 1:
            findings["mediabox origin not at (0,0)"].append(
                f"{r['file']} p{i+1} [{num(mb.left):.0f},{num(mb.bottom):.0f},{num(mb.right):.0f},{num(mb.top):.0f}]")
        # CropBox smaller than MediaBox — readers show the CropBox, we brand the MediaBox
        try:
            cb = p.cropbox
            cw, ch = abs(num(cb.width)), abs(num(cb.height))
            if cw and ch and (mw - cw > 1 or mh - ch > 1):
                findings["CropBox smaller than MediaBox"].append(
                    f"{r['file']} p{i+1} media {mw:.0f}x{mh:.0f} vs crop {cw:.0f}x{ch:.0f}")
        except Exception:
            pass
        if "/UserUnit" in p:
            findings["/UserUnit set (page scaled)"].append(f"{r['file']} p{i+1} = {p.get('/UserUnit')}")
        if mw < 72 or mh < 72:
            findings["tiny page (<1in)"].append(f"{r['file']} p{i+1} {mw:.0f}x{mh:.0f}")
        if mw > 200 * 72 or mh > 200 * 72:
            findings["enormous page (>200in)"].append(f"{r['file']} p{i+1} {mw:.0f}x{mh:.0f}")
        if p.get("/Annots"):
            stats["pages with annotations"] += 1
            findings["_annots"].append(r["file"])
        if page_rotation(p) not in (0, 90, 180, 270):
            findings["odd rotation"].append(f"{r['file']} p{i+1} {page_rotation(p)}")
        if page_rotation(p) == 180:
            findings["/Rotate 180"].append(f"{r['file']} p{i+1}")

    if len(sizes) > 1:
        findings["mixed page SIZES in one document"].append(f"{r['file']} {sorted(sizes)}")
    if len(rots) > 1:
        findings["mixed ROTATION in one document"].append(f"{r['file']} {sorted(rots)}")
    try:
        if rd.trailer.get("/Root", {}).get("/AcroForm"):
            findings["has form fields (AcroForm)"].append(r["file"])
    except Exception:
        pass
    try:
        if rd.outline:
            findings["has bookmarks/outline"].append(r["file"])
    except Exception:
        pass

print(f"scanned {stats['files']} files / {stats['pages']} pages\n")
annots = sorted(set(findings.pop("_annots", [])))
if annots:
    findings["pages carrying annotations (links etc.)"] = annots
for k in sorted(findings, key=lambda k: -len(findings[k])):
    v = findings[k]
    print(f"### {k}: {len(v)}")
    for x in v[:5]: print("      " + str(x))
    if len(v) > 5: print(f"      … and {len(v)-5} more")
    print()
if not findings: print("no structural anomalies found")
