"""Layer 3 — does the branded page LOOK like the source?

Renders each source page and the corresponding content band of the branded page
and compares them as images. Text extraction cannot see cropping, scaling or
displacement; this can. Uses a watermark-free copy of the brand kit so the
comparison is clean.
"""
import random, shutil, sys, logging
from pathlib import Path
import batch_paths                      # also puts the repo on sys.path
logging.getLogger("pypdf").setLevel(logging.CRITICAL)
from PIL import Image, ImageChops, ImageFilter
from pypdf import PdfReader
from docrefine import reviews
from docrefine.rebrand import BrandKit, rebrand_pdf, page_size, _strip_height, _orientation
from docrefine.processing import POPPLER_BIN, convert_from_path

SP = Path(sys.argv[1]); LIMIT = int(sys.argv[2]) if len(sys.argv) > 2 else 60
SRC = batch_paths.SRC
KIT = batch_paths.require_kit()
DPI = 100
# Calibrated on real files: known-bad pages differ by 8.5-9.0%, correct ones by
# 0.5-2.0%. Anything in between is a judgement call, so the threshold is an
# argument and the report also lists what sits in the grey band just under it.
THRESHOLD = float(sys.argv[3]) if len(sys.argv) > 3 else 0.04
GREY_FROM = THRESHOLD * 0.5

qa_kit = SP / "kit_nowatermark"
if qa_kit.exists(): shutil.rmtree(qa_kit)
shutil.copytree(KIT, qa_kit)
for wm in qa_kit.rglob("*atermark*"): wm.unlink()
kit = BrandKit(qa_kit)

rows = [r for r in reviews.read_plan(batch_paths.plan())
        if (r.get("action") or "").lower() == "rebrand"]
names = [r["file"] for r in rows if (SRC / r["file"]).exists()]

FOCUS = [n for n in names if any(k in n for k in (
    "cutsheet", "cut-sheet", "metropolis", "Anchor-Bolt", "keystone_post_cuff",
    "PackageMaster", "4c11d-10", "4c16s-bin", "lock-change", "Imperial-Street-Sign"))]
random.seed(7)
sample = list(dict.fromkeys(FOCUS[:40] + random.sample(names, min(LIMIT, len(names)))))[:LIMIT]

work = SP / "pixel_work"; shutil.rmtree(work, ignore_errors=True); work.mkdir(parents=True)
print(f"comparing {len(sample)} documents at {DPI}dpi (watermark removed for a clean diff)\n")

bad, grey, checked, errs = [], [], 0, []
for name in sample:
    f = SRC / name
    try:
        out = work / "b.pdf"
        rebrand_pdf(f, out, kit, "Specification Sheet")
        s_rd, d_rd = PdfReader(str(f)), PdfReader(str(out))
        readers = kit.live_readers(_orientation(*page_size(s_rd.pages[0])))
    except Exception as e:
        errs.append(f"{name}: {type(e).__name__}: {e}"); continue

    for i in range(len(s_rd.pages)):
        pw, ph = page_size(s_rd.pages[i])
        po = _orientation(pw, ph)
        rs = kit.live_readers(po) if kit.has(po) else readers
        hh = _strip_height(rs["header"], pw) if rs.get("header") else 0.0
        fh = _strip_height(rs["footer"], pw) if rs.get("footer") else 0.0
        try:
            si = convert_from_path(str(f), dpi=DPI, first_page=i+1, last_page=i+1, poppler_path=POPPLER_BIN)[0]
            di = convert_from_path(str(out), dpi=DPI, first_page=i+2, last_page=i+2, poppler_path=POPPLER_BIN)[0]
        except Exception as e:
            errs.append(f"{name} p{i+1}: render: {e}"); continue
        scale = di.height / (ph + hh + fh)
        top = int(round(hh * scale))
        band = di.crop((0, top, di.width, top + int(round(ph * scale))))
        if band.size != si.size:
            band = band.resize(si.size, Image.LANCZOS)
        # Blur first: cropping/scaling/displacement are low-frequency changes,
        # antialiasing on glyph edges is high-frequency. Without this the metric
        # flags every text-heavy page and is useless.
        a = si.convert("L").filter(ImageFilter.GaussianBlur(3))
        b = band.convert("L").filter(ImageFilter.GaussianBlur(3))
        px = ImageChops.difference(a, b).point(lambda v: 255 if v > 40 else 0)
        frac = sum(px.getdata()) / 255 / (px.width * px.height)
        checked += 1
        if frac > THRESHOLD:
            bad.append((frac, name, i + 1, f"{si.size}->{band.size}"))
        elif frac > GREY_FROM:
            grey.append((frac, name, i + 1, f"{si.size}->{band.size}"))

print(f"pages compared: {checked}   documents: {len(sample)}   render/rebrand errors: {len(errs)}")
print(f"threshold: {THRESHOLD*100:.1f}%   (grey band reported from {GREY_FROM*100:.1f}%)")
for e in errs[:5]: print("   ERR " + e)
print()
if bad:
    print(f"PAGES THAT DO NOT MATCH THEIR SOURCE ({len(bad)}):")
    for frac, name, pg, sz in sorted(bad, reverse=True)[:20]:
        print(f"   {frac*100:5.1f}% differing   {name} p{pg}   {sz}")
else:
    print("every page matches its source — no cropping, scaling or displacement detected")
if grey:
    print(f"\nunder the threshold but worth an eye ({len(grey)}):")
    for frac, name, pg, sz in sorted(grey, reverse=True)[:10]:
        print(f"   {frac*100:5.1f}% differing   {name} p{pg}")
