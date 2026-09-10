import json, shutil, sys
from datetime import date
from pathlib import Path
sys.path.insert(0, ".")
from docrefine.rebrand import BrandKit, rebrand_pdf
from docrefine.worker import Worker
from docrefine import stamps as S

SP = Path(sys.argv[1])
OUT = SP / "rebranding-samples"
if OUT.exists(): shutil.rmtree(OUT)
OUT.mkdir(parents=True)

SRC = Path(r"C:\Users\WORK\Documents\Batch 4\_unique-to-rebrand\inst-standardopenaccesslockers-70000-series.pdf")
KIT_SRC = Path(r"C:\Users\WORK\Documents\Batch 4\Template")
MANUF = "Salsbury Industries"
TITLE = "Installation Manual"

DISCLAIMER = ("[PLACEHOLDER — the approved disclaimer wording goes here. This block is set at a "
              "realistic length so you can judge the space it occupies; it is not approved copy.]")

kit_dir = SP / "kit_samples"
if kit_dir.exists(): shutil.rmtree(kit_dir)
shutil.copytree(KIT_SRC, kit_dir)
(kit_dir / "brand.json").write_text(json.dumps({
    "name": "Budget Mailboxes", "slug": "budget-mailboxes",
    "tagline": "Trusted by the Nation",
    "disclaimer": DISCLAIMER,
}, indent=2), encoding="utf-8")
kit = BrandKit(kit_dir)
TODAY = date(2026, 8, 5)

def stamps(**flags):
    return Worker._stamps_for(kit, MANUF, {k: True for k in flags}, TODAY)

variants = [
    ("A - current approved output (no stamps).pdf", None, None),
    ("B - attribution on the cover.pdf", kit.subtitle_for(MANUF), None),
    ("C - attribution in the page footer (SOP).pdf", None, stamps(footer_attribution=1)),
    ("D - full SOP set (all four stamps).pdf", None,
     stamps(footer_attribution=1, stamp_tagline=1, stamp_version=1, stamp_disclaimer=1)),
]
made = []
for name, subtitle, st in variants:
    dst = OUT / name
    info = rebrand_pdf(SRC, dst, kit, TITLE, subtitle=subtitle, stamps=st)
    made.append((name, dst, info))
    print(f"{name:48s} {info['output_pages']}pp  {dst.stat().st_size/1e6:.2f}MB")

# ---- comparison sheets ----
# Two sheets, not one: option B puts the line on the COVER, so in a content-page
# crop it is identical to A. Showing them together would imply B changes nothing.
from PIL import Image, ImageDraw, ImageFont
from docrefine.processing import POPPLER_BIN, convert_from_path

font_path = "docrefine/assets/fonts/Poppins-Bold.ttf"
label_f = ImageFont.truetype(font_path, 20)
by_name = {n.split(" - ")[1].replace(".pdf", ""): d for n, d, _ in made}

def render(dst, page_no, top_frac):
    pg = convert_from_path(str(dst), dpi=120, first_page=page_no, last_page=page_no,
                           poppler_path=POPPLER_BIN)[0]
    return pg.crop((0, int(pg.height * top_frac), pg.width, pg.height))

def sheet(path, panels):
    crops = [(lbl, render(by_name[key], pg, top)) for lbl, key, pg, top in panels]
    W = max(c.width for _, c in crops); LH = 34
    h = sum(c.height + LH + 14 for _, c in crops) + 14
    im = Image.new("RGB", (W, h), "white"); d = ImageDraw.Draw(im); y = 14
    for lbl, band in crops:
        d.rectangle([0, y, W, y + LH], fill="#1F2A5A")
        d.text((14, y + 7), lbl.upper(), font=label_f, fill="white")
        y += LH
        im.paste(band, (0, y))
        d.rectangle([0, y, W - 1, y + band.height - 1], outline="#cccccc")
        y += band.height + 14
    im.save(path)
    print("wrote", path.name, im.size)

sheet(OUT / "comparison 1 - bottom of a content page.png", [
    ("A - current approved output", "current approved output (no stamps)", 2, 0.80),
    ("C - attribution in the page footer (SOP)", "attribution in the page footer (SOP)", 2, 0.80),
    ("D - full SOP set", "full SOP set (all four stamps)", 2, 0.80),
])
sheet(OUT / "comparison 2 - the cover.png", [
    ("A - current approved cover (title only)", "current approved output (no stamps)", 1, 0.0),
    ("B - attribution under the cover title", "attribution on the cover", 1, 0.0),
])

(OUT / "READ ME FIRST.txt").write_text(
    "PDF rebranding — sample set\n"
    "===========================\n\n"
    "All four PDFs are the SAME real document from Batch 4\n"
    "(inst-standardopenaccesslockers-70000-series.pdf, manufacturer: Salsbury Industries),\n"
    "produced by the tool. Only the branding options differ.\n\n"
    "  A  Current approved output. This is what Batch 1, Batch 2 and the\n"
    "     current Batch 4 set look like today. No attribution, no stamps.\n\n"
    "  B  Attribution under the cover title. This is the option that exists\n"
    "     today but is switched off.\n\n"
    "  C  Attribution as a small footer on every page - the placement the\n"
    "     MailboxWorks brief specifies. Nothing else added.\n\n"
    "  D  The full set the Batch 4 brief asks for: footer attribution,\n"
    "     tagline, version / last-updated, and disclaimer.\n\n"
    "  comparison 1 - bottom of a content page.png\n"
    "     A, C and D side by side, cropped to the bottom of a content page, so\n"
    "     the difference is visible without opening each file.\n\n"
    "  comparison 2 - the cover.png\n"
    "     A and B side by side. Option B changes only the COVER, which is why it\n"
    "     is not in the first sheet - on a content page it is identical to A.\n\n"
    "IMPORTANT: the disclaimer text in D is PLACEHOLDER. It is set at a realistic\n"
    "length so you can judge the space it takes, but it is not approved copy -\n"
    "supplying the real wording is one of the things I'm asking for.\n\n"
    "Note what does NOT change between them: the document's own content is never\n"
    "covered, cropped or scaled. The stamps extend the page downward instead.\n",
    encoding="utf-8")

zip_path = SP / "rebranding-samples.zip"
if zip_path.exists(): zip_path.unlink()
shutil.make_archive(str(zip_path.with_suffix("")), "zip", OUT)
print("zip:", zip_path, f"{zip_path.stat().st_size/1e6:.2f}MB")
