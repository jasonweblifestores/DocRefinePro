"""v145: pages whose box is stored top-down are handled like any other page."""
import sys, tempfile, shutil
from pathlib import Path

REPO, BK = Path(sys.argv[1]), Path(sys.argv[2])
sys.path.insert(0, str(REPO))

results = []
def check(n, ok, d=""):
    results.append(bool(ok)); print(f"[{'PASS' if ok else 'FAIL'}] {n}  {d}")

from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas as rlc
from docrefine.rebrand import BrandKit, rebrand_pdf, page_size, _orientation, _dominant_orientation

kit = BrandKit(BK)
work = Path(tempfile.mkdtemp(prefix="drp_box_"))
BODY = "IMPERIAL STREET SIGN SYSTEM. Cast aluminium poles and blades. DOT approved."

def make(path, w=612, h=792, invert=False, landscape=False):
    if landscape:
        w, h = h, w
    c = rlc.Canvas(str(path), pagesize=(w, h))
    c.setFont("Helvetica", 16); c.drawString(50, h - 90, BODY)
    c.showPage(); c.save()
    if invert:
        # rewrite the box top-down: legal PDF, two opposite corners
        r = PdfReader(str(path)); wtr = PdfWriter()
        for pg in r.pages:
            pg.mediabox.lower_left = (0, h)
            pg.mediabox.upper_right = (w, 0)
            wtr.add_page(pg)
        with open(path, "wb") as f:
            wtr.write(f)

normal, inverted = work / "normal.pdf", work / "inverted.pdf"
make(normal); make(inverted, invert=True)

# =====================================================================
#  Reading the size
# =====================================================================
raw = PdfReader(str(inverted)).pages[0]
check("P1 the test file really is stored top-down", float(raw.mediabox.height) < 0,
      f"raw height {float(raw.mediabox.height):.0f}")
w, h = page_size(raw)
check("P2 page_size reports it positive", (w, h) == (612.0, 792.0), f"{w}x{h}")
check("P3 and it is portrait, not landscape", _orientation(w, h) == "portrait")
check("P4 a normal page is unaffected", page_size(PdfReader(str(normal)).pages[0]) == (612.0, 792.0))
check("P5 dominant orientation agrees", _dominant_orientation(PdfReader(str(inverted)).pages) == "portrait")

land = work / "land.pdf"; make(land, landscape=True)
check("P6 genuine landscape still detected",
      _dominant_orientation(PdfReader(str(land)).pages) == "landscape")

# =====================================================================
#  Rebranding it must keep the document's own content
# =====================================================================
def brand(src, dst):
    rebrand_pdf(src, dst, kit, "Specification Sheet")
    r = PdfReader(str(dst))
    cw, ch = page_size(r.pages[0])
    body = " ".join((r.pages[1].extract_text() or "").split())
    cover = " ".join((r.pages[0].extract_text() or "").split())
    return r, cw, ch, body, cover

r_i, cw, ch, body_i, cover_i = brand(inverted, work / "out_inv.pdf")
check("P7 cover dimensions are positive", cw > 0 and ch > 0, f"{cw}x{ch}")
check("P8 the original content survives", "IMPERIAL STREET SIGN" in body_i.upper(), body_i[:50])
check("P9 the cover title renders", "SPECIFICATION SHEET" in cover_i.upper(), cover_i[:40])
check("P10 page count is source + 2", len(r_i.pages) == 3)

r_n, _, _, body_n, cover_n = brand(normal, work / "out_norm.pdf")
check("P11 output matches the equivalent normal page", body_i == body_n, f"{len(body_i)} vs {len(body_n)} chars")
check("P12 covers match too", cover_i == cover_n)

# =====================================================================
#  The real file that exposed this
# =====================================================================
real = Path("C:/Users/WORK/Documents/Batch 4/_unique-to-rebrand/Imperial-Street-Sign-Brochure.pdf")
if real.is_file():
    src_txt = " ".join((PdfReader(str(real)).pages[0].extract_text() or "").split())
    _, cw, ch, body, cover = brand(real, work / "out_real.pdf")
    check("P13 the real file: positive cover", cw > 0 and ch > 0, f"{cw}x{ch}")
    check("P14 the real file: all its text survives",
          len(body) >= len(src_txt) * 0.95, f"{len(body)} of {len(src_txt)} chars")
    check("P15 the real file: cover titled", "SPECIFICATION" in cover.upper(), cover[:34])
else:
    print("[skip] the real sample file is not on this machine")

# =====================================================================
#  /Rotate — a page that DISPLAYS landscape must be treated as landscape
#  (612x792 with /Rotate 270 is landscape to every reader; reading the raw
#   box called it portrait, gave it portrait covers and cropped its content)
# =====================================================================
from docrefine.rebrand import page_rotation, _dominant_orientation
from pypdf import PdfWriter as _W
from pypdf.generic import NameObject, NumberObject

for rot in (90, 270, -90):
    p = work / f"rot{rot}.pdf"
    c = rlc.Canvas(str(p), pagesize=(612, 792))
    c.setFont("Helvetica", 24); c.drawString(60, 700, f"ROTATED {rot} CONTENT MARKER")
    c.showPage(); c.save()
    w = _W(); w.append(PdfReader(str(p)))
    w.pages[0][NameObject("/Rotate")] = NumberObject(rot)
    with open(p, "wb") as fh: w.write(fh)

    pg = PdfReader(str(p)).pages[0]
    sw, sh = page_size(pg)
    check(f"P16.{rot} /Rotate {rot} reports display size 792x612",
          (round(sw), round(sh)) == (792, 612), f"{sw:.0f}x{sh:.0f}")
    check(f"P17.{rot} and is treated as landscape",
          _dominant_orientation(PdfReader(str(p)).pages) == "landscape")

    out = work / f"rot{rot}_out.pdf"
    rebrand_pdf(p, out, kit, "Technical Drawing")
    o = PdfReader(str(out))
    bw, bh = page_size(o.pages[1])
    check(f"P18.{rot} the branded page is landscape, not squeezed portrait",
          bw > bh and round(bw) == 792, f"{bw:.0f}x{bh:.0f}")
    check(f"P19.{rot} the content marker survives",
          "ROTATED" in (o.pages[1].extract_text() or "").upper())

check("P20 an unrotated page is unaffected", page_rotation(PdfReader(str(normal)).pages[0]) == 0)

# the real file that exposed it — 207 of the batch's 215 drawing-named rebrands
rot_real = Path("C:/Users/WORK/Documents/Batch 4/_unique-to-rebrand/4c06d-02-sm_cutsheet_pdp_.pdf")
if rot_real.is_file():
    rr = PdfReader(str(rot_real))
    check("P21 the real rotated file reads as landscape",
          _dominant_orientation(rr.pages) == "landscape")
    o = work / "rot_real.pdf"
    rebrand_pdf(rot_real, o, kit, "Specification Sheet")
    op = PdfReader(str(o))
    bw, bh = page_size(op.pages[1])
    check("P22 and is branded landscape at full width", bw > bh and round(bw) == 792, f"{bw:.0f}x{bh:.0f}")
    # NOTE: the drawing's title block is vector outline art, not extractable text —
    # it is absent from the SOURCE's text too. Which is the point: text extraction
    # cannot see cropping, so geometry (P22) is what actually proves this fixed.
    src_t = "".join((rr.pages[0].extract_text() or "").split())
    out_t = "".join((op.pages[1].extract_text() or "").split())
    check("P23 every extractable character still present", src_t and src_t in out_t,
          f"{len(src_t)} src chars")
    check("P24 the branded content area keeps the source's aspect ratio",
          abs((bw / bh) - (792 / 612)) < 0.25, f"{bw:.0f}x{bh:.0f}")
else:
    print("[skip] the real rotated sample is not on this machine")

shutil.rmtree(work, ignore_errors=True)
print("\n" + "=" * 56)
print(f"RESULT: {sum(results)}/{len(results)} passed")
sys.exit(0 if all(results) else 1)
