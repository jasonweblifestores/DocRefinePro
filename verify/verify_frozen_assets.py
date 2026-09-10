"""Does the RELEASED bundle actually give the rebrand engine what it needs?

The failure this is looking for is silent. `rebrand._draw_overlay` computes the
stamp band as

    sh = stamps.height(...) if (stamps and _ensure_font()) else 0.0

so if the bundled Poppins cannot be found, `_ensure_font()` returns False and
**no stamps are drawn at all** — no error, no warning, just a rebranded file
missing the attribution, tagline, version and disclaimer. That is exactly the
frozen-path risk worth ruling out before a delivery run.

This points sys._MEIPASS at the real extracted release bundle and asks the real
resolution code what it finds, then renders a stamped page with the font it
resolved. Run with the *source* tree on sys.path: the code is identical to what
is frozen, and what is being tested is the bundle layout.

Usage: python verify_frozen_assets.py <repo> <extracted _internal dir>
"""
import sys
from pathlib import Path

REPO = Path(sys.argv[1])
MEIPASS = Path(sys.argv[2])

results = []


def check(n, ok, d=""):
    results.append(bool(ok))
    print(f"[{'PASS' if ok else 'FAIL'}] {n}  {d}")


check("the bundle directory exists", MEIPASS.is_dir(), str(MEIPASS))

# Become a frozen app pointed at the real bundle, BEFORE importing docrefine.
sys.frozen = True
sys._MEIPASS = str(MEIPASS)
sys.path.insert(0, str(REPO))

from docrefine.config import SystemUtils  # noqa: E402
find_binary = SystemUtils.find_binary

check("get_resource_dir follows _MEIPASS when frozen",
      Path(SystemUtils.get_resource_dir()) == MEIPASS, str(SystemUtils.get_resource_dir()))

from docrefine import rebrand, stamps as S  # noqa: E402

# --- the silent-failure check -------------------------------------------------
fp = rebrand._find_font()
check("the cover/stamp font resolves inside the bundle", bool(fp), str(fp))
check("and it is the bundled copy, not a stray system font",
      bool(fp) and str(MEIPASS) in str(fp), str(fp))
check("_ensure_font() succeeds — so stamps will actually be drawn",
      rebrand._ensure_font() is True)

# --- the other bundled assets the rebrand path reaches for --------------------
check("brand.example.json ships with the app",
      (MEIPASS / "docrefine" / "assets" / "brand.example.json").is_file())
check("the Jinja report template ships",
      (MEIPASS / "docrefine" / "templates" / "report.html").is_file())
check("poppler's pdftoppm is found through find_binary",
      bool(find_binary("pdftoppm.exe")), str(find_binary("pdftoppm.exe")))
check("tesseract is found through find_binary",
      bool(find_binary("tesseract.exe")), str(find_binary("tesseract.exe")))

# --- now actually render a stamped page with what the bundle provided ---------
band = S.Stamps(attribution="Manufactured by Salsbury Industries | Sold by Budget Mailboxes",
                tagline="Trusted by the Nation",
                version_line="Version 1.0 · Last Updated: August 2026",
                disclaimer="This guide is provided for reference. Always consult "
                           "manufacturer specifications for complete details.")
band.ink = (0.12, 0.05, 0.4)
h = band.height(612.0, rebrand._FONT_NAME)
check("the stamp band has a real height with the bundled font", h > 20, f"{h:.1f}pt")

import io  # noqa: E402
from reportlab.pdfgen import canvas as rlc  # noqa: E402

buf = io.BytesIO()
c = rlc.Canvas(buf, pagesize=(612, 792))
band.draw(c, 612.0, 0.0, rebrand._FONT_NAME)
c.showPage()
c.save()
pdf = buf.getvalue()
check("a page with the stamp band renders", pdf[:5] == b"%PDF-", f"{len(pdf)} bytes")

from pypdf import PdfReader  # noqa: E402
txt = PdfReader(io.BytesIO(pdf)).pages[0].extract_text() or ""
for label, needle in [("attribution", "Salsbury Industries"),
                      ("tagline", "Trusted by the Nation"),
                      ("version", "Version 1.0"),
                      ("last updated", "Last Updated"),
                      ("disclaimer", "consult manufacturer specifications")]:
    check(f"the {label} is on the rendered page", needle in txt.replace("\n", " "))

# the font must be embedded — a hard limit in the brief
fonts = set()
res = PdfReader(io.BytesIO(pdf)).pages[0].get("/Resources") or {}
for name, f in (res.get("/Font") or {}).items():
    obj = f.get_object()
    fonts.add(str(obj.get("/BaseFont", "")))
    desc = obj.get("/FontDescriptor")
    if desc:
        d = desc.get_object()
        check("the stamp font is embedded in the PDF",
              any(k in d for k in ("/FontFile", "/FontFile2", "/FontFile3")), str(sorted(d.keys())))
check("the stamp text uses the bundled brand font",
      any("Poppins" in f or "BrandTitle" in f for f in fonts), str(sorted(fonts)))

print("\n" + "=" * 50)
print(f"RESULT: {sum(results)}/{len(results)} passed")
sys.exit(0 if all(results) else 1)
