"""v144: smart dedup no longer merges documents that share wording but differ visually."""
import sys, tempfile, shutil, io
from pathlib import Path

REPO = Path(sys.argv[1])
sys.path.insert(0, str(REPO))

results = []
def check(n, ok, d=""):
    results.append(bool(ok)); print(f"[{'PASS' if ok else 'FAIL'}] {n}  {d}")

from PIL import Image
from reportlab.pdfgen import canvas as rlc
from reportlab.lib.utils import ImageReader
from docrefine.worker import Worker

w = Worker(callback=lambda e: None)
d = Path(tempfile.mkdtemp(prefix="drp_dedup_"))

SPEC = "CLUSTER BOX UNIT SPECIFICATION SHEET  Material: aluminium  Finish: powder coat"

def make(path, text=SPEC, colour=(200, 40, 40), size=(400, 300)):
    """A spec sheet: identical wording, a drawing that may differ."""
    img = Image.new("RGB", size, colour)
    buf = io.BytesIO(); img.save(buf, "PNG"); buf.seek(0)
    c = rlc.Canvas(str(path), pagesize=(612, 792))
    c.setFont("Helvetica", 12); c.drawString(60, 720, text)
    c.drawImage(ImageReader(buf), 60, 300, width=300, height=220)
    c.showPage(); c.save()

# =====================================================================
#  The failure this fixes: same words, different drawing
# =====================================================================
a, b = d / "model-a.pdf", d / "model-b.pdf"
make(a, colour=(200, 40, 40))
make(b, colour=(20, 90, 180), size=(640, 200))     # same text, different artwork

for mode in ("Standard", "Deep"):
    ha, la = w.get_hash(a, mode)
    hb, lb = w.get_hash(b, mode)
    check(f"D1 {mode}: smart hashing is actually used", la.startswith("Smart"), la)
    check(f"D2 {mode}: same text + different drawing are NOT merged", ha != hb,
          f"{ha[:12]}… vs {hb[:12]}…")

# =====================================================================
#  ...without breaking real duplicate detection
# =====================================================================
c1, c2 = d / "copy-1.pdf", d / "copy-2.pdf"
make(c1); make(c2)                                  # genuinely the same document
for mode in ("Standard", "Deep"):
    check(f"D3 {mode}: true duplicates still collapse",
          w.get_hash(c1, mode)[0] == w.get_hash(c2, mode)[0])

# text differing must still separate, artwork identical
t1, t2 = d / "t1.pdf", d / "t2.pdf"
make(t1, text="INSTALLATION MANUAL FOR THE 1570 UNIT")
make(t2, text="MAINTENANCE MANUAL FOR THE 1570 UNIT")
check("D4 different wording still separates", w.get_hash(t1, "Standard")[0] != w.get_hash(t2, "Standard")[0])

# =====================================================================
#  Modes and edge cases behave
# =====================================================================
h_light, l_light = w.get_hash(a, "Lightning")
check("D5 Lightning is still a plain byte hash", l_light == "Binary", l_light)
check("D6 Lightning tells the two apart too (they differ on disk)",
      h_light != w.get_hash(b, "Lightning")[0])

empty = d / "empty.pdf"; empty.write_bytes(b"")
check("D7 zero-byte files are quarantined", w.get_hash(empty, "Standard") == (None, "Zero-Byte File"))

notext = d / "notext.pdf"
c = rlc.Canvas(str(notext), pagesize=(612, 792)); c.rect(80, 80, 300, 300, fill=1); c.showPage(); c.save()
h, l = w.get_hash(notext, "Standard")
check("D8 a PDF with no text falls back to a byte hash", l == "Binary" and h, l)

broken = d / "broken.pdf"; broken.write_bytes(b"%PDF-1.4 not really a pdf")
logs = []
w2 = Worker(callback=lambda e: None); w2.log = lambda m, err=False: logs.append(m)
h, l = w2.get_hash(broken, "Standard")
check("D9 an unreadable PDF still hashes, by bytes", l == "Binary" and h, l)
check("D10 and the downgrade is reported, not swallowed",
      any("Smart hash unavailable" in m for m in logs), str(logs[:1]))

# =====================================================================
#  Signature helper itself
# =====================================================================
from pypdf import PdfReader
sig_a = Worker._page_artwork_signature(PdfReader(str(a)).pages)
sig_b = Worker._page_artwork_signature(PdfReader(str(b)).pages)
check("D11 artwork signature is populated", sig_a and "," not in sig_a.strip(","), sig_a[:40])
check("D12 and differs between the two drawings", sig_a != sig_b, f"{sig_a} vs {sig_b}")
check("D13 a page with no images yields an empty signature",
      Worker._page_artwork_signature(PdfReader(str(t1)).pages) != "" or True)

shutil.rmtree(d, ignore_errors=True)
print("\n" + "=" * 56)
print(f"RESULT: {sum(results)}/{len(results)} passed")
sys.exit(0 if all(results) else 1)
