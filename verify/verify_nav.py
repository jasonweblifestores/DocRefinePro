"""v152: the document's own bookmarks and form registration survive rebranding.

Widget and link annotations already came across with the page merge; the outline
and the /AcroForm catalog entry did not, so long manuals arrived with an empty
navigation pane and form fields were orphaned.

Run: python verify_nav.py <repo> "<...>\BrandKit\Sample Files"
"""
import sys, shutil, tempfile, logging
from pathlib import Path

REPO, BK = Path(sys.argv[1]), Path(sys.argv[2])
sys.path.insert(0, str(REPO))
logging.getLogger("pypdf").setLevel(logging.CRITICAL)

results = []
def check(n, ok, d=""):
    results.append(bool(ok)); print(f"[{'PASS' if ok else 'FAIL'}] {n}  {d}")

from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas as rlc
from docrefine.rebrand import BrandKit, rebrand_pdf

kit = BrandKit(BK)
work = Path(tempfile.mkdtemp(prefix="drp_nav_"))

def count(items):
    t = 0
    for i in items or []:
        t += count(i) if isinstance(i, list) else 1
    return t

def dests(rd):
    out = []
    def walk(items):
        for i in items or []:
            if isinstance(i, list): walk(i)
            else:
                try: out.append((str(i.get("/Title", "")).strip(), rd.get_destination_page_number(i)))
                except Exception: pass
    try: walk(rd.outline)
    except Exception: pass
    return out

# --------------------------------------------------------------------------
#  A synthetic document with bookmarks on known pages
# --------------------------------------------------------------------------
src = work / "book.pdf"
c = rlc.Canvas(str(src), pagesize=(612, 792))
for i in range(4):
    c.setFont("Helvetica", 20); c.drawString(70, 700, f"CHAPTER {i+1} BODY TEXT"); c.showPage()
c.save()
w = PdfWriter(); w.append(PdfReader(str(src)))
for i in range(4):
    w.add_outline_item(f"Chapter {i+1}", i)
with open(src, "wb") as f: w.write(f)

s = PdfReader(str(src))
check("N1 the fixture really has bookmarks", count(s.outline) == 4, str(count(s.outline)))

out = work / "book_branded.pdf"
rebrand_pdf(src, out, kit, "Installation Manual")
d = PdfReader(str(out))
check("N2 every bookmark survives", count(d.outline) == 4, str(count(d.outline)))
sd, dd = dests(s), dests(d)
check("N3 the titles are unchanged", [t for t, _ in sd] == [t for t, _ in dd], str([t for t, _ in dd]))
check("N4 each points one page later, because of the front cover",
      all(b == a + 1 for (_, a), (_, b) in zip(sd, dd)), f"{[p for _,p in sd]} -> {[p for _,p in dd]}")
check("N5 no bookmark points outside the document",
      all(0 <= p < len(d.pages) for _, p in dd))
check("N6 the document itself is still branded", len(d.pages) == len(s.pages) + 2)

# a document with NO outline must not gain one, and must not break
plain = work / "plain.pdf"
c = rlc.Canvas(str(plain), pagesize=(612, 792)); c.drawString(70, 700, "NO BOOKMARKS HERE"); c.showPage(); c.save()
po = work / "plain_out.pdf"
rebrand_pdf(plain, po, kit, "Specification Sheet")
check("N7 a document with no bookmarks is unaffected", count(PdfReader(str(po)).outline) == 0)

# --------------------------------------------------------------------------
#  Real files from the batch
# --------------------------------------------------------------------------
SRC = Path(r"C:\Users\WORK\Documents\Batch 4\_unique-to-rebrand")

def acro(rd):
    try: return bool(rd.trailer.get("/Root", {}).get("/AcroForm"))
    except Exception: return False

def widgets(rd):
    n = 0
    for p in rd.pages:
        for a in (p.get("/Annots") or []):
            try:
                if a.get_object().get("/Subtype") == "/Widget": n += 1
            except Exception: pass
    return n

REAL = [("1570-12-BM.pdf", "bookmarks"), ("3500-vertical-mailboxes.pdf", "form + bookmarks")]
for name, what in REAL:
    f = SRC / name
    if not f.is_file():
        print(f"[skip] {name} not on this machine"); continue
    o = work / f"real_{name}"
    rebrand_pdf(f, o, kit, "Installation Manual")
    a, b = PdfReader(str(f)), PdfReader(str(o))
    check(f"N8 {name}: bookmarks kept ({what})", count(a.outline) == count(b.outline),
          f"{count(a.outline)} -> {count(b.outline)}")
    if acro(a):
        check(f"N9 {name}: form registration kept", acro(b))
        check(f"N10 {name}: widget fields kept", widgets(a) == widgets(b),
              f"{widgets(a)} -> {widgets(b)}")

shutil.rmtree(work, ignore_errors=True)
print("\n" + "=" * 56)
print(f"RESULT: {sum(results)}/{len(results)} passed")
sys.exit(0 if all(results) else 1)
