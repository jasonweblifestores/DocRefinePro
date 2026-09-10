"""v157: rescuing a PDF our library cannot open, and one shared failure rule.

Two things are asserted here.

**The rescue is only accepted when it is provably as good as the original.** A
rewrite that silently drops pages or text is worse than shipping the untouched
file, so it must be rejected rather than used. Poppler is stubbed so those
rejections can be forced rather than hoped for.

**A document we cannot brand still reaches the delivery.** That rule used to live
in two places and had already drifted: the pipeline copied the original through,
the delivery path dropped the file. Both now go through one method.

Run: python verify_repair.py <repo> "<...>\BrandKit\Sample Files"
"""
import inspect
import logging
import sys
import tempfile
from pathlib import Path

REPO = Path(sys.argv[1])
sys.path.insert(0, str(REPO))
logging.getLogger("pypdf").setLevel(logging.CRITICAL)

results = []


def check(n, ok, d=""):
    results.append(bool(ok))
    print(f"[{'PASS' if ok else 'FAIL'}] {n}  {d}")


from reportlab.pdfgen import canvas as rlc

from docrefine import processing as P
from docrefine.worker import Worker

work = Path(tempfile.mkdtemp(prefix="drp_rep_"))


def pdf(path, pages=1, body="RESCUE ME"):
    c = rlc.Canvas(str(path), pagesize=(612, 792))
    for _ in range(pages):
        c.setFont("Helvetica", 14)
        c.drawString(40, 700, body)
        c.showPage()
    c.save()
    return path


src = pdf(work / "broken.pdf", pages=1, body="ORIGINAL CONTENT HERE")
good = pdf(work / "rewritten_ok.pdf", pages=1, body="ORIGINAL CONTENT HERE")
short = pdf(work / "rewritten_short.pdf", pages=1, body="X")
extra = pdf(work / "rewritten_pages.pdf", pages=3, body="ORIGINAL CONTENT HERE")

_real_poppler = P._poppler


def stub(pages="1", text="ORIGINAL CONTENT HERE", produce=good, info_rc=0, cairo_rc=0):
    """Pretend to be poppler, so rejection paths can be forced."""
    def fake(tool, *args, **kw):
        if tool == "pdfinfo":
            return info_rc, f"Pages:           {pages}\n", ""
        if tool == "pdftotext":
            return 0, text, ""
        if tool == "pdftocairo":
            out = Path(args[-1])
            if produce is not None:
                out.write_bytes(Path(produce).read_bytes())
            return cairo_rc, "", "" if cairo_rc == 0 else "boom"
        return 1, "", "unexpected tool"
    P._poppler = fake


# =====================================================================
#  A. The rescue only succeeds when the rewrite holds up
# =====================================================================
stub()
out = work / "out_a.pdf"
check("A1 a faithful rewrite is accepted", P.repair_unreadable_pdf(src, out) is True)
check("A2 and it is left on disk to be used", out.is_file() and out.stat().st_size > 0)

stub(produce=extra)
out = work / "out_b.pdf"
check("A3 a rewrite with the wrong page count is rejected",
      P.repair_unreadable_pdf(src, out) is False)
check("A4 and the bad rewrite is not left behind", not out.exists())

stub(produce=short)
out = work / "out_c.pdf"
check("A5 a rewrite that loses the text is rejected",
      P.repair_unreadable_pdf(src, out) is False)
check("A6 and that one is cleaned up too", not out.exists())

stub(info_rc=1)
out = work / "out_d.pdf"
check("A7 a file poppler cannot read either is not attempted",
      P.repair_unreadable_pdf(src, out) is False)

stub(cairo_rc=1, produce=None)
out = work / "out_e.pdf"
check("A8 a failed rewrite is reported, not assumed",
      P.repair_unreadable_pdf(src, out) is False)

stub()
out = work / "sub" / "deep" / "out_f.pdf"
check("A9 the output directory is created as needed",
      P.repair_unreadable_pdf(src, out) is True and out.is_file())

P._poppler = _real_poppler

# =====================================================================
#  B. The shared rule: never drop a document from the delivery
# =====================================================================
class W(Worker):
    def __init__(self):
        self.logs = []

    def log(self, m, err=False):
        self.logs.append(str(m))


# repaired: the rescue works, so the file is branded from the rewrite
stub()
w = W()
dst = work / "d1.pdf"
branded = {}
res = w._fallback_deliver(src, dst, "broken.pdf", ValueError("bad RC4"),
                          rebrand=lambda p: branded.setdefault("from", Path(p).name))
check("B1 a rescued file is branded from the rewrite", res == "repaired", str(res))
check("B2 and branding was handed the rewrite, not the original",
      branded.get("from") == "broken.pdf", str(branded))
check("B3 the log says the rewrite was checked",
      any("checked against the original" in m for m in w.logs))

# unbranded: rescue impossible, so the original must still ship
stub(info_rc=1)
w = W()
dst = work / "d2.pdf"
res = w._fallback_deliver(src, dst, "broken.pdf", ValueError("bad RC4"),
                          rebrand=lambda p: (_ for _ in ()).throw(RuntimeError("nope")))
check("B4 an unrescuable file still reaches the delivery", res == "unbranded", str(res))
check("B5 and it is the original, byte for byte",
      dst.is_file() and dst.read_bytes() == src.read_bytes())
check("B6 and the log says it went out unbranded",
      any("unbranded" in m for m in w.logs))

# with no rescue offered at all, it still copies through
w = W()
dst = work / "d3.pdf"
check("B7 without a rescue callable it still delivers the original",
      w._fallback_deliver(src, dst, "broken.pdf", ValueError("x")) == "unbranded")

# a missing source cannot be delivered, and says so rather than pretending
w = W()
check("B8 a source that isn't there yields nothing, not a false success",
      w._fallback_deliver(work / "nope.pdf", work / "d4.pdf", "nope.pdf",
                          ValueError("x")) is None)

P._poppler = _real_poppler

# =====================================================================
#  C. Both rebrand paths use the one rule
# =====================================================================
apply_src = inspect.getsource(Worker._apply_one_row)
folder_src = inspect.getsource(Worker._folder_rebrand)
check("C1 the delivery path routes failures through _fallback_deliver",
      "_fallback_deliver" in apply_src)
check("C2 the pipeline path routes failures through _fallback_deliver",
      "_fallback_deliver" in folder_src)
check("C3 the delivery path offers the rescue", "rebrand=brand" in apply_src)
check("C4 the pipeline path offers the rescue", "rebrand=brand" in folder_src)
check("C5 the pipeline path can't NameError on an early failure",
      "brand = None" in folder_src)
check("C6 neither path copies through by hand any more",
      "_atomic_copy(p, dst)\n                except Exception: pass" not in folder_src)

print("\n" + "=" * 50)
print(f"RESULT: {sum(results)}/{len(results)} passed")
sys.exit(0 if all(results) else 1)
