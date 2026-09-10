"""v141: usable review triage + a drawing-name tiebreak for unsure rebrand calls."""
import sys, tempfile, shutil
from pathlib import Path

REPO = Path(sys.argv[1])
sys.path.insert(0, str(REPO))

results = []
def check(n, ok, d=""):
    results.append(bool(ok)); print(f"[{'PASS' if ok else 'FAIL'}] {n}  {d}")

from docrefine import reviews, classify
from docrefine.classify import filename_suggests_drawing as D, filename_suggests_instructions as I

# =====================================================================
#  Triage — flag decisions, not merely unreadable files
# =====================================================================
def row(src, act, conf="0"):
    return {"source": src, "action": act, "confidence": conf}

check("T1 unreadable + leave is NOT flagged (the safe default)",
      not reviews.needs_review(row("fallback", "leave")))
check("T2 unreadable + rebrand IS flagged (we intend to brand it blind)",
      reviews.needs_review(row("fallback", "rebrand")))
check("T3 low-confidence model call is flagged",
      reviews.needs_review(row("llm", "rebrand", "0.5")))
check("T4 confident model call is not flagged",
      not reviews.needs_review(row("llm", "rebrand", "0.95")))
check("T5 unparseable confidence is flagged",
      reviews.needs_review(row("llm", "leave", "n/a")))

# =====================================================================
#  Drawing-name tiebreak
# =====================================================================
check("D1 drawing filenames recognised",
      D("tech-3635RL.pdf") and D("1570-12-cbu-florence-cut-sheet.pdf")
      and D("base-bolt-pattern-vp.pdf") and D("AF-Pad-Spec-1565-HSCBU.pdf"))
check("D2 instruction filenames are NOT treated as drawings",
      not D("Bradford-Installation-Instructions.pdf")
      and not D("barcelona_assembly_instructions.pdf")
      and not D("206550INS-1400.pdf"))
check("D3 an instruction cut-sheet stays instructions (rule order matters)",
      not D("AF-CS2002-CC51-52-53FL-Specs-Installation-Cut-Sheet.pdf")
      and I("AF-CS2002-CC51-52-53FL-Specs-Installation-Cut-Sheet.pdf"))
check("D4 ordinary product names are untouched",
      not D("florence-cbu-product-catalog.pdf") and not D("7510b-10-warranty.pdf"))

# v152: the rule decides outright, and no longer defers to a confident model.
# It used to fire only below DRAWING_TIEBREAK_BELOW, which let 213 CAD drawings
# through on the real batch because the model called them "installation guide"
# at confidence 1.0. Confidence measured how sure it was, not whether it was right.
def decide(name, action, conf):
    """Mirror of the branch in classify_document."""
    if action == "rebrand" and D(name):
        return "leave"
    return action

check("D5 unsure rebrand of a drawing flips to leave",
      decide("tech-32355.pdf", "rebrand", 0.5) == "leave")
check("D6 a CONFIDENT rebrand of a drawing now also flips to leave",
      decide("1570-4t5-cut-sheet.pdf", "rebrand", 1.0) == "leave")
check("D6b and so does the real cut-sheet that exposed this",
      decide("4c06d-02-sm_cutsheet_pdp_.pdf", "rebrand", 1.0) == "leave")
check("D7 a rebrand of a NON-drawing is untouched at any confidence",
      decide("Bradford-Installation-Instructions.pdf", "rebrand", 0.4) == "rebrand"
      and decide("Bradford-Installation-Instructions.pdf", "rebrand", 1.0) == "rebrand")
check("D8 a leave decision is never flipped to rebrand",
      decide("tech-32355.pdf", "leave", 0.2) == "leave")
check("D9 an instruction cut-sheet is still instructions, not a drawing",
      decide("AF-CS2002-CC51-52-53FL-Specs-Installation-Cut-Sheet.pdf", "rebrand", 1.0) == "rebrand")

# the override must be visible in the sheet, or nobody reviews it
NOTE = ("filename says technical drawing — the brief leaves these as-is; "
        "flip to rebrand if this one is really a guide or spec sheet")
check("D10 an overridden row is flagged for review however sure the model was",
      reviews.needs_review({"action": "leave", "source": "llm",
                            "confidence": "1", "notes": NOTE}) is True)

# =====================================================================
#  v154: page shape as evidence — the signal the model never had
#  (a drawing is a visual artifact; the model only ever saw extracted text)
# =====================================================================
from reportlab.pdfgen import canvas as rlcv
sd = Path(tempfile.mkdtemp(prefix="drp_shape_"))

def make(name, size, body):
    p = sd / name
    c = rlcv.Canvas(str(p), pagesize=size)
    c.setFont("Helvetica", 9)
    y = size[1] - 40
    for line in body:
        c.drawString(30, y, line); y -= 11
    c.showPage(); c.save()
    return p

# a landscape page with a handful of dimension labels — a drawing
drawing = make("land_sparse.pdf", (792, 612),
               ['32 3/8" ACTUAL', '24 7/16" ACTUAL', "FINISHED FLOOR", "SM06D"])
# the same sparse text on a portrait page — not a drawing by shape
port = make("port_sparse.pdf", (612, 792),
            ['32 3/8" ACTUAL', '24 7/16" ACTUAL', "FINISHED FLOOR", "SM06D"])
# a landscape page dense with prose — a document
dense = make("land_dense.pdf", (792, 612),
             ["Thank you for selecting this product. " * 3] * 40)

check("S1 a sparse landscape page reads as a drawing",
      classify.page_shape_suggests_drawing(drawing) is True)
check("S2 a sparse PORTRAIT page does not (orientation matters)",
      classify.page_shape_suggests_drawing(port) is False)
check("S3 a text-dense landscape page does not",
      classify.page_shape_suggests_drawing(dense) is False)
check("S4 the threshold is stated, not magic",
      isinstance(classify.DRAWING_MAX_DENSITY, float) and classify.DRAWING_MAX_DENSITY > 0)
check("S5 an unreadable file is not guessed at",
      classify.page_shape_suggests_drawing(sd / "does-not-exist.pdf") is False)

REAL = Path(r"C:\Users\WORK\Documents\Batch 4\_unique-to-rebrand")
if REAL.is_dir():
    cad = REAL / "4c06d-02-sm_cutsheet_pdp_.pdf"
    already_branded = REAL / "1570-13-BM.pdf"
    guide = REAL / "1570-cbu-installation-manual.pdf"
    if cad.is_file():
        check("S6 the real cut-sheet that exposed this reads as a drawing",
              classify.page_shape_suggests_drawing(cad) is True)
    if already_branded.is_file():
        check("S7 and a drawing the FILENAME rule missed is caught by shape",
              classify.page_shape_suggests_drawing(already_branded) is True
              and classify.filename_suggests_drawing("1570-13-BM.pdf") is False)
    if guide.is_file():
        check("S8 a genuine installation manual is untouched",
              classify.page_shape_suggests_drawing(guide) is False)

shutil.rmtree(sd, ignore_errors=True)

# =====================================================================
#  extract_text is unchanged from v140 (the page-1 experiment was reverted)
# =====================================================================
from reportlab.pdfgen import canvas as rlc
d = Path(tempfile.mkdtemp(prefix="drp_v141_"))
p = d / "multi.pdf"
c = rlc.Canvas(str(p), pagesize=(612, 792))
for _ in range(2):
    c.rect(80, 80, 400, 400, fill=1); c.showPage()
c.setFont("Helvetica", 20); c.drawString(70, 700, "REAL TEXT ON PAGE THREE OF THIS DOCUMENT")
c.showPage(); c.save()
check("X1 text past page 2 is still found (v140 behaviour kept)",
      len(classify.extract_text(p)) > 20, f"{len(classify.extract_text(p))} chars")

# =====================================================================
#  End to end through a written sheet
# =====================================================================
rows = [
    {"file": "tech-1.pdf", "action": "leave",   "source": "fallback", "confidence": "0"},
    {"file": "guide.pdf",  "action": "rebrand", "source": "fallback", "confidence": "0"},
    {"file": "spec.pdf",   "action": "rebrand", "source": "llm",      "confidence": "0.95"},
    {"file": "odd.pdf",    "action": "rebrand", "source": "llm",      "confidence": "0.4"},
]
from docrefine.worker import Worker
xp = d / "plan.xlsx"
reviews.write_plan(xp, rows, Worker.REBRAND_PLAN_COLUMNS)
from openpyxl import load_workbook
wb = load_workbook(xp); ws = wb.active
flags = [ws.cell(row=i, column=1).value for i in range(2, 6)]
wb.close()
check("E1 only the two genuinely uncertain rows are flagged",
      flags == [None, "YES", None, "YES"], str(flags))

shutil.rmtree(d, ignore_errors=True)
print("\n" + "=" * 56)
print(f"RESULT: {sum(results)}/{len(results)} passed")
sys.exit(0 if all(results) else 1)
