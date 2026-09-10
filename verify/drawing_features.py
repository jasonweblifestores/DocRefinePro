"""Which cheap, deterministic signals separate a CAD drawing from a guide?"""
import re, sys, logging, random, statistics as st
from pathlib import Path
sys.path.insert(0, ".")
logging.getLogger("pypdf").setLevel(logging.CRITICAL)
from pypdf import PdfReader
from docrefine import reviews, classify
from docrefine.rebrand import page_size

SRC = Path(r"C:\Users\WORK\Documents\Batch 4\_unique-to-rebrand")
rows = reviews.read_plan(r"C:\Users\WORK\Documents\DocRefinePro_Data\Rebrand Reviews\Batch 4___unique-to-rebrand_rebrand_plan.xlsx")

TITLEBLOCK = re.compile(
    r"\b(scale|drawn\s*by|dwg|drawing\s*(no|number)|rev(ision)?\b|sheet\s*\d|"
    r"tolerance|do\s*not\s*scale|checked\s*by|material\b|finish\b|part\s*(no|number))",
    re.I)
DIMS = re.compile(r'\d+\s*[\u00bc-\u00be\u2150-\u215e]?\s*"|\d+\s*/\s*\d+\s*"|\bACTUAL\b|\bTYP\b|\bO\.?C\.?\b')

def feats(path):
    try: rd = PdfReader(str(path))
    except Exception: return None
    n = len(rd.pages)
    txt, imgs = [], 0
    for p in rd.pages[:4]:
        txt.append(p.extract_text() or "")
        try:
            xo = p["/Resources"]["/XObject"].get_object()
            imgs += sum(1 for k in xo if xo[k].get_object().get("/Subtype") == "/Image")
        except Exception: pass
    t = " ".join(" ".join(txt).split())
    w, h = page_size(rd.pages[0])
    area = (w * h) / (72.0 * 72.0)
    return dict(pages=n, chars=len(t), area=area,
                density=len(t) / area if area else 0,
                tb=len(TITLEBLOCK.findall(t)), dims=len(DIMS.findall(t)),
                imgs=imgs, landscape=1 if w > h else 0)

disputed = [r for r in rows if (r.get("action") or "").lower() == "rebrand"
            and classify.filename_suggests_drawing(r["file"])]
guides = [r for r in rows if (r.get("action") or "").lower() == "rebrand"
          and not classify.filename_suggests_drawing(r["file"])]
random.seed(11)
groups = {
    "DISPUTED (drawing-named, marked rebrand)": disputed,
    "CONTROL guides (marked rebrand, non-drawing name)": random.sample(guides, min(120, len(guides))),
}
out = {}
for label, rs in groups.items():
    vals = []
    for r in rs:
        f = SRC / r["file"]
        if not f.exists(): continue
        fv = feats(f)
        if fv and fv["chars"] > 0: vals.append(fv)
    out[label] = vals
    print(f"\n=== {label}  (n={len(vals)})")
    for k in ("density", "tb", "dims", "imgs", "chars", "landscape", "pages"):
        xs = [v[k] for v in vals]
        print(f"   {k:10} median={st.median(xs):8.2f}  mean={st.mean(xs):8.2f}  "
              f"p10={sorted(xs)[len(xs)//10]:7.2f}  p90={sorted(xs)[9*len(xs)//10]:8.2f}")

# a simple deterministic rule, scored on both groups
def looks_drawn(v):
    return (v["density"] < 8 and v["imgs"] == 0) or v["tb"] >= 3
for label, vals in out.items():
    hit = sum(1 for v in vals if looks_drawn(v))
    print(f"\nrule fires on {hit}/{len(vals)} of {label}")
