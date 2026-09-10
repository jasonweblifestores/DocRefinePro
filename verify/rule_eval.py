import re, sys, logging, random
from pathlib import Path
sys.path.insert(0, ".")
logging.getLogger("pypdf").setLevel(logging.CRITICAL)
from pypdf import PdfReader
from docrefine import reviews, classify
from docrefine.rebrand import page_size
SRC = Path(r"C:\Users\WORK\Documents\Batch 4\_unique-to-rebrand")
rows = reviews.read_plan(r"C:\Users\WORK\Documents\DocRefinePro_Data\Rebrand Reviews\Batch 4___unique-to-rebrand_rebrand_plan.xlsx")

def feats(path):
    try: rd = PdfReader(str(path))
    except Exception: return None
    txt, imgs = [], 0
    for p in rd.pages[:4]:
        txt.append(p.extract_text() or "")
        try:
            xo = p["/Resources"]["/XObject"].get_object()
            imgs += sum(1 for k in xo if xo[k].get_object().get("/Subtype") == "/Image")
        except Exception: pass
    t = " ".join(" ".join(txt).split())
    w, h = page_size(rd.pages[0]); area = (w*h)/5184.0
    return dict(chars=len(t), density=len(t)/area if area else 0, imgs=imgs,
                landscape=1 if w > h else 0, area=area)

reb = [r for r in rows if (r.get("action") or "").lower() == "rebrand"]
disputed = [r for r in reb if classify.filename_suggests_drawing(r["file"])]
others   = [r for r in reb if not classify.filename_suggests_drawing(r["file"])]
random.seed(11)
ctrl = random.sample(others, min(120, len(others)))

def load(rs):
    o = []
    for r in rs:
        f = SRC / r["file"]
        if not f.exists(): continue
        v = feats(f)
        if v and v["chars"] > 0: o.append((r["file"], v))
    return o

D, C = load(disputed), load(ctrl)
RULES = {
    "landscape AND density<10":              lambda v: v["landscape"] and v["density"] < 10,
    "density<6":                             lambda v: v["density"] < 6,
    "landscape AND density<10 AND imgs==0":  lambda v: v["landscape"] and v["density"] < 10 and v["imgs"] == 0,
    "landscape AND chars<800":               lambda v: v["landscape"] and v["chars"] < 800,
}
print(f'{"rule":44} {"disputed":>14} {"controls":>14}')
for n, fn in RULES.items():
    a = sum(1 for _, v in D if fn(v)); b = sum(1 for _, v in C if fn(v))
    print(f'{n:44} {a:5}/{len(D):<8} {b:5}/{len(C):<8}')
best = RULES["landscape AND density<10"]
hits = [n for n, v in C if best(v)]
print(f"\ncontrol files the rule fires on ({len(hits)}):")
for n in hits[:20]: print("   " + n)
Path(sys.argv[1], "control_hits.txt").write_text("\n".join(hits), encoding="utf-8")
