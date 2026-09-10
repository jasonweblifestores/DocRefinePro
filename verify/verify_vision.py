"""v155: the visual pass — looking at pages the text model cannot read.

Routing is tested with a stub so it runs without the 6GB model; the live model is
exercised separately by verify_vision_live.py.

Run: python verify_vision.py <repo> "<...>\BrandKit\Sample Files"
"""
import sys, shutil, tempfile, inspect, logging, base64
from pathlib import Path

REPO = Path(sys.argv[1])
sys.path.insert(0, str(REPO))
logging.getLogger("pypdf").setLevel(logging.CRITICAL)

results = []
def check(n, ok, d=""):
    results.append(bool(ok)); print(f"[{'PASS' if ok else 'FAIL'}] {n}  {d}")

from reportlab.pdfgen import canvas as rlc
from docrefine import classify
from docrefine.config import ConfigData
from docrefine.worker import Worker

work = Path(tempfile.mkdtemp(prefix="drp_vis_"))

def pdf(name, size=(612, 792), body=("HELLO",), pages=1):
    p = work / name
    c = rlc.Canvas(str(p), pagesize=size)
    for _ in range(pages):
        c.setFont("Helvetica", 14)
        y = size[1] - 40
        for line in body:
            c.drawString(40, y, line); y -= 18
        c.showPage()
    c.save()
    return p

# =====================================================================
#  A. Off by default, and configurable
# =====================================================================
cfg = ConfigData()
check("A1 the visual pass is off by default", cfg.rebrand_vision_pass is False)
check("A2 the model is named in config", cfg.rebrand_vision_model == "qwen2.5vl:7b")
check("A3 classify_document takes vision_model, default None",
      inspect.signature(classify.classify_document).parameters["vision_model"].default is None)
check("A4 run_rebrand_analyze takes vision_pass, default None",
      inspect.signature(Worker.run_rebrand_analyze).parameters["vision_pass"].default is None)
check("A5 the unsure threshold is stated", 0 < classify.VISION_UNSURE_BELOW <= 1)

# =====================================================================
#  B. Rendering pages for the model
# =====================================================================
p = pdf("two_pages.pdf", body=("PAGE CONTENT",), pages=3)
imgs = classify.render_pages_b64(p)
check("B1 pages render to base64 images", len(imgs) == classify.VISION_MAX_PAGES, f"{len(imgs)} images")
check("B2 they are real PNGs", all(base64.b64decode(i)[:4] == b"\x89PNG" for i in imgs))
check("B3 an unreadable file yields no images, rather than raising",
      classify.render_pages_b64(work / "nope.pdf") == [])
big = classify.render_pages_b64(p, pages=1)
check("B4 the page count is honoured", len(big) == 1)

# =====================================================================
#  C. Routing — a stub stands in for the model
# =====================================================================
calls = []
def stub(pdf_path, model=None, url=None):
    calls.append(Path(pdf_path).name)
    return {"action": "rebrand", "doc_type": "installation guide", "product": "Stub",
            "asset_type": "installation-guide", "manufacturer": "", "title": "Installation Manual",
            "confidence": 0.95, "source": "vision", "notes": "read the page: numbered steps"}
real_visually, real_text = classify.classify_visually, classify.extract_text
classify.classify_visually = stub

# a PDF with no extractable text
blank = work / "blank_scan.pdf"
c = rlc.Canvas(str(blank), pagesize=(612, 792)); c.rect(80, 80, 300, 300, fill=0); c.showPage(); c.save()
calls.clear()
out = classify.classify_document(blank, vision_model="stub")
check("C1 a file with no readable text is LOOKED AT", calls == ["blank_scan.pdf"], str(calls))
check("C2 and the visual answer is used", out["source"] == "vision" and out["action"] == "rebrand")
check("C3 the sheet says the page was read", "read the page" in out["notes"])

calls.clear()
out = classify.classify_document(blank, vision_model=None)
check("C4 without the visual pass it falls back to filename guessing",
      calls == [] and out["source"] == "fallback", f"{calls} {out['source']}")
check("C5 and that fallback still defaults to leave", out["action"] == "leave")

# the model declining to answer must not lose the document
classify.classify_visually = lambda *a, **k: None
out = classify.classify_document(blank, vision_model="stub")
check("C6 if the vision model cannot answer, the old fallback still runs",
      out["source"] == "fallback" and out["action"] == "leave")
classify.classify_visually = stub

# =====================================================================
#  D. Vision overrides the page-shape proxy, which exists only for blindness
# =====================================================================
# a landscape page with almost no text: the shape rule calls this a drawing
sparse = pdf("sparse_land.pdf", size=(792, 612), body=('32 3/8" ACTUAL',))
check("D1 the shape rule alone calls it a drawing",
      classify.page_shape_suggests_drawing(sparse) is True)
calls.clear()
out = classify.classify_document(sparse, text="", vision_model="stub")
check("D2 with the visual pass it is looked at instead", calls == ["sparse_land.pdf"])
check("D3 and what the model SAW wins over the proxy",
      out["action"] == "rebrand" and out["source"] == "vision", f"{out['action']}/{out['source']}")

# =====================================================================
#  E. A confident text answer is not second-guessed (keeps the fast path fast)
# =====================================================================
class FakeResp:
    def __init__(self, payload): self._p = payload
    def read(self): return self._p
    def __enter__(self): return self
    def __exit__(self, *a): return False
import json as _json, urllib.request as _u
def fake_urlopen(req, timeout=None):
    body = _json.loads(req.data.decode())
    if body.get("images"):
        raise AssertionError("vision was called for a confident text answer")
    return FakeResp(_json.dumps({"response": _json.dumps({
        "action": "rebrand", "doc_type": "installation guide", "product": "Widget",
        "asset_type": "installation-guide", "manufacturer": "Acme",
        "title": "Installation Manual", "confidence": 0.97})}).encode())
orig_open = _u.urlopen
_u.urlopen = fake_urlopen
try:
    dense = pdf("dense.pdf", body=("Thank you for selecting this product. " * 2,) * 30)
    calls.clear()
    out = classify.classify_document(dense, vision_model="stub")
    check("E1 a confident text answer skips the visual pass", calls == [], str(calls))
    check("E2 and is used as-is", out["source"] == "llm" and out["confidence"] >= 0.9)
finally:
    _u.urlopen = orig_open

# an UNSURE text answer does get looked at
def fake_unsure(req, timeout=None):
    body = _json.loads(req.data.decode())
    if body.get("images"):
        raise AssertionError("should not reach here: stub intercepts vision")
    return FakeResp(_json.dumps({"response": _json.dumps({
        "action": "rebrand", "doc_type": "spec sheet", "product": "", "asset_type": "spec-sheet",
        "manufacturer": "", "title": "Specification Sheet", "confidence": 0.4})}).encode())
_u.urlopen = fake_unsure
try:
    calls.clear()
    out = classify.classify_document(dense, vision_model="stub")
    check("E3 an UNSURE text answer is checked visually", calls == ["dense.pdf"], str(calls))
    check("E4 the sheet records why it was escalated",
          "unsure" in out["notes"].lower(), out["notes"])
finally:
    _u.urlopen = orig_open
    classify.classify_visually = real_visually
    classify.extract_text = real_text

# =====================================================================
#  F. Availability is checked, not assumed
# =====================================================================
real_list = classify.list_models
classify.list_models = lambda *a, **k: ["llama3.2:3b"]
check("F1 a missing vision model is reported missing", classify.has_vision_model("qwen2.5vl:7b") is False)
classify.list_models = lambda *a, **k: ["llama3.2:3b", "qwen2.5vl:7b"]
check("F2 a present one is found", classify.has_vision_model("qwen2.5vl:7b") is True)
classify.list_models = lambda *a, **k: (_ for _ in ()).throw(RuntimeError("server down"))
check("F3 a dead server is not fatal", classify.has_vision_model("qwen2.5vl:7b") is False)
classify.list_models = real_list

shutil.rmtree(work, ignore_errors=True)
print("\n" + "=" * 56)
print(f"RESULT: {sum(results)}/{len(results)} passed")
sys.exit(0 if all(results) else 1)
