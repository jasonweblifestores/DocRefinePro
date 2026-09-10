"""Main-engine audit fixes (A1-A6): stop always unlocks the UI, atomic manifest,
no partial outputs, compressed Office sanitise, live max_pixels setting."""
import sys, ast, json, zipfile, tempfile, threading, shutil, os
from pathlib import Path

REPO = Path(sys.argv[1])
sys.path.insert(0, str(REPO))

results = []
def check(n, ok, d=""):
    results.append(bool(ok)); print(f"[{'PASS' if ok else 'FAIL'}] {n}  {d}")

from docrefine.worker import Worker, _atomic_write_json, _atomic_copy
from docrefine.core.events import EventType
from docrefine.config import Constants, WORKSPACES_ROOT, CFG

# =====================================================================
#  A1/A2 — every run_* exit path must emit DONE, or the UI stays locked
# =====================================================================
src = (REPO / "docrefine" / "worker.py").read_text(encoding="utf-8")
tree = ast.parse(src)
cls = next(n for n in tree.body if isinstance(n, ast.ClassDef) and n.name == "Worker")

def emits_done(node):
    return any(isinstance(n, ast.Attribute) and n.attr == "DONE" for n in ast.walk(node))

offenders = []
for fn in cls.body:
    if not (isinstance(fn, ast.FunctionDef) and fn.name.startswith("run_")):
        continue
    nested = {id(n) for f in ast.walk(fn) if isinstance(f, ast.FunctionDef) and f is not fn
              for n in ast.walk(f)}
    for n in ast.walk(fn):
        if isinstance(n, ast.Return) and id(n) not in nested:
            guard = None
            for p in ast.walk(fn):
                if isinstance(p, ast.If) and any(n is s or n in ast.walk(s) for s in p.body):
                    guard = p
            if guard is not None and not emits_done(guard):
                offenders.append(f"{fn.name}:{n.lineno}")
check("A1 no run_* exit path skips DONE", not offenders, str(offenders))

# the four stop paths now log as well as unlock
for fn_name, phrase in [("run_inventory", "stopped by user during tagging"),
                        ("run_organize", "Unique export stopped"),
                        ("run_distribute", "Distribution stopped"),
                        ("run_full_export", "CSV export stopped")]:
    body = src.split(f"def {fn_name}(")[1].split("\n    def ")[0]
    check(f"A1 {fn_name} explains the stop", phrase in body)

# A2 — export with no manifest reports instead of dying quietly
ws = Path(tempfile.mkdtemp(prefix="drp_eng_")) / "empty_ws"; ws.mkdir(parents=True)
seen = []
w = Worker(callback=lambda e: seen.append(e))
w.run_full_export(str(ws))
kinds = [e.type for e in seen]
check("A2 CSV export with no manifest emits DONE", EventType.DONE in kinds, str(kinds))
check("A2 ...and an explanatory error", EventType.ERROR in kinds, str(kinds))

# =====================================================================
#  A3 — manifest.json is written atomically
# =====================================================================
d = Path(tempfile.mkdtemp(prefix="drp_eng2_"))
big = {f"{i:032x}": {"name": f"f{i}.pdf", "copies": [f"f{i}.pdf"]} for i in range(500)}
_atomic_write_json(d / "manifest.json", big, indent=4)
check("A3 atomic write produces valid JSON",
      len(json.loads((d / "manifest.json").read_text(encoding="utf-8"))) == 500)

good = (d / "manifest.json").read_bytes()
real_dump = json.dump
def exploding_dump(obj, fp, **kw):
    fp.write('{"partial": ')
    raise IOError("simulated crash during serialise")
import docrefine.worker as WK
WK.json.dump = exploding_dump
try:
    _atomic_write_json(d / "manifest.json", big, indent=4)
except IOError:
    pass
WK.json.dump = real_dump
check("A3 a crash mid-write leaves the old manifest intact",
      (d / "manifest.json").read_bytes() == good)
check("A3 no .tmp debris", not list(d.glob("*.tmp")))

# =====================================================================
#  A4 — the refine engine leaves no partial outputs
# =====================================================================
from docrefine.processing import ImageProcessor, OfficeProcessor
from PIL import Image
ev = threading.Event(); ev.set()
img_src = d / "pic.png"; Image.new("RGB", (2400, 1400), (200, 40, 40)).save(img_src)
ip = ImageProcessor(lambda *a, **k: None, lambda: False, ev)

real_save = Image.Image.save
def boom_save(self, fp, *a, **k):
    real_save(self, fp, *a, **k)
    raise IOError("simulated crash after partial save")
Image.Image.save = boom_save
try:
    ip.resize(img_src, d / "out.jpg", 800)
except Exception:
    pass
Image.Image.save = real_save
check("A4 crashed resize leaves no destination file", not (d / "out.jpg").exists())
check("A4 no .part debris", not list(d.glob("*.part")))
check("A4 a normal resize still works",
      ip.resize(img_src, d / "ok.jpg", 800) and (d / "ok.jpg").exists())

# =====================================================================
#  A5 — Office sanitise compresses, and handles UTF-8 metadata
# =====================================================================
doc = d / "in.docx"
body = "<w:p><w:r><w:t>" + "The quick brown fox. " * 400 + "</w:t></w:r></w:p>"
with zipfile.ZipFile(doc, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", '<?xml version="1.0"?><Types/>')
    z.writestr("word/document.xml", "<w:document>" + body * 40 + "</w:document>")
    z.writestr("docProps/core.xml",
               '<?xml version="1.0" encoding="UTF-8"?><cp:coreProperties>'
               '<dc:creator>José Muñoz</dc:creator>'
               '<dc:title>Instrucciones de instalación — Español</dc:title>'
               '</cp:coreProperties>'.encode("utf-8"))
op = OfficeProcessor(lambda *a, **k: None, lambda: False, ev)
out_doc = d / "out.docx"
check("A5 sanitize succeeds", op.sanitize(doc, out_doc))
ratio = out_doc.stat().st_size / doc.stat().st_size
check("A5 output is not inflated", ratio < 2.0, f"{ratio:.2f}x of source")
with zipfile.ZipFile(out_doc) as z:
    types = {i.compress_type for i in z.infolist()}
    core = z.read("docProps/core.xml").decode("utf-8")
check("A5 entries are DEFLATE-compressed", types == {zipfile.ZIP_DEFLATED}, str(types))
check("A5 author stripped", "<dc:creator></dc:creator>" in core.replace(" ", ""))
check("A5 non-ASCII title preserved exactly",
      "instalación — Español" in core, core[core.find("<dc:title>"):][:46])
check("A5 output is a readable zip", zipfile.ZipFile(out_doc).testzip() is None)

# =====================================================================
#  A6 — the max_pixels setting is actually applied
# =====================================================================
import importlib
prev = CFG.get("max_pixels")
try:
    CFG._data.max_pixels = 12345678
    import docrefine.processing as P
    importlib.reload(P)
    check("A6 Pillow limit follows the setting", Image.MAX_IMAGE_PIXELS == 12345678,
          str(Image.MAX_IMAGE_PIXELS))
finally:
    CFG._data.max_pixels = prev
    import docrefine.processing as P
    importlib.reload(P)
check("A6 restored to the configured value", Image.MAX_IMAGE_PIXELS == prev, str(Image.MAX_IMAGE_PIXELS))

shutil.rmtree(d, ignore_errors=True)
shutil.rmtree(ws.parent, ignore_errors=True)
print("\n" + "=" * 56)
print(f"RESULT: {sum(results)}/{len(results)} passed")
sys.exit(0 if all(results) else 1)
