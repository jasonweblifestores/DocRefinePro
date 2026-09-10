import sys, logging
from pathlib import Path
sys.path.insert(0, ".")
logging.getLogger("pypdf").setLevel(logging.CRITICAL)
from pypdf import PdfReader
from docrefine.rebrand import BrandKit, rebrand_pdf, page_size

SRC = Path(r"C:\Users\WORK\Documents\Batch 4\_unique-to-rebrand")
kit = BrandKit(r"C:\Users\WORK\Documents\Batch 4\Template")
W = Path(sys.argv[1]) / "scan_out"; W.mkdir(parents=True, exist_ok=True)

def links(rd):
    n = 0
    for p in rd.pages:
        for a in (p.get("/Annots") or []):
            try:
                o = a.get_object()
                if o.get("/Subtype") == "/Link": n += 1
            except Exception: pass
    return n

def annots(rd):
    return sum(len(p.get("/Annots") or []) for p in rd.pages)

def outline(rd):
    try: return len(rd.outline or [])
    except Exception: return 0

def acroform(rd):
    try: return bool(rd.trailer.get("/Root", {}).get("/AcroForm"))
    except Exception: return False

CASES = [
    ("links + annotations",        "1570-12v-cut-sheet.pdf"),
    ("AcroForm fields",            "1570-cbu-installation-manual.pdf"),
    ("bookmarks/outline",          "1570-12-BM.pdf"),
    ("origin offset [-32,-17]",    "metropolis-style-wall-mount-instructions.pdf"),
    ("origin offset [0,8]",        "3001-Anchor-Bolt-Specifications.pdf"),
    ("CropBox < MediaBox",         "install_keystone_post_cuff.pdf"),
    ("encrypted",                  "30815-lock-change.pdf"),
    ("mixed size + rotation",      "4c11d-10-sm_cutsheet_pdp_.pdf"),
]
print(f"{'case':26} {'annots':>12} {'links':>10} {'outline':>10} {'form':>7}   geometry")
print("-" * 100)
for label, name in CASES:
    f = SRC / name
    if not f.exists(): print(f"{label:26} MISSING {name}"); continue
    try:
        s = PdfReader(str(f))
        out = W / name
        rebrand_pdf(f, out, kit, "Installation Manual")
        d = PdfReader(str(out))
    except Exception as e:
        print(f"{label:26} FAILED TO REBRAND: {type(e).__name__}: {e}"); continue
    sw, sh = page_size(s.pages[0]); dw, dh = page_size(d.pages[1])
    geo = f"src {sw:.0f}x{sh:.0f} -> body {dw:.0f}x{dh:.0f}"
    if abs(dw - sw) > 1: geo += "  <-- WIDTH CHANGED"
    print(f"{label:26} {annots(s):>5}->{annots(d):<6} {links(s):>4}->{links(d):<5} "
          f"{outline(s):>4}->{outline(d):<5} {str(acroform(s))[0]}->{str(acroform(d))[0]}   {geo}")
