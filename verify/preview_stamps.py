import json, shutil, sys
from datetime import date
from pathlib import Path
sys.path.insert(0, ".")
from docrefine.rebrand import BrandKit, rebrand_pdf
from docrefine.worker import Worker
from docrefine.processing import POPPLER_BIN, convert_from_path

SP = Path(sys.argv[1])
BK = Path(r"C:\Users\WORK\Documents\Batch 4\Template")
SAMPLES = Path(r"C:\Users\WORK\Documents\WebLife Labs\PROJECTS\DocRefine Pro\BrandKit\Sample Files")

kit_dir = SP / "kit_preview"
if kit_dir.exists(): shutil.rmtree(kit_dir)
shutil.copytree(BK, kit_dir)
(kit_dir / "brand.json").write_text(json.dumps({
    "name": "Budget Mailboxes", "slug": "budget-mailboxes",
    "tagline": "Trusted by the Nation",
    "disclaimer": "This document is provided for reference only. Product specifications "
                  "and dimensions are subject to change without notice. Refer to the "
                  "manufacturer's instructions for installation and safety requirements.",
}, indent=2), encoding="utf-8")
kit = BrandKit(kit_dir)

ALL_ON = {"footer_attribution": True, "stamp_tagline": True,
          "stamp_version": True, "stamp_disclaimer": True}
st = Worker._stamps_for(kit, "Florence Corporation", ALL_ON, date.today())
out = SP / "preview_stamped.pdf"
info = rebrand_pdf(SAMPLES / "1570_FCBU_Installation_Instructions.pdf", out, kit,
                   "Installation Manual", stamps=st)
print("rebranded:", info)

pages = convert_from_path(str(out), dpi=110, first_page=2, last_page=2,
                          poppler_path=POPPLER_BIN)
im = pages[0]
print("page px:", im.size)
im.crop((0, int(im.height * 0.60), im.width, im.height)).save(SP / "view_stamp_band.png")
im.resize((im.width // 2, im.height // 2)).save(SP / "view_full_page.png")
print("wrote previews")
