import sys
import shutil
import gc
import os
import zipfile
import re
import time
from pathlib import Path
from PIL import Image, ImageFile

# Configure Pillow limits
Image.MAX_IMAGE_PIXELS = 500000000
ImageFile.LOAD_TRUNCATED_IMAGES = True

# Dependency Checks
try:
    from pdf2image import convert_from_path, pdfinfo_from_path
    import pypdf
    import pytesseract
except ImportError:
    pass  # Dependencies handled in main app check or via requirements

from .config import CFG, SystemUtils, log_app

# ==============================================================================
#   BINARY DETECTION (Poppler/Tesseract)
# ==============================================================================
bin_ext = ".exe" if SystemUtils.IS_WIN else ""
poppler_bin_file = SystemUtils.find_binary("pdfinfo" + bin_ext)
POPPLER_BIN = str(Path(poppler_bin_file).parent) if poppler_bin_file else None

tesseract_bin_file = SystemUtils.find_binary("tesseract" + bin_ext)
HAS_TESSERACT = bool(tesseract_bin_file)

if HAS_TESSERACT:
    pytesseract.pytesseract.tesseract_cmd = tesseract_bin_file
    if getattr(sys, 'frozen', False) and SystemUtils.IS_MAC:
        tessdata_path = SystemUtils.get_resource_dir() / "tessdata"
        if tessdata_path.exists():
            os.environ["TESSDATA_PREFIX"] = str(tessdata_path)

def parse_lang_code(selection):
    if "(" in selection and ")" in selection:
        return selection.split("(")[1].replace(")", "")
    return selection

# ==============================================================================
#   PROCESSORS
# ==============================================================================
class BaseProcessor:
    def __init__(self, p_func, s_check, p_event): 
        self.progress = p_func; self.stop_sig_func = s_check; self.pause_event = p_event 
    def check_state(self):
        if self.stop_sig_func(): raise Exception("Stopped")
        if not self.pause_event.is_set():
            self.progress(None, "Paused...", status_only=True)
            self.pause_event.wait() 
            if self.stop_sig_func(): raise Exception("Stopped")

class PdfProcessor(BaseProcessor):
    def flatten_or_ocr(self, src, dest, mode='flatten', dpi=300):
        temp = dest.parent / f"temp_{src.stem}"; temp.mkdir(parents=True, exist_ok=True)
        try:
            info = pdfinfo_from_path(str(src), poppler_path=POPPLER_BIN)
            pages = info.get("Pages", 1)
            imgs = []
            
            ocr_lang = parse_lang_code(CFG.get("ocr_lang"))

            for i in range(1, pages + 1):
                self.check_state() 
                if i % 5 == 0 or i == pages: 
                     self.progress((i/pages)*100, f"Page {i}/{pages}")
                gc.collect() 
                res = convert_from_path(str(src), dpi=dpi, first_page=i, last_page=i, poppler_path=POPPLER_BIN)
                if not res: continue
                img = res[0]
                if mode == 'ocr' and HAS_TESSERACT:
                    t_page = temp / f"page_{i}.jpg"; img.save(t_page, "JPEG", dpi=(int(dpi), int(dpi)))
                    f = temp / f"{i}.pdf"
                    with open(f, "wb") as o: o.write(pytesseract.image_to_pdf_or_hocr(str(t_page), extension='pdf', lang=ocr_lang))
                    imgs.append(str(f))
                else:
                    f = temp / f"{i}.jpg"; img.convert('RGB').save(f, "JPEG", quality=85); imgs.append(str(f))
                del res; del img
            self.check_state(); self.progress(100, "Merging...")
            if mode == 'ocr' and HAS_TESSERACT:
                m = pypdf.PdfWriter(); 
                for f in imgs: m.append(f)
                m.write(dest); m.close()
            else:
                base = Image.open(imgs[0]).convert('RGB')
                base.save(dest, "PDF", resolution=float(dpi), save_all=True, append_images=[Image.open(f).convert('RGB') for f in imgs[1:]])
            return True
        except Exception as e: 
            if str(e) == "Stopped": raise
            return False
        finally: shutil.rmtree(temp, ignore_errors=True); gc.collect()

class ImageProcessor(BaseProcessor):
    def resize(self, src, dest, w):
        try:
            self.check_state(); self.progress(50, "Processing Image...")
            with Image.open(src) as img:
                img.load(); r = min(w / img.width, 1.0)
                img.resize((int(img.width * r), int(img.height * r)), Image.Resampling.LANCZOS).convert('RGB').save(dest, "JPEG", quality=85)
            return True
        except Exception as e:
            if str(e) == "Stopped": raise
            return False
    def convert_to_pdf(self, src, dest):
        try:
            self.check_state(); self.progress(50, "Converting...")
            with Image.open(src) as img: img.load(); img.convert('RGB').save(dest, "PDF")
            return True
        except Exception as e:
            if str(e) == "Stopped": raise
            return False

class OfficeProcessor(BaseProcessor):
    def sanitize(self, src, dest):
        try:
            self.check_state()
            if src.suffix.lower() not in {'.docx', '.xlsx'}: shutil.copy2(src, dest); return False
            if not zipfile.is_zipfile(src): raise Exception("Corrupt File")
            self.progress(50, "Sanitizing...")
            t = dest.parent / f"temp_{src.stem}"; shutil.rmtree(t, ignore_errors=True)
            with zipfile.ZipFile(src) as z: z.extractall(t)
            c = t / "docProps" / "core.xml"
            if c.exists(): c.write_text(re.sub(r'(<dc:creator>).*?(</dc:creator>)', r'\1\2', c.read_text(), flags=re.DOTALL))
            with zipfile.ZipFile(dest, 'w') as z:
                for r, _, fs in os.walk(t):
                    for f in fs: z.write(Path(r)/f, (Path(r)/f).relative_to(t))
            shutil.rmtree(t)
            return True
        except Exception as e:
            if str(e) == "Stopped": raise
            shutil.copy2(src, dest); return False