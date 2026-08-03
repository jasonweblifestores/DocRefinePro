import sys
import shutil
import gc
import os
import tempfile
import zipfile
import re
import time
from pathlib import Path
from PIL import Image, ImageFile

from .config import CFG as _CFG_FOR_LIMITS

# Configure Pillow limits — honour the Settings value rather than a hardcoded one.
try:
    Image.MAX_IMAGE_PIXELS = int(_CFG_FOR_LIMITS.get("max_pixels")) or None
except (TypeError, ValueError):
    Image.MAX_IMAGE_PIXELS = 500000000
ImageFile.LOAD_TRUNCATED_IMAGES = True


class _staged:
    """Write to `dest.part`, swap it in on success, remove it on failure.

    The refine engine resumes by treating any non-empty file at the destination
    as finished, so a half-written output must never appear there.
    """
    def __init__(self, dest):
        self.dest = Path(dest)
        self.path = self.dest.with_suffix(self.dest.suffix + ".part")

    def __enter__(self):
        self.dest.parent.mkdir(parents=True, exist_ok=True)
        return self.path

    def __exit__(self, exc_type, exc, tb):
        if exc_type is None:
            os.replace(self.path, self.dest)
        elif self.path.exists():
            try: os.remove(self.path)
            except OSError: pass
        return False

# Dependency Checks
try:
    from pdf2image import convert_from_path, pdfinfo_from_path
    import pypdf
    import pytesseract
except ImportError:
    pass

from .config import CFG, SystemUtils, log_app, get_hidden_startupinfo

class JobCancelledException(Exception):
    pass

# ==============================================================================
#   BINARY DETECTION
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
        if self.stop_sig_func(): raise JobCancelledException()
        if not self.pause_event.is_set():
            self.progress(None, "Paused...", status_only=True)
            self.pause_event.wait() 
            if self.stop_sig_func(): raise JobCancelledException()

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
                
                # UPDATE: Report EVERY page. 
                # The worker.py throttler will ensure the UI doesn't freeze.
                self.progress((i/pages)*100, f"Page {i}/{pages}")
                
                res = convert_from_path(str(src), dpi=dpi, first_page=i, last_page=i, poppler_path=POPPLER_BIN)
                if not res: continue
                img = res[0]
                if mode == 'ocr' and HAS_TESSERACT:
                    t_page = temp / f"page_{i}.jpg"; img.save(t_page, "JPEG", dpi=(int(dpi), int(dpi)))
                    f = temp / f"{i}.pdf"
                    with open(f, "wb") as o: o.write(pytesseract.image_to_pdf_or_hocr(str(t_page), extension='pdf', lang=ocr_lang))
                    imgs.append(str(f))
                else:
                    # Flatten: render each page to its own single-page PDF so the
                    # final merge streams from disk instead of loading every page
                    # into memory at once (prevents OOM on large multi-page PDFs).
                    f = temp / f"{i}.pdf"
                    img.convert('RGB').save(f, "PDF", resolution=float(dpi))
                    imgs.append(str(f))
                del res; del img

            self.check_state(); self.progress(100, "Merging...")

            if not imgs:
                return False

            # Both modes merge page PDFs incrementally, so peak memory stays
            # bounded to a single page regardless of document length.
            # Merge into the temp folder and swap in: a run killed mid-write must
            # not leave a truncated PDF, because the resume check treats any
            # non-empty file at the destination as finished work.
            merged = temp / "_merged.pdf"
            m = pypdf.PdfWriter()
            for f in imgs: m.append(f)
            m.write(str(merged)); m.close()
            dest.parent.mkdir(parents=True, exist_ok=True)
            os.replace(merged, dest)
            return True
        except JobCancelledException:
            raise
        except Exception as e: 
            log_app(f"PDF Processor Error: {e}", "ERROR")
            return False
        finally: 
            shutil.rmtree(temp, ignore_errors=True)
            gc.collect()

class ImageProcessor(BaseProcessor):
    def resize(self, src, dest, w):
        try:
            self.check_state(); self.progress(50, "Processing...")
            with _staged(dest) as staged:
                with Image.open(src) as img:
                    img.load(); r = min(w / img.width, 1.0)
                    img.resize((int(img.width * r), int(img.height * r)), Image.Resampling.LANCZOS).convert('RGB').save(staged, "JPEG", quality=85)
            return True
        except JobCancelledException:
            raise
        except Exception as e:
            log_app(f"Image Resize Error: {e}", "ERROR")
            return False
    def convert_to_pdf(self, src, dest):
        try:
            self.check_state(); self.progress(50, "Converting...")
            with _staged(dest) as staged:
                with Image.open(src) as img: img.load(); img.convert('RGB').save(staged, "PDF")
            return True
        except JobCancelledException:
            raise
        except Exception as e:
            log_app(f"Image PDF Convert Error: {e}", "ERROR")
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
            if c.exists():
                # Office XML is UTF-8; reading it with the machine's locale
                # encoding either mangles non-ASCII metadata or throws and
                # silently skips sanitising the file altogether.
                xml = c.read_text(encoding="utf-8", errors="surrogatepass")
                c.write_text(re.sub(r'(<dc:creator>).*?(</dc:creator>)', r'\1\2', xml, flags=re.DOTALL),
                             encoding="utf-8", errors="surrogatepass")
            # ZipFile defaults to ZIP_STORED — re-zipping without DEFLATE inflated
            # every sanitised Office file (measured 225x on compressible XML).
            with _staged(dest) as staged:
                with zipfile.ZipFile(staged, 'w', zipfile.ZIP_DEFLATED) as z:
                    for r, _, fs in os.walk(t):
                        for f in fs: z.write(Path(r)/f, (Path(r)/f).relative_to(t))
            shutil.rmtree(t, ignore_errors=True)
            return True
        except JobCancelledException:
            raise
        except Exception as e:
            log_app(f"Office Sanitize Error: {e}", "ERROR")
            shutil.copy2(src, dest); return False