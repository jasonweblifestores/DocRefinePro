import sys
import os
import subprocess
import queue
import time
import shutil
import hashlib
import json
import traceback
import threading
import zipfile
import re
import platform
import glob
import gc
import uuid
import logging
import urllib.request
import webbrowser
import csv
import random 
import ssl 
import concurrent.futures
from logging.handlers import RotatingFileHandler
from pathlib import Path
from datetime import datetime, timedelta
from tkinter import filedialog, scrolledtext, messagebox, ttk, Menu
import tkinter as tk
from PIL import Image, ImageFile, ImageTk

# ==============================================================================
#   WINDOWS GHOST WINDOW FIX
# ==============================================================================
if os.name == 'nt':
    try:
        _original_popen = subprocess.Popen
        def safe_popen(*args, **kwargs):
            if 'startupinfo' not in kwargs:
                si = subprocess.STARTUPINFO()
                si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                si.wShowWindow = subprocess.SW_HIDE
                kwargs['startupinfo'] = si
                if 'creationflags' not in kwargs:
                    kwargs['creationflags'] = 0x08000000 
            return _original_popen(*args, **kwargs)
        subprocess.Popen = safe_popen
    except Exception as e: print(f"Warning: Could not patch subprocess: {e}")

# Global defaults
Image.MAX_IMAGE_PIXELS = 500000000 
ImageFile.LOAD_TRUNCATED_IMAGES = True 
try: import psutil; HAS_PSUTIL = True
except ImportError: HAS_PSUTIL = False

# ==============================================================================
#   DOCREFINE PRO v112
# ==============================================================================

# --- 1. SYSTEM ABSTRACTION & CONFIG ---
class SystemUtils:
    IS_WIN = platform.system() == 'Windows'
    IS_MAC = platform.system() == 'Darwin'
    CURRENT_VERSION = "v112"
    UPDATE_MANIFEST_URL = "https://gist.githubusercontent.com/jasonweblifestores/53752cda3c39550673fc5dafb96c4bed/raw/docrefine_version.json"

    @staticmethod
    def get_resource_dir():
        if getattr(sys, 'frozen', False): return Path(sys._MEIPASS)
        return Path(__file__).parent

    @staticmethod
    def get_user_data_dir():
        if SystemUtils.IS_MAC or SystemUtils.IS_WIN:
            p = Path.home() / "Documents" / "DocRefinePro_Data"
            p.mkdir(parents=True, exist_ok=True)
            return p
        if getattr(sys, 'frozen', False): return Path(sys.executable).parent
        return Path(__file__).parent

    @staticmethod
    def open_file(path):
        p = str(path)
        try:
            if not Path(p).exists(): return
            if SystemUtils.IS_WIN: os.startfile(p)
            elif SystemUtils.IS_MAC: subprocess.call(['open', p])
            else: subprocess.call(['xdg-open', p])
        except Exception as e: print(f"Error opening file: {e}")

    @staticmethod
    def reveal_file(path):
        p = str(Path(path).resolve())
        try:
            if not Path(p).exists(): return
            if SystemUtils.IS_WIN:
                subprocess.Popen(f'explorer /select,"{p}"')
            elif SystemUtils.IS_MAC:
                subprocess.call(["open", "-R", p])
            else:
                subprocess.call(['xdg-open', str(Path(p).parent)])
        except Exception as e:
            print(f"Error revealing file: {e}")
            SystemUtils.open_file(Path(p).parent)

    @staticmethod
    def find_binary(bin_name):
        res_dir = SystemUtils.get_resource_dir()
        if (res_dir / bin_name).exists(): return str(res_dir / bin_name)
        if (res_dir / "bin" / bin_name).exists(): return str(res_dir / "bin" / bin_name)
        
        portable_target = res_dir / "DocRefine_Portable"
        if portable_target.exists():
             if (portable_target / bin_name).exists(): return str(portable_target)
             if (portable_target / "bin" / bin_name).exists(): return str(portable_target / "bin")

        sys_path = shutil.which(bin_name)
        if sys_path: return str(Path(sys_path).resolve())

        if SystemUtils.IS_MAC:
            for loc in ["/opt/homebrew/bin", "/usr/local/bin"]:
                brew_path = Path(loc) / bin_name
                if brew_path.exists(): return str(brew_path)
        return None

class Config:
    GITHUB_REPO = "jasonweblifestores/DocRefinePro" 
    DEFAULTS = { 
        "ram_warning_mb": 1024, 
        "resize_width": 1920, 
        "log_level": "INFO",
        "max_pixels": 500000000,
        "max_threads": 0, 
        "default_export_prio": "Auto (Best Available)",
        "default_ingest_mode": "Standard", 
        "ocr_lang": "eng",
        "last_workspace": "",
        "last_geometry": "1024x700",
        "last_tab": 0
    }
    
    def __init__(self):
        self.data = self.DEFAULTS.copy()
        self.path = SystemUtils.get_user_data_dir() / "config.json"
        if self.path.exists():
            try:
                with open(self.path, 'r') as f: self.data.update(json.load(f))
            except: pass
        try: Image.MAX_IMAGE_PIXELS = int(self.data.get("max_pixels", 500000000))
        except: pass

    def get(self, key): return self.data.get(key, self.DEFAULTS.get(key))
    def set(self, key, val): self.data[key] = val; self.save()
    def reset(self): self.data = self.DEFAULTS.copy(); self.save()
    def save(self):
        try:
            with open(self.path, 'w') as f: json.dump(self.data, f, indent=4)
        except Exception as e: print(f"Config Save Error: {e}")

CFG = Config()

# --- 2. LOGGING ---
USER_DIR = SystemUtils.get_user_data_dir()
LOG_PATH = USER_DIR / "app_debug.log"
JSON_LOG_PATH = USER_DIR / "app_events.jsonl"
WORKSPACES_ROOT = USER_DIR / "Workspaces"
WORKSPACES_ROOT.mkdir(parents=True, exist_ok=True)

SUPPORTED_EXTENSIONS = {'.pdf', '.doc', '.docx', '.jpg', '.png', '.xls', '.xlsx', '.csv', '.jpeg'}

logger = logging.getLogger("DocRefine")
logger.setLevel(getattr(logging, CFG.get("log_level").upper(), logging.INFO))
c_handler = logging.StreamHandler()
c_handler.setFormatter(logging.Formatter('[%(levelname)s] %(message)s'))
logger.addHandler(c_handler)

try:
    f_handler = RotatingFileHandler(LOG_PATH, maxBytes=1024*1024, backupCount=5, encoding='utf-8')
    f_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
    logger.addHandler(f_handler)
except: pass

def log_app(msg, level="INFO", structured_data=None):
    if level == "ERROR": logger.error(msg)
    elif level == "WARN": logger.warning(msg)
    else: logger.info(msg)
    try:
        entry = {
            "timestamp": datetime.now().isoformat(),
            "level": level,
            "message": msg,
            "os": platform.system(),
            "version": SystemUtils.CURRENT_VERSION
        }
        if structured_data: entry.update(structured_data)
        with open(JSON_LOG_PATH, "a", encoding="utf-8") as f:
            f.write(json.dumps(entry) + "\n")
    except: pass

log_app(f"=== STARTUP {SystemUtils.CURRENT_VERSION} ===", structured_data={"event": "startup", "path": str(USER_DIR)})

# --- 3. STARTUP HYGIENE ---
def clean_temp_files():
    try:
        limit = time.time() - 86400
        for ws in WORKSPACES_ROOT.iterdir():
            if ws.is_dir():
                for item in ws.glob("temp_*"):
                    if item.is_dir() and item.stat().st_mtime < limit:
                        shutil.rmtree(item, ignore_errors=True)
    except: pass
clean_temp_files()

# --- 4. DEPENDENCIES ---
bin_ext = ".exe" if SystemUtils.IS_WIN else ""
poppler_bin_file = SystemUtils.find_binary("pdfinfo" + bin_ext)
POPPLER_BIN = str(Path(poppler_bin_file).parent) if poppler_bin_file else None
if not POPPLER_BIN: log_app("CRITICAL: Poppler not found.", "ERROR")

tesseract_bin_file = SystemUtils.find_binary("tesseract" + bin_ext)
HAS_TESSERACT = bool(tesseract_bin_file)

if HAS_TESSERACT:
    import pytesseract
    pytesseract.pytesseract.tesseract_cmd = tesseract_bin_file
    if getattr(sys, 'frozen', False) and SystemUtils.IS_MAC:
        tessdata_path = SystemUtils.get_resource_dir() / "tessdata"
        if tessdata_path.exists():
            os.environ["TESSDATA_PREFIX"] = str(tessdata_path)

def get_tesseract_langs():
    if not HAS_TESSERACT: return ["N/A"]
    try:
        raw_langs = pytesseract.get_languages(config='')
        friendly_map = {
            'eng': 'English', 'spa': 'Spanish', 'fra': 'French', 'deu': 'German',
            'ita': 'Italian', 'por': 'Portuguese', 'chi_sim': 'Chinese (Simp)',
            'jpn': 'Japanese', 'rus': 'Russian'
        }
        clean = []
        for l in raw_langs:
            if l == 'osd': continue
            name = friendly_map.get(l, l)
            if name != l: clean.append(f"{name} ({l})")
            else: clean.append(l)
        return sorted(clean)
    except: return ["eng"]

def parse_lang_code(selection):
    if "(" in selection and ")" in selection:
        return selection.split("(")[1].replace(")", "")
    return selection

from pdf2image import convert_from_path, pdfinfo_from_path
import pypdf
from pypdf import PdfReader, PdfWriter

# --- 5. UTILS ---
def sanitize_filename(name): return re.sub(r'[<>:"/\\|?*]', '_', name)

def update_stats_time(ws, cat, sec):
    try:
        p = Path(ws)/"stats.json"
        if not p.exists(): return
        with open(p,'r') as f: s=json.load(f)
        s[cat] = s.get(cat,0.0)+sec
        with open(p,'w') as f: json.dump(s,f,indent=4)
    except: pass

def check_memory():
    if not HAS_PSUTIL: return True
    try:
        free_mb = psutil.virtual_memory().available / (1024 * 1024)
        if free_mb < CFG.get("ram_warning_mb"): return False
    except: pass
    return True

def generate_job_report(ws_path, action_name, file_results=None):
    try:
        ws = Path(ws_path)
        rpt_dir = ws / "04_Reports"
        rpt_dir.mkdir(parents=True, exist_ok=True)
        
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M')
        file_name = f"Audit_Certificate_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html"
        
        s = {}
        if (ws/"stats.json").exists():
            with open(ws/"stats.json") as f: s = json.load(f)
        
        total_orig = 0
        total_new = 0
        errors = []
        
        if file_results:
            for res in file_results:
                total_orig += res.get('orig_size', 0)
                total_new += res.get('new_size', 0)
                if not res.get('ok', True):
                    errors.append(res)
        
        saved_bytes = total_orig - total_new
        saved_mb = round(saved_bytes / (1024*1024), 2)
        saved_pct = round((saved_bytes / total_orig * 100), 1) if total_orig > 0 else 0

        error_rows = ""
        if errors:
            rows = []
            for e in errors:
                fname = e.get('file', '?')
                err_msg = e.get('error', 'Unknown')
                rows.append(f"<tr class='error-row'><td>{fname}</td><td>FAILED</td><td>{err_msg}</td></tr>")
            error_rows = f"<table><thead><tr><th>File</th><th>Status</th><th>Error Details</th></tr></thead><tbody>{''.join(rows)}</tbody></table>"
        else:
            error_rows = "<p>No errors reported. Clean run.</p>"

        html = f"""
        <html>
        <head>
            <style>
                body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; padding: 40px; background: #f0f2f5; color: #333; }}
                .container {{ background: white; padding: 30px; border-radius: 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.05); max-width: 900px; margin: auto; }}
                .header {{ border-bottom: 2px solid #0078d7; padding-bottom: 20px; margin-bottom: 30px; display: flex; justify-content: space-between; align-items: center; }}
                .title h1 {{ margin: 0; color: #2c3e50; font-size: 24px; }}
                .title span {{ color: #7f8c8d; font-size: 14px; }}
                .badge {{ background: #0078d7; color: white; padding: 5px 10px; border-radius: 4px; font-weight: bold; font-size: 12px; }}
                
                .grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 20px; margin-bottom: 30px; }}
                .card {{ background: #f8f9fa; padding: 15px; border-radius: 8px; border: 1px solid #e9ecef; }}
                .card-label {{ font-size: 11px; color: #6c757d; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 5px; }}
                .card-value {{ font-size: 20px; font-weight: 600; color: #212529; }}
                .highlight {{ color: #28a745; }}
                
                table {{ width: 100%; border-collapse: collapse; margin-top: 20px; font-size: 13px; }}
                th {{ text-align: left; border-bottom: 2px solid #dee2e6; padding: 10px; color: #495057; }}
                td {{ border-bottom: 1px solid #dee2e6; padding: 10px; }}
                .error-row {{ background-color: #fff5f5; color: #c0392b; }}
                
                .footer {{ margin-top: 40px; font-size: 11px; color: #adb5bd; text-align: center; border-top: 1px solid #eee; padding-top: 20px; }}
            </style>
        </head>
        <body>
            <div class="container">
                <div class="header">
                    <div class="title">
                        <h1>Processing Audit Certificate</h1>
                        <span>DocRefine Pro {SystemUtils.CURRENT_VERSION}</span>
                    </div>
                    <span class="badge">COMPLETED</span>
                </div>
                
                <p><strong>Operation:</strong> {action_name}<br><strong>Timestamp:</strong> {timestamp}</p>
                
                <div class="grid">
                    <div class="card">
                        <div class="card-label">Files Processed</div>
                        <div class="card-value">{len(file_results) if file_results else s.get('total_scanned', 0)}</div>
                    </div>
                    <div class="card">
                        <div class="card-label">Storage Reclaimed</div>
                        <div class="card-value highlight">{saved_mb} MB ({saved_pct}%)</div>
                    </div>
                    <div class="card">
                        <div class="card-label">Failed / Skipped</div>
                        <div class="card-value" style="color: {'red' if errors else '#212529'}">{len(errors)}</div>
                    </div>
                    <div class="card">
                        <div class="card-label">Duration</div>
                        <div class="card-value">{str(timedelta(seconds=int(s.get('batch_time', 0) + s.get('ingest_time', 0))))}</div>
                    </div>
                </div>

                <h3>Exceptions & Errors</h3>
                {error_rows}
                
                <div class="footer">
                    This document certifies that the files listed above were processed by the DocRefine Engine.<br>
                    Generated automatically on {timestamp}
                </div>
            </div>
        </body>
        </html>
        """
        
        with open(rpt_dir / file_name, "w", encoding="utf-8") as f:
            f.write(html)
        return str(rpt_dir / file_name)
    except Exception as e:
        print(f"Report Gen Error: {e}")
        return None

# --- 6. PROCESSORS ---
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

# --- 7. WORKER ---
class Worker:
    def __init__(self, q): 
        self.q = q; self.stop_sig = False; self.pause_event = threading.Event(); self.pause_event.set(); self.current_ws = None 
    def stop(self): self.stop_sig = True; self.pause_event.set()
    def pause(self): self.pause_event.clear()
    def resume(self): self.pause_event.set()

    def log(self, m, err=False):
        self.q.put(("log", m, err))
        log_app(m, "ERROR" if err else "INFO", structured_data={"ws": self.current_ws})

    def set_job_status(self, ws, stage, details=""):
        try:
            data = { "stage": stage, "last_update": datetime.now().strftime('%Y-%m-%d %H:%M:%S'), "details": details }
            with open(Path(ws) / "status.json", 'w') as f: json.dump(data, f, indent=4)
        except: pass

    def prog_main(self, v, t): self.q.put(("main_p", v, t))
    
    def prog_sub(self, v, t, status_only=False): 
        tid = threading.get_ident()
        self.q.put(("slot_update", tid, t, v))

    def get_hash(self, path, mode):
        if os.path.getsize(path) == 0: return None, "Zero-Byte File"
        if path.suffix.lower() == '.pdf' and mode != "Lightning":
            try:
                r = PdfReader(str(path), strict=False) 
                if len(r.pages) == 0: return None, "PDF has 0 Pages"
                if mode == "Standard":
                    txt = "".join([r.pages[i].extract_text() for i in range(min(3, len(r.pages)))])
                    if len(txt.strip()) > 10: return hashlib.md5(f"{txt}{len(r.pages)}".encode()).hexdigest(), "Smart-Standard"
                elif mode == "Deep":
                    txt = "".join([p.extract_text() for p in r.pages])
                    if len(txt.strip()) > 10: return hashlib.md5(f"{txt}{len(r.pages)}".encode()).hexdigest(), "Smart-Deep"
            except: pass 
        try:
            h = hashlib.md5()
            with open(path, 'rb') as f:
                for chunk in iter(lambda: f.read(65536), b""): h.update(chunk)
            return h.hexdigest(), "Binary"
        except Exception as e: return None, f"Read-Error: {str(e)[:20]}"

    def get_best_source(self, ws, file_uid, priority_mode="Auto (Best Available)"):
        master = ws / "01_Master_Files" / file_uid
        base_cache = ws / "02_Ready_For_Redistribution"
        
        def find_in_dir(d, stem):
            if d.exists():
                if (d / file_uid).exists(): return d / file_uid
                match = next((f for f in d.iterdir() if f.stem == stem), None)
                if match: return match
            return None

        stem = Path(file_uid).stem
        
        if "Force: OCR" in priority_mode:
            f = find_in_dir(base_cache/"OCR", stem)
            return f if f else master
            
        elif "Force: Flattened" in priority_mode:
            f = find_in_dir(base_cache/"Flattened", stem)
            return f if f else master
            
        elif "Force: Original" in priority_mode:
            return master
            
        else: 
            for sub in ["OCR", "Flattened", "Resized", "Sanitized", "Standard"]:
                f = find_in_dir(base_cache/sub, stem)
                if f: return f
            return master if master.exists() else None

    # v104: Restored Single Threaded Ingest
    def run_inventory(self, d_str, ingest_mode):
        try:
            # v110: Fix "Dead Worker" bug by resetting stop signal
            self.stop_sig = False
            self.resume()
            
            d = Path(d_str); start_time = time.time()
            ws = WORKSPACES_ROOT / f"{d.name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            m_dir = ws / "01_Master_Files"; m_dir.mkdir(parents=True); (ws/"00_Quarantine").mkdir()
            self.current_ws = str(ws); self.log(f"Inventory Start: {d}")
            self.q.put(("job", str(ws))); self.q.put(("ws", str(ws)))
            self.set_job_status(ws, "SCANNING", "Ingesting...")
            
            files = [Path(r)/f for r,_,fs in os.walk(d) for f in fs]
            files = [f for f in files if f.suffix.lower() in SUPPORTED_EXTENSIONS]
            
            seen = {}; quarantined = 0
            self.q.put(("slot_config", 1))

            for i, f in enumerate(files):
                if self.stop_sig: break
                if not self.pause_event.is_set(): self.prog_sub(None, "Paused...", True); self.pause_event.wait()
                
                self.prog_main((i/len(files))*100, f"Scanning {i}/{len(files)}")
                self.prog_sub(None, f"Hashing: {f.name}", True)
                
                try:
                    h, method = self.get_hash(f, ingest_mode)
                    if not h: 
                        self.log(f"⚠️ Quarantine: {f.name}", True)
                        shutil.copy2(f, ws/"00_Quarantine"/f"{uuid.uuid4()}_{sanitize_filename(f.name)}")
                        quarantined += 1; continue
                    
                    rel = str(f.relative_to(d))
                    if h in seen: seen[h]['copies'].append(rel)
                    else: seen[h] = {'master': rel, 'copies': [rel], 'name': f.name, 'root': str(d)}
                except Exception as e:
                    self.log(f"Hash Error: {e}", True)

            # Check stop before finalizing
            if self.stop_sig: 
                self.log("Ingest Stopped by User.")
                self.q.put(("done",))
                return

            self.log("Tagging..."); total = len(seen)
            for i, (h, data) in enumerate(seen.items()):
                if self.stop_sig: break
                safe_name = f"[{i+1:04d}]_{sanitize_filename(data['name'])}"
                shutil.copy2(d / data['master'], m_dir / safe_name)
                data['uid'] = safe_name; data['id'] = f"[{i+1:04d}]"
            
            if self.stop_sig: return

            stats = {
                "ingest_time": time.time()-start_time, 
                "masters": total, 
                "quarantined": quarantined,
                "total_scanned": len(files)
            }
            with open(ws/"manifest.json", 'w') as f: json.dump(seen, f, indent=4)
            with open(ws/"stats.json", 'w') as f: json.dump(stats, f)
            self.set_job_status(ws, "INGESTED", f"Masters: {total}")
            self.log(f"Done. Masters: {total}"); self.q.put(("job", str(ws))); self.q.put(("done",))
        except Exception as e: self.log(f"Error: {e}", True); self.q.put(("done",))

    def process_file_task(self, f, bots, options, base_dst):
        if self.stop_sig: return None
        result = {'file': f.name, 'orig_size': f.stat().st_size, 'new_size': 0, 'ok': False}
        try:
            self.q.put(("status_blue", f"Refining: {f.name}"))
            ext = f.suffix.lower()
            ok = False
            dpi_val = int(options.get('dpi', 300))
            
            target_folder = "Standard"
            
            if ext == '.pdf':
                mode = options.get('pdf_mode', 'none')
                if mode == 'flatten': target_folder = "Flattened"
                elif mode == 'ocr': target_folder = "OCR"
            elif ext in {'.jpg','.png'}:
                if options.get('resize'): target_folder = "Resized"
                if options.get('img2pdf'): target_folder = "Resized" 
            elif ext in {'.docx','.xlsx'}:
                if options.get('sanitize'): target_folder = "Sanitized"
            
            final_dst_dir = base_dst / target_folder
            final_dst_dir.mkdir(parents=True, exist_ok=True)
            
            dst_file = final_dst_dir / f.name

            if ext == '.pdf':
                mode = options.get('pdf_mode', 'none')
                if mode == 'flatten': ok = bots['pdf'].flatten_or_ocr(f, dst_file, 'flatten', dpi=dpi_val)
                elif mode == 'ocr': ok = bots['pdf'].flatten_or_ocr(f, dst_file, 'ocr', dpi=dpi_val)
            elif ext in {'.jpg','.png'}:
                if options.get('resize'): ok = bots['img'].resize(f, dst_file, CFG.get('resize_width'))
                if options.get('img2pdf'): ok = bots['img'].convert_to_pdf(f, final_dst_dir/f"{f.stem}.pdf")
            elif ext in {'.docx','.xlsx'}:
                if options.get('sanitize'): ok = bots['office'].sanitize(f, dst_file)

            if not ok and not dst_file.exists(): 
                 shutil.copy2(f, dst_file)
            
            if dst_file.exists():
                result['new_size'] = dst_file.stat().st_size
                result['ok'] = True
            
            return result
                 
        except Exception as e:
            self.log(f"Err {f.name}: {e}", True)
            result['error'] = str(e)
            return result

    def run_batch(self, ws_p, options):
        try:
            # v110: Reset stop signal
            self.stop_sig = False
            self.resume()
            
            ws = Path(ws_p); self.current_ws = str(ws)
            start_time = time.time(); src = ws/"01_Master_Files"; dst = ws/"02_Ready_For_Redistribution"; dst.mkdir(exist_ok=True)
            self.log(f"Refinement Start. Opts: {options}")
            self.set_job_status(ws, "PROCESSING", "Refining...")

            bots = {
                'pdf': PdfProcessor(lambda v,t,s=False: self.prog_sub(v,t,s), lambda: self.stop_sig, self.pause_event),
                'img': ImageProcessor(lambda v,t,s=False: self.prog_sub(v,t,s), lambda: self.stop_sig, self.pause_event),
                'office': OfficeProcessor(lambda v,t,s=False: self.prog_sub(v,t,s), lambda: self.stop_sig, self.pause_event)
            }
            fs = list(src.iterdir())
            
            forced_workers = int(CFG.get("max_threads"))
            if forced_workers > 0:
                max_workers = forced_workers
                self.log(f"Manual Worker Override: {max_workers}")
            else:
                max_workers = 2
                if HAS_PSUTIL:
                    try:
                        total_ram_gb = psutil.virtual_memory().total / (1024 ** 3)
                        if total_ram_gb < 8: max_workers = 1
                        elif total_ram_gb < 16: max_workers = 2
                        else: max_workers = 4
                    except: pass
                
                max_workers = min(max_workers, os.cpu_count() or 1)
                max_workers = max(1, max_workers)
                self.log(f"Auto-Throttled Workers: {max_workers}")

            self.q.put(("slot_config", max_workers))
            
            file_results = []
            with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
                futures = {executor.submit(self.process_file_task, f, bots, options, dst): f for f in fs}
                for i, future in enumerate(concurrent.futures.as_completed(futures)):
                    if self.stop_sig: break
                    self.prog_main((i/len(fs))*100, f"Refining {i+1}/{len(fs)}")
                    try: 
                        r = future.result()
                        if r: file_results.append(r)
                    except Exception as e: self.log(f"Thread Err: {e}", True)

            # v110: "Fake Completion" Fix - Abort if stopped
            if self.stop_sig: 
                self.log("Batch Stopped by User.")
                self.q.put(("done",))
                return

            update_stats_time(ws, "batch_time", time.time() - start_time)
            self.set_job_status(ws, "PROCESSED", "Complete")
            
            rpt = generate_job_report(ws, "Content Refinement Batch", file_results)
            if rpt: self.log(f"Receipt Generated: {Path(rpt).name}")
            
            self.q.put(("job", str(ws))) 
            self.prog_main(100, "Done"); self.q.put(("done",)); SystemUtils.open_file(dst)
        except Exception as e: self.log(f"Err: {e}", True); self.q.put(("done",))

    def run_organize(self, ws_p, priority_mode):
        try:
            self.stop_sig = False; self.resume()
            ws = Path(ws_p); self.current_ws = str(ws)
            start_time = time.time()
            out = ws / "03_Organized_Output"; m = out/"Unique_Masters"; q = out/"Quarantine"
            for p in [m,q]: p.mkdir(parents=True, exist_ok=True)
            
            self.log(f"Unique Export ({priority_mode})")
            with open(ws/"manifest.json") as f: man = json.load(f)
            total = len(man)
            
            self.q.put(("slot_config", 1))

            dup_csv = out / "duplicates_report.csv"
            with open(dup_csv, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(["Master_Filename", "Duplicate_Location"])
                
                for i, (h, data) in enumerate(man.items()):
                    if self.stop_sig: break
                    self.prog_main((i/total)*100, "Exporting Unique...")
                    self.q.put(("slot_update", threading.get_ident(), f"Exporting: {data['name']}", None))
                    
                    if data.get("status") == "QUARANTINE": 
                        for f in (ws/"00_Quarantine").glob("*"):
                            if data['orig_name'] in f.name: shutil.copy2(f, q/f.name)
                    else:
                        src = self.get_best_source(ws, data['uid'], priority_mode)
                        if src and src.exists():
                            clean_name = data['name']
                            if src.suffix != Path(clean_name).suffix:
                                clean_name = Path(clean_name).stem + src.suffix

                            tgt = m / clean_name
                            ctr = 1
                            while tgt.exists():
                                tgt = m / f"{Path(clean_name).stem}_{ctr}{Path(clean_name).suffix}"
                                ctr += 1
                            shutil.copy2(src, tgt)

                        if len(data.get('copies', [])) > 1:
                            for c in data['copies']:
                                if c != data.get('master'):
                                    writer.writerow([data['name'], c])

            if self.stop_sig: return

            update_stats_time(ws, "organize_time", time.time() - start_time)
            self.set_job_status(ws, "ORGANIZED", "Done")
            
            rpt = generate_job_report(ws, f"Unique Export ({priority_mode})")
            
            self.q.put(("job", str(ws))) 
            self.prog_main(100, "Done"); self.q.put(("done",)); SystemUtils.open_file(out)
        except Exception as e: self.log(f"Err: {e}", True); self.q.put(("done",))

    def run_distribute(self, ws_p, ext_src, priority_mode):
        try:
            self.stop_sig = False; self.resume()
            ws = Path(ws_p); self.current_ws = str(ws)
            if not (ws/"manifest.json").exists():
                 self.log("CRITICAL: Manifest missing.", True)
                 self.q.put(("error", "Manifest missing.")); self.q.put(("done",)); return

            start_time = time.time(); 
            dst = ws / "Final_Delivery"
            self.log(f"Reconstruction Start ({priority_mode})")
            self.set_job_status(ws, "DISTRIBUTING", "Reconstructing...")
            
            with open(ws/"manifest.json") as f: man = json.load(f)
            
            orphans = {}
            if ext_src:
                 orphans = {f.name: f for f in Path(ext_src).iterdir()}

            self.q.put(("slot_config", 1))

            for i, (h, d) in enumerate(man.items()):
                if self.stop_sig: break
                self.prog_main((i/len(man))*100, f"Recon {i+1}")
                self.q.put(("slot_update", threading.get_ident(), f"Copying: {d['name']}", None))
                
                if d.get("status") == "QUARANTINE": continue
                
                src = None
                if ext_src:
                    src = next((v for k,v in orphans.items() if k.startswith(d['id'])), None)
                else:
                    src = self.get_best_source(ws, d['uid'], priority_mode)
                
                if not src: continue
                
                for c in d['copies']:
                    t = dst / c; t.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(src, t.with_suffix(src.suffix))
            
            if self.stop_sig: return

            q_src = ws / "00_Quarantine"
            if q_src.exists():
                q_dst = dst / "_QUARANTINED_FILES"; 
                q_dst.mkdir(parents=True, exist_ok=True) 
                for qf in q_src.iterdir(): shutil.copy2(qf, q_dst / qf.name)

            update_stats_time(ws, "dist_time", time.time() - start_time)
            self.set_job_status(ws, "DISTRIBUTED", "Done")
            
            rpt = generate_job_report(ws, "Full Reconstruction")
            
            self.q.put(("job", str(ws))) 
            self.prog_main(100, "Done"); self.q.put(("done",)); SystemUtils.open_file(dst)
        except Exception as e: self.log(f"Err: {e}", True); self.q.put(("done",))

    def run_full_export(self, ws_p):
        try:
            self.stop_sig = False; self.resume()
            ws = Path(ws_p); self.current_ws = str(ws)
            if not (ws/"manifest.json").exists(): return

            rpt_dir = ws / "04_Reports"
            rpt_dir.mkdir(parents=True, exist_ok=True)
            csv_path = rpt_dir / "Full_Inventory_Manifest.csv"

            self.log("Generating Full Inventory CSV...")
            
            with open(ws/"manifest.json") as f: man = json.load(f)
            
            try:
                # utf-8-sig for Excel compatibility with special chars
                with open(csv_path, 'w', newline='', encoding='utf-8-sig') as csvfile:
                    writer = csv.writer(csvfile)
                    writer.writerow(["ID", "Status", "Original_Filename", "Original_Path_Structure", "Master_Location_In_Workplace", "Hash_Type", "Hash", "Copy_Count", "Error_Details"])
                    
                    total = len(man)
                    for i, (h, data) in enumerate(man.items()):
                        if self.stop_sig: break
                        self.prog_main((i/total)*100, "Writing CSV...")
                        
                        uid = data.get('id', '?')
                        status = data.get('status', 'OK')
                        name = data.get('name', '?')
                        master_rel = data.get('master', '')
                        
                        if status == "QUARANTINE":
                            orig = data.get('orig_name', name)
                            writer.writerow([uid, status, orig, "N/A - Quarantined", "00_Quarantine", "Binary", h, 0, data.get('error_reason', '')])
                        else:
                            copies = data.get('copies', [])
                            for copy_path in copies:
                                writer.writerow([
                                    uid, 
                                    status, 
                                    name, 
                                    copy_path, 
                                    master_rel, 
                                    "MD5", 
                                    h, 
                                    len(copies), 
                                    ""
                                ])
            except PermissionError:
                self.q.put(("error", "Could not write CSV.\nPlease close the file in Excel and try again."))
                self.q.put(("done",))
                return

            if self.stop_sig: return

            self.log(f"Exported: {csv_path.name}")
            self.q.put(("job", str(ws))) 
            self.prog_main(100, "Done"); self.q.put(("done",)); SystemUtils.open_file(rpt_dir)

        except Exception as e: self.log(f"Err: {e}", True); self.q.put(("done",))

    def run_preview(self, ws_p, dpi):
        try:
            self.stop_sig = False; self.resume()
            ws = Path(ws_p); self.current_ws = str(ws)
            src = ws/"01_Master_Files"; pdf = next(src.glob("*.pdf"), None)
            if not pdf: self.q.put(("preview_done",)); return
            for old in ws.glob("PREVIEW_*.pdf"): 
                try: os.remove(old)
                except: pass
            out = ws / f"PREVIEW_{int(time.time())}.pdf"
            imgs = convert_from_path(str(pdf), dpi=int(dpi), first_page=1, last_page=1, poppler_path=POPPLER_BIN)
            if imgs: 
                imgs[0].save(out, "PDF", resolution=float(dpi))
                SystemUtils.open_file(out)
            # v110: Send specific signal to avoid UI reset
            self.q.put(("preview_done",))
        except: self.q.put(("preview_done",))

    # v109: Threaded Export Logic
    def _export_debug_bundle_task(self):
        try:
            ts = datetime.now().strftime('%Y%m%d_%H%M%S')
            
            # v112: Mac Permissions Fallback
            base_dir = SystemUtils.get_user_data_dir()
            
            # Create a write test
            try:
                test_file = base_dir / "write_test.tmp"
                test_file.touch()
                test_file.unlink()
            except PermissionError:
                # Fallback to Downloads or Tmp
                if SystemUtils.IS_MAC:
                    base_dir = Path.home() / "Downloads"
                else:
                    base_dir = Path(os.getenv('TEMP', '/tmp'))
            
            dest_zip = base_dir / f"Debug_Bundle_{ts}.zip"
            temp_dir = base_dir / f"temp_debug_{ts}"
            temp_dir.mkdir(parents=True, exist_ok=True)
            
            def safe_copy(src, dst_name):
                try:
                    if not src or not Path(src).exists(): return
                    try:
                        shutil.copy2(src, temp_dir / dst_name)
                    except PermissionError:
                        with open(src, 'rb') as f_in:
                            content = f_in.read()
                        with open(temp_dir / dst_name, 'wb') as f_out:
                            f_out.write(content)
                except Exception as e:
                    with open(temp_dir / f"{dst_name}_ERROR.txt", 'w') as err_f:
                        err_f.write(str(e))

            # Core Logs
            safe_copy(LOG_PATH, "app_debug.log")
            safe_copy(JSON_LOG_PATH, "app_events.jsonl")
            safe_copy(CFG.path, "config.json")
            
            # Current WS
            ws = self.get_ws()
            if ws:
                safe_copy(ws/"session_log.txt", "current_job_log.txt")
                safe_copy(ws/"stats.json", "current_job_stats.json")
            
            shutil.make_archive(str(dest_zip).replace(".zip", ""), 'zip', temp_dir)
            shutil.rmtree(temp_dir, ignore_errors=True)
            
            self.q.put(("export_success", str(dest_zip)))
        except Exception as e:
            self.q.put(("error", f"Export Failed: {e}"))

    def start_debug_export_thread(self, btn_ref, win_ref):
        def _run():
            self._export_debug_bundle_task()
            self.q.put(("export_reset_btn", btn_ref))
        
        btn_ref.config(text="Exporting...", state="disabled")
        threading.Thread(target=_run, daemon=True).start()

# --- 8. UI ---
class ForensicComparator:
    def __init__(self, root, ws_path, manifest, master_path, dup_candidates):
        self.win = tk.Toplevel(root)
        self.win.title("Forensic Verification (Sync View)")
        self.win.geometry("1400x800")
        
        # v104: Fixed Center Logic - NO CLAMPING
        App.center_toplevel(self.win, root)
        
        self.ws_path = ws_path
        self.manifest = manifest
        self.master_path = master_path
        self.dups = dup_candidates
        self.dup_idx = 0
        self.page = 1
        self.total_pages = 1
        self.zoom = 1.0
        self.last_scroll_time = 0
        
        # UI Structure
        self.top = tk.Frame(self.win, bg="#eee", pady=5)
        self.top.pack(fill="x")
        
        # Added Filenames
        self.lbl_info = tk.Frame(self.win)
        self.lbl_info.pack(fill="x", pady=2)
        tk.Label(self.lbl_info, text=f"MASTER: {master_path.name}", font=("Consolas", 9, "bold"), fg="#2c3e50").pack(side="left", fill="x", expand=True)
        tk.Label(self.lbl_info, text="CANDIDATE (DUPLICATE)", font=("Consolas", 9, "bold"), fg="#c0392b").pack(side="left", fill="x", expand=True)

        self.mid = tk.Frame(self.win)
        self.mid.pack(fill="both", expand=True)
        
        self.c1 = tk.Canvas(self.mid, bg="#444", scrollregion=(0,0,1000,1000))
        self.c1.pack(side="left", fill="both", expand=True)
        
        self.sep = ttk.Separator(self.mid, orient="vertical")
        self.sep.pack(side="left", fill="y", padx=5)
        
        self.c2 = tk.Canvas(self.mid, bg="#444", scrollregion=(0,0,1000,1000))
        self.c2.pack(side="left", fill="both", expand=True)
        
        self._build_toolbar()
        
        # Init
        self.is_pdf = self.master_path.suffix.lower() == '.pdf'
        if self.is_pdf:
            try: self.total_pages = pdfinfo_from_path(str(self.master_path), poppler_path=POPPLER_BIN).get('Pages', 1)
            except: self.total_pages = 1
            
        self.load_images()
        self.bind_events()

    def _build_toolbar(self):
        # Page Nav
        f_page = tk.Frame(self.top); f_page.pack(side="left", padx=10)
        tk.Button(f_page, text="< Prev Page", command=self.prev_page).pack(side="left")
        self.lbl_page = tk.Label(f_page, text="Page 1/1", width=10)
        self.lbl_page.pack(side="left", padx=5)
        tk.Button(f_page, text="Next Page >", command=self.next_page).pack(side="left")
        
        # Zoom
        f_zoom = tk.Frame(self.top); f_zoom.pack(side="left", padx=20)
        tk.Button(f_zoom, text="- Zoom", command=lambda: self.do_zoom(0.8)).pack(side="left")
        self.lbl_zoom = tk.Label(f_zoom, text="100%", width=6)
        self.lbl_zoom.pack(side="left")
        tk.Button(f_zoom, text="Zoom +", command=lambda: self.do_zoom(1.2)).pack(side="left")
        # Fit Width
        tk.Button(f_zoom, text="[Fit Width]", command=self.fit_width).pack(side="left", padx=5)
        
        # Duplicate Nav
        f_dup = tk.Frame(self.top); f_dup.pack(side="right", padx=10)
        tk.Button(f_dup, text="< Prev Copy", command=self.prev_dup).pack(side="left")
        self.lbl_dup = tk.Label(f_dup, text="Copy 1/1", width=15)
        self.lbl_dup.pack(side="left", padx=5)
        tk.Button(f_dup, text="Next Copy >", command=self.next_dup).pack(side="left")
        # v110: Open Candidate Button
        tk.Button(f_dup, text="Open File", command=self.open_current_dup).pack(side="left", padx=5)
        
        # v112: Mac Safety Color
        kw_uniq = {"bg": "green", "fg": "white"} if not SystemUtils.IS_MAC else {}
        tk.Button(f_dup, text="MARK AS UNIQUE", command=self.mark_as_unique, **kw_uniq).pack(side="left", padx=10)

    def load_images(self):
        # Update Labels
        self.lbl_page.config(text=f"Page {self.page}/{self.total_pages}")
        self.lbl_zoom.config(text=f"{int(self.zoom*100)}%")
        self.lbl_dup.config(text=f"Copy {self.dup_idx+1}/{len(self.dups)}")
        
        if self.dups:
            # v110: Show Full Path
            path_str = str(self.dups[self.dup_idx])
            # Truncate if too long
            if len(path_str) > 60: path_str = "..." + path_str[-57:]
            
            tk.Label(self.lbl_info.winfo_children()[1], text=path_str).pack_forget() # Refresh info
            self.lbl_info.winfo_children()[1].config(text=f"CANDIDATE: {path_str}")

        # Render Master
        self.img1 = self._render(self.master_path)
        self.show_img(self.c1, self.img1)
        
        # Render Duplicate
        if self.dups:
            dup_path = self.dups[self.dup_idx]
            self.img2 = self._render(dup_path)
            self.show_img(self.c2, self.img2)
        else:
            self.c2.delete("all")

    def _render(self, path):
        try:
            if not Path(path).exists(): return None
            img = None
            if Path(path).suffix.lower() == '.pdf':
                imgs = convert_from_path(str(path), dpi=int(72*self.zoom), first_page=self.page, last_page=self.page, poppler_path=POPPLER_BIN)
                if imgs: img = imgs[0]
            else:
                img = Image.open(str(path))
                w, h = img.size
                img = img.resize((int(w*self.zoom), int(h*self.zoom)))
            
            return ImageTk.PhotoImage(img) if img else None
        except: return None

    def show_img(self, cv, photo):
        cv.delete("all")
        if photo:
            cv.create_image(0,0, image=photo, anchor="nw")
            cv.config(scrollregion=cv.bbox("all"))

    def bind_events(self):
        # Sync Scroll (Linux/Win/Mac handling)
        # v111: Mouse Wheel flips pages
        self.c1.bind("<MouseWheel>", self.on_scroll_page)
        self.c2.bind("<MouseWheel>", self.on_scroll_page)
        self.c1.bind("<Button-4>", self.on_scroll_page)
        self.c1.bind("<Button-5>", self.on_scroll_page)
        
        # Pan (Drag)
        self.c1.bind("<ButtonPress-1>", self.scroll_start)
        self.c1.bind("<B1-Motion>", self.scroll_move)
        self.c2.bind("<ButtonPress-1>", self.scroll_start)
        self.c2.bind("<B1-Motion>", self.scroll_move)

    def on_scroll_page(self, event):
        # v112: Debounce to prevent Mac "Super Scroll"
        now = time.time()
        if now - self.last_scroll_time < 0.4: return
        self.last_scroll_time = now

        d = 0
        if event.num == 5 or event.delta < 0: d = 1 # Down/Next
        elif event.num == 4 or event.delta > 0: d = -1 # Up/Prev
        
        if d == 1: self.next_page()
        elif d == -1: self.prev_page()

    def scroll_start(self, event):
        # Correct Tkinter Panning
        self.c1.scan_mark(event.x, event.y)
        self.c2.scan_mark(event.x, event.y)

    def scroll_move(self, event):
        # Correct Tkinter Panning
        self.c1.scan_dragto(event.x, event.y, gain=1)
        self.c2.scan_dragto(event.x, event.y, gain=1)

    def do_zoom(self, factor):
        self.zoom *= factor
        self.load_images()

    def fit_width(self):
        try:
            # Simple fit: assuming standard letter width ~600px at 72dpi
            cw = self.c1.winfo_width()
            if cw > 50:
                self.zoom = (cw - 20) / 600.0
                self.load_images()
        except: pass

    def next_page(self):
        if self.page < self.total_pages:
            self.page += 1
            self.load_images()

    def prev_page(self):
        if self.page > 1:
            self.page -= 1
            self.load_images()

    def next_dup(self):
        if self.dup_idx < len(self.dups)-1:
            self.dup_idx += 1
            self.load_images()

    def prev_dup(self):
        if self.dup_idx > 0:
            self.dup_idx -= 1
            self.load_images()

    def open_current_dup(self):
        if self.dups:
            # v111: Reveal instead of just open
            SystemUtils.reveal_file(self.dups[self.dup_idx])

    def mark_as_unique(self):
        # Safety Pivot - Promote to Master instead of Delete
        if not self.dups: return
        target_path = self.dups[self.dup_idx]
        
        if messagebox.askyesno("Promote File", f"Mark '{target_path.name}' as a unique Master file?\n\nIt will be removed from this duplicate list and treated as a distinct document."):
            try:
                # 1. Identify current entry in manifest
                root_path = None
                original_hash = None
                
                # Reverse lookup path in manifest (finding key by value in copy list)
                for h, data in self.manifest.items():
                    r_p = data.get('root')
                    if r_p:
                        try:
                            rel_p = str(target_path.relative_to(r_p))
                            if rel_p in data['copies']:
                                original_hash = h
                                root_path = r_p
                                break
                        except: continue

                if original_hash:
                    # 2. Modify Manifest
                    # A. Remove from original copies
                    rel_p = str(target_path.relative_to(root_path))
                    if rel_p in self.manifest[original_hash]['copies']:
                        self.manifest[original_hash]['copies'].remove(rel_p)

                    # B. Create new entry
                    new_uid = f"{uuid.uuid4()}_{sanitize_filename(target_path.name)}"
                    new_key = f"PROMOTED_{uuid.uuid4()}"
                    
                    self.manifest[new_key] = {
                        "master": rel_p,
                        "copies": [rel_p],
                        "name": target_path.name,
                        "root": root_path,
                        "uid": new_uid,
                        "id": "PROMOTED",
                        "status": "OK"
                    }
                    
                    # 3. Create Physical Master Copy (Required for Export logic)
                    m_dir = self.ws_path / "01_Master_Files"
                    shutil.copy2(target_path, m_dir / new_uid)

                    # 4. Save JSON
                    with open(self.ws_path / "manifest.json", 'w') as f:
                        json.dump(self.manifest, f, indent=4)

                    # 5. Update UI
                    del self.dups[self.dup_idx]
                    if not self.dups:
                        messagebox.showinfo("Done", "All duplicates handled.")
                        self.win.destroy()
                    else:
                        if self.dup_idx >= len(self.dups): self.dup_idx -= 1
                        self.load_images()
                else:
                    messagebox.showerror("Error", "Could not locate file in manifest structure.")

            except Exception as e:
                messagebox.showerror("Error", str(e))

class App:
    def __init__(self, root):
        self.root = root
        self.root.title(f"DocRefine Pro {SystemUtils.CURRENT_VERSION} ({platform.system()})")
        
        # v102: Smart Window Logic
        self.apply_smart_geometry(CFG.get("last_geometry"))
        
        self.q = queue.Queue(); self.worker = Worker(self.q)
        self.start_t = 0; self.running = False; self.paused = False
        self.current_manifest = {}
        self.slot_widgets = {} 
        self.slot_frames = []
        self.context_menu = None 
        
        self.is_mac = SystemUtils.IS_MAC
        self.Btn = ttk.Button if self.is_mac else tk.Button
        self.style = ttk.Style()
        
        if self.is_mac: 
            self.style.theme_use('clam')
            self.style.configure("Treeview", background="white", foreground="black", fieldbackground="white")
            self.style.map("Treeview", background=[('selected', '#0078d7')])

        # --- LAYOUT ---
        left = tk.Frame(root, width=350); left.pack(side="left", fill="both", padx=10, pady=10)
        tk.Label(left, text="Workspace Dashboard", font=("Segoe UI", 12, "bold")).pack(pady=5)
        
        kw_new = {"bg": "#e3f2fd"} if not self.is_mac else {}
        self.btn_new = self.Btn(left, text="+ New Ingest Job", command=self.ask_ingest_mode, **kw_new)
        self.btn_new.pack(fill="x", pady=2)
        
        btn_row = tk.Frame(left); btn_row.pack(fill="x", pady=5)
        self.btn_refresh = self.Btn(btn_row, text="↻ Refresh", command=self.load_jobs); self.btn_refresh.pack(side="left", fill="x", expand=True)
        kw_del = {"bg": "#ffcdd2"} if not self.is_mac else {}
        self.btn_del = self.Btn(btn_row, text="🗑 Delete", command=self.safe_delete_job, **kw_del); self.btn_del.pack(side="right")
        
        self.btn_upd = self.Btn(left, text="Check Updates", command=lambda: threading.Thread(target=self.check_updates, args=(True,), daemon=True).start())
        self.btn_upd.pack(anchor="w", pady=2)
        
        self.btn_settings = self.Btn(left, text="⚙ Settings", command=self.open_settings)
        self.btn_settings.pack(anchor="w", pady=2)

        self.btn_log = self.Btn(left, text="View App Log", command=self.open_app_log); self.btn_log.pack(anchor="w", pady=5)
        self.btn_open = self.Btn(left, text="Open Folder", command=self.open_f, state="disabled"); self.btn_open.pack(fill="x", pady=5)
        
        self.stats_fr = tk.LabelFrame(left, text="Stats", padx=5, pady=5); self.stats_fr.pack(fill="x")
        self.lbl_stats = tk.Label(self.stats_fr, text="Select a job...", anchor="w", justify="left"); self.lbl_stats.pack(fill="x")
        
        self.tree = ttk.Treeview(left, columns=("Name","Status","LastActive"), show="headings")
        for c, w in [("Name",140),("Status",70),("LastActive",110)]:
            self.tree.heading(c, text=c, command=lambda _c=c: self.sort_tree(self.tree,_c,False)); self.tree.column(c, width=w)
        self.tree.pack(fill="both", expand=True); self.tree.bind("<<TreeviewSelect>>", self.on_sel)
        
        right = tk.Frame(root); right.pack(side="right", fill="both", expand=True, padx=10, pady=10)
        self.nb = ttk.Notebook(right); self.nb.pack(fill="both", expand=True)
        
        self.tab_process = tk.Frame(self.nb); self.nb.add(self.tab_process, text=" 1. Refine ")
        self._build_refine()
        self.tab_dist = tk.Frame(self.nb); self.nb.add(self.tab_dist, text=" 2. Export ")
        self._build_export() 
        self.tab_inspect = tk.Frame(self.nb); self.nb.add(self.tab_inspect, text=" 🔍 Inspector ")
        self._build_inspect()
        
        mon = tk.LabelFrame(right, text="Process Monitor", padx=10, pady=10); mon.pack(fill="x", pady=10)
        head = tk.Frame(mon); head.pack(fill="x")
        
        # v112: Remove Blue/Colors for Mac Legibility
        status_fg = "blue" if not self.is_mac else "black"
        self.lbl_status = tk.Label(head, text="Ready", fg=status_fg, anchor="w", font=("Segoe UI", 9, "bold")); 
        self.lbl_status.pack(side="left", fill="x", expand=True)
        self.lbl_timer = tk.Label(head, text="00:00:00", font=("Consolas", 10)); self.lbl_timer.pack(side="right", padx=10)
        
        self.btn_receipt = self.Btn(head, text="View Receipt", command=self.open_receipt, state="disabled")
        self.btn_receipt.pack(side="right", padx=10)
        
        kw_stop = {"bg": "red", "fg": "white"} if not self.is_mac else {}
        self.btn_stop = self.Btn(head, text="STOP", command=self.stop, state="disabled", **kw_stop); self.btn_stop.pack(side="right")
        self.btn_pause = self.Btn(head, text="PAUSE", command=self.toggle_pause, state="disabled", width=8); self.btn_pause.pack(side="right", padx=5)

        tk.Label(mon, text="Overall Batch:", font=("Segoe UI", 8), anchor="w").pack(fill="x", pady=(5,0))
        self.p_main = tk.DoubleVar(); ttk.Progressbar(mon, variable=self.p_main).pack(fill="x", pady=2)
        
        self.frame_slots = tk.Frame(mon, pady=5)
        self.frame_slots.pack(fill="x")
        
        self.log_box = scrolledtext.ScrolledText(mon, height=8, bg="white", fg="black", insertbackground="black")
        self.log_box.pack(fill="both", expand=True)
        
        self.load_jobs()
        self.restore_session()
        
        # v112: Smarter Poll loop
        self.poll()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        
        threading.Thread(target=self.check_updates, args=(False,), daemon=True).start()

    # v108: Aggressive Startup Resize logic
    def apply_smart_geometry(self, saved_geo):
        try:
            sw = self.root.winfo_screenwidth()
            sh = self.root.winfo_screenheight()
            
            # Mac Dock Safety Buffer
            if SystemUtils.IS_MAC: sh -= 120 
            else: sh -= 60 # Win Taskbar Safety
            
            # Default fallback
            # v112: Larger default for Mac
            if SystemUtils.IS_MAC: w, h = 1280, 850
            else: w, h = 1024, 700
            
            x, y = 0, 0
            
            if saved_geo:
                parts = re.split(r'[x+]', saved_geo)
                if len(parts) == 4:
                    w, h, x, y = map(int, parts)
            
            # Aggressive Sanity Check
            if w > (sw * 0.95): 
                w = int(sw * 0.8)
                h = int(sh * 0.8)
                x = (sw // 2) - (w // 2)
                y = (sh // 2) - (h // 2)
            else:
                w = min(w, sw)
                h = min(h, sh)
            
            self.root.geometry(f"{w}x{h}+{x}+{y}")
        except:
            if SystemUtils.IS_MAC: self.root.geometry("1280x850")
            else: self.root.geometry("1024x700")

    # v104: Fixed Center Logic - NO CLAMPING (Allow negative coords for left monitors)
    @staticmethod
    def center_toplevel(win, parent):
        try:
            win.withdraw() # Hide
            win.update_idletasks() # Force size calc
            
            pw = parent.winfo_width()
            ph = parent.winfo_height()
            px = parent.winfo_rootx()
            py = parent.winfo_rooty()
            
            cw = win.winfo_width()
            ch = win.winfo_height()
            
            x = px + (pw // 2) - (cw // 2)
            y = py + (ph // 2) - (ch // 2)
            
            # Removed the max(0, x) clamp to support left-side monitors
            win.geometry(f"+{x}+{y}")
            win.deiconify() # Show
        except: 
            win.deiconify()

    def _build_refine(self):
        tk.Label(self.tab_process, text="Content Refinement (Modifies Files)", font=("Segoe UI",10,"bold")).pack(anchor="w",pady=(10,5),padx=10)
        
        self.chk_frame = tk.Frame(self.tab_process); self.chk_frame.pack(fill="x",padx=10)
        self.chk_vars = {} 
        
        # v110: PDF Controls in container for dynamic hiding
        self.f_pdf_ctrl = tk.Frame(self.tab_process)
        tk.Label(self.f_pdf_ctrl, text="PDF Action:", font=("Segoe UI", 9)).pack(anchor="w", padx=10, pady=(10,0))
        self.pdf_mode_var = tk.StringVar(value="No Action")
        self.cb_pdf = ttk.Combobox(self.f_pdf_ctrl, textvariable=self.pdf_mode_var, values=["No Action", "Flatten Only (Fast)", "Flatten + OCR (Slow)"], state="readonly")
        self.cb_pdf.pack(fill="x", padx=10, pady=2)
        # It will be packed in on_sel if needed
        
        ctrl = tk.Frame(self.tab_process); ctrl.pack(fill="x",pady=15,padx=10)
        
        # v110: Renamed Label
        tk.Label(ctrl, text="Processing & Preview Quality:").pack(side="left")
        self.dpi_var = tk.StringVar(value="Medium (Standard)")
        self.cb_dpi = ttk.Combobox(ctrl, textvariable=self.dpi_var, values=["Low (Fast)", "Medium (Standard)", "High (Slow)"], width=18, state="readonly")
        self.cb_dpi.pack(side="left",padx=5)
        
        self.btn_prev = self.Btn(ctrl, text="Generate Preview", command=self.safe_preview, state="disabled"); self.btn_prev.pack(side="left",padx=15)
        
        kw_run = {"bg": "#e8f5e9"} if not self.is_mac else {}
        self.btn_run = self.Btn(ctrl, text="Run Refinement", command=self.safe_start_batch, state="disabled", **kw_run); self.btn_run.pack(side="right")

    def _build_export(self):
        tk.Label(self.tab_dist, text="Final Export Strategies", font=("Segoe UI",10,"bold")).pack(anchor="w",pady=10,padx=10)
        f = tk.Frame(self.tab_dist); f.pack(fill="x",padx=10)
        
        tk.Label(f, text="Source Priority:", font=("Segoe UI", 9)).pack(anchor="w", pady=(0,2))
        
        default_prio = CFG.get("default_export_prio")
        self.prio_var = tk.StringVar(value=default_prio)
        self.cb_prio = ttk.Combobox(f, textvariable=self.prio_var, values=["Auto (Best Available)", "Force: OCR (Searchable)", "Force: Flattened (Visual)", "Force: Original Masters"], state="readonly", width=30)
        self.cb_prio.pack(anchor="w", pady=(0,10))
        
        self.var_ext = tk.BooleanVar(); tk.Checkbutton(f, text="Override Source: External Folder", variable=self.var_ext).pack(anchor="w")
        
        f_a = tk.LabelFrame(self.tab_dist, text="Option A: Unique Masters", padx=10, pady=10)
        f_a.pack(fill="x", padx=10, pady=5)
        tk.Label(f_a, text="Export a clean folder containing one copy of every unique file.", justify="left", fg="#555").pack(anchor="w")
        
        ra = tk.Frame(f_a); ra.pack(fill="x", pady=5)
        kw_org = {"bg": "#fff8e1"} if not self.is_mac else {}
        self.btn_org = self.Btn(ra, text="Export Unique Files", command=self.safe_start_organize, state="disabled", **kw_org); self.btn_org.pack(side="right")
        self.btn_dup_rpt = self.Btn(ra, text="View Dup Report", command=self.open_dup_report, state="disabled"); self.btn_dup_rpt.pack(side="right", padx=10)

        f_b = tk.LabelFrame(self.tab_dist, text="Option B: Reconstruct Original Structure", padx=10, pady=10)
        f_b.pack(fill="x", padx=10, pady=5)
        tk.Label(f_b, text="Re-create the original folder structure.", justify="left", fg="#555").pack(anchor="w")
        kw_dist = {"bg": "#fff3e0"} if not self.is_mac else {}
        self.btn_dist = self.Btn(f_b, text="Run Reconstruction", command=self.safe_start_dist, state="disabled", **kw_dist); self.btn_dist.pack(anchor="e", pady=5)

        # v101 - Option C: Reports
        f_c = tk.LabelFrame(self.tab_dist, text="Option C: Reports & Logs", padx=10, pady=10)
        f_c.pack(fill="x", padx=10, pady=5)
        
        rc = tk.Frame(f_c); rc.pack(fill="x", pady=5)
        self.btn_open_csv = self.Btn(rc, text="Open Full CSV", command=self.open_full_csv, state="disabled"); self.btn_open_csv.pack(side="right", padx=(10,0))
        self.btn_full_csv = self.Btn(rc, text="Export Full Inventory CSV", command=self.safe_start_full_export, state="disabled"); self.btn_full_csv.pack(side="right")
        tk.Label(rc, text="List every single scanned file and its location.", fg="#555").pack(side="left")

    def _build_inspect(self):
        h = tk.Frame(self.tab_inspect); h.pack(fill="x", padx=5, pady=5)
        tk.Label(h, text="Filter:").pack(side="left")
        
        self.search_var = tk.StringVar()
        self.search_var.trace("w", self.filter_inspection)
        self.entry_search = ttk.Entry(h, textvariable=self.search_var)
        self.entry_search.pack(side="left", fill="x", expand=True, padx=5)
        
        tk.Label(self.tab_inspect, text="Double-click a row to open file / view error.", font=("Segoe UI", 8), fg="#666").pack(anchor="w", padx=5)

        f = tk.Frame(self.tab_inspect); f.pack(fill="both",expand=True,padx=5,pady=(0,5))
        self.insp_tree = ttk.Treeview(f, columns=("ID","Name","Status","Copies"), show="headings")
        for c,w in [("ID",60),("Name",200),("Status",80),("Copies",50)]:
            self.insp_tree.heading(c, text=c, command=lambda _c=c: self.sort_tree(self.insp_tree,_c,False)); self.insp_tree.column(c, width=w)
        vsb = ttk.Scrollbar(f, orient="vertical", command=self.insp_tree.yview); self.insp_tree.configure(yscrollcommand=vsb.set)
        self.insp_tree.pack(side="left",fill="both",expand=True); vsb.pack(side="right",fill="y")
        
        self.insp_tree.bind("<Double-1>", self.on_inspect_click)
        
        self.context_menu = Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="Open File", command=self.open_selected_file) 
        self.context_menu.add_command(label="Reveal in Folder", command=self.reveal_in_folder)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Compare Duplicates...", command=self.on_compare_click)
        
        if self.is_mac:
            self.insp_tree.bind("<Button-2>", self.show_context_menu)
            self.insp_tree.bind("<Control-1>", self.show_context_menu)
        else:
            self.insp_tree.bind("<Button-3>", self.show_context_menu)

    def show_context_menu(self, event):
        try:
            item = self.insp_tree.identify_row(event.y)
            if item:
                self.insp_tree.selection_set(item)
                self.context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.context_menu.grab_release()

    def open_selected_file(self):
        item_id = self.insp_tree.selection()
        if not item_id: return
        self.on_inspect_click(None)

    def reveal_in_folder(self):
        ws = self.get_ws()
        if not ws: return
        item_id = self.insp_tree.selection()
        if not item_id: return
        name = self.insp_tree.item(item_id[0], 'values')[1]
        data = next((v for k,v in self.current_manifest.items() if v.get('name') == name or v.get('orig_name') == name), None)
        if data:
            path = ws/"01_Master_Files"/data['uid']
            if path.exists(): SystemUtils.open_file(path.parent)

    def on_compare_click(self):
        ws = self.get_ws()
        item_id = self.insp_tree.selection()
        if not ws or not item_id: return
        
        name = self.insp_tree.item(item_id[0], 'values')[1]
        data = next((v for k,v in self.current_manifest.items() if v.get('name') == name), None)
        
        if not data or len(data.get('copies', [])) < 2:
            messagebox.showinfo("Compare", "No duplicates to compare.")
            return
            
        master_file = ws/"01_Master_Files"/data['uid']
        
        root = data.get('root')
        if not root or not Path(root).exists():
            root = filedialog.askdirectory(title=f"Locate Source Folder for {name}")
            if not root: return
        
        duplicates = []
        for c in data['copies']:
            if c != data['master']:
                path = Path(root) / c
                if path.exists(): duplicates.append(path)
        
        if duplicates:
            ForensicComparator(self.root, ws, self.current_manifest, master_file, duplicates)
        else:
            messagebox.showerror("Error", "Could not locate any duplicate files.")

    def open_settings(self):
        win = tk.Toplevel(self.root)
        win.title("Preferences")
        win.geometry("600x600")
        
        # v103: Smart Centering
        App.center_toplevel(win, self.root)
        
        lf_perf = tk.LabelFrame(win, text="Processing Engine", padx=10, pady=10)
        lf_perf.pack(fill="x", padx=10, pady=5)
        
        # v112: Thread Clarity
        tk.Label(lf_perf, text="Max Worker Threads (0 = Auto):", font=("Segoe UI", 9, "bold")).grid(row=0,column=0,sticky="w")
        v_threads = tk.IntVar(value=CFG.get("max_threads"))
        tk.Spinbox(lf_perf, from_=0, to=32, textvariable=v_threads, width=5).grid(row=0,column=1,sticky="e")
        
        tk.Label(lf_perf, text="Safety Cap (Max Pixels):", font=("Segoe UI", 9, "bold")).grid(row=2,column=0,sticky="w", pady=(10,0))
        v_pixels = tk.StringVar(value=str(CFG.get("max_pixels")))
        tk.Entry(lf_perf, textvariable=v_pixels, width=15).grid(row=2,column=1,sticky="e", pady=(10,0))

        lf_def = tk.LabelFrame(win, text="Workflow Defaults", padx=10, pady=10)
        lf_def.pack(fill="x", padx=10, pady=5)
        
        tk.Label(lf_def, text="Default Ingest Mode:", font=("Segoe UI", 9, "bold")).grid(row=0,column=0,sticky="w")
        v_ingest = tk.StringVar(value=CFG.get("default_ingest_mode"))
        ttk.Combobox(lf_def, textvariable=v_ingest, values=["Standard", "Lightning", "Deep"], state="readonly").grid(row=0,column=1,sticky="e")

        tk.Label(lf_def, text="Default Export Priority:", font=("Segoe UI", 9, "bold")).grid(row=2,column=0,sticky="w", pady=(10,0))
        v_export = tk.StringVar(value=CFG.get("default_export_prio"))
        ttk.Combobox(lf_def, textvariable=v_export, values=["Auto (Best Available)", "Force: OCR (Searchable)", "Force: Flattened (Visual)", "Force: Original Masters"], state="readonly", width=30).grid(row=2,column=1,sticky="e", pady=(10,0))

        lf_ocr = tk.LabelFrame(win, text="Optical Character Recognition (OCR)", padx=10, pady=10)
        lf_ocr.pack(fill="x", padx=10, pady=5)
        
        tk.Label(lf_ocr, text="Tesseract Language:", font=("Segoe UI", 9, "bold")).grid(row=0,column=0,sticky="w")
        langs = get_tesseract_langs()
        v_lang = tk.StringVar(value=CFG.get("ocr_lang"))
        
        current = v_lang.get()
        if "(" not in current and current in langs: pass 
        else: 
            match = next((l for l in langs if f"({current})" in l), None)
            if match: v_lang.set(match)
            
        ttk.Combobox(lf_ocr, textvariable=v_lang, values=langs, state="readonly").grid(row=0,column=1,sticky="e")
        
        fr_btns = tk.Frame(lf_ocr)
        fr_btns.grid(row=1, column=0, columnspan=2, pady=10, sticky="e")
        
        def open_lang_folder():
            path = os.environ.get("TESSDATA_PREFIX")
            if not path and HAS_TESSERACT:
                try: path = str(Path(pytesseract.pytesseract.tesseract_cmd).parent / "tessdata")
                except: path = None
            if path and Path(path).exists(): SystemUtils.open_file(path)
            else: messagebox.showerror("Error", "Could not locate tessdata folder automatically.")

        def open_help():
            webbrowser.open("https://github.com/tesseract-ocr/tessdata_best")

        self.Btn(fr_btns, text="Open Language Folder", command=open_lang_folder).pack(side="left", padx=5)
        self.Btn(fr_btns, text="Get Languages (Web)", command=open_help).pack(side="left", padx=5)

        # v105: Support Section
        lf_support = tk.LabelFrame(win, text="Support & Diagnostics", padx=10, pady=10)
        lf_support.pack(fill="x", padx=10, pady=5)
        
        # v109: Threaded Launch
        def do_export_debug():
            self.start_debug_export_thread(btn_export, win)

        btn_export = self.Btn(lf_support, text="Export Debug Bundle (Zipped Logs)", command=do_export_debug)
        btn_export.pack(fill="x")
        tk.Label(lf_support, text="Use this if you need to report a bug.", font=("Segoe UI", 8), fg="#555").pack()

        def save():
            CFG.set("max_threads", v_threads.get())
            try: CFG.set("max_pixels", int(v_pixels.get()))
            except: pass
            CFG.set("default_ingest_mode", v_ingest.get())
            CFG.set("default_export_prio", v_export.get())
            
            code = parse_lang_code(v_lang.get())
            CFG.set("ocr_lang", code)
            
            Image.MAX_IMAGE_PIXELS = int(CFG.get("max_pixels"))
            self.prio_var.set(v_export.get())
            messagebox.showinfo("Saved", "Preferences updated.")
            win.destroy()
            
        btn_fr = tk.Frame(win); btn_fr.pack(pady=20)
        self.Btn(btn_fr, text="Save & Close", command=save, bg="#e8f5e9").pack(side="left", padx=10)

    # v109: Threaded Export Wrapper
    def start_debug_export_thread(self, btn_ref, win_ref):
        def _run():
            self._export_debug_bundle_task()
            self.q.put(("export_reset_btn", btn_ref))
        
        btn_ref.config(text="Exporting...", state="disabled")
        threading.Thread(target=_run, daemon=True).start()

    # --- ACTIONS ---
    def on_close(self):
        try:
            CFG.set("last_geometry", self.root.geometry())
            ws = self.get_ws()
            if ws: CFG.set("last_workspace", str(ws))
            current_tab = self.nb.index(self.nb.select())
            CFG.set("last_tab", current_tab)
        except: pass
        self.root.destroy()
        
    def restore_session(self):
        try:
            last_ws = CFG.get("last_workspace")
            if last_ws:
                for item in self.tree.get_children():
                    val = self.tree.item(item)['values'][0]
                    if Path(last_ws).name == val:
                        self.tree.selection_set(item)
                        self.tree.see(item)
                        break
            tab_idx = CFG.get("last_tab")
            if tab_idx is not None and tab_idx < 3: self.nb.select(tab_idx)
        except: pass

    def open_receipt(self):
        ws = self.get_ws()
        if not ws: return
        rpt_dir = ws / "04_Reports"
        if rpt_dir.exists():
            files = sorted(rpt_dir.glob("*.html"), key=os.path.getmtime, reverse=True)
            if files: SystemUtils.open_file(files[0])
            else: messagebox.showinfo("No Receipt", "No job receipt found yet.")
        else: messagebox.showinfo("No Receipt", "No job receipt found.")

    def open_dup_report(self):
        ws = self.get_ws()
        if not ws: return
        p = ws / "03_Organized_Output" / "duplicates_report.csv"
        if p.exists(): SystemUtils.open_file(p)
        else: messagebox.showinfo("Not Found", "No duplicates report found.")

    def open_full_csv(self):
        ws = self.get_ws()
        if not ws: return
        p = ws / "04_Reports" / "Full_Inventory_Manifest.csv"
        if p.exists(): SystemUtils.open_file(p)
        else: messagebox.showinfo("Not Found", "CSV not generated yet.")

    def filter_inspection(self, *args):
        query = self.search_var.get().lower()
        self.insp_tree.delete(*self.insp_tree.get_children())
        for k, v in self.current_manifest.items():
            name = v.get('name', '').lower()
            uid = v.get('id', '').lower()
            if query in name or query in uid:
                self._insert_inspect_row(v)

    def _insert_inspect_row(self, v):
        if v.get("status") == "QUARANTINE":
            st = "⛔ Quarantined"
            self.insp_tree.insert("", "end", values=("Q", v.get('orig_name', v.get('name','?')), st, "-"), tags=('q',))
        else:
            st = "Duplicate" if len(v.get('copies',[]))>1 else "Master"
            self.insp_tree.insert("", "end", values=(v.get('id','?'), v.get('name','?'), st, len(v.get('copies',[]))), tags=('ok',))

    def on_inspect_click(self, event):
        item_id = self.insp_tree.selection()
        if not item_id: return
        vals = self.insp_tree.item(item_id[0], 'values')
        name = vals[1]
        data = next((v for k,v in self.current_manifest.items() if v.get('name') == name or v.get('orig_name') == name), None)
        if not data: return
        ws = self.get_ws()
        if not ws: return
        if data.get("status") == "QUARANTINE":
            reason = data.get('error_reason', 'Unknown Error')
            messagebox.showerror("Quarantine Info", f"File: {name}\nReason: {reason}")
            q_file = next((f for f in (ws/"00_Quarantine").iterdir() if name in f.name), None)
            if q_file: SystemUtils.open_file(q_file.parent)
        else:
            # v100 FIX: Removed 'if event:' check to allow Context Menu access
            f_path = ws/"01_Master_Files"/data['uid']
            if f_path.exists():
                 SystemUtils.open_file(f_path)

    # v110: "Ghost Runner" Fix
    def check_run_btn(self, *args):
        if self.running: 
            self.btn_run.config(state="disabled")
            return
            
        any_checked = any(v.get() for v in self.chk_vars.values())
        pdf_active = self.pdf_mode_var.get() != "No Action"
        self.btn_run.config(state="normal" if (any_checked or pdf_active) else "disabled")

    def check_updates(self, manual=False):
        try:
            base_url = SystemUtils.UPDATE_MANIFEST_URL
            if "REPLACE" in base_url:
                if manual: messagebox.showinfo("Update Check", "Update URL not configured.")
                return
            url = f"{base_url}?t={int(time.time())}" if "?" not in base_url else f"{base_url}&t={int(time.time())}"
            ctx = ssl.create_default_context()
            ctx.check_hostname = False
            ctx.verify_mode = ssl.CERT_NONE
            with urllib.request.urlopen(url, context=ctx, timeout=5) as r:
                if r.status == 200:
                    data = json.loads(r.read().decode())
                    rem_ver = data.get("latest_version", "v0")
                    if rem_ver > SystemUtils.CURRENT_VERSION:
                        self.q.put(("update_avail", rem_ver, data.get("download_url")))
                    elif manual:
                         messagebox.showinfo("Update Check", f"You are up to date ({SystemUtils.CURRENT_VERSION}).")
        except Exception as e: 
            if manual: messagebox.showerror("Update Check Failed", str(e))

    def ask_ingest_mode(self):
        top = tk.Toplevel(self.root)
        top.title("New Job Setup")
        top.geometry("450x500") 
        
        # v103: Smart Centering
        App.center_toplevel(top, self.root)
        
        tk.Label(top, text="Select Mode", font=("Segoe UI", 12, "bold")).pack(pady=15)
        
        mode = tk.StringVar(value=CFG.get("default_ingest_mode"))
        modes = [
            ("Standard (Recommended)", "Smart Text Hash (PDFs).\nStrict Binary Hash (Others)."),
            ("Lightning (Fastest)", "Strict Binary Hash (All Files).\nExact digital copies only."),
            ("Deep Scan (Slowest)", "Full Text Scan (PDFs).\nStrict Binary Hash (Others).")
        ]
        for m, desc in modes:
            f = tk.Frame(top, pady=10, relief="groove", bd=1); f.pack(fill="x", padx=20, pady=5)
            tk.Radiobutton(f, text=m, variable=mode, value=m.split(" ")[0], font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=5)
            tk.Label(f, text=desc, fg="#555", justify="left").pack(anchor="w", padx=25)
        
        def go():
            # v111: ID tracking for jobs
            new_id = uuid.uuid4()
            self.current_job_id = new_id
            
            top.destroy(); d = filedialog.askdirectory()
            if d: 
                self.toggle(False)
                threading.Thread(target=self.wrap, args=(self.worker.run_inventory, d, mode.get(), new_id), daemon=True).start()
        
        self.Btn(top, text="Select Folder & Start", command=go).pack(fill="x", padx=20, pady=20, side="bottom")

    # v110: Added 'reset' parameter to control UI clearing
    def toggle(self, enable, reset=True):
        self.running = not enable; s = "normal" if enable else "disabled"
        if not enable: self.start_t = time.time(); self.paused = False
        self.tree.config(selectmode="browse" if enable else "none")
        for c in [self.btn_new, self.btn_refresh, self.btn_del, self.btn_dist, self.btn_prev, self.btn_open, self.btn_org, self.btn_full_csv, self.btn_open_csv]: 
            try: c.configure(state=s) 
            except: pass
        if enable: self.check_run_btn() 
        else: self.btn_run.config(state="disabled")
        self.btn_stop.config(state="normal" if not enable else "disabled")
        self.btn_pause.config(state="normal" if not enable else "disabled")
        if enable and reset: self.on_sel(None)

    def stop(self):
        # v111: Warning Dialog
        if messagebox.askyesno("Confirm Stop", "Stopping now will abort the current operation completely.\nYou will need to restart the job.\n\nAre you sure?"):
            self.worker.stop()
            self.btn_stop.config(state="disabled")
            self.btn_pause.config(state="disabled")

    def toggle_pause(self):
        if self.paused:
            self.worker.resume()
            if not self.is_mac: self.btn_pause.config(text="PAUSE", bg="SystemButtonFace")
            else: self.btn_pause.config(text="PAUSE")
            self.lbl_status.config(text="Resuming...", fg="blue")
        else:
            self.worker.pause()
            if not self.is_mac: self.btn_pause.config(text="RESUME", bg="yellow")
            else: self.btn_pause.config(text="RESUME")
        self.paused = not self.paused

    def load_jobs(self, sel=None):
        self.tree.delete(*self.tree.get_children())
        try: items = sorted(WORKSPACES_ROOT.iterdir(), key=lambda x: x.stat().st_mtime, reverse=True)
        except: items = []
        for d in items:
            if d.is_dir():
                s = "Empty"
                if (d/"status.json").exists(): 
                     with open(d/"status.json") as f: s = json.load(f).get("stage","?")
                elif (d/"Final_Delivery").exists(): s = "DISTRIBUTED"
                elif (d/"03_Organized_Output").exists(): s = "ORGANIZED"
                elif (d/"02_Ready_For_Redistribution").exists(): s = "PROCESSED"
                elif (d/"01_Master_Files").exists(): s = "INGESTED"
                
                item_id = self.tree.insert("", "end", values=(d.name, s, datetime.fromtimestamp(d.stat().st_mtime).strftime('%Y-%m-%d %H:%M')))
                
                if sel and str(d) == str(Path(sel)):
                     self.tree.selection_set(item_id)
                     self.tree.see(item_id)

    def safe_start_batch(self):
        ws = self.get_ws()
        opts = {k: v.get() for k,v in self.chk_vars.items()}
        
        p_mode = self.pdf_mode_var.get()
        if "OCR" in p_mode: opts['pdf_mode'] = 'ocr'
        elif "Flatten" in p_mode: opts['pdf_mode'] = 'flatten'
        else: opts['pdf_mode'] = 'none'
        
        d_raw = self.dpi_var.get()
        if "Low" in d_raw: opts['dpi'] = 150
        elif "High" in d_raw: opts['dpi'] = 600
        else: opts['dpi'] = 300
        
        if ws: self.toggle(False); threading.Thread(target=self.wrap, args=(self.worker.run_batch, str(ws), opts), daemon=True).start()

    def safe_start_organize(self):
        ws = self.get_ws()
        prio = self.prio_var.get()
        if ws: self.toggle(False); threading.Thread(target=self.wrap, args=(self.worker.run_organize, str(ws), prio), daemon=True).start()
        
    def safe_start_dist(self):
        ws=self.get_ws(); src=filedialog.askdirectory() if self.var_ext.get() else None
        prio = self.prio_var.get()
        if ws: self.toggle(False); threading.Thread(target=self.wrap, args=(self.worker.run_distribute, str(ws), src, prio), daemon=True).start()

    def safe_start_full_export(self):
        ws = self.get_ws()
        if ws: self.toggle(False); threading.Thread(target=self.wrap, args=(self.worker.run_full_export, str(ws)), daemon=True).start()
        
    def safe_preview(self):
        ws=self.get_ws()
        d_raw = self.dpi_var.get()
        dpi = 150 if "Low" in d_raw else (600 if "High" in d_raw else 300)
        if ws: self.toggle(False); threading.Thread(target=self.wrap, args=(self.worker.run_preview, str(ws), dpi), daemon=True).start()
    
    def safe_delete_job(self):
        ws = self.get_ws()
        if ws and messagebox.askyesno("Confirm", "Delete Job?"):
            def _del():
                try: shutil.rmtree(ws)
                except Exception as e: self.q.put(("error", f"Could not delete: {e}"))
                finally: self.q.put(("job", None))
            threading.Thread(target=_del, daemon=True).start()

    def on_sel(self, e):
        for w in self.chk_frame.winfo_children(): w.destroy()
        self.chk_vars.clear()
        self.btn_run.config(state="disabled")
        self.btn_prev.config(state="disabled")
        self.cb_dpi.config(state="disabled")
        self.cb_pdf.config(state="disabled")
        self.lbl_stats.config(text="Stats: Select a job...")
        self.btn_receipt.config(state="disabled")
        
        self.dpi_var.set("Medium (Standard)")
        self.pdf_mode_var.set("No Action")
        
        self.btn_dist.config(state="disabled")
        self.btn_org.config(state="disabled")
        self.btn_dup_rpt.config(state="disabled")
        
        self.btn_full_csv.config(state="disabled")
        self.btn_open_csv.config(state="disabled")
        
        self.current_manifest = {}
        self.insp_tree.delete(*self.insp_tree.get_children())
        self.search_var.set("")

        if self.running: return
        ws = self.get_ws()
        
        if not ws: 
            self.btn_open.config(state="disabled")
            return
        
        self.btn_open.config(state="normal")
        self.btn_dist.config(state="normal")
        self.btn_org.config(state="normal")
        self.btn_full_csv.config(state="normal")
        
        if (ws/"04_Reports").exists():
            if list((ws/"04_Reports").glob("Audit_Certificate_*.html")):
                self.btn_receipt.config(state="normal")
            if (ws/"04_Reports"/"Full_Inventory_Manifest.csv").exists():
                self.btn_open_csv.config(state="normal")
            
        if (ws/"03_Organized_Output"/"duplicates_report.csv").exists():
            self.btn_dup_rpt.config(state="normal")

        try:
            with open(ws/"stats.json") as f: s = json.load(f)
            total_seconds = int(
                s.get('ingest_time',0) + 
                s.get('batch_time',0) + 
                s.get('dist_time',0) + 
                s.get('organize_time',0)
            )
            self.lbl_stats.config(text=f"Files: {s.get('total_scanned',0)} | Masters: {s.get('masters',0)} | Q: {s.get('quarantined',0)} | Time: {str(timedelta(seconds=total_seconds))}")
        except: self.lbl_stats.config(text="Stats: N/A")
        
        types = set()
        if (ws/"01_Master_Files").exists(): types = {f.suffix.lower() for f in (ws/"01_Master_Files").rglob('*') if f.is_file()}
        
        def ac(l,k,desc): 
            f = tk.Frame(self.chk_frame); f.pack(fill="x", pady=2)
            v=tk.BooleanVar(); v.trace_add("write", self.check_run_btn)
            c=tk.Checkbutton(f,text=l,variable=v, font=("Segoe UI", 9, "bold")); c.pack(anchor="w")
            tk.Label(f, text=desc, font=("Segoe UI", 8), fg="#555").pack(anchor="w", padx=20)
            self.chk_vars[k]=v
        
        # v110: Dynamic PDF UI
        if '.pdf' in types: 
             self.f_pdf_ctrl.pack(fill="x", padx=10, pady=2) # Show
             self.cb_pdf.config(state="readonly")
             self.pdf_mode_var.trace_add("write", self.check_run_btn)
             self.btn_prev.config(state="normal")
        else:
             self.f_pdf_ctrl.pack_forget() # Hide

        if any(x in types for x in ['.jpg','.png']): ac("Resize Images","resize","Resize to 1920px (HD Standard)."); ac("Images to PDF","img2pdf","Bundle loose images into one PDF.")
        if any(x in types for x in ['.docx','.xlsx']): ac("Sanitize Office","sanitize","Remove author metadata and revision history.")
        
        if types:
             self.cb_dpi.config(state="readonly")
        
        self.log_box.delete(1.0, tk.END)
        if (ws/"session_log.txt").exists(): self.log_box.insert(tk.END, (ws/"session_log.txt").read_text(encoding="utf-8"))
        
        if (ws/"manifest.json").exists():
            try:
                with open(ws/"manifest.json") as f:
                    self.current_manifest = json.load(f)
                    for k,v in self.current_manifest.items():
                        self._insert_inspect_row(v)
            except: pass

    def sort_tree(self, t, c, r):
        l = [(t.set(k,c),k) for k in t.get_children('')]
        try: l.sort(key=lambda x: int(x[0]), reverse=r)
        except: l.sort(reverse=r)
        for i, (_,k) in enumerate(l): t.move(k,'',i)
        t.heading(c, command=lambda: self.sort_tree(t,c,not r))

    def open_app_log(self): SystemUtils.open_file(LOG_PATH)
    def open_f(self): SystemUtils.open_file(self.get_ws())
    def get_ws(self): s = self.tree.selection(); return WORKSPACES_ROOT / self.tree.item(s[0])['values'][0] if s else None
    def wrap(self, func, *args):
        try: func(*args)
        except Exception as e: self.q.put(("error", str(e))); self.q.put(("done",))
    def poll(self):
        try:
            for _ in range(50):
                m = self.q.get_nowait()
                if m[0]=='log': self.log_box.insert(tk.END, m[1]+"\n"); self.log_box.see(tk.END)
                elif m[0]=='main_p': self.p_main.set(m[1]); self.lbl_status.config(text=m[2], fg="blue")
                elif m[0]=='status_blue': self.lbl_status.config(text=m[1], fg="blue")
                elif m[0]=='status': self.lbl_status.config(text=m[1], fg="orange")
                elif m[0]=='job': self.load_jobs(m[1]) 
                # v111: Clean up pause state on Done
                elif m[0]=='done': 
                    self.toggle(True); 
                    self.lbl_status.config(text="Done", fg="green"); 
                    self.btn_pause.config(text="PAUSE", bg="SystemButtonFace", state="disabled")
                
                # v110: Smart Preview Complete Handler (Does not reset UI)
                elif m[0]=='preview_done':
                    self.toggle(True, reset=False) 
                    self.lbl_status.config(text="Preview Ready", fg="green")
                elif m[0]=='update_avail':
                    if messagebox.askyesno("Update Available", f"Version {m[1]} is available.\nDownload now?"): webbrowser.open(m[2])
                elif m[0]=='auto_open': SystemUtils.open_file(m[1])
                elif m[0]=='error': messagebox.showerror("Error", m[1])
                elif m[0]=='slot_config':
                    count = m[1]
                    for w in self.frame_slots.winfo_children(): w.destroy()
                    self.slot_widgets = {}
                    self.slot_frames = []
                    for i in range(count):
                        f = tk.Frame(self.frame_slots, relief="sunken", bd=1, bg="white")
                        f.pack(fill="x", padx=5, pady=1)
                        self.slot_frames.append(f)
                        tk.Label(f, text=f"Worker {i+1}:", font=("Consolas", 8), bg="#eee", width=10).pack(side="left")
                        lbl = tk.Label(f, text="Idle", font=("Segoe UI", 8), anchor="w", bg="white")
                        lbl.pack(side="left", fill="x", expand=True)
                elif m[0]=='slot_update':
                    tid, txt, _ = m[1], m[2], m[3]
                    if tid not in self.slot_widgets:
                        idx = len(self.slot_widgets)
                        if idx < len(self.slot_frames):
                             lbl = self.slot_frames[idx].winfo_children()[1]
                             self.slot_widgets[tid] = lbl
                    if tid in self.slot_widgets:
                        self.slot_widgets[tid].config(text=txt)
                elif m[0]=='export_success':
                    messagebox.showinfo("Export Successful", f"Debug bundle saved to:\n{Path(m[1]).name}")
                    SystemUtils.open_file(Path(m[1]).parent)
                elif m[0]=='export_reset_btn':
                    try: m[1].config(text="Export Debug Bundle (Zipped Logs)", state="normal")
                    except: pass

        except: pass
        if self.running and not self.paused: self.lbl_timer.config(text=str(timedelta(seconds=int(time.time()-self.start_t))))
        
        # v112: Sequential Polling (prevents queue flood on slow Macs)
        self.root.after(300, self.poll)

if __name__ == "__main__":
    try: root = tk.Tk(); App(root); root.mainloop()
    except Exception as e: messagebox.showerror("Fatal", str(e))