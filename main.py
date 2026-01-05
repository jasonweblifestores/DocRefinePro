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
import concurrent.futures
from logging.handlers import RotatingFileHandler
from pathlib import Path
from datetime import datetime, timedelta
from tkinter import filedialog, scrolledtext, messagebox, ttk
import tkinter as tk
from PIL import Image, ImageFile

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

# v87: Memory Safety Cap (500MP) to prevent decompression bombs
Image.MAX_IMAGE_PIXELS = 500000000 
ImageFile.LOAD_TRUNCATED_IMAGES = True 
try: import psutil; HAS_PSUTIL = True
except ImportError: HAS_PSUTIL = False

# ==============================================================================
#   DOCREFINE PRO v87 (MULTITHREADING & UX LOGIC UPDATE)
# ==============================================================================

# --- 1. SYSTEM ABSTRACTION & CONFIG ---
class SystemUtils:
    IS_WIN = platform.system() == 'Windows'
    IS_MAC = platform.system() == 'Darwin'
    CURRENT_VERSION = "v87"
    UPDATE_MANIFEST_URL = "https://gist.githubusercontent.com/jasonweblifestores/53752cda3c39550673fc5dafb96c4bed/raw/docrefine_version.json"

    @staticmethod
    def get_resource_dir():
        if getattr(sys, 'frozen', False): return Path(sys._MEIPASS)
        return Path(__file__).parent

    @staticmethod
    def get_user_data_dir():
        if SystemUtils.IS_MAC:
            p = Path.home() / "Documents" / "DocRefinePro_Data"
            p.mkdir(parents=True, exist_ok=True)
            return p
        if getattr(sys, 'frozen', False): return Path(sys.executable).parent
        return Path(__file__).parent

    @staticmethod
    def open_file(path):
        p = str(path)
        try:
            if SystemUtils.IS_WIN: os.startfile(p)
            elif SystemUtils.IS_MAC: subprocess.call(['open', p])
            else: subprocess.call(['xdg-open', p])
        except Exception as e: print(f"Error opening file: {e}")

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
    DEFAULTS = { "ram_warning_mb": 1024, "resize_width": 1920, "log_level": "INFO" }
    
    def __init__(self):
        self.data = self.DEFAULTS.copy()
        p = SystemUtils.get_user_data_dir() / "config.json"
        if p.exists():
            try:
                with open(p, 'r') as f: self.data.update(json.load(f))
            except: pass 
    def get(self, key): return self.data.get(key, self.DEFAULTS.get(key))

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

from pdf2image import convert_from_path, pdfinfo_from_path
import pypdf
from pypdf import PdfReader, PdfWriter

# --- 5. UTILS ---
class ToolTip:
    def __init__(self, w, t):
        self.w, self.t, self.tip = w, t, None
        w.bind("<Enter>", self.s); w.bind("<Leave>", self.h)
    def s(self, e=None):
        x, y, _, _ = self.w.bbox("insert"); x += self.w.winfo_rootx() + 25; y += self.w.winfo_rooty() + 25
        self.tip = tk.Toplevel(self.w); self.tip.wm_overrideredirect(True); self.tip.wm_geometry(f"+{x}+{y}")
        tk.Label(self.tip, text=self.t, bg="#ffffe0", relief="solid", bd=1, font=("tahoma","8")).pack()
    def h(self, e=None):
        if self.tip: self.tip.destroy(); self.tip=None

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
            for i in range(1, pages + 1):
                self.check_state() 
                # Reduced granularity for multithreading
                if i % 5 == 0 or i == pages: 
                     self.progress((i/pages)*100, f"Page {i}/{pages}")
                gc.collect() 
                res = convert_from_path(str(src), dpi=dpi, first_page=i, last_page=i, poppler_path=POPPLER_BIN)
                if not res: continue
                img = res[0]
                if mode == 'ocr' and HAS_TESSERACT:
                    t_page = temp / f"page_{i}.jpg"; img.save(t_page, "JPEG", dpi=(int(dpi), int(dpi)))
                    f = temp / f"{i}.pdf"
                    with open(f, "wb") as o: o.write(pytesseract.image_to_pdf_or_hocr(str(t_page), extension='pdf'))
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
        if status_only: self.q.put(("status_blue", t)) # Special handler for blue text
        else: self.q.put(("sub_p", v, t))
    
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

    def get_best_source(self, ws, file_uid):
        processed = ws / "02_Ready_For_Redistribution" / file_uid
        master = ws / "01_Master_Files" / file_uid
        if processed.parent.exists():
            if processed.exists(): return processed
            p_match = next((f for f in processed.parent.iterdir() if f.stem == Path(file_uid).stem), None)
            if p_match: return p_match
        return master if master.exists() else None

    def run_inventory(self, d_str, ingest_mode):
        try:
            d = Path(d_str); start_time = time.time()
            ws = WORKSPACES_ROOT / f"{d.name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            m_dir = ws / "01_Master_Files"; m_dir.mkdir(parents=True); (ws/"00_Quarantine").mkdir()
            self.current_ws = str(ws); self.log(f"Inventory Start: {d}")
            self.q.put(("job", str(ws))); self.q.put(("ws", str(ws)))
            self.set_job_status(ws, "SCANNING", "Ingesting...")
            files = [Path(r)/f for r,_,fs in os.walk(d) for f in fs]
            seen = {}; quarantined = 0; file_types = {}

            for i, f in enumerate(files):
                if self.stop_sig: break
                if not self.pause_event.is_set(): self.prog_sub(None, "Paused...", True); self.pause_event.wait()
                self.prog_main((i/len(files))*100, f"Scanning {i}/{len(files)}")
                if f.suffix.lower() not in SUPPORTED_EXTENSIONS: continue
                file_types[f.suffix.lower()] = file_types.get(f.suffix.lower(), 0) + 1
                
                h, method = self.get_hash(f, ingest_mode)
                if not h: 
                    self.log(f"⚠️ Quarantine: {f.name}", True)
                    shutil.copy2(f, ws/"00_Quarantine"/f"{uuid.uuid4()}_{sanitize_filename(f.name)}")
                    quarantined += 1; continue
                
                rel = str(f.relative_to(d))
                if h in seen: seen[h]['copies'].append(rel)
                else: seen[h] = {'master': rel, 'copies': [rel], 'name': f.name}
            
            self.log("Tagging..."); total = len(seen)
            for i, (h, data) in enumerate(seen.items()):
                if self.stop_sig: break
                safe_name = f"[{i+1:04d}]_{sanitize_filename(data['name'])}"
                shutil.copy2(d / data['master'], m_dir / safe_name)
                data['uid'] = safe_name; data['id'] = f"[{i+1:04d}]"
            
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

    # v87: Process Single File (Helper for ThreadPool)
    def process_file_task(self, f, bots, options, dst):
        if self.stop_sig: return
        try:
            # Thread-safe logging requires simplified messages
            self.q.put(("status_blue", f"Refining: {f.name}"))
            
            ext = f.suffix.lower()
            ok = False
            
            # v87: Logic Mapped from UI Dropdowns
            dpi_val = int(options.get('dpi', 300))
            
            if ext == '.pdf':
                mode = options.get('pdf_mode', 'none')
                if mode == 'flatten': ok = bots['pdf'].flatten_or_ocr(f, dst/f.name, 'flatten', dpi=dpi_val)
                elif mode == 'ocr': ok = bots['pdf'].flatten_or_ocr(f, dst/f.name, 'ocr', dpi=dpi_val)
            
            elif ext in {'.jpg','.png'}:
                if options.get('resize'): ok = bots['img'].resize(f, dst/f.name, CFG.get('resize_width'))
                if options.get('img2pdf'): ok = bots['img'].convert_to_pdf(f, dst/f"{f.stem}.pdf")
            
            elif ext in {'.docx','.xlsx'}:
                if options.get('sanitize'): ok = bots['office'].sanitize(f, dst/f.name)

            if not ok and not (dst/f.name).exists(): shutil.copy2(f, dst/f.name)
        except Exception as e:
            self.log(f"Err {f.name}: {e}", True)

    def run_batch(self, ws_p, options):
        try:
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
            
            # v87: Multithreading Implementation
            # Limit workers to 4 to prevent UI freeze on dual-core machines
            max_workers = min(4, os.cpu_count() or 2)
            
            with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
                futures = {executor.submit(self.process_file_task, f, bots, options, dst): f for f in fs}
                
                for i, future in enumerate(concurrent.futures.as_completed(futures)):
                    if self.stop_sig: break
                    self.prog_main((i/len(fs))*100, f"Refining {i+1}/{len(fs)}")
                    try: future.result()
                    except Exception as e: self.log(f"Thread Err: {e}", True)

            update_stats_time(ws, "batch_time", time.time() - start_time)
            self.set_job_status(ws, "PROCESSED", "Complete")
            self.q.put(("job", str(ws))) 
            self.prog_main(100, "Done"); self.q.put(("done",)); SystemUtils.open_file(dst)
        except Exception as e: self.log(f"Err: {e}", True); self.q.put(("done",))

    def run_organize(self, ws_p):
        try:
            ws = Path(ws_p); self.current_ws = str(ws)
            start_time = time.time()
            out = ws / "03_Organized_Output"; m = out/"Unique_Masters"; q = out/"Quarantine"
            for p in [m,q]: p.mkdir(parents=True, exist_ok=True)
            
            self.log("Unique Export Start")
            with open(ws/"manifest.json") as f: man = json.load(f)
            total = len(man)
            
            dup_csv = out / "duplicates_report.csv"
            with open(dup_csv, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(["Master_Filename", "Duplicate_Location"])
                
                for i, (h, data) in enumerate(man.items()):
                    if self.stop_sig: break
                    self.prog_main((i/total)*100, "Exporting Unique...")
                    self.q.put(("status_blue", f"Exporting: {data['name']}"))
                    
                    if data.get("status") == "QUARANTINE": 
                        for f in (ws/"00_Quarantine").glob("*"):
                            if data['orig_name'] in f.name: shutil.copy2(f, q/f.name)
                    else:
                        src = self.get_best_source(ws, data['uid'])
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

            update_stats_time(ws, "organize_time", time.time() - start_time)
            self.set_job_status(ws, "ORGANIZED", "Done")
            self.q.put(("job", str(ws))) 
            self.prog_main(100, "Done"); self.q.put(("done",)); SystemUtils.open_file(out)
        except Exception as e: self.log(f"Err: {e}", True); self.q.put(("done",))

    def run_distribute(self, ws_p, ext_src):
        try:
            ws = Path(ws_p); self.current_ws = str(ws)
            if not (ws/"manifest.json").exists():
                 self.log("CRITICAL: Manifest missing.", True)
                 self.q.put(("error", "Manifest missing.")); self.q.put(("done",)); return

            start_time = time.time(); 
            dst = ws / "Final_Delivery"
            self.log("Reconstruction Start")
            self.set_job_status(ws, "DISTRIBUTING", "Reconstructing...")
            
            with open(ws/"manifest.json") as f: man = json.load(f)
            
            orphans = {}
            if ext_src:
                 orphans = {f.name: f for f in Path(ext_src).iterdir()}

            for i, (h, d) in enumerate(man.items()):
                if self.stop_sig: break
                self.prog_main((i/len(man))*100, f"Recon {i+1}")
                self.q.put(("status_blue", f"Copying: {d['name']}"))
                
                if d.get("status") == "QUARANTINE": continue
                
                src = None
                if ext_src:
                    src = next((v for k,v in orphans.items() if k.startswith(d['id'])), None)
                else:
                    src = self.get_best_source(ws, d['uid'])
                
                if not src: continue
                
                for c in d['copies']:
                    t = dst / c; t.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(src, t.with_suffix(src.suffix))
            
            q_src = ws / "00_Quarantine"
            if q_src.exists():
                q_dst = dst / "_QUARANTINED_FILES"; 
                q_dst.mkdir(parents=True, exist_ok=True) 
                for qf in q_src.iterdir(): shutil.copy2(qf, q_dst / qf.name)

            update_stats_time(ws, "dist_time", time.time() - start_time)
            self.set_job_status(ws, "DISTRIBUTED", "Done")
            self.q.put(("job", str(ws))) 
            self.prog_main(100, "Done"); self.q.put(("done",)); SystemUtils.open_file(dst)
        except Exception as e: self.log(f"Err: {e}", True); self.q.put(("done",))

    def run_preview(self, ws_p, dpi):
        try:
            ws = Path(ws_p); self.current_ws = str(ws)
            src = ws/"01_Master_Files"; pdf = next(src.glob("*.pdf"), None)
            if not pdf: self.q.put(("done",)); return
            for old in ws.glob("PREVIEW_*.pdf"): 
                try: os.remove(old)
                except: pass
            out = ws / f"PREVIEW_{int(time.time())}.pdf"
            imgs = convert_from_path(str(pdf), dpi=int(dpi), first_page=1, last_page=1, poppler_path=POPPLER_BIN)
            if imgs: 
                imgs[0].save(out, "PDF", resolution=float(dpi))
                SystemUtils.open_file(out)
            self.q.put(("done",))
        except: self.q.put(("done",))

# --- 8. UI (MAC NATIVE) ---
class App:
    def __init__(self, root):
        self.root = root
        self.root.title(f"DocRefine Pro {SystemUtils.CURRENT_VERSION} ({platform.system()})")
        self.root.geometry("1150x900")
        self.q = queue.Queue(); self.worker = Worker(self.q)
        self.start_t = 0; self.running = False; self.paused = False
        
        self.is_mac = SystemUtils.IS_MAC
        self.Btn = ttk.Button if self.is_mac else tk.Button
        self.style = ttk.Style()
        if self.is_mac: self.style.theme_use('clam') 

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
        self.lbl_status = tk.Label(head, text="Ready", fg="blue", anchor="w", font=("Segoe UI", 9, "bold")); 
        self.lbl_status.pack(side="left", fill="x", expand=True)
        self.lbl_timer = tk.Label(head, text="00:00:00", font=("Consolas", 10)); self.lbl_timer.pack(side="right", padx=10)
        
        kw_stop = {"bg": "red", "fg": "white"} if not self.is_mac else {}
        self.btn_stop = self.Btn(head, text="STOP", command=self.stop, state="disabled", **kw_stop); self.btn_stop.pack(side="right")
        self.btn_pause = self.Btn(head, text="PAUSE", command=self.toggle_pause, state="disabled", width=8); self.btn_pause.pack(side="right", padx=5)

        tk.Label(mon, text="Overall Batch:", font=("Segoe UI", 8), anchor="w").pack(fill="x", pady=(5,0))
        self.p_main = tk.DoubleVar(); ttk.Progressbar(mon, variable=self.p_main).pack(fill="x", pady=2)
        
        self.lbl_sub_stats = tk.Label(mon, text="Waiting...", font=("Segoe UI", 8), anchor="w", fg="#666")
        self.lbl_sub_stats.pack(fill="x", pady=(5,0))
        self.p_sub = tk.DoubleVar(); ttk.Progressbar(mon, variable=self.p_sub).pack(fill="x", pady=2)
        
        self.log_box = scrolledtext.ScrolledText(mon, height=8); self.log_box.pack(fill="both", expand=True)
        self.load_jobs(); self.root.after(100, self.poll)
        
        threading.Thread(target=self.check_updates, args=(False,), daemon=True).start()

    def _build_refine(self):
        # v87: Logic Refactor for Dropdowns
        tk.Label(self.tab_process, text="Content Refinement (Modifies Files)", font=("Segoe UI",10,"bold")).pack(anchor="w",pady=(10,5),padx=10)
        
        self.chk_frame = tk.Frame(self.tab_process); self.chk_frame.pack(fill="x",padx=10)
        self.chk_vars = {} 
        
        # New: PDF Mode Dropdown
        tk.Label(self.tab_process, text="PDF Action:", font=("Segoe UI", 9)).pack(anchor="w", padx=10, pady=(10,0))
        self.pdf_mode_var = tk.StringVar(value="No Action")
        self.cb_pdf = ttk.Combobox(self.tab_process, textvariable=self.pdf_mode_var, values=["No Action", "Flatten Only (Fast)", "Flatten + OCR (Slow)"], state="readonly")
        self.cb_pdf.pack(fill="x", padx=10, pady=2)
        
        # New Control Row: Quality + Preview + Run
        ctrl = tk.Frame(self.tab_process); ctrl.pack(fill="x",pady=15,padx=10)
        
        tk.Label(ctrl, text="Quality:").pack(side="left")
        self.dpi_var = tk.StringVar(value="Medium (Standard)")
        # v87: Semantic Quality Labels
        self.cb_dpi = ttk.Combobox(ctrl, textvariable=self.dpi_var, values=["Low (Fast)", "Medium (Standard)", "High (Slow)"], width=18, state="readonly")
        self.cb_dpi.pack(side="left",padx=5)
        
        self.btn_prev = self.Btn(ctrl, text="Generate Preview", command=self.safe_preview, state="disabled"); self.btn_prev.pack(side="left",padx=15)
        
        kw_run = {"bg": "#e8f5e9"} if not self.is_mac else {}
        self.btn_run = self.Btn(ctrl, text="Run Refinement", command=self.safe_start_batch, state="disabled", **kw_run); self.btn_run.pack(side="right")

    def _build_export(self):
        tk.Label(self.tab_dist, text="Final Export Strategies", font=("Segoe UI",10,"bold")).pack(anchor="w",pady=10,padx=10)
        f = tk.Frame(self.tab_dist); f.pack(fill="x",padx=10)
        self.var_ext = tk.BooleanVar(); tk.Checkbutton(f, text="Override Source: External Folder", variable=self.var_ext).pack(anchor="w")
        
        f_a = tk.LabelFrame(self.tab_dist, text="Option A: Unique Masters", padx=10, pady=10)
        f_a.pack(fill="x", padx=10, pady=5)
        tk.Label(f_a, text="Export a clean folder containing one copy of every unique file.\n(Uses refined versions if available).", justify="left", fg="#555").pack(anchor="w")
        kw_org = {"bg": "#fff8e1"} if not self.is_mac else {}
        self.btn_org = self.Btn(f_a, text="Export Unique Files", command=self.safe_start_organize, state="disabled", **kw_org); self.btn_org.pack(anchor="e", pady=5)

        f_b = tk.LabelFrame(self.tab_dist, text="Option B: Reconstruct Original Structure", padx=10, pady=10)
        f_b.pack(fill="x", padx=10, pady=5)
        tk.Label(f_b, text="Re-create the original folder structure using refined files.", justify="left", fg="#555").pack(anchor="w")
        kw_dist = {"bg": "#fff3e0"} if not self.is_mac else {}
        self.btn_dist = self.Btn(f_b, text="Run Reconstruction", command=self.safe_start_dist, state="disabled", **kw_dist); self.btn_dist.pack(anchor="e", pady=5)

    def _build_inspect(self):
        f = tk.Frame(self.tab_inspect); f.pack(fill="both",expand=True,padx=5,pady=5)
        self.insp_tree = ttk.Treeview(f, columns=("ID","Name","Status","Copies"), show="headings")
        for c,w in [("ID",60),("Name",200),("Status",80),("Copies",50)]:
            self.insp_tree.heading(c, text=c, command=lambda _c=c: self.sort_tree(self.insp_tree,_c,False)); self.insp_tree.column(c, width=w)
        vsb = ttk.Scrollbar(f, orient="vertical", command=self.insp_tree.yview); self.insp_tree.configure(yscrollcommand=vsb.set)
        self.insp_tree.pack(side="left",fill="both",expand=True); vsb.pack(side="right",fill="y")

    # --- ACTIONS ---
    def check_run_btn(self, *args):
        # v87: Enable button if any checkbox OR PDF mode is active
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
            with urllib.request.urlopen(url, timeout=5) as r:
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
        tk.Label(top, text="Select Mode", font=("Segoe UI", 12, "bold")).pack(pady=15)
        mode = tk.StringVar(value="Standard")
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
            top.destroy(); d = filedialog.askdirectory()
            if d: self.toggle(False); threading.Thread(target=self.wrap, args=(self.worker.run_inventory,d,mode.get()), daemon=True).start()
        
        self.Btn(top, text="Select Folder & Start", command=go).pack(fill="x", padx=20, pady=20, side="bottom")

    def toggle(self, enable):
        self.running = not enable; s = "normal" if enable else "disabled"
        if not enable: self.start_t = time.time(); self.paused = False
        self.tree.config(selectmode="browse" if enable else "none")
        for c in [self.btn_new, self.btn_refresh, self.btn_del, self.btn_dist, self.btn_prev, self.btn_open, self.btn_org]: 
            try: c.configure(state=s) 
            except: pass
        if enable: self.check_run_btn() 
        else: self.btn_run.config(state="disabled")
        self.btn_stop.config(state="normal" if not enable else "disabled")
        self.btn_pause.config(state="normal" if not enable else "disabled")
        if enable: self.on_sel(None)

    def stop(self):
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
                
                # RE-SELECT LOGIC
                if sel and str(d) == str(Path(sel)):
                     self.tree.selection_set(item_id)
                     self.tree.see(item_id)

    def safe_start_batch(self):
        ws = self.get_ws()
        # v87: Build Options Dictionary
        opts = {k: v.get() for k,v in self.chk_vars.items()}
        
        # Parse PDF Mode
        p_mode = self.pdf_mode_var.get()
        if "OCR" in p_mode: opts['pdf_mode'] = 'ocr'
        elif "Flatten" in p_mode: opts['pdf_mode'] = 'flatten'
        else: opts['pdf_mode'] = 'none'
        
        # Parse DPI
        d_raw = self.dpi_var.get()
        if "Low" in d_raw: opts['dpi'] = 150
        elif "High" in d_raw: opts['dpi'] = 600
        else: opts['dpi'] = 300
        
        if ws: self.toggle(False); threading.Thread(target=self.wrap, args=(self.worker.run_batch, str(ws), opts), daemon=True).start()

    def safe_start_organize(self):
        ws = self.get_ws()
        if ws: self.toggle(False); threading.Thread(target=self.wrap, args=(self.worker.run_organize, str(ws)), daemon=True).start()
    def safe_start_dist(self):
        ws=self.get_ws(); src=filedialog.askdirectory() if self.var_ext.get() else None
        if ws: self.toggle(False); threading.Thread(target=self.wrap, args=(self.worker.run_distribute, str(ws), src), daemon=True).start()
    def safe_preview(self):
        ws=self.get_ws()
        # Parse DPI for preview
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
        
        self.dpi_var.set("Medium (Standard)")
        self.pdf_mode_var.set("No Action")
        
        self.btn_dist.config(state="disabled")
        self.btn_org.config(state="disabled")

        if self.running: return
        ws = self.get_ws()
        
        if not ws: 
            self.btn_open.config(state="disabled")
            self.insp_tree.delete(*self.insp_tree.get_children())
            return
        
        self.btn_open.config(state="normal")
        self.btn_dist.config(state="normal")
        self.btn_org.config(state="normal")

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
        
        def ac(l,k,t): 
            v=tk.BooleanVar(); v.trace_add("write", self.check_run_btn)
            c=tk.Checkbutton(self.chk_frame,text=l,variable=v); c.pack(side="left",padx=10); self.chk_vars[k]=v; ToolTip(c,t)
        
        # v87: UI Control Enabling
        if '.pdf' in types: 
             self.cb_pdf.config(state="readonly")
             self.pdf_mode_var.trace_add("write", self.check_run_btn)
             self.btn_prev.config(state="normal")

        if any(x in types for x in ['.jpg','.png']): ac("Resize Images","resize","Resize to 1920px."); ac("Images to PDF","img2pdf","Bundle images.")
        if any(x in types for x in ['.docx','.xlsx']): ac("Sanitize Office","sanitize","Remove metadata.")
        
        if types:
             self.cb_dpi.config(state="readonly")
        
        self.log_box.delete(1.0, tk.END)
        if (ws/"session_log.txt").exists(): self.log_box.insert(tk.END, (ws/"session_log.txt").read_text(encoding="utf-8"))
        self.insp_tree.delete(*self.insp_tree.get_children())
        if (ws/"manifest.json").exists():
            try:
                with open(ws/"manifest.json") as f:
                    for k,v in json.load(f).items():
                        if v.get("status") == "QUARANTINE":
                            st = f"⛔ {v.get('error_reason')}"
                            self.insp_tree.insert("", "end", values=("Q", v.get('orig_name', v.get('name','?')), st, "-"))
                        else:
                            st = "Duplicate" if len(v.get('copies',[]))>1 else "Master"
                            self.insp_tree.insert("", "end", values=(v.get('id','?'), v.get('name','?'), st, len(v.get('copies',[]))))
            except: pass

    def sort_tree(self, t, c, r):
        l = [(t.set(k,c),k) for k in t.get_children('')]
        try: l.sort(key=lambda x: int(x[0]), reverse=r)
        except: l.sort(reverse=r)
        for i, (_,k) in enumerate(l): t.move(k,'',i)
        t.heading(c, command=lambda: self.sort_tree(t,c,not r))
    
    def warn_ocr(self): 
        # v87: Deprecated in favor of dropdown logic, keeping for safety if reused
        pass

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
                elif m[0]=='sub_p': self.p_sub.set(m[1]); self.lbl_sub_stats.config(text=m[2]) # Update DETAIL label
                elif m[0]=='status_blue': self.lbl_status.config(text=m[1], fg="blue") # EXPLICIT BLUE UPDATE
                elif m[0]=='status': self.lbl_status.config(text=m[1], fg="orange")
                elif m[0]=='job': self.load_jobs(m[1]) # Pass ID for reselection
                elif m[0]=='done': self.toggle(True); self.lbl_status.config(text="Done", fg="green"); self.p_sub.set(0)
                elif m[0]=='update_avail':
                    if messagebox.askyesno("Update Available", f"Version {m[1]} is available.\nDownload now?"): webbrowser.open(m[2])
                elif m[0]=='auto_open': SystemUtils.open_file(m[1])
                elif m[0]=='error': messagebox.showerror("Error", m[1])
        except: pass
        if self.running and not self.paused: self.lbl_timer.config(text=str(timedelta(seconds=int(time.time()-self.start_t))))
        self.root.after(100, self.poll)

if __name__ == "__main__":
    try: root = tk.Tk(); App(root); root.mainloop()
    except Exception as e: messagebox.showerror("Fatal", str(e))