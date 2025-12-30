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
from logging.handlers import RotatingFileHandler
from pathlib import Path
from datetime import datetime, timedelta
from tkinter import filedialog, scrolledtext, messagebox, ttk
import tkinter as tk
from PIL import Image, ImageFile

# ==============================================================================
#   WINDOWS GHOST WINDOW FIX (CRITICAL FOR PDF2IMAGE/TESSERACT)
# ==============================================================================
if os.name == 'nt':
    try:
        import subprocess
        _original_popen = subprocess.Popen

        def safe_popen(*args, **kwargs):
            if 'startupinfo' not in kwargs:
                si = subprocess.STARTUPINFO()
                si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                si.wShowWindow = subprocess.SW_HIDE
                kwargs['startupinfo'] = si
                if 'creationflags' not in kwargs:
                    kwargs['creationflags'] = 0x08000000 # CREATE_NO_WINDOW
            return _original_popen(*args, **kwargs)
        subprocess.Popen = safe_popen
    except Exception as e:
        print(f"Warning: Could not patch subprocess: {e}")

# Increase PIL limit & allow truncated images
Image.MAX_IMAGE_PIXELS = None
ImageFile.LOAD_TRUNCATED_IMAGES = True 

# TRY IMPORT PSUTIL (Graceful Degradation)
try:
    import psutil
    HAS_PSUTIL = True
except ImportError:
    HAS_PSUTIL = False

# ==============================================================================
#   DOCREFINE PRO v74 (UPDATE CHECKER)
# ==============================================================================

# --- 1. SYSTEM ABSTRACTION & CONFIG ---
class SystemUtils:
    IS_WIN = platform.system() == 'Windows'
    IS_MAC = platform.system() == 'Darwin'
    CURRENT_VERSION = "v74" # Update this to match your git tag

    @staticmethod
    def get_base_dir():
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
    def find_binary(portable_dir, bin_name):
        base = SystemUtils.get_base_dir()
        tgt = base / portable_dir
        if tgt.exists():
            if (tgt / bin_name).exists(): return str(tgt)
            if (tgt / "bin" / bin_name).exists(): return str(tgt / "bin")
        sys_path = shutil.which(bin_name)
        if sys_path: return str(Path(sys_path).parent)
        return None

class Config:
    # --- CONFIGURE YOUR REPO HERE ---
    GITHUB_REPO = "jasonweblifestores/DocRefinePro" 
    # --------------------------------
    
    DEFAULTS = {
        "ram_warning_mb": 1024,
        "resize_width": 1920,
        "log_level": "INFO",
        "log_max_bytes": 1024 * 1024,
        "log_backup_count": 5
    }
    def __init__(self):
        self.data = self.DEFAULTS.copy()
        p = SystemUtils.get_base_dir() / "config.json"
        if p.exists():
            try:
                with open(p, 'r') as f: self.data.update(json.load(f))
            except: pass 

    def get(self, key): return self.data.get(key, self.DEFAULTS.get(key))

CFG = Config()

# --- 2. LOGGING SETUP ---
BASE_DIR = SystemUtils.get_base_dir()
LOG_PATH = BASE_DIR / "app_debug.log"
WORKSPACES_ROOT = BASE_DIR / "Workspaces"
WORKSPACES_ROOT.mkdir(exist_ok=True)

SUPPORTED_EXTENSIONS = {'.pdf', '.doc', '.docx', '.jpg', '.png', '.xls', '.xlsx', '.csv', '.jpeg'}

logger = logging.getLogger("DocRefine")
logger.setLevel(getattr(logging, CFG.get("log_level").upper(), logging.INFO))
c_handler = logging.StreamHandler()
c_handler.setFormatter(logging.Formatter('[%(levelname)s] %(message)s'))
logger.addHandler(c_handler)

try:
    f_handler = RotatingFileHandler(
        LOG_PATH, maxBytes=CFG.get("log_max_bytes"), backupCount=CFG.get("log_backup_count"), encoding='utf-8'
    )
    f_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
    logger.addHandler(f_handler)
except: pass

def log_app(msg, level="INFO"):
    if level == "ERROR": logger.error(msg)
    elif level == "WARN": logger.warning(msg)
    else: logger.info(msg)

log_app(f"=== APP STARTUP {SystemUtils.CURRENT_VERSION} ({platform.system()}) ===")
if HAS_PSUTIL: log_app(f"Memory Monitor: ACTIVE (Threshold: {CFG.get('ram_warning_mb')}MB)")
else: log_app("Memory Monitor: UNAVAILABLE (psutil missing)", "WARN")

# --- 3. STARTUP HYGIENE ---
def clean_temp_files():
    try:
        limit = time.time() - 86400
        for ws in WORKSPACES_ROOT.iterdir():
            if ws.is_dir():
                for item in ws.glob("temp_*"):
                    if item.is_dir() and item.stat().st_mtime < limit:
                        shutil.rmtree(item, ignore_errors=True)
                        log_app(f"Cleaned zombie temp: {item.name}")
    except: pass
clean_temp_files()

# --- 4. DEPENDENCIES ---
POPPLER_BIN = SystemUtils.find_binary("poppler", "pdfinfo" + (".exe" if SystemUtils.IS_WIN else ""))
TESSERACT_BIN = SystemUtils.find_binary("Tesseract-OCR", "tesseract" + (".exe" if SystemUtils.IS_WIN else ""))
HAS_TESSERACT = bool(TESSERACT_BIN)
if TESSERACT_BIN:
    import pytesseract
    exe_name = "tesseract.exe" if SystemUtils.IS_WIN else "tesseract"
    pytesseract.pytesseract.tesseract_cmd = str(Path(TESSERACT_BIN) / exe_name)

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

# --- 6. PROCESSORS (DEEP VALIDATION) ---
class BaseProcessor:
    def __init__(self, p_func, s_check, p_event): 
        self.progress = p_func 
        self.stop_sig_func = s_check
        self.pause_event = p_event 

    def check_state(self):
        if self.stop_sig_func(): raise Exception("Stopped")
        if not self.pause_event.is_set():
            self.progress(None, "Paused... (Waiting to Resume)", status_only=True)
            self.pause_event.wait() 
            if self.stop_sig_func(): raise Exception("Stopped")

class PdfProcessor(BaseProcessor):
    def flatten_or_ocr(self, src, dest, mode='flatten', dpi=300):
        temp = dest.parent / f"temp_{src.stem}"
        temp.mkdir(parents=True, exist_ok=True)
        try:
            info = pdfinfo_from_path(str(src), poppler_path=POPPLER_BIN)
            pages = info.get("Pages", 1)
            imgs = []
            
            for i in range(1, pages + 1):
                self.check_state() 
                self.progress((i/pages)*100, f"Page {i}/{pages}")
                gc.collect() 
                
                res = convert_from_path(str(src), dpi=dpi, first_page=i, last_page=i, poppler_path=POPPLER_BIN)
                if not res: continue
                img = res[0]
                
                if mode == 'ocr' and HAS_TESSERACT:
                    t_page = temp / f"page_{i}.jpg"
                    img.save(t_page, "JPEG", dpi=(int(dpi), int(dpi)))
                    f = temp / f"{i}.pdf"
                    with open(f, "wb") as o: 
                        o.write(pytesseract.image_to_pdf_or_hocr(str(t_page), extension='pdf'))
                    imgs.append(str(f))
                else:
                    f = temp / f"{i}.jpg"
                    img.convert('RGB').save(f, "JPEG", quality=85)
                    imgs.append(str(f))
                del res; del img
            
            self.check_state()
            self.progress(100, "Merging...")
            
            if mode == 'ocr' and HAS_TESSERACT:
                m = pypdf.PdfWriter()
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
            self.check_state()
            self.progress(50, "Processing Image...")
            with Image.open(src) as img:
                img.load() 
                r = min(w / img.width, 1.0)
                img.resize((int(img.width * r), int(img.height * r)), Image.Resampling.LANCZOS).convert('RGB').save(dest, "JPEG", quality=85)
            return True
        except Exception as e:
            if str(e) == "Stopped": raise
            return False
            
    def convert_to_pdf(self, src, dest):
        try:
            self.check_state()
            self.progress(50, "Converting to PDF...")
            with Image.open(src) as img:
                img.load() 
                img.convert('RGB').save(dest, "PDF")
            return True
        except Exception as e:
            if str(e) == "Stopped": raise
            return False

class OfficeProcessor(BaseProcessor):
    def sanitize(self, src, dest):
        try:
            self.check_state()
            if src.suffix.lower() not in {'.docx', '.xlsx'}:
                shutil.copy2(src, dest); return False
            
            if not zipfile.is_zipfile(src): raise Exception("Corrupt Office File")

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
        self.q = q
        self.stop_sig = False
        self.pause_event = threading.Event()
        self.pause_event.set()
        self.current_ws = None 

    def stop(self): self.stop_sig = True; self.pause_event.set()
    def pause(self): self.pause_event.clear()
    def resume(self): self.pause_event.set()

    def log(self, m, err=False):
        self.q.put(("log", m, err))
        if self.current_ws:
            try:
                with open(Path(self.current_ws) / "session_log.txt", "a", encoding="utf-8") as f:
                    f.write(f"[{datetime.now().strftime('%H:%M:%S')}] {m}\n")
            except: pass
        log_app(m, "ERROR" if err else "INFO")

    # --- STATE MANAGER ---
    def set_job_status(self, ws, stage, details=""):
        try:
            data = {
                "stage": stage,
                "last_update": datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                "details": details
            }
            with open(Path(ws) / "status.json", 'w') as f:
                json.dump(data, f, indent=4)
        except Exception as e:
            self.log(f"Status Write Error: {e}", True)

    def prog_main(self, v, t): self.q.put(("main_p", v, t))
    def prog_sub(self, v, t, status_only=False): 
        if status_only: self.q.put(("status", t))
        else: self.q.put(("sub_p", v, t))
    
    def get_hash(self, path, mode):
        if os.path.getsize(path) == 0: return None, "Zero-Byte File"
        
        if path.suffix.lower() == '.pdf' and mode != "Lightning":
            try:
                r = PdfReader(str(path), strict=False) 
                if len(r.pages) == 0: return None, "PDF has 0 Pages"
                
                if mode == "Standard":
                    txt = ""
                    for i in range(min(3, len(r.pages))): txt += r.pages[i].extract_text()
                    if len(r.pages) > 3:
                        for i in range(len(r.pages)-3, len(r.pages)): txt += r.pages[i].extract_text()
                    if len(txt.strip()) > 10: return hashlib.md5(f"{txt}{len(r.pages)}".encode()).hexdigest(), "Smart-Standard"
                elif mode == "Deep":
                    txt = "".join([p.extract_text() for p in r.pages])
                    if len(txt.strip()) > 10: return hashlib.md5(f"{txt}{len(r.pages)}".encode()).hexdigest(), "Smart-Deep"
            except Exception as e:
                if mode == "Deep": return None, f"PDF Corrupt: {str(e)[:20]}"
                pass 

        try:
            h = hashlib.md5()
            with open(path, 'rb') as f:
                for chunk in iter(lambda: f.read(4096), b""): h.update(chunk)
            if path.suffix.lower() == '.pdf' and mode != "Lightning": return h.hexdigest(), "Binary-Fallback"
            return h.hexdigest(), "Binary"
        except Exception as e: return None, f"Read-Error: {str(e)[:20]}"

    def run_inventory(self, d_str, ingest_mode):
        try:
            d = Path(d_str); start_time = time.time()
            ws = WORKSPACES_ROOT / f"{d.name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            m_dir = ws / "01_Master_Files"; m_dir.mkdir(parents=True); (ws/"00_Quarantine").mkdir()
            self.current_ws = str(ws); self.log(f"Inventory Start: {d} Mode:{ingest_mode}")
            self.q.put(("job", str(ws))); self.q.put(("ws", str(ws)))
            
            # Set Initial Status
            self.set_job_status(ws, "SCANNING", "Ingesting files...")

            files = [Path(r)/f for r,_,fs in os.walk(d) for f in fs]
            seen = {}; quarantined = 0; file_types = {}

            for i, f in enumerate(files):
                if self.stop_sig: break
                if not self.pause_event.is_set(): self.prog_sub(None, "Paused...", True); self.pause_event.wait()
                
                self.prog_main((i/len(files))*100, f"Scanning {i}/{len(files)}")
                ext = f.suffix.lower()
                if ext not in SUPPORTED_EXTENSIONS: continue
                file_types[ext] = file_types.get(ext, 0) + 1
                
                h, method = self.get_hash(f, ingest_mode)
                if not h: 
                    self.log(f"⚠️ Quarantine: {f.name} ({method})", True)
                    q_name = f"{uuid.uuid4()}_{sanitize_filename(f.name)}"
                    shutil.copy2(f, ws/"00_Quarantine"/q_name)
                    quarantined += 1
                    seen[q_name] = {"status": "QUARANTINE", "error_reason": method, "orig_name": f.name}
                    continue
                
                rel_path = str(f.relative_to(d))
                if h in seen: seen[h]['copies'].append(rel_path)
                else: seen[h] = {'master': rel_path, 'copies': [rel_path], 'name': f.name}
            
            self.log("Tagging...")
            total = len(seen)
            for i, (h, data) in enumerate(seen.items()):
                if self.stop_sig: break
                if data.get("status") == "QUARANTINE": continue
                safe_name = f"[{i+1:04d}]_{sanitize_filename(data['name'])}"
                shutil.copy2(d / data['master'], m_dir / safe_name)
                data['uid'] = safe_name; data['id'] = f"[{i+1:04d}]"
            
            stats = {"ingest_time": time.time()-start_time, "masters": total-quarantined, "quarantined": quarantined, "total_scanned": len(files), "types": file_types}
            with open(ws/"manifest.json", 'w') as f: json.dump(seen, f, indent=4)
            with open(ws/"stats.json", 'w') as f: json.dump(stats, f)
            
            # Update Final Status
            self.set_job_status(ws, "INGESTED", f"Masters: {stats['masters']}")
            
            self.log(f"Done. Masters: {stats['masters']}"); self.q.put(("job", str(ws))); self.q.put(("auto_open", str(m_dir))); self.q.put(("done",))
        except Exception as e: self.log(f"Error: {e}", True); self.q.put(("done",))

    def run_batch(self, ws_p, active_modes, val):
        try:
            ws = Path(ws_p); self.current_ws = str(ws)
            start_time = time.time(); src = ws/"01_Master_Files"; dst = ws/"02_Ready_For_Redistribution"; dst.mkdir(exist_ok=True)
            self.q.put(("ws", ws_p)); self.log(f"Batch Start: {active_modes}")
            
            self.set_job_status(ws, "PROCESSING", f"Running: {', '.join(active_modes)}")

            if 'flatten' in active_modes and not check_memory():
                self.log("⚠️ LOW MEMORY DETECTED - OCR may lag.", True)

            bots = {
                'pdf': PdfProcessor(lambda v,t,s=False: self.prog_sub(v,t,s), lambda: self.stop_sig, self.pause_event),
                'img': ImageProcessor(lambda v,t,s=False: self.prog_sub(v,t,s), lambda: self.stop_sig, self.pause_event),
                'office': OfficeProcessor(lambda v,t,s=False: self.prog_sub(v,t,s), lambda: self.stop_sig, self.pause_event)
            }
            
            fs = list(src.iterdir())
            for i, f in enumerate(fs):
                if self.stop_sig: break
                if not self.pause_event.is_set(): self.prog_sub(None, "Paused...", True); self.pause_event.wait()
                
                self.log(f"Processing: {f.name}") 
                self.prog_main((i/len(fs))*100, f"File {i+1}/{len(fs)}") 
                self.prog_sub(0, "Starting...") 
                
                ext = f.suffix.lower(); ok = False
                
                if 'flatten' in active_modes and ext == '.pdf': ok = bots['pdf'].flatten_or_ocr(f, dst/f.name, dpi=int(val))
                elif 'resize' in active_modes and ext in {'.jpg','.png','.jpeg'}: ok = bots['img'].resize(f, dst/f.name, CFG.get('resize_width')) 
                elif 'img2pdf' in active_modes and ext in {'.jpg','.png','.jpeg'}: ok = bots['img'].convert_to_pdf(f, dst/f"{f.stem}.pdf")
                elif 'sanitize' in active_modes and ext in {'.docx','.xlsx'}: ok = bots['office'].sanitize(f, dst/f.name)
                
                if not ok and not (dst/f.name).exists(): 
                    shutil.copy2(f, dst/f.name)
            
            update_stats_time(ws, "batch_time", time.time() - start_time)
            
            self.set_job_status(ws, "PROCESSED", f"Complete: {', '.join(active_modes)}")
            self.q.put(("job", str(ws))) 
            
            self.prog_main(100, "Done"); self.prog_sub(0, ""); self.q.put(("done",)); SystemUtils.open_file(dst)
        except Exception as e: self.log(f"Err: {e}", True); self.q.put(("done",))

    def run_distribute(self, ws_p, ext_src, ocr):
        try:
            ws = Path(ws_p); self.current_ws = str(ws)
            if not (ws/"manifest.json").exists():
                 self.log("CRITICAL: Manifest missing. Cannot distribute.", True)
                 self.q.put(("error", "Manifest missing. Re-run Ingest.")); self.q.put(("done",)); return

            start_time = time.time(); src_dir = Path(ext_src) if ext_src else ws/"02_Ready_For_Redistribution"; dst = ws/"Final_Delivery"
            self.log("Distribute Start"); 
            self.set_job_status(ws, "DISTRIBUTING", "Copying files...")

            if ocr and not check_memory(): self.log("⚠️ LOW MEMORY - OCR Warning", True)
            
            with open(ws/"manifest.json") as f: man = json.load(f)
            orphans = {f.name: f for f in src_dir.iterdir()}
            bot = PdfProcessor(lambda v,t,s=False: self.prog_sub(v,t,s), lambda: self.stop_sig, self.pause_event)
            
            for i, (h, d) in enumerate(man.items()):
                if self.stop_sig: break
                if not self.pause_event.is_set(): self.prog_sub(None, "Paused...", True); self.pause_event.wait()
                
                self.prog_main((i/len(man))*100, f"Dist {i+1}")
                if d.get("status") == "QUARANTINE": continue
                
                src = next((v for k,v in orphans.items() if k.startswith(d['id'])), None)
                if not src: continue
                
                if ocr and src.suffix=='.pdf':
                    c = ws/"OCR_Cache"; c.mkdir(exist_ok=True); f = c/src.name
                    if not f.exists(): bot.flatten_or_ocr(src, f, 'ocr')
                    src = f
                
                for c in d['copies']:
                    t = dst / c; t.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(src, t.with_suffix(src.suffix))
            
            q_src = ws / "00_Quarantine"
            if q_src.exists():
                q_dst = dst / "_QUARANTINED_FILES"; q_dst.mkdir(exist_ok=True)
                for qf in q_src.iterdir(): shutil.copy2(qf, q_dst / qf.name)

            update_stats_time(ws, "dist_time", time.time() - start_time)
            
            self.set_job_status(ws, "DISTRIBUTED", "Delivered to Final folder")
            self.q.put(("job", str(ws))) 
            
            self.prog_main(100, "Done"); self.prog_sub(0, ""); self.q.put(("done",)); SystemUtils.open_file(dst)
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

# --- 8. UI ---
class App:
    def __init__(self, root):
        self.root = root; self.root.title(f"DocRefine Pro {SystemUtils.CURRENT_VERSION} ({platform.system()})")
        self.root.geometry("1150x900")
        self.q = queue.Queue(); self.worker = Worker(self.q)
        self.start_t = 0; self.running = False; self.paused = False
        
        # --- UI LAYOUT ---
        left = tk.Frame(root, width=350); left.pack(side="left", fill="both", padx=10, pady=10)
        tk.Label(left, text="Workspace Dashboard", font=("Segoe UI", 12, "bold")).pack(pady=5)
        self.btn_new = tk.Button(left, text="+ New Ingest Job", command=self.ask_ingest_mode, bg="#e3f2fd", height=2); self.btn_new.pack(fill="x")
        
        btn_row = tk.Frame(left); btn_row.pack(fill="x", pady=5)
        self.btn_refresh = tk.Button(btn_row, text="↻ Refresh", command=self.load_jobs); self.btn_refresh.pack(side="left", fill="x", expand=True)
        self.btn_del = tk.Button(btn_row, text="🗑 Delete", command=self.safe_delete_job, bg="#ffcdd2"); self.btn_del.pack(side="right")
        self.btn_log = tk.Button(left, text="View App Log", command=self.open_app_log); self.btn_log.pack(anchor="w", pady=5)
        self.btn_open = tk.Button(left, text="Open Folder", command=self.open_f, state="disabled"); self.btn_open.pack(fill="x", pady=5)
        
        self.stats_fr = tk.LabelFrame(left, text="Stats", padx=5, pady=5); self.stats_fr.pack(fill="x")
        self.lbl_stats = tk.Label(self.stats_fr, text="Select a job...", anchor="w", justify="left"); self.lbl_stats.pack(fill="x")
        
        self.tree = ttk.Treeview(left, columns=("Name","Status","LastActive"), show="headings")
        for c, w in [("Name",140),("Status",70),("LastActive",110)]:
            self.tree.heading(c, text=c, command=lambda _c=c: self.sort_tree(self.tree,_c,False)); self.tree.column(c, width=w)
        self.tree.pack(fill="both", expand=True); self.tree.bind("<<TreeviewSelect>>", self.on_sel)
        
        right = tk.Frame(root); right.pack(side="right", fill="both", expand=True, padx=10, pady=10)
        self.nb = ttk.Notebook(right); self.nb.pack(fill="both", expand=True)
        
        self.tab_process = tk.Frame(self.nb); self.nb.add(self.tab_process, text=" ⚙️ Process ")
        self._build_process()
        self.tab_dist = tk.Frame(self.nb); self.nb.add(self.tab_dist, text=" 🚀 Distribute ")
        self._build_dist()
        self.tab_inspect = tk.Frame(self.nb); self.nb.add(self.tab_inspect, text=" 🔍 Inspector ")
        self._build_inspect()

        mon = tk.LabelFrame(right, text="Process Monitor", padx=10, pady=10); mon.pack(fill="x", pady=10)
        head = tk.Frame(mon); head.pack(fill="x")
        self.lbl_status = tk.Label(head, text="Ready", fg="blue", anchor="w"); self.lbl_status.pack(side="left", fill="x", expand=True)
        self.lbl_timer = tk.Label(head, text="00:00:00", font=("Consolas", 10)); self.lbl_timer.pack(side="right", padx=10)
        
        self.btn_stop = tk.Button(head, text="STOP", bg="red", fg="white", command=self.stop, state="disabled"); self.btn_stop.pack(side="right")
        self.btn_pause = tk.Button(head, text="PAUSE", command=self.toggle_pause, state="disabled", width=8); self.btn_pause.pack(side="right", padx=5)

        self.p_main = tk.DoubleVar(); ttk.Progressbar(mon, variable=self.p_main).pack(fill="x", pady=2)
        self.p_sub = tk.DoubleVar(); ttk.Progressbar(mon, variable=self.p_sub).pack(fill="x", pady=2)
        self.log_box = scrolledtext.ScrolledText(mon, height=8); self.log_box.pack(fill="both", expand=True)
        self.load_jobs(); self.root.after(100, self.poll)
        
        # --- AUTO-CHECK UPDATES ---
        threading.Thread(target=self.check_updates, daemon=True).start()

    def check_updates(self):
        """Secure, notification-only update check using standard lib."""
        repo = CFG.get("GITHUB_REPO")
        if not repo or "username" in repo: return # Skip if not configured
        
        url = f"https://api.github.com/repos/{repo}/releases/latest"
        try:
            # Use urllib to avoid dependency on requests
            with urllib.request.urlopen(url, timeout=5) as response:
                if response.status == 200:
                    data = json.loads(response.read().decode())
                    remote_ver = data.get("tag_name", "").strip().lower()
                    local_ver = SystemUtils.CURRENT_VERSION.strip().lower()
                    
                    # Simple string comparison (v74 vs v73)
                    # For robust semver, we'd need a parser, but this works for basic tags
                    if remote_ver > local_ver:
                        self.q.put(("update_avail", remote_ver, data.get("html_url")))
        except:
            # Silent fail for firewall/offline
            pass

    def _build_process(self):
        tk.Label(self.tab_process, text="Step 2: Batch Processing", font=("Segoe UI",10,"bold")).pack(anchor="w",pady=10,padx=10)
        self.chk_frame = tk.Frame(self.tab_process); self.chk_frame.pack(fill="x",padx=10)
        self.chk_vars = {} 
        ctrl = tk.Frame(self.tab_process); ctrl.pack(fill="x",pady=10,padx=10)
        tk.Label(ctrl, text="Quality (DPI):").pack(side="left")
        self.dpi_var = tk.StringVar(value="300")
        self.cb_dpi = ttk.Combobox(ctrl, textvariable=self.dpi_var, values=["150","300","600"], width=5, state="readonly"); self.cb_dpi.pack(side="left",padx=5)
        self.btn_prev = tk.Button(ctrl, text="Generate Preview", command=self.safe_preview, state="disabled"); self.btn_prev.pack(side="left",padx=15)
        self.btn_run = tk.Button(ctrl, text="Run Actions", command=self.safe_start_batch, bg="#e8f5e9", height=2, state="disabled"); self.btn_run.pack(side="right")

    def _build_dist(self):
        tk.Label(self.tab_dist, text="Step 3: Distribution", font=("Segoe UI",10,"bold")).pack(anchor="w",pady=10,padx=10)
        f = tk.Frame(self.tab_dist); f.pack(fill="x",padx=10)
        self.var_ext = tk.BooleanVar(); tk.Checkbutton(f, text="Source: External Folder", variable=self.var_ext).pack(anchor="w")
        self.var_ocr = tk.BooleanVar(); tk.Checkbutton(f, text="OCR (Searchable)", variable=self.var_ocr, command=self.warn_ocr).pack(anchor="w")
        self.btn_dist = tk.Button(self.tab_dist, text="Run Distribution", command=self.safe_start_dist, bg="#fff3e0", height=2, state="disabled"); self.btn_dist.pack(fill="x",padx=10,pady=20)

    def _build_inspect(self):
        f = tk.Frame(self.tab_inspect); f.pack(fill="both",expand=True,padx=5,pady=5)
        self.insp_tree = ttk.Treeview(f, columns=("ID","Name","Status","Copies"), show="headings")
        for c,w in [("ID",60),("Name",200),("Status",80),("Copies",50)]:
            self.insp_tree.heading(c, text=c, command=lambda _c=c: self.sort_tree(self.insp_tree,_c,False)); self.insp_tree.column(c, width=w)
        vsb = ttk.Scrollbar(f, orient="vertical", command=self.insp_tree.yview); self.insp_tree.configure(yscrollcommand=vsb.set)
        self.insp_tree.pack(side="left",fill="both",expand=True); vsb.pack(side="right",fill="y")

    # --- ACTIONS ---
    def check_run_btn(self, *args): self.btn_run.config(state="normal" if any(v.get() for v in self.chk_vars.values()) else "disabled")
    def sort_tree(self, t, c, r):
        l = [(t.set(k,c),k) for k in t.get_children('')]
        try: l.sort(key=lambda x: int(x[0]), reverse=r)
        except: l.sort(reverse=r)
        for i, (_,k) in enumerate(l): t.move(k,'',i)
        t.heading(c, command=lambda: self.sort_tree(t,c,not r))
    def warn_ocr(self): 
        if self.var_ocr.get():
            m = "Warning: OCR flattens pages.\n"
            if not HAS_PSUTIL: m += "NOTE: Low memory detection is unavailable.\n"
            elif not check_memory(): m += "CRITICAL: System memory is low!\n"
            if not messagebox.askyesno("Confirm", m+"Continue?"): self.var_ocr.set(False)
    def wrap(self, func, *args):
        try: func(*args)
        except Exception as e: self.q.put(("error", str(e))); self.q.put(("done",))

    # --- POPUP ---
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
            top.destroy()
            d = filedialog.askdirectory()
            if d: self.toggle(False); threading.Thread(target=self.wrap, args=(self.worker.run_inventory,d,mode.get()), daemon=True).start()
        tk.Button(top, text="Select Folder & Start", command=go, bg="#2196f3", fg="white", height=2).pack(fill="x", padx=20, pady=20, side="bottom")

    def safe_start_batch(self):
        ws = self.get_ws()
        if ws: self.toggle(False); threading.Thread(target=self.wrap, args=(self.worker.run_batch, str(ws), [k for k,v in self.chk_vars.items() if v.get()], self.cb_dpi.get()), daemon=True).start()
    def safe_start_dist(self):
        ws=self.get_ws(); src=filedialog.askdirectory() if self.var_ext.get() else None
        if ws and (not self.var_ext.get() or src): self.toggle(False); threading.Thread(target=self.wrap, args=(self.worker.run_distribute, str(ws), src, self.var_ocr.get()), daemon=True).start()
    def safe_preview(self):
        ws=self.get_ws()
        if ws: self.toggle(False); threading.Thread(target=self.wrap, args=(self.worker.run_preview, str(ws), self.cb_dpi.get()), daemon=True).start()

    # --- CONTROL ---
    def stop(self): self.worker.stop(); self.btn_stop.config(state="disabled"); self.btn_pause.config(state="disabled")
    def toggle_pause(self):
        if self.paused:
            self.worker.resume(); self.btn_pause.config(text="PAUSE", bg="SystemButtonFace")
            self.lbl_status.config(text="Resuming...", fg="blue")
        else:
            self.worker.pause(); self.btn_pause.config(text="RESUME", bg="yellow")
        self.paused = not self.paused

    def poll(self):
        try:
            for _ in range(50):
                m = self.q.get_nowait()
                if m[0]=='log': self.log_box.insert(tk.END, m[1]+"\n"); self.log_box.see(tk.END)
                elif m[0]=='ws': 
                    self.log_box.delete(1.0, tk.END)
                    if (Path(m[1])/"session_log.txt").exists(): self.log_box.insert(tk.END, (Path(m[1])/"session_log.txt").read_text(encoding="utf-8"))
                elif m[0]=='main_p': self.p_main.set(m[1]); self.lbl_status.config(text=m[2], fg="blue")
                elif m[0]=='status': self.lbl_status.config(text=m[1], fg="orange")
                elif m[0]=='sub_p': self.p_sub.set(m[1])
                elif m[0]=='job': self.load_jobs(m[1])
                elif m[0]=='auto_open': SystemUtils.open_file(m[1])
                elif m[0]=='error': messagebox.showerror("Error", m[1])
                elif m[0]=='update_avail':
                    if messagebox.askyesno("Update Available", f"Version {m[1]} is available.\nDownload now?"):
                        webbrowser.open(m[2])
                elif m[0]=='done': self.toggle(True); self.lbl_status.config(text="Done", fg="green"); self.p_sub.set(0)
        except: pass
        if self.running and not self.paused: self.lbl_timer.config(text=str(timedelta(seconds=int(time.time()-self.start_t))))
        self.root.after(100, self.poll)

    def load_jobs(self, sel=None):
        self.tree.delete(*self.tree.get_children())
        try: items = sorted(WORKSPACES_ROOT.iterdir(), key=lambda x: x.stat().st_mtime, reverse=True)
        except: items = []
        for d in items:
            if d.is_dir():
                s = "Empty"
                stat_file = d / "status.json"
                if stat_file.exists():
                    try:
                        with open(stat_file, 'r') as f: s = json.load(f).get("stage", "Unknown")
                    except: pass
                elif (d/"Final_Delivery").exists(): s = "DISTRIBUTED"
                elif (d/"02_Ready_For_Redistribution").exists(): s = "PROCESSED"
                elif (d/"01_Master_Files").exists(): s = "INGESTED"
                
                try: ts = (d/"session_log.txt").stat().st_mtime if (d/"session_log.txt").exists() else d.stat().st_mtime
                except: ts = time.time()
                i = self.tree.insert("", "end", values=(d.name, s, datetime.fromtimestamp(ts).strftime('%Y-%m-%d %H:%M')))
                if sel and d.name == Path(sel).name: self.tree.selection_set(i)

    def on_sel(self, e):
        if self.running: return
        ws = self.get_ws()
        if not ws: self.btn_open.config(state="disabled"); self.insp_tree.delete(*self.insp_tree.get_children()); return
        self.btn_open.config(state="normal")
        try:
            with open(ws/"stats.json") as f: s = json.load(f)
            self.lbl_stats.config(text=f"Files: {s.get('total_scanned',0)} | Masters: {s.get('masters',0)} | Time: {str(timedelta(seconds=int(s.get('ingest_time',0)+s.get('batch_time',0)+s.get('dist_time',0))))}")
        except: self.lbl_stats.config(text="Stats: N/A")
        
        for w in self.chk_frame.winfo_children(): w.destroy()
        self.chk_vars.clear(); types = set()
        if (ws/"01_Master_Files").exists(): types = {f.suffix.lower() for f in (ws/"01_Master_Files").rglob('*') if f.is_file()}
        
        def ac(l,k,t): 
            v=tk.BooleanVar(); v.trace_add("write", self.check_run_btn)
            c=tk.Checkbutton(self.chk_frame,text=l,variable=v); c.pack(side="left",padx=10); self.chk_vars[k]=v; ToolTip(c,t)
        if '.pdf' in types: ac("Flatten PDFs","flatten","Convert pages to images.")
        if any(x in types for x in ['.jpg','.png']): ac("Resize Images","resize","Resize to 1920px."); ac("Images to PDF","img2pdf","Bundle images.")
        if any(x in types for x in ['.docx','.xlsx']): ac("Sanitize Office","sanitize","Remove metadata.")
        self.btn_run.config(state="disabled"); self.btn_prev.config(state="normal" if '.pdf' in types else "disabled")
        
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

    def get_ws(self): s = self.tree.selection(); return WORKSPACES_ROOT / self.tree.item(s[0])['values'][0] if s else None
    def toggle(self, enable):
        self.running = not enable; s = "normal" if enable else "disabled"
        if not enable: self.start_t = time.time(); self.paused = False; self.btn_pause.config(text="PAUSE", bg="SystemButtonFace")
        self.tree.config(selectmode="browse" if enable else "none")
        for b in [self.btn_new, self.btn_refresh, self.btn_del, self.cb_dpi]: b.config(state="readonly" if b==self.cb_dpi and enable else s)
        for c in self.chk_frame.winfo_children() + [self.btn_dist, self.btn_prev, self.btn_open]: 
            try: c.configure(state=s) 
            except: pass
        if enable: self.check_run_btn() 
        else: self.btn_run.config(state="disabled")
        self.btn_stop.config(state="normal" if not enable else "disabled")
        self.btn_pause.config(state="normal" if not enable else "disabled")
        if enable: self.on_sel(None)
    def open_f(self): SystemUtils.open_file(self.get_ws())
    def open_app_log(self): SystemUtils.open_file(LOG_PATH)
    def safe_delete_job(self):
        ws = self.get_ws()
        if ws and messagebox.askyesno("Confirm", "Delete Job?"):
            def _del():
                try: shutil.rmtree(ws)
                except Exception as e: self.q.put(("error", f"Could not delete: {e}"))
                finally: self.q.put(("job", None))
            threading.Thread(target=_del, daemon=True).start()

if __name__ == "__main__":
    try: root = tk.Tk(); App(root); root.mainloop()
    except Exception as e: messagebox.showerror("Fatal", str(e))