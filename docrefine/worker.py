# SAVE AS: docrefine/worker.py
import threading
import time
import json
import shutil
import hashlib
import uuid
import os
import csv
import re
import tempfile
import concurrent.futures
from pathlib import Path
from datetime import datetime

# Local Package Imports
from .config import CFG, SystemUtils, log_app, WORKSPACES_ROOT, LOG_PATH, JSON_LOG_PATH, Constants
from .reporting import generate_job_report
from .core.events import AppEvent, EventType
from .processing import (
    PdfProcessor, 
    ImageProcessor, 
    OfficeProcessor, 
    POPPLER_BIN, 
    convert_from_path
)

try:
    import psutil
    HAS_PSUTIL = True
except ImportError:
    HAS_PSUTIL = False

try:
    from pypdf import PdfReader
except ImportError:
    PdfReader = None

SUPPORTED_EXTENSIONS = {'.pdf', '.doc', '.docx', '.jpg', '.png', '.xls', '.xlsx', '.csv', '.jpeg'}

# ==============================================================================
#   HELPER FUNCTIONS
# ==============================================================================

def sanitize_filename(name):
    return re.sub(r'[<>:"/\\|?*]', '_', name)

def promote_duplicate_to_master(ws_path, master_uid, dup_path):
    """Promote a duplicate file into its own standalone master entry.

    Detaches the chosen copy from the group it was filed under, copies it into
    the Master Files folder with the next available [NNNN] id, and records it as
    a new unique master in manifest.json. Returns the new master id.
    """
    ws = Path(ws_path)
    dup_path = Path(dup_path)
    manifest_path = ws / "manifest.json"

    with open(manifest_path, "r", encoding="utf-8") as f:
        manifest = json.load(f)

    entry = next((v for v in manifest.values() if v.get("uid") == master_uid), None)
    if entry is None:
        raise ValueError("Master entry not found in manifest.")

    root = Path(entry.get("root", ""))
    try:
        dup_rel = str(dup_path.relative_to(root))
    except ValueError:
        dup_rel = dup_path.name

    # Next available master number, e.g. [0007].
    next_idx = 1
    for v in manifest.values():
        vid = str(v.get("id", ""))
        if vid.startswith("[") and vid.endswith("]"):
            try:
                next_idx = max(next_idx, int(vid.strip("[]")) + 1)
            except ValueError:
                pass

    new_id = f"[{next_idx:04d}]"
    new_uid = f"{new_id}_{sanitize_filename(dup_path.name)}"
    master_dir = ws / Constants.DIR_MASTER
    master_dir.mkdir(parents=True, exist_ok=True)
    shutil.copy2(dup_path, master_dir / new_uid)

    # Detach the promoted copy from its old group.
    if dup_rel in entry.get("copies", []):
        entry["copies"].remove(dup_rel)

    # Register the new standalone master.
    manifest[f"PROMOTED::{new_uid}"] = {
        "master": dup_rel,
        "copies": [dup_rel],
        "name": dup_path.name,
        "root": entry.get("root", ""),
        "uid": new_uid,
        "id": new_id,
    }

    _atomic_write_json(manifest_path, manifest, indent=4)

    return new_id

STATS_LOCK = threading.Lock()

def _atomic_copy(src, dst):
    """Copy so the destination only ever exists complete.

    Delivery folders are resumable: any non-empty file at the destination is
    taken as finished. A copy interrupted part-way would therefore be shipped as
    though it were the real thing, so the bytes land in a temp file first.
    """
    fd, tmp = tempfile.mkstemp(dir=str(Path(dst).parent), suffix=".part")
    os.close(fd)
    try:
        shutil.copyfile(src, tmp)
        shutil.copystat(src, tmp)
        os.replace(tmp, dst)
        tmp = None
    finally:
        if tmp and os.path.exists(tmp):
            try: os.remove(tmp)
            except OSError: pass


def _atomic_write_json(path, data, indent=None):
    """Write JSON so the file is either the old version or the new one, never half.

    manifest.json is the single source of truth for organize/distribute/export —
    a 37k-file manifest takes ~0.5s to serialise, and an interruption inside that
    window used to leave truncated JSON that made the whole ingest unreadable.
    """
    path = Path(path)
    fd, tmp = tempfile.mkstemp(dir=str(path.parent), suffix=".tmp")
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=indent)
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp, path)
        tmp = None
    finally:
        if tmp and os.path.exists(tmp):
            try: os.remove(tmp)
            except OSError: pass


def update_stats_time(ws, cat, sec):
    with STATS_LOCK:
        try:
            p = Path(ws) / "stats.json"
            s = {}
            if p.exists():
                with open(p, 'r') as f: s = json.load(f)
            s[cat] = s.get(cat, 0.0) + sec
            _atomic_write_json(p, s, indent=4)
        except: pass

class Worker:
    def __init__(self, callback): 
        self.callback = callback 
        self.stop_sig = False
        self.pause_event = threading.Event()
        self.pause_event.set()
        self.current_ws = None 
        self._last_update = {}

    def emit(self, event: AppEvent):
        """Bridge to the observer (UI/CLI)"""
        if self.callback:
            self.callback(event)

    def stop(self): 
        self.stop_sig = True
        self.pause_event.set()

    def pause(self): 
        self.pause_event.clear()

    def resume(self): 
        self.pause_event.set()

    def log(self, m, err=False):
        level = "ERROR" if err else "INFO"
        self.emit(AppEvent.log(m, level))
        log_app(m, level, structured_data={"ws": self.current_ws})

    def set_job_status(self, ws, stage, details=""):
        try:
            data = { "stage": stage, "last_update": datetime.now().strftime('%Y-%m-%d %H:%M:%S'), "details": details }
            _atomic_write_json(Path(ws) / "status.json", data, indent=4)
        except: pass

    def prog_main(self, v, t): 
        self.emit(AppEvent.progress(v, t))
    
    def prog_sub(self, v, t, status_only=False): 
        tid = threading.get_ident()
        now = time.time()
        
        if tid not in self._last_update:
            self._last_update[tid] = 0
            
        if (now - self._last_update[tid]) > 0.1: 
            self.emit(AppEvent(EventType.SLOT_UPDATE, {"tid": tid, "text": t, "percent": v}))
            self._last_update[tid] = now

    def get_hash(self, path, mode):
        if os.path.getsize(path) == 0: return None, "Zero-Byte File"
        if path.suffix.lower() == '.pdf' and mode != "Lightning":
            try:
                if PdfReader is None: raise Exception("pypdf not available")
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
        master = ws / Constants.DIR_MASTER / file_uid
        base_cache = ws / Constants.DIR_READY
        
        # EXTRACT ID: Safely grab the [0001] tag to allow for wildcard matching
        file_id = None
        if file_uid.startswith("[") and "]" in file_uid:
            file_id = file_uid.split("]")[0] + "]"
        
        def find_in_dir(d, stem):
            if d.exists():
                # 1. Exact Name Match
                if (d / file_uid).exists(): return d / file_uid
                
                # 2. Smart Wildcard Match
                for f in d.iterdir():
                    if f.stem == stem: return f
                    if file_id and f.name.startswith(file_id): return f
            return None

        stem = Path(file_uid).stem
        
        if "Force: OCR" in priority_mode:
            f = find_in_dir(base_cache/"OCR", stem)
            return f if f else (master if master.exists() else None)
            
        elif "Force: Flattened" in priority_mode:
            f = find_in_dir(base_cache/"Flattened", stem)
            return f if f else (master if master.exists() else None)
            
        elif "Force: Original" in priority_mode:
            return master if master.exists() else None
            
        else: 
            for sub in ["OCR", "Flattened", "Resized", "Sanitized", "Standard"]:
                f = find_in_dir(base_cache/sub, stem)
                if f: return f
            return master if master.exists() else None

    def run_inventory(self, d_str, ingest_mode):
        try:
            self.stop_sig = False
            self.resume()
            
            d = Path(d_str)
            start_time = time.time()
            ws = WORKSPACES_ROOT / f"{d.name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            m_dir = ws / Constants.DIR_MASTER
            m_dir.mkdir(parents=True); (ws/Constants.DIR_QUARANTINE).mkdir()
            self.current_ws = str(ws)
            self.log(f"Inventory Start: {d}")
            
            self.emit(AppEvent(EventType.JOB_DATA, str(ws)))
            self.set_job_status(ws, "SCANNING", "Ingesting...")
            
            files = [Path(r)/f for r,_,fs in os.walk(d) for f in fs]
            files = [f for f in files if f.suffix.lower() in SUPPORTED_EXTENSIONS]
            
            seen = {}; quarantined = 0
            
            self.emit(AppEvent(EventType.WORKER_CONFIG, 1))

            for i, f in enumerate(files):
                if self.stop_sig: break
                if not self.pause_event.is_set(): 
                    self.prog_sub(None, "Paused...", True)
                    self.pause_event.wait()
                
                # FIX: 100% Math Calculation
                self.prog_main(((i+1)/len(files))*100, f"Scanning {i+1}/{len(files)}")
                self.prog_sub(None, f"Hashing: {f.name}", True)
                
                try:
                    h, method = self.get_hash(f, ingest_mode)
                    if not h:
                        self.log(f"⚠️ Quarantine: {f.name}", True)
                        q_name = f"{uuid.uuid4()}_{sanitize_filename(f.name)}"
                        shutil.copy2(f, ws/Constants.DIR_QUARANTINE/q_name)
                        quarantined += 1
                        # Record quarantined files in the manifest so they stay
                        # visible in the Inspector and the exported CSV.
                        seen[f"QUARANTINE::{q_name}"] = {
                            'status': 'QUARANTINE',
                            'id': f"Q{quarantined:04d}",
                            'name': f.name,
                            'orig_name': q_name,
                            'error_reason': method,
                            'root': str(d),
                            'master': str(f.relative_to(d)),
                            'copies': [str(f.relative_to(d))],
                        }
                        continue
                    
                    rel = str(f.relative_to(d))
                    if h in seen: seen[h]['copies'].append(rel)
                    else: seen[h] = {'master': rel, 'copies': [rel], 'name': f.name, 'root': str(d)}
                except Exception as e:
                    self.log(f"Hash Error: {e}", True)

            if self.stop_sig: 
                self.log("Ingest Stopped by User.")
                self.emit(AppEvent(EventType.DONE))
                return

            self.log("Tagging...")
            master_count = 0
            for h, data in seen.items():
                if self.stop_sig: break
                if data.get('status') == 'QUARANTINE': continue  # not a master copy
                master_count += 1
                safe_name = f"[{master_count:04d}]_{sanitize_filename(data['name'])}"
                shutil.copy2(d / data['master'], m_dir / safe_name)
                data['uid'] = safe_name; data['id'] = f"[{master_count:04d}]"
            total = master_count
            
            if self.stop_sig:
                self.log("Ingest stopped by user during tagging.")
                self.emit(AppEvent(EventType.DONE)); return

            stats = {
                "ingest_time": time.time()-start_time, 
                "masters": total, 
                "quarantined": quarantined,
                "total_scanned": len(files)
            }
            _atomic_write_json(ws/"manifest.json", seen, indent=4)
            _atomic_write_json(ws/"stats.json", stats)
            self.set_job_status(ws, "INGESTED", f"Masters: {total}")
            self.log(f"Done. Masters: {total}")
            
            self.emit(AppEvent(EventType.JOB_DATA, str(ws)))
            self.emit(AppEvent(EventType.DONE))
            
        except Exception as e: 
            self.log(f"Error: {e}", True)
            self.emit(AppEvent(EventType.DONE))

    def process_file_task(self, f, bots, options, base_dst):
        if self.stop_sig: return None
        result = {'file': f.name, 'orig_size': f.stat().st_size, 'new_size': 0, 'ok': False, 'skipped': False}
        try:
            self.emit(AppEvent.status("PROCESSING", f"Refining: {f.name}", "blue"))
            
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

            # FIX: Skip if already exists
            if dst_file.exists() and dst_file.stat().st_size > 0:
                result['new_size'] = dst_file.stat().st_size
                result['ok'] = True
                result['skipped'] = True
                return result

            if ext == '.pdf':
                mode = options.get('pdf_mode', 'none')
                if mode == 'flatten': ok = bots['pdf'].flatten_or_ocr(f, dst_file, 'flatten', dpi=dpi_val)
                elif mode == 'ocr': ok = bots['pdf'].flatten_or_ocr(f, dst_file, 'ocr', dpi=dpi_val)
            elif ext in {'.jpg','.png'}:
                if options.get('resize'): ok = bots['img'].resize(f, dst_file, CFG.get('resize_width'))
                if options.get('img2pdf'): ok = bots['img'].convert_to_pdf(f, final_dst_dir/f"{f.stem}.pdf")
            elif ext in {'.docx','.xlsx'}:
                if options.get('sanitize'): ok = bots['office'].sanitize(f, dst_file)

            # FIX: Pause/Stop Safety Net - DO NOT COPY if stopped
            if self.stop_sig or not self.pause_event.is_set():
                # If we stopped, we don't copy the original. We just fail the task safely.
                return None

            if not ok and not dst_file.exists():
                _atomic_copy(f, dst_file)
            
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
            self.stop_sig = False
            self.resume()
            
            ws = Path(ws_p); self.current_ws = str(ws)
            start_time = time.time()
            
            # --- CHAINED WORKFLOW ROUTING ---
            if options.get("chain_flattened"):
                src = ws / Constants.DIR_READY / "Flattened"
                if not src.exists() or not any(src.iterdir()):
                    self.log("CRITICAL: Flattened cache is empty. Please run Flatten first.", True)
                    self.emit(AppEvent(EventType.DONE))
                    return
                self.log("Routing: Sourcing files from Flattened cache (Chained Workflow).")
            else:
                src = ws / Constants.DIR_MASTER
            # --------------------------------
            
            dst = ws/Constants.DIR_READY; dst.mkdir(exist_ok=True)
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

            self.emit(AppEvent(EventType.WORKER_CONFIG, max_workers))
            
            file_results = []
            with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
                futures = {executor.submit(self.process_file_task, f, bots, options, dst): f for f in fs}
                for i, future in enumerate(concurrent.futures.as_completed(futures)):
                    if self.stop_sig: break
                    self.prog_main(((i+1)/len(fs))*100, f"Refining {i+1}/{len(fs)}")
                    try: 
                        r = future.result()
                        if r: file_results.append(r)
                    except Exception as e: self.log(f"Thread Err: {e}", True)

            if self.stop_sig: 
                self.log("Batch Stopped by User.")
                self.emit(AppEvent(EventType.DONE))
                return

            update_stats_time(ws, "batch_time", time.time() - start_time)
            self.set_job_status(ws, "PROCESSED", "Complete")
            
            rpt = generate_job_report(ws, "Content Refinement Batch", file_results)
            if rpt: self.log(f"Receipt Generated: {Path(rpt).name}")
            
            self.emit(AppEvent(EventType.JOB_DATA, str(ws))) 
            self.prog_main(100, "Done")
            self.emit(AppEvent(EventType.DONE))
            self.emit(AppEvent(EventType.NOTIFICATION, {"title": "Batch Complete", "msg": "Batch processing finished.", "open_path": str(dst)}))
            
        except Exception as e: 
            self.log(f"Err: {e}", True)
            self.emit(AppEvent(EventType.DONE))

    def run_organize(self, ws_p, priority_mode):
        try:
            self.stop_sig = False; self.resume()
            ws = Path(ws_p); self.current_ws = str(ws)
            start_time = time.time()
            out = ws / Constants.DIR_ORGANIZED
            m = out/"Unique_Masters"; q = out/"Quarantine"
            for p in [m,q]: p.mkdir(parents=True, exist_ok=True)
            
            self.log(f"Unique Export ({priority_mode})")
            with open(ws/"manifest.json") as f: man = json.load(f)
            total = len(man)
            
            self.emit(AppEvent(EventType.WORKER_CONFIG, 1))

            dup_csv = out / "duplicates_report.csv"
            with open(dup_csv, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(["Master_Filename", "Duplicate_Location"])
                
                for i, (h, data) in enumerate(man.items()):
                    if self.stop_sig: break
                    self.prog_main(((i+1)/total)*100, "Exporting Unique...")
                    self.emit(AppEvent(EventType.SLOT_UPDATE, {"tid": threading.get_ident(), "text": f"Exporting: {data['name']}", "percent": None}))
                    
                    if data.get("status") == "QUARANTINE": 
                        for f in (ws/Constants.DIR_QUARANTINE).glob("*"):
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

            if self.stop_sig:
                self.log("Unique export stopped by user.")
                self.emit(AppEvent(EventType.DONE)); return

            update_stats_time(ws, "organize_time", time.time() - start_time)
            self.set_job_status(ws, "ORGANIZED", "Done")
            
            rpt = generate_job_report(ws, f"Unique Export ({priority_mode})")
            
            self.emit(AppEvent(EventType.JOB_DATA, str(ws))) 
            self.prog_main(100, "Done")
            self.emit(AppEvent(EventType.DONE))
            self.emit(AppEvent(EventType.NOTIFICATION, {"title": "Organization Complete", "msg": "Files organized.", "open_path": str(out)}))
            
        except Exception as e: 
            self.log(f"Err: {e}", True)
            self.emit(AppEvent(EventType.DONE))

    def run_distribute(self, ws_p, ext_src, priority_mode):
        try:
            self.stop_sig = False; self.resume()
            ws = Path(ws_p); self.current_ws = str(ws)
            if not (ws/"manifest.json").exists():
                 self.log("CRITICAL: Manifest missing.", True)
                 self.emit(AppEvent(EventType.ERROR, "Manifest missing."))
                 self.emit(AppEvent(EventType.DONE))
                 return

            start_time = time.time(); 
            dst = ws / "Final_Delivery"
            self.log(f"Reconstruction Start ({priority_mode})")
            self.set_job_status(ws, "DISTRIBUTING", "Reconstructing...")
            
            with open(ws/"manifest.json") as f: man = json.load(f)
            
            orphans = {}
            if ext_src:
                 orphans = {f.name: f for f in Path(ext_src).iterdir()}

            self.emit(AppEvent(EventType.WORKER_CONFIG, 1))

            for i, (h, d) in enumerate(man.items()):
                if self.stop_sig: break
                self.prog_main(((i+1)/len(man))*100, f"Recon {i+1}")
                self.emit(AppEvent(EventType.SLOT_UPDATE, {"tid": threading.get_ident(), "text": f"Copying: {d['name']}", "percent": None}))
                
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
            
            if self.stop_sig:
                self.log("Distribution stopped by user.")
                self.emit(AppEvent(EventType.DONE)); return

            q_src = ws / Constants.DIR_QUARANTINE
            if q_src.exists():
                q_dst = dst / "_QUARANTINED_FILES"; q_dst.mkdir(parents=True, exist_ok=True) 
                for qf in q_src.iterdir(): shutil.copy2(qf, q_dst / qf.name)

            update_stats_time(ws, "dist_time", time.time() - start_time)
            self.set_job_status(ws, "DISTRIBUTED", "Done")
            
            rpt = generate_job_report(ws, "Full Reconstruction")
            
            self.emit(AppEvent(EventType.JOB_DATA, str(ws))) 
            self.prog_main(100, "Done")
            self.emit(AppEvent(EventType.DONE))
            self.emit(AppEvent(EventType.NOTIFICATION, {"title": "Distribution Complete", "msg": "Reconstruction finished.", "open_path": str(dst)}))
            
        except Exception as e: 
            self.log(f"Err: {e}", True)
            self.emit(AppEvent(EventType.DONE))

    def run_full_export(self, ws_p):
        try:
            self.stop_sig = False; self.resume()
            ws = Path(ws_p); self.current_ws = str(ws)
            if not (ws/"manifest.json").exists():
                self.log("No manifest for this job — ingest it first.", True)
                self.emit(AppEvent(EventType.ERROR, "This job has no manifest yet — run an ingest first."))
                self.emit(AppEvent(EventType.DONE)); return

            rpt_dir = ws / Constants.DIR_REPORTS
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
                        self.prog_main(((i+1)/total)*100, "Writing CSV...")
                        
                        uid = data.get('id', '?')
                        status = data.get('status', 'OK')
                        name = data.get('name', '?')
                        master_rel = data.get('master', '')
                        
                        if status == "QUARANTINE":
                            orig = data.get('orig_name', name)
                            writer.writerow([uid, status, orig, "N/A - Quarantined", Constants.DIR_QUARANTINE, "Binary", h, 0, data.get('error_reason', '')])
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
                self.emit(AppEvent(EventType.ERROR, "Could not write CSV.\nPlease close the file in Excel and try again."))
                self.emit(AppEvent(EventType.DONE))
                return

            if self.stop_sig:
                self.log("CSV export stopped by user.")
                self.emit(AppEvent(EventType.DONE)); return

            self.log(f"Exported: {csv_path.name}")
            self.emit(AppEvent(EventType.JOB_DATA, str(ws))) 
            self.prog_main(100, "Done")
            self.emit(AppEvent(EventType.DONE))
            self.emit(AppEvent(EventType.NOTIFICATION, {"title": "CSV Exported", "msg": "Inventory saved.", "open_path": str(rpt_dir)}))

        except Exception as e: 
            self.log(f"Err: {e}", True)
            self.emit(AppEvent(EventType.DONE))

    def run_preview(self, ws_p, dpi):
        try:
            self.stop_sig = False; self.resume()
            ws = Path(ws_p); self.current_ws = str(ws)
            src = ws/Constants.DIR_MASTER; pdf = next(src.glob("*.pdf"), None)
            
            if not pdf: 
                self.emit(AppEvent.status("PREVIEW", "No PDF found.", "red"))
                self.emit(AppEvent(EventType.DONE))
                return
                
            for old in ws.glob("PREVIEW_*.pdf"): 
                try: os.remove(old)
                except: pass
            
            out = ws / f"PREVIEW_{int(time.time())}.pdf"
            imgs = convert_from_path(str(pdf), dpi=int(dpi), first_page=1, last_page=1, poppler_path=POPPLER_BIN)
            if imgs: 
                imgs[0].save(out, "PDF", resolution=float(dpi))
                self.emit(AppEvent(EventType.NOTIFICATION, {"title": "Preview Ready", "msg": "Opening preview...", "open_path": str(out)}))
            
            self.emit(AppEvent.status("PREVIEW", "Preview Generated", "green"))
            self.emit(AppEvent(EventType.DONE))
            
        except: 
            self.emit(AppEvent(EventType.DONE))

    REBRAND_PLAN_COLUMNS = ["file", "action", "doc_type", "product", "asset_type",
                            "manufacturer", "title", "pages", "confidence", "source", "notes"]

    def run_rebrand_analyze(self, src_dir, out_csv=None):
        """Read every PDF in a folder with the local model and write a review sheet."""
        try:
            self.stop_sig = False
            self.resume()
            from .classify import (classify_document, ollama_status, start_server,
                                   DEFAULT_MODEL, OLLAMA_DOWNLOAD_URL)
            from .reviews import plan_path_for
            src = Path(src_dir)
            plan = Path(out_csv) if out_csv else plan_path_for(src)
            self.current_ws = str(src)

            pdfs = [Path(r) / f for r, _, fs in os.walk(src) for f in fs if f.lower().endswith(".pdf")]
            if not pdfs:
                self.log("No PDFs found in the selected folder.", True)
                self.emit(AppEvent(EventType.DONE)); return

            # Make the local AI usable: auto-start the server if it's installed but idle.
            status = ollama_status()
            if status["state"] == "not_running":
                self.log("Ollama is installed but not running — starting it…")
                start_server()
                status = ollama_status()
            have_llm = status["state"] == "ready"
            if not have_llm:
                if status["state"] == "no_model":
                    self.log(f"Ollama is running but model '{DEFAULT_MODEL}' isn't downloaded — "
                             f"using filenames. Get it via Rebrand → Analyze (offer to download), "
                             f"or run: ollama pull {DEFAULT_MODEL}", True)
                elif status["state"] == "not_installed":
                    self.log(f"Ollama not installed — using filenames. Install it (free, offline) "
                             f"from {OLLAMA_DOWNLOAD_URL}", True)
                else:
                    self.log("Ollama unavailable — using filenames. Review the sheet carefully.", True)
            else:
                self.log("Local AI ready — classifying with Ollama.")
            self.log(f"Analyzing {len(pdfs)} PDFs for the review sheet")
            self.emit(AppEvent(EventType.WORKER_CONFIG, 1))

            rows = []
            for i, pdf in enumerate(pdfs):
                if self.stop_sig: break
                if not self.pause_event.is_set():
                    self.prog_sub(None, "Paused...", True); self.pause_event.wait()
                self.prog_main(((i + 1) / len(pdfs)) * 100, f"Analyzing {i + 1}/{len(pdfs)}")
                self.emit(AppEvent(EventType.SLOT_UPDATE, {"tid": threading.get_ident(), "text": pdf.name, "percent": None}))

                info = classify_document(pdf)
                try:
                    pages = len(PdfReader(str(pdf)).pages)
                except Exception:
                    pages = ""
                rows.append({
                    "file": str(pdf.relative_to(src)), "action": info["action"],
                    "doc_type": info.get("doc_type", ""), "product": info.get("product", ""),
                    "asset_type": info.get("asset_type", ""), "manufacturer": info.get("manufacturer", ""),
                    "title": info.get("title", ""), "pages": pages,
                    "confidence": info.get("confidence", 0), "source": info.get("source", ""),
                    "notes": info.get("notes", ""),
                })

            if self.stop_sig:
                self.log("Analysis stopped by user.")
                self.emit(AppEvent(EventType.DONE)); return

            # Never silently discard a sheet someone already reviewed.
            plan.parent.mkdir(parents=True, exist_ok=True)
            if plan.exists():
                bak = plan.with_name(plan.stem + ".previous" + plan.suffix)
                try:
                    shutil.copy2(plan, bak)
                    self.log(f"Previous review sheet kept as {bak.name}")
                except Exception:
                    pass

            from .reviews import write_plan
            from .rebrand import ASSET_TYPE_TITLES
            try:
                write_plan(plan, rows, self.REBRAND_PLAN_COLUMNS, src_root=src,
                           asset_types=sorted(ASSET_TYPE_TITLES))
            except Exception as e:
                # Never lose an hour of classification to a spreadsheet problem.
                plan = plan.with_suffix(".csv")
                self.log(f"Could not write the Excel sheet ({e}); wrote CSV instead.", True)
                write_plan(plan, rows, self.REBRAND_PLAN_COLUMNS)

            n_reb = sum(1 for r in rows if r["action"] == "rebrand")
            n_leave = len(rows) - n_reb
            self.log(f"Review sheet ready: {plan}  ({n_reb} to rebrand, {n_leave} to leave as-is)")
            self.prog_main(100, "Done")
            self.emit(AppEvent(EventType.DONE))
            self.emit(AppEvent(EventType.NOTIFICATION, {
                "title": "Review Sheet Ready",
                "msg": f"{n_reb} to rebrand, {n_leave} to leave as-is.\n"
                       f"Saved to {plan.parent}\nEdit it in Excel, then run Rebrand → Apply.",
                "open_path": str(plan)}))
        except Exception as e:
            self.log(f"Analyze error: {e}", True)
            self.emit(AppEvent(EventType.DONE))

    def _auto_workers(self):
        """Thread count: manual override, else scaled by RAM (like the refine engine)."""
        try:
            forced = int(CFG.get("max_threads"))
        except Exception:
            forced = 0
        if forced > 0:
            return max(1, forced)
        mw = 2
        if HAS_PSUTIL:
            try:
                gb = psutil.virtual_memory().total / (1024 ** 3)
                mw = 1 if gb < 8 else (2 if gb < 16 else 4)
            except Exception:
                pass
        return max(1, min(mw, os.cpu_count() or 1))

    def _parallel_map(self, items, fn, label):
        """Run fn over items on the worker pool, updating progress. Returns results in order."""
        n = len(items)
        if n == 0:
            return []
        max_workers = self._auto_workers()
        self.emit(AppEvent(EventType.WORKER_CONFIG, max_workers))
        results = [None] * n
        with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as ex:
            futures = {ex.submit(fn, it): i for i, it in enumerate(items)}
            completed = 0
            for fut in concurrent.futures.as_completed(futures):
                if self.stop_sig:
                    break
                completed += 1
                self.prog_main((completed / n) * 100, f"{label} {completed}/{n}")
                try:
                    results[futures[fut]] = fut.result()
                except Exception as e:
                    self.log(f"{label} error: {e}", True)
        return results

    def _assign_output_names(self, rows, out, kit=None):
        """Assign each row a unique output path up front (single-threaded) so parallel
        writes never collide, and re-runs stay deterministic (resume-safe)."""
        from .rebrand import delivery_filename, numbered_filename as _numbered
        brand = {"brand_slug": kit.brand_slug} if kit is not None else {}

        def preferred(row, rel_path):
            return delivery_filename(rel_path.stem, row.get("product"),
                                     row.get("asset_type"), **brand)

        # First work out which product/asset-type names more than one document
        # wants. Many documents share both fields, and the model leaves product
        # blank on a large share of them, so without this a whole group collapses
        # onto one name and gets meaningless -2/-3 suffixes. Whether a given file
        # kept the clean name would come down to its position in the sheet, so
        # when a name is contested EVERY member falls back to its source name.
        wanted = {}
        for row in rows:
            rel = (row.get("file") or "").strip().replace("\\", "/")
            if rel and (row.get("action") or "").strip().lower() != "leave":
                wanted[preferred(row, Path(rel))] = wanted.get(preferred(row, Path(rel)), 0) + 1

        taken = set()
        for row in rows:
            rel = (row.get("file") or "").strip().replace("\\", "/")
            if not rel:
                row["_dst"] = None
                continue
            rel_path = Path(rel)
            if (row.get("action") or "").strip().lower() == "leave":
                dst = out / rel_path  # unbranded, original name/structure
            else:
                name = preferred(row, rel_path)
                if wanted.get(name, 0) > 1:
                    # The source name identifies the document; "-2" tells the
                    # reader nothing.
                    name = delivery_filename(rel_path.stem, "", row.get("asset_type"), **brand)
                dst = out / rel_path.parent / name
            base, n = dst, 1
            while str(dst).lower() in taken:   # last resort, still deterministic
                n += 1
                dst = base.with_name(_numbered(base.name, n))
            taken.add(str(dst).lower())
            row["_dst"] = str(dst)

    def _apply_one_row(self, row, src, out, kit):
        """Process one review-sheet row (output path pre-assigned in row['_dst'])."""
        from .rebrand import rebrand_pdf, title_for
        rel = (row.get("file") or "").strip().replace("\\", "/")
        dst_str = row.get("_dst")
        if not rel or not dst_str:
            return "skip"
        self.emit(AppEvent(EventType.SLOT_UPDATE, {"tid": threading.get_ident(), "text": Path(rel).name, "percent": None}))
        pdf = src / rel
        if not pdf.exists():
            self.log(f"Missing source: {rel}", True)
            return "fail"
        dst = Path(dst_str)
        dst.parent.mkdir(parents=True, exist_ok=True)
        if dst.exists() and dst.stat().st_size > 0:   # resume: already produced
            return "skip"
        if (row.get("action") or "").strip().lower() == "leave":
            _atomic_copy(pdf, dst)
            return "leave"
        title = (row.get("title") or "").strip() or title_for(
            row.get("asset_type"), fallback_stem=pdf.stem, doc_type=row.get("doc_type"))
        subtitle = kit.subtitle_for(row.get("manufacturer"))
        stats = rebrand_pdf(pdf, dst, kit, title, subtitle=subtitle)
        if stats["size_mb"] >= 50:
            self.log(f"⚠️ {dst.name} is {stats['size_mb']} MB (over 50 MB)", True)
            return "oversize"
        return "done"

    @staticmethod
    def _kit_problem(kit):
        """Explain exactly why a brand kit can't be used, naming the missing art."""
        if not (kit.root.exists() and kit.root.is_dir()):
            return f"Brand kit folder not found: {kit.root}"
        gaps = []
        for o in ("portrait", "landscape"):
            if not kit.has_folder(o):
                gaps.append(f"no '{o.title()}' folder")
            elif kit.missing(o):
                gaps.append(f"'{o.title()}' has no {' or '.join(kit.missing(o))} image")
        detail = "; ".join(gaps) or "no usable Portrait/Landscape assets"
        return (f"Brand kit is unusable — {detail}. Each orientation folder needs at "
                f"least a cover and a back-cover PNG.")

    def _copy_extras(self, src, out, rows):
        """Complete set: copy every non-PDF source file into the mirrored output.

        PDFs are the review sheet's business; everything else (images, Office
        docs, spreadsheets) is copied through unchanged so the output tree is the
        whole upload set. Review sheets themselves are never copied."""
        from .reviews import is_plan_file
        planned = {(r.get("file") or "").strip().replace("\\", "/").lower() for r in rows}
        try:
            out_res = out.resolve()
        except OSError:
            out_res = out

        extras, unplanned_pdfs = [], 0
        for root, dirs, files in os.walk(src):
            rp = Path(root)
            try:
                rr = rp.resolve()
                if rr == out_res or out_res in rr.parents:
                    dirs[:] = []          # never re-ingest our own output
                    continue
            except OSError:
                pass
            for name in files:
                p = rp / name
                if p.suffix.lower() == ".pdf":
                    rel = str(p.relative_to(src)).replace("\\", "/").lower()
                    if rel not in planned:
                        unplanned_pdfs += 1
                    continue
                if is_plan_file(name):
                    continue
                extras.append(p)

        if unplanned_pdfs:
            self.log(f"Note: {unplanned_pdfs} PDF(s) in the source are not in the review sheet "
                     f"and were not produced — re-run Analyze to include them.", True)
        if not extras:
            self.log("Complete set: no non-PDF files to copy.")
            return 0
        self.log(f"Complete set: copying {len(extras)} non-PDF files")

        def cp(p):
            if self.stop_sig:
                return None
            if not self.pause_event.is_set():
                self.pause_event.wait()
            self.emit(AppEvent(EventType.SLOT_UPDATE, {"tid": threading.get_ident(), "text": p.name, "percent": None}))
            dst = out / p.relative_to(src)
            try:
                dst.parent.mkdir(parents=True, exist_ok=True)
                if dst.exists() and dst.stat().st_size == p.stat().st_size:
                    return "skip"     # already there (resume-safe)
                _atomic_copy(p, dst)
                return "copied"
            except Exception as e:
                self.log(f"Copy failed: {p.name}: {e}", True)
                return None

        res = self._parallel_map(extras, cp, "Copying")
        return sum(1 for r in res if r in ("copied", "skip"))

    def run_rebrand_apply(self, src_dir, kit_dir, plan_csv=None, out_dir=None, complete_set=False):
        """Apply an approved review sheet: rebrand the 'rebrand' rows, copy the rest as-is.

        With complete_set, every non-PDF file in the source is copied through to
        the mirrored output too, so the output tree is the whole upload set."""
        try:
            self.stop_sig = False
            self.resume()
            from .rebrand import BrandKit
            from .reviews import find_plan
            src = Path(src_dir)
            out = Path(out_dir) if out_dir else src.parent / f"{src.name}_rebranded"
            self.current_ws = str(out)

            kit = BrandKit(kit_dir)
            if not (kit.has("portrait") or kit.has("landscape")):
                why = self._kit_problem(kit)
                self.log(f"CRITICAL: {why}", True)
                self.emit(AppEvent(EventType.ERROR, why))
                self.emit(AppEvent(EventType.DONE)); return

            if not plan_csv:   # find the sheet Analyze wrote for this folder
                found = find_plan(src)
                if not found:
                    self.log("No review sheet found for this folder — run Analyze first.", True)
                    self.emit(AppEvent(EventType.ERROR, "No review sheet found — run Analyze first."))
                    self.emit(AppEvent(EventType.DONE)); return
                plan_csv = str(found)
                self.log(f"Using review sheet: {plan_csv}")

            from .reviews import read_plan
            rows = read_plan(plan_csv)
            if not rows:
                self.log("Review sheet is empty.", True)
                self.emit(AppEvent(EventType.DONE)); return

            self.log(f"Applying review sheet: {len(rows)} entries -> {out}")
            start_time = time.time()

            # Assign unique output names single-threaded (collision- and resume-safe).
            self._assign_output_names(rows, out, kit)

            # Warm shared state once so the worker threads don't race on it.
            from .rebrand import _ensure_font
            _ensure_font()
            for o in ("portrait", "landscape"):
                if kit.has(o):
                    kit.specs(o)

            def apply_row(row):
                if self.stop_sig:
                    return "skip"
                if not self.pause_event.is_set():
                    self.pause_event.wait()
                try:
                    return self._apply_one_row(row, src, out, kit)
                except Exception as e:
                    self.log(f"Failed: {(row.get('file') or '?')}: {e}", True)
                    return "fail"

            statuses = self._parallel_map(rows, apply_row, "Applying")
            done = sum(1 for s in statuses if s in ("done", "oversize"))
            oversize = sum(1 for s in statuses if s == "oversize")
            left = sum(1 for s in statuses if s == "leave")
            skipped = sum(1 for s in statuses if s == "skip")
            failed = sum(1 for s in statuses if s == "fail")

            if self.stop_sig:
                self.log("Apply stopped by user.")
                self.emit(AppEvent(EventType.DONE)); return

            copied = self._copy_extras(src, out, rows) if complete_set else 0

            # Timing goes to the log, not a stats.json inside the tree — the output
            # folder is the delivery set and must contain nothing but the documents.
            self.log(f"Rebrand took {(time.time() - start_time) / 60:.1f} min")
            msg = f"Done. Rebranded {done}, left as-is {left}, skipped {skipped}, failed {failed}."
            if complete_set: msg += f" Copied {copied} non-PDF files."
            if oversize: msg += f" ({oversize} over 50 MB — review.)"
            self.log(msg)
            self.prog_main(100, "Done")
            self.emit(AppEvent(EventType.DONE))
            self.emit(AppEvent(EventType.NOTIFICATION, {"title": "Rebrand Complete", "msg": msg, "open_path": str(out)}))
        except Exception as e:
            self.log(f"Apply error: {e}", True)
            self.emit(AppEvent(EventType.DONE))

    # ==========================================================================
    #   COMPOSABLE PIPELINE  (start from any folder; pick steps; OCR runs last)
    # ==========================================================================
    def _load_plan(self, src):
        """Load this folder's review sheet (wherever it lives) into {relative_path: row}."""
        from .reviews import find_plan, read_plan
        p = find_plan(src)
        out = {}
        if p:
            try:
                for row in read_plan(p):
                    key = (row.get("file") or "").strip().replace("\\", "/")
                    if key:
                        out[key] = row
            except Exception:
                pass
        return out

    def _folder_pdf_op(self, src, out, mode, dpi, label, only_no_text=False, complete_set=True):
        """Flatten or OCR every PDF from src into out (mirrored).

        Non-PDFs pass through unchanged with complete_set; without it they are
        left behind and the output holds PDFs only."""
        from .processing import PdfProcessor
        bot = PdfProcessor(lambda v, t, s=False: self.prog_sub(v, t, s), lambda: self.stop_sig, self.pause_event)
        files = [Path(r) / f for r, _, fs in os.walk(src) for f in fs
                 if complete_set or f.lower().endswith(".pdf")]

        def op(p):
            if self.stop_sig:
                return
            if not self.pause_event.is_set():
                self.pause_event.wait()
            rel = p.relative_to(src); dst = out / rel
            dst.parent.mkdir(parents=True, exist_ok=True)
            self.emit(AppEvent(EventType.SLOT_UPDATE, {"tid": threading.get_ident(), "text": p.name, "percent": None}))
            if p.suffix.lower() != ".pdf":
                shutil.copy2(p, dst); return
            if only_no_text:
                from .classify import extract_text
                if len(extract_text(p) or "") >= 20:   # already searchable → leave it
                    shutil.copy2(p, dst); return
            try:
                ok = bot.flatten_or_ocr(p, dst, mode=mode, dpi=dpi)
                if not ok and not dst.exists():
                    shutil.copy2(p, dst)
            except Exception as e:
                self.log(f"{mode} failed: {rel}: {e}", True)
                if not dst.exists():
                    shutil.copy2(p, dst)

        self._parallel_map(files, op, label)

    def _folder_rebrand(self, src, plan_src, out, kit, label, complete_set=True):
        """Rebrand every PDF from src into out, honouring a review plan from plan_src if present.

        With complete_set, non-PDF files are copied through so the output is the
        full set; without it, only PDFs reach the output."""
        from .rebrand import (rebrand_pdf, delivery_filename, title_for, _ensure_font)
        from .reviews import is_plan_file
        brand = {"brand_slug": kit.brand_slug}
        _ensure_font()
        for o in ("portrait", "landscape"):
            if kit.has(o):
                kit.specs(o)
        plan = self._load_plan(plan_src)
        files = [Path(r) / f for r, _, fs in os.walk(src) for f in fs
                 if (complete_set or f.lower().endswith(".pdf")) and not is_plan_file(f)]

        # Pre-assign a unique output path per file (single-threaded → collision-safe).
        dsts = {}
        used = {}
        for p in files:
            rel = p.relative_to(src)
            row = plan.get(str(rel).replace("\\", "/")) or plan.get(p.name)
            is_leave = (row and (row.get("action") or "").strip().lower() == "leave")
            if p.suffix.lower() != ".pdf" or is_leave:
                dst = out / rel
            elif row:
                dst = out / rel.parent / delivery_filename(
                    p.stem, row.get("product"), row.get("asset_type"), **brand)
            else:
                dst = out / rel.parent / delivery_filename(p.stem, **brand)
            key = str(dst).lower()
            n = used.get(key, 0) + 1
            used[key] = n
            if n > 1:
                dst = dst.with_name(f"{dst.stem}-{n}{dst.suffix}")
            dsts[p] = dst

        def one(p):
            if self.stop_sig:
                return
            if not self.pause_event.is_set():
                self.pause_event.wait()
            rel = p.relative_to(src)
            self.emit(AppEvent(EventType.SLOT_UPDATE, {"tid": threading.get_ident(), "text": p.name, "percent": None}))
            row = plan.get(str(rel).replace("\\", "/")) or plan.get(p.name)
            is_leave = (row and (row.get("action") or "").strip().lower() == "leave")
            dst = dsts[p]; dst.parent.mkdir(parents=True, exist_ok=True)
            try:
                if p.suffix.lower() != ".pdf" or is_leave:
                    _atomic_copy(p, dst); return
                title = ((row.get("title") if row else "") or "").strip() or title_for(
                    row.get("asset_type") if row else None, fallback_stem=p.stem,
                    doc_type=row.get("doc_type") if row else None)
                subtitle = kit.subtitle_for(row.get("manufacturer") if row else "")
                rebrand_pdf(p, dst, kit, title, subtitle=subtitle)
            except Exception as e:
                self.log(f"Rebrand failed: {rel}: {e}", True)
                try: _atomic_copy(p, dst)
                except Exception: pass

        self._parallel_map(files, one, label)

    def run_pipeline(self, src_dir, do_flatten, do_rebrand, do_ocr, kit_dir=None, dpi=300,
                     out_dir=None, complete_set=True):
        """Run selected steps over a folder, in the fixed safe order Flatten → Rebrand → OCR.

        complete_set carries non-PDF files through every stage into the output."""
        import tempfile
        try:
            self.stop_sig = False
            self.resume()
            src = Path(src_dir)
            out = Path(out_dir) if out_dir else src.parent / f"{src.name}_processed"
            self.current_ws = str(out)

            active = [n for n, on in (("Flatten", do_flatten), ("Rebrand", do_rebrand), ("Make Searchable", do_ocr)) if on]
            if not active:
                self.log("No pipeline steps selected.", True); self.emit(AppEvent(EventType.DONE)); return

            kit = None
            if do_rebrand:
                from .rebrand import BrandKit
                kit = BrandKit(kit_dir) if kit_dir else None
                if not (kit and (kit.has("portrait") or kit.has("landscape"))):
                    why = self._kit_problem(kit) if kit else "The Rebrand step needs a brand kit."
                    self.log(f"CRITICAL: {why}", True)
                    self.emit(AppEvent(EventType.ERROR, why))
                    self.emit(AppEvent(EventType.DONE)); return

            self.log(f"Pipeline: {' → '.join(active)}   ({src.name} → {out.name})")
            self.emit(AppEvent(EventType.WORKER_CONFIG, 1))
            start_time = time.time()
            work = Path(tempfile.mkdtemp(prefix="drp_pipe_"))
            current = src
            step = 0; n = len(active)

            if do_flatten and not self.stop_sig:
                step += 1; nxt = work / "1_flattened"
                self._folder_pdf_op(current, nxt, "flatten", dpi, f"[{step}/{n}] Flatten",
                                    complete_set=complete_set); current = nxt
            if do_rebrand and not self.stop_sig:
                step += 1; nxt = work / "2_rebranded"
                self._folder_rebrand(current, src, nxt, kit, f"[{step}/{n}] Rebrand",
                                     complete_set=complete_set); current = nxt
            if do_ocr and not self.stop_sig:
                step += 1; nxt = work / "3_searchable"
                self._folder_pdf_op(current, nxt, "ocr", dpi, f"[{step}/{n}] OCR", only_no_text=True,
                                    complete_set=complete_set); current = nxt

            if self.stop_sig:
                self.log("Pipeline stopped by user."); shutil.rmtree(work, ignore_errors=True)
                self.emit(AppEvent(EventType.DONE)); return

            if out.exists():
                shutil.rmtree(out, ignore_errors=True)
            out.parent.mkdir(parents=True, exist_ok=True)
            shutil.move(str(current), str(out))
            shutil.rmtree(work, ignore_errors=True)

            self.log(f"Pipeline complete: {' → '.join(active)}"
                     f"  ({(time.time() - start_time) / 60:.1f} min)")
            self.prog_main(100, "Done")
            self.emit(AppEvent(EventType.DONE))
            self.emit(AppEvent(EventType.NOTIFICATION, {"title": "Pipeline Complete",
                       "msg": f"{' → '.join(active)} finished.", "open_path": str(out)}))
        except Exception as e:
            self.log(f"Pipeline error: {e}", True)
            self.emit(AppEvent(EventType.DONE))

    def run_pull_model(self, model="llama3.2:3b"):
        """Download an Ollama model with progress (used by the analyze setup flow)."""
        try:
            self.stop_sig = False
            self.resume()
            from .classify import start_server, pull_model, has_model
            if not start_server():
                self.log("Could not start Ollama — is it installed?", True)
                self.emit(AppEvent(EventType.ERROR, "Could not start Ollama."))
                self.emit(AppEvent(EventType.DONE)); return
            if has_model(model):
                self.log(f"Model '{model}' is already downloaded.")
                self.emit(AppEvent(EventType.DONE))
                self.emit(AppEvent(EventType.NOTIFICATION, {"title": "Model Ready",
                           "msg": f"'{model}' is ready. Click Analyze to continue."}))
                return

            self.log(f"Downloading model '{model}' (~2 GB, one-time)…")
            self.emit(AppEvent(EventType.WORKER_CONFIG, 1))
            self._last_pull_emit = 0.0

            def on_prog(msg):
                status_txt = msg.get("status", "")
                total = msg.get("total") or 0
                completed = msg.get("completed") or 0
                now = time.time()
                if status_txt != "success" and now - getattr(self, "_last_pull_emit", 0.0) < 0.25:
                    return
                self._last_pull_emit = now
                if total:
                    self.prog_main(completed / total * 100,
                                   f"{status_txt}: {completed // (1024 * 1024)}/{total // (1024 * 1024)} MB")
                else:
                    self.prog_main(0, status_txt or "Downloading…")

            ok = pull_model(model, on_progress=on_prog)
            if ok or has_model(model):
                self.log(f"Model '{model}' downloaded.")
                self.prog_main(100, "Done")
                self.emit(AppEvent(EventType.DONE))
                self.emit(AppEvent(EventType.NOTIFICATION, {"title": "Model Ready",
                           "msg": f"'{model}' downloaded. Click Rebrand → Analyze to continue."}))
            else:
                self.log(f"Model download failed for '{model}'.", True)
                self.emit(AppEvent(EventType.ERROR, "Model download failed."))
                self.emit(AppEvent(EventType.DONE))
        except Exception as e:
            self.log(f"Model download error: {e}", True)
            self.emit(AppEvent(EventType.DONE))

    def run_debug_export(self, ws_path_str):
        try:
            ts = datetime.now().strftime('%Y%m%d_%H%M%S')
            base_dir = SystemUtils.get_user_data_dir()
            
            try:
                test_file = base_dir / "write_test.tmp"
                test_file.touch()
                test_file.unlink()
            except PermissionError:
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
            if ws_path_str:
                ws = Path(ws_path_str)
                if ws.exists():
                    safe_copy(ws/"session_log.txt", "current_job_log.txt")
                    safe_copy(ws/"stats.json", "current_job_stats.json")
            
            shutil.make_archive(str(dest_zip).replace(".zip", ""), 'zip', temp_dir)
            shutil.rmtree(temp_dir, ignore_errors=True)
            
            self.emit(AppEvent(EventType.NOTIFICATION, {"title": "Debug Export", "msg": f"Saved to {dest_zip.name}", "open_path": str(base_dir)}))
            
        except Exception as e:
            self.emit(AppEvent(EventType.ERROR, f"Export Failed: {e}"))