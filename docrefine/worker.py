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
import concurrent.futures
import platform
from pathlib import Path
from datetime import datetime, timedelta

# Local Package Imports
from .config import CFG, SystemUtils, log_app, WORKSPACES_ROOT, LOG_PATH, JSON_LOG_PATH
from .core.events import AppEvent, EventType
from .processing import (
    PdfProcessor, 
    ImageProcessor, 
    OfficeProcessor, 
    POPPLER_BIN, 
    HAS_TESSERACT,
    pdfinfo_from_path,
    convert_from_path
)

# 3rd Party Dependencies
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
#   HELPER FUNCTIONS (Preserved from v118)
# ==============================================================================

def sanitize_filename(name):
    return re.sub(r'[<>:"/\\|?*]', '_', name)

def update_stats_time(ws, cat, sec):
    try:
        p = Path(ws) / "stats.json"
        if not p.exists(): return
        with open(p, 'r') as f: s = json.load(f)
        s[cat] = s.get(cat, 0.0) + sec
        with open(p, 'w') as f: json.dump(s, f, indent=4)
    except: pass

def generate_job_report(ws_path, action_name, file_results=None):
    # Logic identical to v118, preserved for report generation
    try:
        ws = Path(ws_path)
        rpt_dir = ws / "04_Reports"
        rpt_dir.mkdir(parents=True, exist_ok=True)
        
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M')
        file_name = f"Audit_Certificate_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html"
        
        s = {}
        if (ws / "stats.json").exists():
            with open(ws / "stats.json") as f: s = json.load(f)
        
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
        saved_mb = round(saved_bytes / (1024 * 1024), 2)
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

        # (HTML generation logic preserved)
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

# ==============================================================================
#   WORKER CLASS (REFACTORED FOR EVENTS)
# ==============================================================================
class Worker:
    def __init__(self, callback): 
        # REFACTORED: Accepts a generic callback instead of a queue
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
        # REFACTORED: Uses Event System
        level = "ERROR" if err else "INFO"
        self.emit(AppEvent.log(m, level))
        log_app(m, level, structured_data={"ws": self.current_ws})

    def set_job_status(self, ws, stage, details=""):
        try:
            data = { "stage": stage, "last_update": datetime.now().strftime('%Y-%m-%d %H:%M:%S'), "details": details }
            with open(Path(ws) / "status.json", 'w') as f: json.dump(data, f, indent=4)
        except: pass

    def prog_main(self, v, t): 
        # REFACTORED
        self.emit(AppEvent.progress(v, t))
    
    def prog_sub(self, v, t, status_only=False): 
        # REFACTORED
        tid = threading.get_ident()
        now = time.time()
        
        if tid not in self._last_update:
            self._last_update[tid] = 0
            
        if (now - self._last_update[tid]) > 0.1: # Max 10 updates/sec per thread
            # Emitting structured slot data
            self.emit(AppEvent(EventType.SLOT_UPDATE, {"tid": tid, "text": t, "percent": v}))
            self._last_update[tid] = now

    def get_hash(self, path, mode):
        # (Logic preserved from v118)
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
        # (Logic preserved from v118)
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

    def run_inventory(self, d_str, ingest_mode):
        try:
            self.stop_sig = False
            self.resume()
            
            d = Path(d_str); start_time = time.time()
            ws = WORKSPACES_ROOT / f"{d.name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            m_dir = ws / "01_Master_Files"; m_dir.mkdir(parents=True); (ws/"00_Quarantine").mkdir()
            self.current_ws = str(ws); self.log(f"Inventory Start: {d}")
            
            # REFACTORED
            self.emit(AppEvent(EventType.JOB_DATA, str(ws)))
            self.set_job_status(ws, "SCANNING", "Ingesting...")
            
            files = [Path(r)/f for r,_,fs in os.walk(d) for f in fs]
            files = [f for f in files if f.suffix.lower() in SUPPORTED_EXTENSIONS]
            
            seen = {}; quarantined = 0
            
            # REFACTORED
            self.emit(AppEvent(EventType.WORKER_CONFIG, 1))

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

            if self.stop_sig: 
                self.log("Ingest Stopped by User.")
                self.emit(AppEvent(EventType.DONE))
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
            self.log(f"Done. Masters: {total}")
            
            # REFACTORED
            self.emit(AppEvent(EventType.JOB_DATA, str(ws)))
            self.emit(AppEvent(EventType.DONE))
            
        except Exception as e: 
            self.log(f"Error: {e}", True)
            self.emit(AppEvent(EventType.DONE))

    def process_file_task(self, f, bots, options, base_dst):
        # (Logic preserved from v118, updated logs/events only)
        if self.stop_sig: return None
        result = {'file': f.name, 'orig_size': f.stat().st_size, 'new_size': 0, 'ok': False}
        try:
            # REFACTORED
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
            self.stop_sig = False
            self.resume()
            
            ws = Path(ws_p); self.current_ws = str(ws)
            start_time = time.time(); src = ws/"01_Master_Files"; dst = ws/"02_Ready_For_Redistribution"; dst.mkdir(exist_ok=True)
            self.log(f"Refinement Start. Opts: {options}")
            self.set_job_status(ws, "PROCESSING", "Refining...")

            # Use imported Processors
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

            # REFACTORED
            self.emit(AppEvent(EventType.WORKER_CONFIG, max_workers))
            
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
            
            # REFACTORED: Send notification instead of direct open_file (though standard behavior can be handled in adapter)
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
            out = ws / "03_Organized_Output"; m = out/"Unique_Masters"; q = out/"Quarantine"
            for p in [m,q]: p.mkdir(parents=True, exist_ok=True)
            
            self.log(f"Unique Export ({priority_mode})")
            with open(ws/"manifest.json") as f: man = json.load(f)
            total = len(man)
            
            # REFACTORED
            self.emit(AppEvent(EventType.WORKER_CONFIG, 1))

            dup_csv = out / "duplicates_report.csv"
            with open(dup_csv, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(["Master_Filename", "Duplicate_Location"])
                
                for i, (h, data) in enumerate(man.items()):
                    if self.stop_sig: break
                    self.prog_main((i/total)*100, "Exporting Unique...")
                    # REFACTORED
                    self.emit(AppEvent(EventType.SLOT_UPDATE, {"tid": threading.get_ident(), "text": f"Exporting: {data['name']}", "percent": None}))
                    
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
                self.prog_main((i/len(man))*100, f"Recon {i+1}")
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
            
            if self.stop_sig: return

            q_src = ws / "00_Quarantine"
            if q_src.exists():
                q_dst = dst / "_QUARANTINED_FILES"; 
                q_dst.mkdir(parents=True, exist_ok=True) 
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
                self.emit(AppEvent(EventType.ERROR, "Could not write CSV.\nPlease close the file in Excel and try again."))
                self.emit(AppEvent(EventType.DONE))
                return

            if self.stop_sig: return

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
            src = ws/"01_Master_Files"; pdf = next(src.glob("*.pdf"), None)
            
            if not pdf: 
                # REFACTORED: Status + Done
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
            
            # REFACTORED
            self.emit(AppEvent(EventType.NOTIFICATION, {"title": "Debug Export", "msg": f"Saved to {dest_zip.name}", "open_path": str(base_dir)}))
            
        except Exception as e:
            self.emit(AppEvent(EventType.ERROR, f"Export Failed: {e}"))