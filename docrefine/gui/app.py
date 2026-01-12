import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox, ttk, Menu
import threading
import queue
import time
import json
import os
import shutil
import ssl
import urllib.request
import webbrowser
import platform
import uuid
import re
from pathlib import Path
from datetime import datetime, timedelta
from PIL import Image, ImageTk

# 3rd Party (Optional UI dependencies)
try:
    import pytesseract
except ImportError:
    pass

# Local Package Imports
from ..config import CFG, SystemUtils, LOG_PATH, WORKSPACES_ROOT
from ..worker import Worker, HAS_TESSERACT, POPPLER_BIN
from ..processing import convert_from_path, pdfinfo_from_path

# ==============================================================================
#   UI HELPERS
# ==============================================================================
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

def sanitize_filename(name):
    return re.sub(r'[<>:"/\\|?*]', '_', name)

# ==============================================================================
#   SUB-WINDOWS
# ==============================================================================
class DocViewer:
    def __init__(self, root, filename, title):
        self.win = tk.Toplevel(root)
        self.win.title(title)
        self.win.geometry("800x600")
        App.center_toplevel(self.win, root)
        
        target = SystemUtils.find_doc_file(filename)
        
        if not target:
            tk.Label(self.win, text=f"File '{filename}' not found.", font=("Segoe UI", 12, "bold"), fg="red").pack(pady=20)
            tk.Label(self.win, text="Please ensure the file is present in the application folder.", font=("Segoe UI", 10)).pack()
            return
            
        txt = scrolledtext.ScrolledText(self.win, wrap="word", font=("Consolas", 10), padx=10, pady=10)
        txt.pack(fill="both", expand=True)
        
        try:
            content = target.read_text(encoding='utf-8')
            txt.insert("1.0", content)
        except Exception as e:
            txt.insert("1.0", f"Error reading file: {e}")
            
        txt.config(state="disabled")

class ForensicComparator:
    def __init__(self, root, ws_path, manifest, master_path, dup_candidates):
        self.win = tk.Toplevel(root)
        self.win.title("Forensic Verification (Sync View)")
        self.win.geometry("1400x800")
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
        
        self.top = tk.Frame(self.win, bg="#eee", pady=5)
        self.top.pack(fill="x")
        
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
        
        self.is_pdf = self.master_path.suffix.lower() == '.pdf'
        if self.is_pdf:
            try: self.total_pages = pdfinfo_from_path(str(self.master_path), poppler_path=POPPLER_BIN).get('Pages', 1)
            except: self.total_pages = 1
            
        self.load_images()
        self.bind_events()

    def _build_toolbar(self):
        f_page = tk.Frame(self.top); f_page.pack(side="left", padx=10)
        tk.Button(f_page, text="< Prev Page", command=self.prev_page).pack(side="left")
        self.lbl_page = tk.Label(f_page, text="Page 1/1", width=10)
        self.lbl_page.pack(side="left", padx=5)
        tk.Button(f_page, text="Next Page >", command=self.next_page).pack(side="left")
        
        f_zoom = tk.Frame(self.top); f_zoom.pack(side="left", padx=20)
        tk.Button(f_zoom, text="- Zoom", command=lambda: self.do_zoom(0.8)).pack(side="left")
        self.lbl_zoom = tk.Label(f_zoom, text="100%", width=6)
        self.lbl_zoom.pack(side="left")
        tk.Button(f_zoom, text="Zoom +", command=lambda: self.do_zoom(1.2)).pack(side="left")
        tk.Button(f_zoom, text="[Fit Width]", command=self.fit_width).pack(side="left", padx=5)
        
        f_dup = tk.Frame(self.top); f_dup.pack(side="right", padx=10)
        tk.Button(f_dup, text="< Prev Copy", command=self.prev_dup).pack(side="left")
        self.lbl_dup = tk.Label(f_dup, text="Copy 1/1", width=15)
        self.lbl_dup.pack(side="left", padx=5)
        tk.Button(f_dup, text="Next Copy >", command=self.next_dup).pack(side="left")
        tk.Button(f_dup, text="Open File", command=self.open_current_dup).pack(side="left", padx=5)
        
        kw_uniq = {"bg": "green", "fg": "white"} if not SystemUtils.IS_MAC else {}
        tk.Button(f_dup, text="MARK AS UNIQUE", command=self.mark_as_unique, **kw_uniq).pack(side="left", padx=10)

    def load_images(self):
        self.lbl_page.config(text=f"Page {self.page}/{self.total_pages}")
        self.lbl_zoom.config(text=f"{int(self.zoom*100)}%")
        self.lbl_dup.config(text=f"Copy {self.dup_idx+1}/{len(self.dups)}")
        
        if self.dups:
            path_str = str(self.dups[self.dup_idx])
            if len(path_str) > 60: path_str = "..." + path_str[-57:]
            tk.Label(self.lbl_info.winfo_children()[1], text=path_str).pack_forget() 
            self.lbl_info.winfo_children()[1].config(text=f"CANDIDATE: {path_str}")

        self.img1 = self._render(self.master_path)
        self.show_img(self.c1, self.img1)
        
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
        self.c1.bind("<MouseWheel>", self.on_scroll_page)
        self.c2.bind("<MouseWheel>", self.on_scroll_page)
        self.c1.bind("<Button-4>", self.on_scroll_page)
        self.c1.bind("<Button-5>", self.on_scroll_page)
        self.c1.bind("<ButtonPress-1>", self.scroll_start)
        self.c1.bind("<B1-Motion>", self.scroll_move)
        self.c2.bind("<ButtonPress-1>", self.scroll_start)
        self.c2.bind("<B1-Motion>", self.scroll_move)

    def on_scroll_page(self, event):
        now = time.time()
        if now - self.last_scroll_time < 0.4: return
        self.last_scroll_time = now

        d = 0
        if event.num == 5 or event.delta < 0: d = 1 
        elif event.num == 4 or event.delta > 0: d = -1 
        
        if d == 1: self.next_page()
        elif d == -1: self.prev_page()

    def scroll_start(self, event):
        self.c1.scan_mark(event.x, event.y)
        self.c2.scan_mark(event.x, event.y)

    def scroll_move(self, event):
        self.c1.scan_dragto(event.x, event.y, gain=1)
        self.c2.scan_dragto(event.x, event.y, gain=1)

    def do_zoom(self, factor):
        self.zoom *= factor
        self.load_images()

    def fit_width(self):
        try:
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
            SystemUtils.reveal_file(self.dups[self.dup_idx])

    def mark_as_unique(self):
        if not self.dups: return
        target_path = self.dups[self.dup_idx]
        
        if messagebox.askyesno("Promote File", f"Mark '{target_path.name}' as a unique Master file?\n\nIt will be removed from this duplicate list and treated as a distinct document."):
            try:
                root_path = None
                original_hash = None
                
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
                    rel_p = str(target_path.relative_to(root_path))
                    if rel_p in self.manifest[original_hash]['copies']:
                        self.manifest[original_hash]['copies'].remove(rel_p)

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
                    
                    m_dir = self.ws_path / "01_Master_Files"
                    shutil.copy2(target_path, m_dir / new_uid)

                    with open(self.ws_path / "manifest.json", 'w') as f:
                        json.dump(self.manifest, f, indent=4)

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

# ==============================================================================
#   MAIN APPLICATION
# ==============================================================================
class App:
    def __init__(self, root):
        self.root = root
        self.root.title(f"DocRefine Pro {SystemUtils.CURRENT_VERSION} ({platform.system()})")
        
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
        
        self.poll()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        
        threading.Thread(target=self.check_updates, args=(False,), daemon=True).start()

    def apply_smart_geometry(self, saved_geo):
        try:
            sw = self.root.winfo_screenwidth()
            sh = self.root.winfo_screenheight()
            
            if SystemUtils.IS_MAC: sh -= 120 
            else: sh -= 60 
            
            if SystemUtils.IS_MAC: w, h = 1280, 850
            else: w, h = 1024, 700
            
            x, y = 0, 0
            
            if saved_geo:
                parts = re.split(r'[x+]', saved_geo)
                if len(parts) == 4:
                    w, h, x, y = map(int, parts)
            
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

    @staticmethod
    def center_toplevel(win, parent):
        try:
            win.withdraw() 
            win.update_idletasks() 
            
            pw = parent.winfo_width()
            ph = parent.winfo_height()
            px = parent.winfo_rootx()
            py = parent.winfo_rooty()
            
            cw = win.winfo_width()
            ch = win.winfo_height()
            
            x = px + (pw // 2) - (cw // 2)
            y = py + (ph // 2) - (ch // 2)
            
            win.geometry(f"+{x}+{y}")
            win.deiconify() 
        except: 
            win.deiconify()

    def _build_refine(self):
        tk.Label(self.tab_process, text="Content Refinement (Modifies Files)", font=("Segoe UI",10,"bold")).pack(anchor="w",pady=(10,5),padx=10)
        
        self.chk_frame = tk.Frame(self.tab_process); self.chk_frame.pack(fill="x",padx=10)
        self.chk_vars = {} 
        
        self.f_pdf_ctrl = tk.Frame(self.tab_process)
        tk.Label(self.f_pdf_ctrl, text="PDF Action:", font=("Segoe UI", 9)).pack(anchor="w", padx=10, pady=(10,0))
        self.pdf_mode_var = tk.StringVar(value="No Action")
        self.cb_pdf = ttk.Combobox(self.f_pdf_ctrl, textvariable=self.pdf_mode_var, values=["No Action", "Flatten Only (Fast)", "Flatten + OCR (Slow)"], state="readonly")
        self.cb_pdf.pack(fill="x", padx=10, pady=2)
        
        ctrl = tk.Frame(self.tab_process); ctrl.pack(fill="x",pady=15,padx=10)
        
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
            if path.exists(): SystemUtils.reveal_file(path)

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
        win.geometry("600x700")
        
        App.center_toplevel(win, self.root)
        
        lf_perf = tk.LabelFrame(win, text="Processing Engine", padx=10, pady=10)
        lf_perf.pack(fill="x", padx=10, pady=5)
        
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

        lf_support = tk.LabelFrame(win, text="Support & Diagnostics", padx=10, pady=10)
        lf_support.pack(fill="x", padx=10, pady=5)
        
        def do_export_debug():
            self.start_debug_export_thread(btn_export, win)

        btn_export = self.Btn(lf_support, text="Export Debug Bundle (Zipped Logs)", command=do_export_debug)
        btn_export.pack(fill="x", pady=2)
        
        # --- DOCUMENTATION BUTTONS ---
        f_docs = tk.Frame(lf_support); f_docs.pack(fill="x", pady=5)
        self.Btn(f_docs, text="View Changelog", command=lambda: DocViewer(win, "CHANGELOG.md", "Version History")).pack(side="left", fill="x", expand=True, padx=(0,5))
        self.Btn(f_docs, text="View User Guide", command=lambda: DocViewer(win, "README.md", "User Guide")).pack(side="left", fill="x", expand=True, padx=(5,0))

        def save():
            CFG.set("max_threads", v_threads.get())
            try: 
                px_val = int(v_pixels.get())
                if px_val <= 0: raise ValueError
                CFG.set("max_pixels", px_val)
            except: 
                 CFG.set("max_pixels", 500000000)
            
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

    def start_debug_export_thread(self, btn_ref, win_ref):
        def _run():
            ws = self.get_ws()
            # Capture the path as a string safely on the main thread
            ws_str = str(ws) if ws else None
            # Delegate to worker without touching UI
            self.worker.run_debug_export(ws_str)
            self.q.put(("export_reset_btn", btn_ref))
        
        btn_ref.config(text="Exporting...", state="disabled")
        threading.Thread(target=_run, daemon=True).start()

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
            f_path = ws/"01_Master_Files"/data['uid']
            if f_path.exists():
                 SystemUtils.open_file(f_path)

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
            # new_id = uuid.uuid4() # (Unused, clean removal)
            top.destroy(); d = filedialog.askdirectory()
            if d: 
                self.toggle(False)
                threading.Thread(target=self.wrap, args=(self.worker.run_inventory, d, mode.get()), daemon=True).start()
        
        self.Btn(top, text="Select Folder & Start", command=go).pack(fill="x", padx=20, pady=20, side="bottom")

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
        
        if '.pdf' in types: 
             self.f_pdf_ctrl.pack(fill="x", padx=10, pady=2) # Show
             self.cb_pdf.config(state="readonly")
             self.pdf_mode_var.trace_add("write", self.check_run_btn)
             self.btn_prev.config(state="normal")
        else:
             self.f_pdf_ctrl.pack_forget() 

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
            for _ in range(20):
                if self.q.empty(): break
                m = self.q.get_nowait()
                if m[0]=='log': self.log_box.insert(tk.END, m[1]+"\n"); self.log_box.see(tk.END)
                elif m[0]=='main_p': self.p_main.set(m[1]); self.lbl_status.config(text=m[2], fg="blue")
                elif m[0]=='status_blue': self.lbl_status.config(text=m[1], fg="blue")
                elif m[0]=='status': self.lbl_status.config(text=m[1], fg="orange")
                elif m[0]=='job': self.load_jobs(m[1]) 
                elif m[0]=='done': 
                    self.toggle(True); 
                    self.lbl_status.config(text="Done", fg="green"); 
                    self.btn_pause.config(text="PAUSE", bg="SystemButtonFace", state="disabled")
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
        self.root.after(100, self.poll)