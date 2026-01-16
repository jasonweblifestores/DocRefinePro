# SAVE AS: docrefine/gui/main_window.py
import json
from pathlib import Path
from datetime import datetime, timedelta
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QSplitter, QTreeWidget, QTreeWidgetItem, QPushButton, 
    QLabel, QProgressBar, QTextEdit, QTabWidget, QFrame,
    QCheckBox, QComboBox, QGroupBox, QHeaderView, QLineEdit,
    QMenu, QAbstractItemView, QGridLayout
)
from PySide6.QtCore import Qt, Slot, Signal, QTimer
from PySide6.QtGui import QColor
from docrefine.config import WORKSPACES_ROOT

class NumericTreeWidgetItem(QTreeWidgetItem):
    def __lt__(self, other):
        column = self.treeWidget().sortColumn()
        text1 = self.text(column)
        text2 = other.text(column)
        try: return float(text1) < float(text2)
        except ValueError: return text1 < text2

class MainWindow(QMainWindow):
    req_open_file = Signal(str) 
    req_reveal_file = Signal(str)
    req_compare = Signal(str)

    def __init__(self):
        super().__init__()
        self.setWindowTitle("DocRefine Pro (PySide6 Era)")
        self.resize(1300, 900)
        self.current_manifest = {}
        
        # Timer
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_timer)
        self.start_time = None
        self.accumulated_time = 0
        
        # Worker Slots Logic
        self.slot_map = {} 
        self.slot_widgets = []
        
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QHBoxLayout(central)
        splitter = QSplitter(Qt.Horizontal)
        main_layout.addWidget(splitter)
        
        # --- LEFT PANEL ---
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0,0,0,0)
        
        lbl_dash = QLabel("Workspace Dashboard")
        lbl_dash.setStyleSheet("font-weight: bold; font-size: 14px;")
        left_layout.addWidget(lbl_dash)
        
        self.btn_new_job = QPushButton("+ New Ingest Job")
        self.btn_new_job.setStyleSheet("padding: 6px; font-weight: bold;")
        left_layout.addWidget(self.btn_new_job)
        
        self.job_tree = QTreeWidget()
        self.job_tree.setHeaderLabels(["Name", "Status", "Date"])
        self.job_tree.header().setSectionResizeMode(0, QHeaderView.Stretch)
        self.job_tree.setSortingEnabled(True)
        self.job_tree.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.job_tree.setSelectionMode(QAbstractItemView.SingleSelection)
        self.job_tree.itemSelectionChanged.connect(self.on_job_selected)
        left_layout.addWidget(self.job_tree)
        
        self.gb_stats = QGroupBox("Job Statistics")
        self.gb_stats.setVisible(False)
        v_stats = QVBoxLayout(self.gb_stats)
        self.lbl_stat_files = QLabel("Files: -")
        self.lbl_stat_masters = QLabel("Masters: -")
        self.lbl_stat_q = QLabel("Quarantined: -")
        
        # Time breakdown
        self.lbl_t_ingest = QLabel("Ingest: -")
        self.lbl_t_refine = QLabel("Refine: -")
        self.lbl_t_total = QLabel("Total: -")
        
        v_stats.addWidget(self.lbl_stat_files)
        v_stats.addWidget(self.lbl_stat_masters)
        v_stats.addWidget(self.lbl_stat_q)
        v_stats.addWidget(QLabel("--- Timing ---"))
        v_stats.addWidget(self.lbl_t_ingest)
        v_stats.addWidget(self.lbl_t_refine)
        v_stats.addWidget(self.lbl_t_total)
        left_layout.addWidget(self.gb_stats)

        btn_row = QHBoxLayout()
        self.btn_refresh = QPushButton("Refresh")
        self.btn_refresh.clicked.connect(lambda: self.refresh_job_list(None))
        self.btn_delete = QPushButton("Delete Job")
        self.btn_delete.setStyleSheet("color: #ff5555; font-weight: bold;") 
        btn_row.addWidget(self.btn_refresh)
        btn_row.addWidget(self.btn_delete)
        left_layout.addLayout(btn_row)
        
        self.btn_settings = QPushButton("⚙ Settings")
        self.btn_open_folder = QPushButton("📂 Open Folder")
        self.btn_open_folder.setEnabled(False)
        self.btn_logs = QPushButton("📜 App Log")
        left_layout.addWidget(self.btn_settings)
        left_layout.addWidget(self.btn_open_folder)
        left_layout.addWidget(self.btn_logs)
        splitter.addWidget(left_panel)
        
        # --- RIGHT PANEL ---
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        self.tabs = QTabWidget()
        self.tab_refine = QWidget(); self._build_refine_tab()
        self.tab_export = QWidget(); self._build_export_tab()
        self.tab_inspect = QWidget(); self._build_inspector_tab()
        
        # FIX: Reordered Tabs (Inspector First)
        self.tabs.addTab(self.tab_inspect, "🔍 1. Inspector")
        self.tabs.addTab(self.tab_refine, "2. Refine")
        self.tabs.addTab(self.tab_export, "3. Export")
        right_layout.addWidget(self.tabs)
        
        monitor_frame = QFrame()
        monitor_frame.setFrameShape(QFrame.StyledPanel)
        mon_layout = QVBoxLayout(monitor_frame)
        
        # FIX: Stable Monitor Layout
        head_layout = QHBoxLayout()
        self.lbl_status = QLabel("Ready")
        self.lbl_status.setStyleSheet("color: #0078d7; font-weight: bold;")
        head_layout.addWidget(self.lbl_status)
        
        # Spacer pushes everything right
        head_layout.addStretch()
        
        self.btn_receipt = QPushButton("View Receipt")
        self.btn_receipt.setEnabled(False)
        head_layout.addWidget(self.btn_receipt)
        
        self.btn_pause = QPushButton("Pause")
        self.btn_pause.setCheckable(True)
        self.btn_pause.setEnabled(False)
        head_layout.addWidget(self.btn_pause)
        
        self.btn_stop = QPushButton("STOP")
        self.btn_stop.setStyleSheet("background-color: #ff4444; color: white; font-weight: bold;")
        self.btn_stop.setEnabled(False)
        head_layout.addWidget(self.btn_stop)
        
        # Timer at the very end (fixed width to prevent jitter)
        self.lbl_timer = QLabel("00:00:00")
        self.lbl_timer.setStyleSheet("font-family: Consolas; font-weight: bold; padding-left: 10px;")
        self.lbl_timer.setFixedWidth(80)
        self.lbl_timer.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        head_layout.addWidget(self.lbl_timer)

        mon_layout.addLayout(head_layout)
        
        self.progress_main = QProgressBar()
        mon_layout.addWidget(self.progress_main)
        
        self.slot_group = QGroupBox("Active Workers")
        self.slot_layout = QGridLayout(self.slot_group)
        mon_layout.addWidget(self.slot_group)
        
        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setMaximumHeight(150)
        mon_layout.addWidget(self.log_box)
        
        right_layout.addWidget(monitor_frame)
        splitter.addWidget(right_panel)
        splitter.setSizes([350, 950])

    def _build_refine_tab(self):
        layout = QVBoxLayout(self.tab_refine)
        layout.setAlignment(Qt.AlignTop)
        layout.addWidget(QLabel("Content Refinement (Modifies Files)"))
        self.gb_pdf = QGroupBox("PDF Actions")
        gl_pdf = QVBoxLayout(self.gb_pdf)
        self.cb_pdf_mode = QComboBox()
        self.cb_pdf_mode.addItems(["No Action", "Flatten Only (Fast)", "Flatten + OCR (Slow)"])
        gl_pdf.addWidget(QLabel("Mode:"))
        gl_pdf.addWidget(self.cb_pdf_mode)
        layout.addWidget(self.gb_pdf)
        self.gb_gen = QGroupBox("General Actions")
        gl_gen = QVBoxLayout(self.gb_gen)
        self.chk_resize = QCheckBox("Resize Images (1920px HD Standard)")
        self.chk_img2pdf = QCheckBox("Bundle Images to PDF")
        self.chk_sanitize = QCheckBox("Sanitize Office Docs (Remove Metadata)")
        gl_gen.addWidget(self.chk_resize)
        gl_gen.addWidget(self.chk_img2pdf)
        gl_gen.addWidget(self.chk_sanitize)
        layout.addWidget(self.gb_gen)
        gb_qual = QGroupBox("Processing Quality")
        gl_qual = QHBoxLayout(gb_qual)
        self.cb_dpi = QComboBox()
        self.cb_dpi.addItems(["Low (Fast)", "Medium (Standard)", "High (Slow)"])
        self.cb_dpi.setCurrentIndex(1)
        gl_qual.addWidget(QLabel("DPI:"))
        gl_qual.addWidget(self.cb_dpi)
        self.btn_preview = QPushButton("Generate Preview")
        self.btn_preview.setEnabled(False)
        gl_qual.addWidget(self.btn_preview)
        layout.addWidget(gb_qual)
        self.btn_run_refine = QPushButton("Run Refinement Batch")
        self.btn_run_refine.setStyleSheet("font-weight: bold; padding: 10px;")
        self.btn_run_refine.setEnabled(False)
        layout.addWidget(self.btn_run_refine)

    def _build_export_tab(self):
        layout = QVBoxLayout(self.tab_export)
        layout.setAlignment(Qt.AlignTop)
        gb_prio = QGroupBox("Source Priority")
        vb_prio = QVBoxLayout(gb_prio)
        self.cb_prio = QComboBox()
        self.cb_prio.addItems(["Auto (Best Available)", "Force: OCR (Searchable)", "Force: Flattened (Visual)", "Force: Original Masters"])
        vb_prio.addWidget(self.cb_prio)
        self.chk_ext_src = QCheckBox("Override Source: External Folder")
        vb_prio.addWidget(self.chk_ext_src)
        layout.addWidget(gb_prio)
        gb_a = QGroupBox("Option A: Unique Masters")
        vb_a = QVBoxLayout(gb_a)
        self.btn_org = QPushButton("Export Unique Files")
        self.btn_org.setEnabled(False)
        vb_a.addWidget(self.btn_org)
        layout.addWidget(gb_a)
        gb_b = QGroupBox("Option B: Reconstruction")
        vb_b = QVBoxLayout(gb_b)
        self.btn_dist = QPushButton("Run Reconstruction")
        self.btn_dist.setEnabled(False)
        vb_b.addWidget(self.btn_dist)
        layout.addWidget(gb_b)
        gb_c = QGroupBox("Option C: Reports")
        vb_c = QVBoxLayout(gb_c)
        self.btn_csv = QPushButton("Export Full Inventory CSV")
        self.btn_csv.setEnabled(False)
        vb_c.addWidget(self.btn_csv)
        layout.addWidget(gb_c)

    def _build_inspector_tab(self):
        layout = QVBoxLayout(self.tab_inspect)
        search_layout = QHBoxLayout()
        search_layout.addWidget(QLabel("Filter:"))
        self.txt_search = QLineEdit()
        self.txt_search.textChanged.connect(self.filter_inspector)
        search_layout.addWidget(self.txt_search)
        layout.addLayout(search_layout)
        self.insp_tree = QTreeWidget()
        self.insp_tree.setHeaderLabels(["ID", "Name", "Status", "Copies"])
        self.insp_tree.header().setSectionResizeMode(1, QHeaderView.Stretch)
        self.insp_tree.setSortingEnabled(True)
        self.insp_tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.insp_tree.customContextMenuRequested.connect(self.show_insp_context_menu)
        layout.addWidget(self.insp_tree)
        layout.addWidget(QLabel("Double-click a file to inspect details."))

    def show_insp_context_menu(self, pos):
        item = self.insp_tree.itemAt(pos)
        if not item: return
        menu = QMenu(self)
        a_open = menu.addAction("Open File")
        a_reveal = menu.addAction("Reveal in Folder")
        menu.addSeparator()
        a_comp = menu.addAction("Compare Duplicates (Forensic)")
        action = menu.exec_(self.insp_tree.viewport().mapToGlobal(pos))
        file_id = item.text(0)
        if action == a_open: self.req_open_file.emit(file_id)
        elif action == a_reveal: self.req_reveal_file.emit(file_id)
        elif action == a_comp: self.req_compare.emit(file_id)

    def refresh_job_list(self, auto_select_path=None):
        selected_path = auto_select_path
        if not selected_path:
            items = self.job_tree.selectedItems()
            if items: selected_path = items[0].data(0, Qt.UserRole)

        self.job_tree.blockSignals(True)
        disk_jobs = {}
        if WORKSPACES_ROOT.exists():
            for d in WORKSPACES_ROOT.iterdir():
                if d.is_dir():
                    status = "Empty"
                    if (d/"status.json").exists():
                        try: 
                            with open(d/"status.json") as f: status = json.load(f).get("stage", "?")
                        except: pass
                    elif (d/"Final_Delivery").exists(): status = "DISTRIBUTED"
                    elif (d/"01_Master_Files").exists(): status = "INGESTED"
                    date_str = datetime.fromtimestamp(d.stat().st_mtime).strftime('%Y-%m-%d %H:%M')
                    disk_jobs[str(d)] = {"name": d.name, "status": status, "date": date_str}

        root = self.job_tree.invisibleRootItem()
        existing_paths = set()
        to_remove = []
        for i in range(root.childCount()):
            item = root.child(i)
            path = item.data(0, Qt.UserRole)
            if path in disk_jobs:
                data = disk_jobs[path]
                if item.text(1) != data['status']: item.setText(1, data['status'])
                existing_paths.add(path)
            else: to_remove.append(item)
        for item in to_remove: root.removeChild(item)
        for path, data in disk_jobs.items():
            if path not in existing_paths:
                item = QTreeWidgetItem([data['name'], data['status'], data['date']])
                item.setData(0, Qt.UserRole, path)
                self.job_tree.addTopLevelItem(item)
        self.job_tree.blockSignals(False)
        
        if selected_path:
            for i in range(root.childCount()):
                item = root.child(i)
                if item.data(0, Qt.UserRole) == selected_path:
                    item.setSelected(True); break
        self.on_job_selected()

    def on_job_selected(self):
        items = self.job_tree.selectedItems()
        enabled = bool(items)
        self.btn_run_refine.setEnabled(enabled)
        self.btn_org.setEnabled(enabled)
        self.btn_dist.setEnabled(enabled)
        self.btn_csv.setEnabled(enabled)
        self.btn_open_folder.setEnabled(enabled)
        self.btn_preview.setEnabled(enabled)
        self.gb_stats.setVisible(enabled)
        self.btn_delete.setEnabled(enabled)
        self.insp_tree.clear()
        self.current_manifest = {}
        if enabled:
            path_str = items[0].data(0, Qt.UserRole)
            if not path_str: return
            ws_path = Path(path_str)
            self.update_refine_context(ws_path)
            self.load_stats(ws_path)
            rpt_dir = ws_path / "04_Reports"
            self.btn_receipt.setEnabled(rpt_dir.exists() and any(rpt_dir.glob("*.html")))
            if (ws_path / "manifest.json").exists():
                try: 
                    with open(ws_path / "manifest.json") as f: self.current_manifest = json.load(f)
                    self.filter_inspector("")
                except: pass
        else:
            self.reset_stats()

    def reset_stats(self):
        self.lbl_stat_files.setText("Files: -")
        self.lbl_stat_masters.setText("Masters: -")
        self.lbl_stat_q.setText("Quarantined: -")
        self.lbl_t_ingest.setText("Ingest: -")
        self.lbl_t_refine.setText("Refine: -")
        self.lbl_t_total.setText("Total: -")

    def load_stats(self, ws_path):
        try:
            with open(ws_path / "stats.json") as f: s = json.load(f)
            self.lbl_stat_files.setText(f"Files Scanned: {s.get('total_scanned', 0)}")
            self.lbl_stat_masters.setText(f"Unique Masters: {s.get('masters', 0)}")
            self.lbl_stat_q.setText(f"Quarantined: {s.get('quarantined', 0)}")
            
            t_ing = s.get('ingest_time', 0)
            t_ref = s.get('batch_time', 0)
            t_tot = t_ing + t_ref + s.get('dist_time',0) + s.get('organize_time',0)
            
            self.lbl_t_ingest.setText(f"Ingest: {str(timedelta(seconds=int(t_ing)))}")
            self.lbl_t_refine.setText(f"Refine: {str(timedelta(seconds=int(t_ref)))}")
            self.lbl_t_total.setText(f"Total: {str(timedelta(seconds=int(t_tot)))}")
        except: self.reset_stats()

    def update_refine_context(self, ws_path):
        m = ws_path / "01_Master_Files"
        has_pdf = False; has_img = False; has_office = False
        if m.exists():
            for f in m.rglob('*'):
                e = f.suffix.lower()
                if e == '.pdf': has_pdf = True
                elif e in {'.jpg','.png'}: has_img = True
                elif e in {'.docx','.xlsx'}: has_office = True
        self.gb_pdf.setVisible(has_pdf); self.gb_gen.setVisible(has_img or has_office)

    def filter_inspector(self, text):
        self.insp_tree.clear()
        query = text.lower()
        for k, v in self.current_manifest.items():
            name = v.get('name', '').lower()
            if query in name or query in v.get('id','').lower():
                st = "Duplicate" if len(v.get('copies', []))>1 else "Master"
                if v.get('status') == 'QUARANTINE': st = "⛔ Quarantined"
                item = NumericTreeWidgetItem([v.get('id','?'), v.get('name','?'), st, str(len(v.get('copies',[])))])
                if "Quar" in st: item.setForeground(2, QColor("#e74c3c"))
                elif "Dup" in st: item.setForeground(2, QColor("#3498db"))
                self.insp_tree.addTopLevelItem(item)

    # --- STATE ---
    def set_processing_state(self, active, multi_threaded=False):
        self.btn_new_job.setEnabled(not active)
        self.btn_delete.setEnabled(not active)
        self.btn_run_refine.setEnabled(not active)
        self.btn_pause.setEnabled(active)
        self.btn_stop.setEnabled(active)
        self.slot_group.setVisible(active and multi_threaded)
        
        if active:
            self.reset_stats()
            self.start_time = datetime.now()
            self.timer.start(1000)
        else:
            self.timer.stop()
            self.accumulated_time = 0
            self.lbl_timer.setText("00:00:00")
            self.refresh_job_list(None)

    def pause_timer(self, paused):
        if paused:
            self.timer.stop()
            self.accumulated_time += (datetime.now() - self.start_time).total_seconds()
        else:
            self.start_time = datetime.now()
            self.timer.start(1000)

    def update_timer(self):
        if self.start_time:
            now = datetime.now()
            delta = (now - self.start_time).total_seconds() + self.accumulated_time
            self.lbl_timer.setText(str(timedelta(seconds=int(delta))))

    @Slot(int)
    def setup_slots(self, count):
        for i in reversed(range(self.slot_layout.count())): 
            self.slot_layout.itemAt(i).widget().setParent(None)
        
        self.slot_map = {} 
        self.slot_widgets = [] 
        
        row = 0; col = 0
        for i in range(count):
            lbl = QLabel(f"W{i+1}: Idle")
            lbl.setStyleSheet("background-color: #333; color: white; padding: 4px; border-radius: 4px;")
            self.slot_layout.addWidget(lbl, row, col)
            self.slot_widgets.append(lbl)
            col += 1
            if col > 3: col = 0; row += 1

    @Slot(dict)
    def update_slot(self, data):
        tid = data.get('tid')
        if tid not in self.slot_map:
            next_idx = len(self.slot_map)
            if next_idx < len(self.slot_widgets):
                self.slot_map[tid] = next_idx
            else: return 
        
        idx = self.slot_map[tid]
        self.slot_widgets[idx].setText(data.get('text'))

    @Slot(str, str)
    def update_log(self, msg, level):
        c = "#ff5555" if level == "ERROR" else "#ccc"
        if level == "INFO": c = "#ddd"
        self.log_box.append(f'<span style="color:{c}">[{level}] {msg}</span>')

    @Slot(float, str)
    def update_progress(self, percent, text):
        self.progress_main.setValue(int(percent))
        self.lbl_status.setText(text)
        
    @Slot(str, str, str)
    def update_status_label(self, stage, msg, color):
        self.lbl_status.setText(msg)
        c = {"blue": "#55aaff", "green": "#55ff55", "red": "#ff5555", "orange": "#ffaa00"}.get(color, "#ffffff")
        self.lbl_status.setStyleSheet(f"color: {c}; font-weight: bold;")

    @Slot()
    def job_done(self):
        self.progress_main.setValue(100)
        self.lbl_status.setText("Completed")
        self.set_processing_state(False)