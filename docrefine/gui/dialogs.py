import os
import sys
import webbrowser
from pathlib import Path
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QLabel, QRadioButton, 
    QButtonGroup, QPushButton, QFileDialog, QFrame,
    QSpinBox, QComboBox, QGroupBox, QGridLayout, QLineEdit,
    QHBoxLayout, QMessageBox, QWidget, QTextEdit, QDialogButtonBox
)
from PySide6.QtCore import Qt, QSize
from PySide6.QtGui import QFont, QColor, QPalette
from docrefine.config import CFG, SystemUtils

# --- HELPER: Tesseract ---
def get_tesseract_langs():
    try:
        import pytesseract
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

# --- DIALOGS ---

class InternalViewerDialog(QDialog):
    """Generic Read-Only Text Viewer for Docs/Logs"""
    def __init__(self, title, content, parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(800, 600)
        
        layout = QVBoxLayout(self)
        
        self.txt = QTextEdit()
        self.txt.setReadOnly(True)
        self.txt.setPlainText(content)
        # Set a monospaced font for logs/markdown
        font = QFont("Consolas", 10)
        font.setStyleHint(QFont.Monospace)
        self.txt.setFont(font)
        
        layout.addWidget(self.txt)
        
        btns = QDialogButtonBox(QDialogButtonBox.Close)
        btns.rejected.connect(self.close)
        layout.addWidget(btns)

class NewJobDialog(QDialog):
    def __init__(self, parent=None, default_mode="Standard"):
        super().__init__(parent)
        self.setWindowTitle("New Job Setup")
        self.resize(450, 450)
        self.selected_mode = default_mode
        self.selected_path = None
        
        layout = QVBoxLayout(self)
        lbl = QLabel("Select Ingest Mode")
        lbl.setStyleSheet("font-weight: bold; font-size: 11pt;")
        layout.addWidget(lbl)
        
        self.bg = QButtonGroup(self)
        self.modes = [
            ("Standard (Recommended)", "Smart Text Hash (PDFs).\nStrict Binary Hash (Others).", "Standard"),
            ("Lightning (Fastest)", "Strict Binary Hash (All Files).\nExact digital copies only.", "Lightning"),
            ("Deep Scan (Slowest)", "Full Text Scan (PDFs).\nStrict Binary Hash (Others).", "Deep")
        ]
        
        for text, desc, val in self.modes:
            frame = QFrame()
            frame.setFrameShape(QFrame.StyledPanel)
            fl = QVBoxLayout(frame)
            fl.setSpacing(2)
            rb = QRadioButton(text)
            if val == default_mode: rb.setChecked(True)
            fl.addWidget(rb)
            self.bg.addButton(rb)
            lbl_d = QLabel(desc)
            # Use standard palette color for description instead of hardcoded grey
            lbl_d.setStyleSheet("margin-left: 20px; font-size: 9pt;")
            # Manually dim it slightly if needed, or rely on opacity
            opacity_eff = QPalette()
            opacity_eff.setColor(QPalette.WindowText, QColor(128, 128, 128))
            
            fl.addWidget(lbl_d)
            layout.addWidget(frame)
            rb.mode_value = val

        layout.addStretch()
        btn = QPushButton("Select Folder && Start") # && escapes to &
        btn.setStyleSheet("padding: 8px; font-weight: bold;")
        btn.clicked.connect(self.on_submit)
        layout.addWidget(btn)

    def on_submit(self):
        self.selected_mode = self.bg.checkedButton().mode_value
        d = QFileDialog.getExistingDirectory(self, "Select Source Folder")
        if d:
            self.selected_path = d
            self.accept()

class SettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Preferences")
        self.resize(600, 650)
        
        layout = QVBoxLayout(self)
        
        # Performance
        gb_perf = QGroupBox("Processing Engine")
        gl_perf = QGridLayout(gb_perf)
        gl_perf.addWidget(QLabel("Max Threads (0=Auto):"), 0, 0)
        self.spin_threads = QSpinBox()
        self.spin_threads.setRange(0, 32)
        self.spin_threads.setValue(int(CFG.get("max_threads")))
        gl_perf.addWidget(self.spin_threads, 0, 1)
        
        gl_perf.addWidget(QLabel("Safety Cap (Max Pixels):"), 1, 0)
        self.txt_pixels = QLineEdit(str(CFG.get("max_pixels")))
        gl_perf.addWidget(self.txt_pixels, 1, 1)
        layout.addWidget(gb_perf)
        
        # Defaults
        gb_def = QGroupBox("Workflow Defaults")
        gl_def = QGridLayout(gb_def)
        gl_def.addWidget(QLabel("Default Ingest:"), 0, 0)
        self.cb_ingest = QComboBox()
        self.cb_ingest.addItems(["Standard", "Lightning", "Deep"])
        self.cb_ingest.setCurrentText(CFG.get("default_ingest_mode"))
        gl_def.addWidget(self.cb_ingest, 0, 1)
        
        gl_def.addWidget(QLabel("Default Export:"), 1, 0)
        self.cb_export = QComboBox()
        self.cb_export.addItems(["Auto (Best Available)", "Force: OCR (Searchable)", "Force: Flattened (Visual)", "Force: Original Masters"])
        self.cb_export.setCurrentText(CFG.get("default_export_prio"))
        gl_def.addWidget(self.cb_export, 1, 1)
        layout.addWidget(gb_def)
        
        # OCR
        gb_ocr = QGroupBox("Optical Character Recognition (OCR)")
        gl_ocr = QVBoxLayout(gb_ocr)
        
        row_lang = QHBoxLayout()
        row_lang.addWidget(QLabel("Tesseract Language:"))
        self.cb_lang = QComboBox()
        langs = get_tesseract_langs()
        self.cb_lang.addItems(langs)
        
        cur = CFG.get("ocr_lang")
        idx = self.cb_lang.findText(cur, Qt.MatchContains)
        if idx >= 0: self.cb_lang.setCurrentIndex(idx)
        
        row_lang.addWidget(self.cb_lang)
        gl_ocr.addLayout(row_lang)
        
        row_btns = QHBoxLayout()
        btn_open_tess = QPushButton("Open Language Folder")
        btn_open_tess.clicked.connect(self.open_tess_folder)
        row_btns.addWidget(btn_open_tess)
        
        btn_get_langs = QPushButton("Get Languages (Web)")
        btn_get_langs.clicked.connect(lambda: webbrowser.open("https://github.com/tesseract-ocr/tessdata_best"))
        row_btns.addWidget(btn_get_langs)
        gl_ocr.addLayout(row_btns)
        layout.addWidget(gb_ocr)
        
        # Support
        gb_supp = QGroupBox("Support & Diagnostics")
        gl_supp = QVBoxLayout(gb_supp)
        
        self.btn_export_debug = QPushButton("Export Debug Bundle (Zipped Logs)")
        gl_supp.addWidget(self.btn_export_debug) # Connected in parent
        
        row_docs = QHBoxLayout()
        self.btn_cl = QPushButton("View Changelog")
        self.btn_ug = QPushButton("View User Guide")
        row_docs.addWidget(self.btn_cl)
        row_docs.addWidget(self.btn_ug)
        gl_supp.addLayout(row_docs)
        layout.addWidget(gb_supp)
        
        # Save
        btn_save = QPushButton("Save && Close") # FIX: Escape ampersand
        btn_save.setStyleSheet("font-weight: bold; padding: 8px;")
        btn_save.clicked.connect(self.save)
        layout.addWidget(btn_save)
        
    def open_tess_folder(self):
        try:
            import pytesseract
            path = os.environ.get("TESSDATA_PREFIX")
            if not path:
                path = str(Path(pytesseract.pytesseract.tesseract_cmd).parent / "tessdata")
            SystemUtils.open_file(path)
        except:
            QMessageBox.warning(self, "Error", "Could not locate Tesseract folder.")

    def save(self):
        CFG.set("max_threads", self.spin_threads.value())
        try:
            CFG.set("max_pixels", int(self.txt_pixels.text()))
        except: pass
        CFG.set("default_ingest_mode", self.cb_ingest.currentText())
        CFG.set("default_export_prio", self.cb_export.currentText())
        
        txt = self.cb_lang.currentText()
        if "(" in txt: code = txt.split("(")[1].replace(")", "")
        else: code = txt
        CFG.set("ocr_lang", code)
        
        self.accept()