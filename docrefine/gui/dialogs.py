import os
import sys
import webbrowser
from pathlib import Path
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QLabel, QRadioButton, 
    QButtonGroup, QPushButton, QFileDialog, QFrame,
    QSpinBox, QComboBox, QGroupBox, QGridLayout, QLineEdit,
    QHBoxLayout, QMessageBox, QWidget, QTextEdit, QDialogButtonBox,
    QCheckBox
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

class StampOptions(QGroupBox):
    """The per-page stamps the rebranding SOP asks for, each its own toggle.

    Shared by both rebranding dialogs so the two can never drift apart. Every
    stamp is off by default — the output then matches the signed-off Batch 1 and
    2 sets — and each remembers its own setting. The *wording* is not set here:
    it comes from the brand kit's brand.json, so a stamp with nothing to say
    prints nothing and the run log points at the file to edit.
    """
    KEYS = ("footer_attribution", "stamp_tagline", "stamp_version", "stamp_disclaimer")

    def __init__(self, parent=None):
        super().__init__("Page stamps — added to every page (per SOP)", parent)
        v = QVBoxLayout(self)
        hint = QLabel("Wording comes from brand.json in the brand kit. Anything not "
                      "written there is not printed.")
        hint.setStyleSheet("color: #888; font-size: 9pt;")
        hint.setWordWrap(True)
        v.addWidget(hint)

        self.chk_footer_attribution = QCheckBox(
            "Manufacturer attribution in the page footer")
        self.chk_footer_attribution.setToolTip(
            "The SOP puts the “Manufactured by … | Sold by …” line in a small footer\n"
            "on every page, rather than under the cover title.\n"
            "A manufacturer value that names a website, the seller or the brand itself\n"
            "is left off that document rather than printed wrong — the log counts them.")
        self.chk_tagline = QCheckBox("Tagline")
        self.chk_tagline.setToolTip(
            "The brand tagline (for Budget Mailboxes, “Trusted by the Nation”).\n"
            "Set “tagline” in the kit's brand.json.")
        self.chk_version = QCheckBox("Version and last-updated line")
        self.chk_version.setToolTip(
            "“Version 1.0 · Last Updated [Month Year]”. The month is the run date,\n"
            "the same for every file in a batch; override with “last_updated”.")
        self.chk_disclaimer = QCheckBox("Standard disclaimer")
        self.chk_disclaimer.setToolTip(
            "Neither task brief spells the disclaimer out — the wording lives in the\n"
            "SOP. Paste it into “disclaimer” in the kit's brand.json and it is stamped\n"
            "on every page; leave it empty and nothing is printed.")
        for c in (self.chk_footer_attribution, self.chk_tagline,
                  self.chk_version, self.chk_disclaimer):
            v.addWidget(c)
        self.load()

    def load(self):
        """Show each stamp's remembered setting."""
        self.chk_footer_attribution.setChecked(bool(CFG.get("rebrand_footer_attribution")))
        self.chk_tagline.setChecked(bool(CFG.get("rebrand_stamp_tagline")))
        self.chk_version.setChecked(bool(CFG.get("rebrand_stamp_version")))
        self.chk_disclaimer.setChecked(bool(CFG.get("rebrand_stamp_disclaimer")))

    def values(self):
        """The toggles as the dict the worker takes."""
        return {
            "footer_attribution": self.chk_footer_attribution.isChecked(),
            "stamp_tagline": self.chk_tagline.isChecked(),
            "stamp_version": self.chk_version.isChecked(),
            "stamp_disclaimer": self.chk_disclaimer.isChecked(),
        }


class RebrandDialog(QDialog):
    """Two-step rebrand: (1) Analyze a folder into a review sheet, then
    (2) Apply the reviewed sheet to produce the branded PDFs."""
    def __init__(self, parent=None, default_kit="", default_source=""):
        super().__init__(parent)
        self.setWindowTitle("Rebrand a Folder")
        self.resize(620, 320)
        self.source_path = default_source or None
        self.kit_path = default_kit or None
        self.plan_path = None
        self.complete_set = bool(CFG.get("rebrand_complete_set"))
        self.show_attribution = bool(CFG.get("rebrand_show_attribution"))
        self.keep_original_names = bool(CFG.get("rebrand_keep_original_names"))
        self.stamp_opts = {}
        self.vision_pass = bool(CFG.get("rebrand_vision_pass"))
        self.mode = None  # "analyze" or "apply"

        layout = QVBoxLayout(self)
        title = QLabel("Rebrand a folder of PDFs")
        title.setStyleSheet("font-weight: bold; font-size: 12pt;")
        layout.addWidget(title)
        layout.addWidget(QLabel("Original technical content is preserved — only Budget Mailboxes branding is added."))

        self.txt_src = QLineEdit(); self.txt_src.setReadOnly(True)
        self.txt_src.setPlaceholderText("Folder of PDFs…")
        self.txt_src.setText(default_source or "")
        btn_src = QPushButton("Browse…"); btn_src.clicked.connect(self.pick_source)
        row_src = QHBoxLayout(); row_src.addWidget(QLabel("Source folder:")); row_src.addWidget(self.txt_src, 1); row_src.addWidget(btn_src)
        layout.addLayout(row_src)

        # --- Step 1 ---
        gb1 = QGroupBox("Step 1 — Analyze (creates a review sheet)")
        v1 = QVBoxLayout(gb1)
        v1.addWidget(QLabel("The local model reads each PDF and drafts what to rebrand vs. leave as-is."))
        self.chk_vision = QCheckBox("Also look at pages that have no readable text (slower)")
        self.chk_vision.setChecked(bool(CFG.get("rebrand_vision_pass")))
        self.chk_vision.setToolTip(
            "Roughly half of a typical batch has no extractable text — a scanned guide,\n"
            "a drawing and a certificate all look identical to a text-only model, so those\n"
            "files are otherwise classified from their filename alone.\n\n"
            "With this on, a local vision model looks at the rendered page for any file the\n"
            "text pass cannot read or is unsure about, and its answer is used instead.\n\n"
            "SPEED DEPENDS ENTIRELY ON YOUR MACHINE. The vision model is about 6 GB; if it\n"
            "fits in your graphics memory expect a few seconds per file, and if it does not\n"
            "it runs partly or wholly on the processor and can be many times slower.\n"
            "The run measures your machine after a few files and logs an estimate, and you\n"
            "can Stop at any point — the text results are kept either way.")
        v1.addWidget(self.chk_vision)
        self.lbl_vision_note = QLabel(
            "Looking at pages is much slower than reading text, and how much slower "
            "depends on your graphics memory. The run logs a measured estimate early on.")
        self.lbl_vision_note.setStyleSheet("color: #888; font-size: 9pt;")
        self.lbl_vision_note.setWordWrap(True)
        self.lbl_vision_note.setVisible(self.chk_vision.isChecked())
        self.chk_vision.toggled.connect(self.lbl_vision_note.setVisible)
        v1.addWidget(self.lbl_vision_note)
        self.btn_analyze = QPushButton("Analyze → Create Review Sheet")
        self.btn_analyze.clicked.connect(self.on_analyze)
        v1.addWidget(self.btn_analyze)
        layout.addWidget(gb1)

        # --- Step 2 ---
        gb2 = QGroupBox("Step 2 — Apply (after you review the sheet in Excel)")
        v2 = QVBoxLayout(gb2)
        self.txt_kit = QLineEdit(); self.txt_kit.setReadOnly(True); self.txt_kit.setText(default_kit or "")
        self.txt_kit.setPlaceholderText("Brand kit folder (Portrait / Landscape assets)…")
        btn_kit = QPushButton("Browse…"); btn_kit.clicked.connect(self.pick_kit)
        row_kit = QHBoxLayout(); row_kit.addWidget(QLabel("Brand kit:")); row_kit.addWidget(self.txt_kit, 1); row_kit.addWidget(btn_kit)
        v2.addLayout(row_kit)
        self.txt_plan = QLineEdit(); self.txt_plan.setReadOnly(True)
        self.txt_plan.setPlaceholderText("Found automatically after Analyze…")
        btn_plan = QPushButton("Browse…"); btn_plan.clicked.connect(self.pick_plan)
        btn_reviews = QPushButton("Open folder"); btn_reviews.clicked.connect(self.open_reviews_folder)
        btn_reviews.setToolTip("Open the folder where review sheets are saved.")
        row_plan = QHBoxLayout(); row_plan.addWidget(QLabel("Review sheet:")); row_plan.addWidget(self.txt_plan, 1)
        row_plan.addWidget(btn_plan); row_plan.addWidget(btn_reviews)
        v2.addLayout(row_plan)
        self.chk_complete = QCheckBox("Complete set — also copy non-PDF files into the output")
        self.chk_complete.setChecked(self.complete_set)
        self.chk_complete.setToolTip(
            "On: the output tree is the whole upload set — images, Office docs and\n"
            "anything else are copied through unchanged alongside the branded PDFs.\n"
            "Off: PDFs only (non-PDF files stay in the source folder).")
        v2.addWidget(self.chk_complete)
        self.chk_attrib = QCheckBox("Manufacturer line on covers  (“Manufactured by … | Sold by …”)")
        self.chk_attrib.setChecked(self.show_attribution)
        self.chk_attrib.setToolTip(
            "Off: covers show the title only, matching the Batch 1 and 2 sets.\n"
            "On: adds the attribution line the task brief asks for, taken from\n"
            "the review sheet's manufacturer column.")
        v2.addWidget(self.chk_attrib)
        self.chk_keepnames = QCheckBox("Keep the original filenames")
        self.chk_keepnames.setChecked(self.keep_original_names)
        self.chk_keepnames.setToolTip(
            "Off: rename to the brief's pattern,\n"
            "     product-asset-type-budget-mailboxes.pdf (lowercase, ≤60 chars).\n"
            "On:  every file keeps the name it came in with, as Batch 1 and 2 did.\n"
            "Files left as-is always keep their original name either way.")
        v2.addWidget(self.chk_keepnames)
        self.stamps = StampOptions(self)
        v2.addWidget(self.stamps)
        self.btn_apply = QPushButton("Apply Reviewed Sheet")
        self.btn_apply.setStyleSheet("font-weight: bold;")
        self.btn_apply.clicked.connect(self.on_apply)
        v2.addWidget(self.btn_apply)
        layout.addWidget(gb2)

        note = QLabel("Output goes to a “_rebranded” folder beside the source, mirroring its structure.\n"
                      "Review sheets (Excel) are saved in Documents\\DocRefinePro_Data\\Rebrand Reviews.")
        note.setStyleSheet("color: #888; font-size: 9pt;")
        layout.addWidget(note)

        self._autofill_plan()

    def _autofill_plan(self):
        """Point Step 2 at the sheet Analyze wrote for this source, if there is one."""
        if self.source_path:
            from docrefine.reviews import find_plan
            found = find_plan(self.source_path)
            if found:
                self.plan_path = str(found); self.txt_plan.setText(str(found))
        self._sync_apply()

    def _sync_apply(self):
        """Apply only makes sense once a sheet exists — say so instead of failing later."""
        ready = bool(self.source_path and self.plan_path)
        self.btn_apply.setEnabled(ready)
        self.btn_apply.setToolTip("" if ready else
                                  "Run Analyze first — the review sheet is picked up automatically.")

    def open_reviews_folder(self):
        from docrefine.config import REVIEWS_ROOT
        REVIEWS_ROOT.mkdir(parents=True, exist_ok=True)
        SystemUtils.open_file(REVIEWS_ROOT)

    def pick_source(self):
        d = QFileDialog.getExistingDirectory(self, "Select Source Folder", self.source_path or "")
        if d:
            self.source_path = d; self.txt_src.setText(d)
            self.plan_path = None; self.txt_plan.clear()
            self._autofill_plan()

    def pick_kit(self):
        d = QFileDialog.getExistingDirectory(self, "Select Brand Kit Folder")
        if d:
            self.kit_path = d; self.txt_kit.setText(d)

    def pick_plan(self):
        from docrefine.config import REVIEWS_ROOT
        start = str(Path(self.plan_path).parent) if self.plan_path else str(REVIEWS_ROOT)
        f, _ = QFileDialog.getOpenFileName(self, "Select Review Sheet", start,
                                           "Review sheets (*.xlsx *.csv);;All files (*)")
        if f:
            self.plan_path = f; self.txt_plan.setText(f)
            self._sync_apply()

    def on_analyze(self):
        if not self.source_path:
            QMessageBox.warning(self, "Missing selection", "Please choose a source folder to analyze.")
            return
        self.vision_pass = self.chk_vision.isChecked()
        self.mode = "analyze"; self.accept()

    def on_apply(self):
        if not self.plan_path:
            self._autofill_plan()      # sheet may have been created since the dialog opened
        if not (self.source_path and self.kit_path and self.plan_path):
            QMessageBox.warning(self, "Missing selection",
                                "Apply needs a source folder, a brand kit, and a reviewed sheet.\n\n"
                                "Run Analyze first — its sheet is picked up automatically.")
            return
        self.complete_set = self.chk_complete.isChecked()
        self.show_attribution = self.chk_attrib.isChecked()
        self.keep_original_names = self.chk_keepnames.isChecked()
        self.stamp_opts = self.stamps.values()
        self.mode = "apply"; self.accept()

class PipelineDialog(QDialog):
    """Run a chain of steps over a folder, in the fixed order Flatten → Rebrand → OCR."""
    def __init__(self, parent=None, default_kit=""):
        super().__init__(parent)
        self.setWindowTitle("Process a Folder")
        self.resize(620, 360)
        self.source_path = None
        self.kit_path = default_kit or None
        self.do_flatten = self.do_rebrand = self.do_ocr = False
        self.complete_set = bool(CFG.get("rebrand_complete_set"))
        self.show_attribution = bool(CFG.get("rebrand_show_attribution"))
        self.keep_original_names = bool(CFG.get("rebrand_keep_original_names"))
        self.stamp_opts = {}

        layout = QVBoxLayout(self)
        title = QLabel("Process a folder of PDFs")
        title.setStyleSheet("font-weight: bold; font-size: 12pt;")
        layout.addWidget(title)
        layout.addWidget(QLabel("Pick the steps to run. They execute in order — OCR always runs last."))

        self.txt_src = QLineEdit(); self.txt_src.setReadOnly(True); self.txt_src.setPlaceholderText("Folder of PDFs…")
        btn_src = QPushButton("Browse…"); btn_src.clicked.connect(self.pick_source)
        row_src = QHBoxLayout(); row_src.addWidget(QLabel("Source folder:")); row_src.addWidget(self.txt_src, 1); row_src.addWidget(btn_src)
        layout.addLayout(row_src)

        gb = QGroupBox("Steps (run in this order)")
        v = QVBoxLayout(gb)
        self.chk_flatten = QCheckBox("1 · Flatten pages  (only for problem PDFs)")
        self.chk_rebrand = QCheckBox("2 · Rebrand  (uses a review sheet in the folder if present)")
        self.chk_rebrand.setChecked(True)
        self.chk_ocr = QCheckBox("3 · Make searchable — OCR  (for scanned files; runs last)")
        for c in (self.chk_flatten, self.chk_rebrand, self.chk_ocr):
            v.addWidget(c)

        self.lbl_kit = QLabel("Brand kit:")
        self.txt_kit = QLineEdit(); self.txt_kit.setReadOnly(True); self.txt_kit.setText(default_kit or "")
        self.txt_kit.setPlaceholderText("Brand kit folder (needed for the Rebrand step)…")
        self.btn_kit = QPushButton("Browse…"); self.btn_kit.clicked.connect(self.pick_kit)
        row_kit = QHBoxLayout(); row_kit.addWidget(self.lbl_kit); row_kit.addWidget(self.txt_kit, 1); row_kit.addWidget(self.btn_kit)
        v.addLayout(row_kit)
        self.chk_rebrand.toggled.connect(self._sync_kit)
        layout.addWidget(gb)

        self.chk_complete = QCheckBox("Complete set — carry non-PDF files through to the output")
        self.chk_complete.setChecked(self.complete_set)
        self.chk_complete.setToolTip(
            "On: images, Office docs and anything else are copied through unchanged,\n"
            "so the output tree is the whole set.\n"
            "Off: PDFs only (non-PDF files stay in the source folder).")
        layout.addWidget(self.chk_complete)

        self.chk_attrib = QCheckBox("Manufacturer line on covers  (“Manufactured by … | Sold by …”)")
        self.chk_attrib.setChecked(self.show_attribution)
        self.chk_attrib.setToolTip("Off: covers show the title only, matching the Batch 1 and 2 sets.")
        layout.addWidget(self.chk_attrib)

        self.chk_keepnames = QCheckBox("Keep the original filenames")
        self.chk_keepnames.setChecked(self.keep_original_names)
        self.chk_keepnames.setToolTip("On: files keep the name they came in with, as Batch 1 and 2 did.")
        layout.addWidget(self.chk_keepnames)

        self.stamps = StampOptions(self)
        layout.addWidget(self.stamps)
        self.chk_rebrand.toggled.connect(self.stamps.setEnabled)
        self.stamps.setEnabled(self.chk_rebrand.isChecked())

        note = QLabel("Output goes to a “_processed” folder beside the source.")
        note.setStyleSheet("color: #888; font-size: 9pt;")
        layout.addWidget(note)
        layout.addStretch()

        self.btn_run = QPushButton("Run Pipeline")
        self.btn_run.setStyleSheet("font-weight: bold; padding: 8px;")
        self.btn_run.clicked.connect(self.on_run)
        layout.addWidget(self.btn_run)

    def _sync_kit(self, on):
        for w in (self.lbl_kit, self.txt_kit, self.btn_kit):
            w.setEnabled(on)

    def pick_source(self):
        d = QFileDialog.getExistingDirectory(self, "Select Source Folder")
        if d:
            self.source_path = d; self.txt_src.setText(d)

    def pick_kit(self):
        d = QFileDialog.getExistingDirectory(self, "Select Brand Kit Folder")
        if d:
            self.kit_path = d; self.txt_kit.setText(d)

    def on_run(self):
        if not self.source_path:
            QMessageBox.warning(self, "Missing selection", "Please choose a source folder.")
            return
        self.do_flatten = self.chk_flatten.isChecked()
        self.do_rebrand = self.chk_rebrand.isChecked()
        self.do_ocr = self.chk_ocr.isChecked()
        self.complete_set = self.chk_complete.isChecked()
        self.show_attribution = self.chk_attrib.isChecked()
        self.keep_original_names = self.chk_keepnames.isChecked()
        self.stamp_opts = self.stamps.values()
        if not (self.do_flatten or self.do_rebrand or self.do_ocr):
            QMessageBox.warning(self, "No steps selected", "Please select at least one step.")
            return
        if self.do_rebrand and not self.kit_path:
            QMessageBox.warning(self, "Brand kit needed", "The Rebrand step needs a brand kit.")
            return
        self.accept()