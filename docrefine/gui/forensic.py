from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, 
    QGraphicsView, QGraphicsScene, QGraphicsPixmapItem, QSplitter,
    QFrame, QMessageBox, QWidget, QSizePolicy
)
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QPixmap, QImage, QColor, QBrush
from docrefine.processing import convert_from_path, POPPLER_BIN
import time

class SyncGraphicsView(QGraphicsView):
    req_zoom = Signal(float)
    req_page = Signal(int)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setDragMode(QGraphicsView.ScrollHandDrag)
        self.setTransformationAnchor(QGraphicsView.AnchorUnderMouse)
        self.setResizeAnchor(QGraphicsView.AnchorUnderMouse)
        self.setBackgroundBrush(QBrush(QColor("#303030")))
        self.setScene(QGraphicsScene())
        self.last_scroll = 0

    def set_image(self, pixmap):
        self.scene().clear()
        item = QGraphicsPixmapItem(pixmap)
        self.scene().addItem(item)
        self.fitInView(self.scene().itemsBoundingRect(), Qt.KeepAspectRatio)

    def wheelEvent(self, event):
        modifiers = event.modifiers()
        if modifiers & Qt.ControlModifier:
            factor = 1.25 if event.angleDelta().y() > 0 else 0.8
            self.apply_zoom(factor)
            self.req_zoom.emit(factor)
        else:
            now = time.time()
            if now - self.last_scroll > 0.4: # Debounce page turns
                delta = 1 if event.angleDelta().y() < 0 else -1
                self.req_page.emit(delta)
                self.last_scroll = now

    def apply_zoom(self, factor):
        self.scale(factor, factor)

class ForensicDialog(QDialog):
    def __init__(self, ws_path, manifest, master_path, dup_candidates, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Forensic Verification (Sync View)")
        self.resize(1400, 900)
        self.setWindowFlags(self.windowFlags() | Qt.WindowMaximizeButtonHint)
        
        self.ws_path = ws_path
        self.manifest = manifest
        self.master_path = master_path
        self.dups = dup_candidates
        self.dup_idx = 0
        self.page = 1
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        
        # --- Toolbar ---
        toolbar = QHBoxLayout()
        toolbar.addWidget(QLabel("MASTER:"))
        lbl_m = QLabel(master_path.name)
        lbl_m.setStyleSheet("font-weight: bold; color: #2ecc71; font-size: 11pt;")
        toolbar.addWidget(lbl_m)
        toolbar.addStretch()
        
        btn_prev = QPushButton("◀")
        btn_prev.setFixedWidth(40)
        btn_prev.clicked.connect(lambda: self.change_page(-1))
        self.lbl_page = QLabel("Page 1")
        self.lbl_page.setStyleSheet("font-weight: bold; padding: 0 15px;")
        btn_next = QPushButton("▶")
        btn_next.setFixedWidth(40)
        btn_next.clicked.connect(lambda: self.change_page(1))
        
        toolbar.addWidget(btn_prev)
        toolbar.addWidget(self.lbl_page)
        toolbar.addWidget(btn_next)
        toolbar.addSpacing(30)
        
        btn_zo = QPushButton("−")
        btn_zo.setFixedWidth(30)
        btn_zo.clicked.connect(lambda: self.manual_zoom(0.8))
        btn_zi = QPushButton("+")
        btn_zi.setFixedWidth(30)
        btn_zi.clicked.connect(lambda: self.manual_zoom(1.25))
        toolbar.addWidget(btn_zo)
        toolbar.addWidget(btn_zi)
        
        toolbar.addStretch()
        toolbar.addWidget(QLabel("CANDIDATE:"))
        self.lbl_dup_name = QLabel("")
        self.lbl_dup_name.setStyleSheet("font-weight: bold; color: #e74c3c; font-size: 11pt;")
        toolbar.addWidget(self.lbl_dup_name)
        layout.addLayout(toolbar)
        
        self.lbl_full_path = QLabel("Path: ...")
        self.lbl_full_path.setStyleSheet("color: #888; font-family: Consolas; margin-bottom: 5px;")
        layout.addWidget(self.lbl_full_path)
        
        # --- Split View ---
        splitter = QSplitter(Qt.Horizontal)
        splitter.setHandleWidth(12)
        
        self.view_master = SyncGraphicsView()
        self.view_dup = SyncGraphicsView()
        
        # Syncing
        self.view_master.verticalScrollBar().valueChanged.connect(self.view_dup.verticalScrollBar().setValue)
        self.view_dup.verticalScrollBar().valueChanged.connect(self.view_master.verticalScrollBar().setValue)
        self.view_master.horizontalScrollBar().valueChanged.connect(self.view_dup.horizontalScrollBar().setValue)
        self.view_dup.horizontalScrollBar().valueChanged.connect(self.view_master.horizontalScrollBar().setValue)
        
        self.view_master.req_zoom.connect(self.view_dup.apply_zoom)
        self.view_dup.req_zoom.connect(self.view_master.apply_zoom)
        self.view_master.req_page.connect(self.change_page)
        self.view_dup.req_page.connect(self.change_page)

        splitter.addWidget(self.view_master)
        splitter.addWidget(self.view_dup)
        
        # FIX: Force splitter to expand
        layout.addWidget(splitter, 1)
        
        # Bottom
        bot = QHBoxLayout()
        btn_prev_d = QPushButton("Previous Candidate")
        btn_prev_d.clicked.connect(lambda: self.change_dup(-1))
        self.lbl_dup_idx = QLabel("1 / N")
        btn_next_d = QPushButton("Next Candidate")
        btn_next_d.clicked.connect(lambda: self.change_dup(1))
        
        bot.addWidget(btn_prev_d)
        bot.addWidget(self.lbl_dup_idx)
        bot.addWidget(btn_next_d)
        bot.addStretch()
        
        btn_mark = QPushButton("MARK UNIQUE")
        btn_mark.setStyleSheet("background-color: green; color: white; font-weight: bold; padding: 6px 15px;")
        btn_mark.clicked.connect(self.mark_unique)
        bot.addWidget(btn_mark)
        layout.addLayout(bot)
        
        self.load_images()

    def render_file(self, path):
        if not path.exists(): return None
        if path.suffix.lower() == '.pdf':
            try:
                images = convert_from_path(str(path), first_page=self.page, last_page=self.page, poppler_path=POPPLER_BIN)
                if images:
                    im = images[0].convert("RGBA")
                    data = im.tobytes("raw", "RGBA")
                    qim = QImage(data, im.size[0], im.size[1], QImage.Format_RGBA8888)
                    return QPixmap.fromImage(qim)
            except: pass
        else:
            return QPixmap(str(path))
        return None

    def load_images(self):
        if not self.dups: return
        current_dup = self.dups[self.dup_idx]
        self.lbl_dup_name.setText(current_dup.name)
        self.lbl_full_path.setText(f"Candidate: {current_dup}")
        self.lbl_dup_idx.setText(f"{self.dup_idx + 1} / {len(self.dups)}")
        self.lbl_page.setText(f"Page {self.page}")

        pix_m = self.render_file(self.master_path)
        if pix_m: self.view_master.set_image(pix_m)
        pix_d = self.render_file(current_dup)
        if pix_d: self.view_dup.set_image(pix_d)

    def change_page(self, delta):
        if self.page + delta > 0:
            self.page += delta
            self.load_images()

    def change_dup(self, delta):
        new_idx = self.dup_idx + delta
        if 0 <= new_idx < len(self.dups):
            self.dup_idx = new_idx
            self.load_images()
            
    def manual_zoom(self, factor):
        self.view_master.apply_zoom(factor)
        self.view_dup.apply_zoom(factor)

    def mark_unique(self):
        QMessageBox.information(self, "Info", "Promotion logic pending backend connection.")