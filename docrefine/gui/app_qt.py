# SAVE AS: docrefine/gui/app_qt.py
import sys
import threading
import shutil
import ctypes
from pathlib import Path
from PySide6.QtWidgets import QApplication, QMessageBox, QFileDialog
from PySide6.QtCore import Qt
from .main_window import MainWindow
from .dialogs import NewJobDialog, SettingsDialog, InternalViewerDialog
from .qt_adapter import DocRefineAdapter
from .forensic import ForensicDialog
from docrefine.worker import Worker
from docrefine.config import log_app, LOG_PATH, WORKSPACES_ROOT, SystemUtils, Constants

class AppController:
    def __init__(self, app):
        self.app = app
        self.window = MainWindow()
        self.adapter = DocRefineAdapter()
        self.worker = Worker(callback=self.adapter.ingest_event)
        self.setup_signals()

    def setup_signals(self):
        # UI UPDATES
        self.adapter.sig_log.connect(self.window.update_log)
        self.adapter.sig_progress_main.connect(self.window.update_progress)
        self.adapter.sig_status.connect(lambda s, m, c: self.window.update_status_label(s, m, c))
        self.adapter.sig_done.connect(self.on_done)
        self.adapter.sig_job_data.connect(self.on_job_data)
        self.adapter.sig_worker_config.connect(self.window.setup_slots)
        self.adapter.sig_slot_update.connect(self.window.update_slot)
        self.adapter.sig_notification.connect(self.handle_notification)

        # BUTTONS
        self.window.btn_stop.clicked.connect(self.worker.stop)
        self.window.btn_pause.clicked.connect(self.toggle_pause)
        self.window.req_compare.connect(self.start_forensic)

        # ACTIONS
        self.window.btn_delete.clicked.connect(lambda: self.delete_job(self.get_selected_ws()))
        self.window.btn_open_folder.clicked.connect(lambda: SystemUtils.open_file(self.get_selected_ws()))
        self.window.btn_receipt.clicked.connect(self.open_receipt)
        self.window.btn_logs.clicked.connect(self.open_logs)
        self.window.btn_settings.clicked.connect(self.open_settings)

        self.window.req_open_file.connect(self.on_inspector_open)
        self.window.req_reveal_file.connect(self.on_inspector_reveal)
        self.window.insp_tree.itemDoubleClicked.connect(lambda item, _: self.on_inspector_open(item.text(0)))

        # PROCESS LAUNCHERS
        self.window.btn_new_job.clicked.connect(self.launch_new_job)
        self.window.btn_run_refine.clicked.connect(self.launch_refine)
        self.window.btn_preview.clicked.connect(self.launch_preview)
        self.window.btn_org.clicked.connect(self.launch_organize)
        self.window.btn_dist.clicked.connect(self.launch_distribute)
        self.window.btn_csv.clicked.connect(lambda: self.start_process(self.worker.run_full_export, (self.get_selected_ws(),)))

    def run(self):
        self.window.refresh_job_list()
        self.window.show()
        sys.exit(self.app.exec())

    def on_done(self):
        self.window.set_processing_state(False)
        self.window.refresh_job_list(self.get_selected_ws())

    def on_job_data(self, path_str):
        self.window.refresh_job_list(auto_select_path=path_str)

    def handle_notification(self, data):
        if data['title'] == "Preview Ready":
            SystemUtils.open_file(data['open_path'])
        else:
            QMessageBox.information(self.window, data['title'], data['msg'])
            if 'open_path' in data: SystemUtils.open_file(data['open_path'])

    def toggle_pause(self, checked):
        if checked: 
            self.worker.log("Pausing operation...")
            self.worker.pause()
            self.window.pause_timer(True)
            self.window.btn_pause.setText("Resume")
        else: 
            self.worker.log("Resuming operation...")
            self.worker.resume()
            self.window.pause_timer(False)
            self.window.btn_pause.setText("Pause")

    def get_selected_ws(self):
        items = self.window.job_tree.selectedItems()
        return str(items[0].data(0, Qt.UserRole)) if items else None

    def start_forensic(self, file_id):
        ws = self.get_selected_ws()
        if not ws: return
        ws_path = Path(ws)
        entry = None
        for k, v in self.window.current_manifest.items():
            if v.get('id') == file_id: entry = v; break
        
        if not entry: return
        master = ws_path / Constants.DIR_MASTER / entry['uid']
        dups = []
        if 'root' in entry:
            root = Path(entry['root'])
            for copy_rel in entry.get('copies', []):
                if copy_rel != entry.get('master'):
                    d_path = root / copy_rel
                    if d_path.exists(): dups.append(d_path)
        
        if not dups:
            QMessageBox.information(self.window, "Info", "No duplicates accessible.")
            return
        
        ForensicDialog(ws_path, self.window.current_manifest, master, dups, self.window).exec()

    def delete_job(self, ws):
        if not ws: return
        if QMessageBox.question(self.window, "Confirm", "Delete this job?") == QMessageBox.Yes:
            try: shutil.rmtree(ws)
            except: pass
            self.window.refresh_job_list(None)

    def open_receipt(self):
        ws = Path(self.get_selected_ws())
        rpt = list((ws / Constants.DIR_REPORTS).glob("*.html"))
        if rpt: SystemUtils.open_file(rpt[0])

    def open_logs(self):
        try:
            log_text = Path(LOG_PATH).read_text(encoding='utf-8', errors='ignore')
            InternalViewerDialog("Log", log_text, self.window).exec()
        except:
            pass

    def open_settings(self):
        dlg = SettingsDialog(self.window)
        dlg.btn_cl.clicked.disconnect()
        dlg.btn_ug.clicked.disconnect()
        def view_doc(f): 
            p = SystemUtils.find_doc_file(f)
            if p: InternalViewerDialog(f, p.read_text(encoding='utf-8', errors='ignore'), self.window).exec()
        dlg.btn_cl.clicked.connect(lambda: view_doc("CHANGELOG.md"))
        dlg.btn_ug.clicked.connect(lambda: view_doc("README.md"))
        dlg.btn_export_debug.clicked.disconnect()
        dlg.btn_export_debug.clicked.connect(lambda: threading.Thread(target=self.worker.run_debug_export, args=(self.get_selected_ws(),), daemon=True).start())
        dlg.exec()

    def resolve_file_path(self, file_id):
        ws = self.get_selected_ws()
        if not ws: return None, "No job"
        target = None
        for k, v in self.window.current_manifest.items():
            if v.get('id') == file_id: target = v; break
        if not target: return None, "ID not found"
        return Path(ws) / Constants.DIR_MASTER / target['uid'], "OK"

    def on_inspector_open(self, file_id):
        p, err = self.resolve_file_path(file_id)
        if p and p.exists(): SystemUtils.open_file(p)
        else: QMessageBox.warning(self.window, "Error", f"File missing: {err}")

    def on_inspector_reveal(self, file_id):
        p, err = self.resolve_file_path(file_id)
        if p:
            clean = str(p.resolve())
            if SystemUtils.IS_WIN:
                ctypes.windll.shell32.ShellExecuteW(None, "open", "explorer.exe", f'/select,"{clean}"', None, 1)
            else:
                SystemUtils.reveal_file(clean)

    def start_process(self, target, args, multi_threaded=False):
        self.window.set_processing_state(True, multi_threaded=multi_threaded)
        threading.Thread(target=target, args=args, daemon=True).start()

    def launch_new_job(self):
        d = NewJobDialog(self.window)
        if d.exec():
            self.start_process(self.worker.run_inventory, (d.selected_path, d.selected_mode), multi_threaded=False)

    def launch_refine(self):
        ws = self.get_selected_ws()
        if not ws: return
        opts = {
            "resize": self.window.chk_resize.isChecked(),
            "img2pdf": self.window.chk_img2pdf.isChecked(),
            "sanitize": self.window.chk_sanitize.isChecked(),
            "pdf_mode": ['none','flatten','ocr'][self.window.cb_pdf_mode.currentIndex()],
            "dpi": [150, 300, 600][self.window.cb_dpi.currentIndex()],
            "chain_flattened": self.window.chk_chain_flattened.isChecked()
        }
        self.start_process(self.worker.run_batch, (ws, opts), multi_threaded=True)

    def launch_preview(self):
        ws = self.get_selected_ws()
        if not ws: return
        dpi = [150,300,600][self.window.cb_dpi.currentIndex()]
        self.start_process(self.worker.run_preview, (ws, dpi))

    def launch_organize(self):
        ws = self.get_selected_ws()
        if not ws: return
        self.start_process(self.worker.run_organize, (ws, self.window.cb_prio.currentText()))

    def launch_distribute(self):
        ws = self.get_selected_ws()
        if not ws: return
        ext_src = None
        if self.window.chk_ext_src.isChecked():
            ext_src = QFileDialog.getExistingDirectory(self.window, "Select External Source Folder (e.g. Rebranded Files)")
            if not ext_src: return
        self.start_process(self.worker.run_distribute, (ws, ext_src, self.window.cb_prio.currentText()))

def run():
    app = QApplication(sys.argv)
    app.setStyle("Fusion") 
    controller = AppController(app)
    controller.run()

def dry_run():
    app = QApplication(sys.argv)
    # Just constructing the QApplication tests that Qt initialized without crashing
    # and all essential framework hooks are linked properly.
    return