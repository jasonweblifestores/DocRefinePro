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
from docrefine.config import log_app, LOG_PATH, WORKSPACES_ROOT, SystemUtils

def run():
    app = QApplication(sys.argv)
    app.setStyle("Fusion") 
    
    window = MainWindow()
    adapter = DocRefineAdapter()
    worker = Worker(callback=adapter.ingest_event)
    
    # --- UI UPDATES ---
    adapter.sig_log.connect(window.update_log)
    adapter.sig_progress_main.connect(window.update_progress)
    adapter.sig_status.connect(lambda s, m, c: window.update_status_label(s, m, c))
    
    def on_done():
        window.set_processing_state(False)
        window.refresh_job_list(get_selected_ws())
    adapter.sig_done.connect(on_done)
    
    def on_job_data(path_str):
        window.refresh_job_list(auto_select_path=path_str)
    adapter.sig_job_data.connect(on_job_data)
    
    adapter.sig_worker_config.connect(window.setup_slots)
    adapter.sig_slot_update.connect(window.update_slot)
    
    def handle_notification(data):
        if data['title'] == "Preview Ready":
             SystemUtils.open_file(data['open_path'])
        else:
             QMessageBox.information(window, data['title'], data['msg'])
             if 'open_path' in data: SystemUtils.open_file(data['open_path'])
    adapter.sig_notification.connect(handle_notification)

    # --- BUTTONS ---
    window.btn_stop.clicked.connect(worker.stop)
    
    def toggle_pause(checked):
        if checked: 
            worker.log("Pausing operation...")
            worker.pause()
            window.pause_timer(True)
            window.btn_pause.setText("Resume")
        else: 
            worker.log("Resuming operation...")
            worker.resume()
            window.pause_timer(False)
            window.btn_pause.setText("Pause")
    window.btn_pause.clicked.connect(toggle_pause)

    # --- FORENSIC ---
    def start_forensic(file_id):
        ws = window.job_tree.selectedItems()[0].data(0, Qt.UserRole)
        ws_path = Path(ws)
        entry = None
        for k, v in window.current_manifest.items():
            if v.get('id') == file_id: entry = v; break
        
        if not entry: return
        master = ws_path / "01_Master_Files" / entry['uid']
        dups = []
        if 'root' in entry:
            root = Path(entry['root'])
            for copy_rel in entry.get('copies', []):
                if copy_rel != entry.get('master'):
                    d_path = root / copy_rel
                    if d_path.exists(): dups.append(d_path)
        
        if not dups:
            QMessageBox.information(window, "Info", "No duplicates accessible.")
            return
        
        ForensicDialog(ws_path, window.current_manifest, master, dups, window).exec()

    window.req_compare.connect(start_forensic)

    # --- ACTIONS ---
    def get_selected_ws():
        items = window.job_tree.selectedItems()
        return str(items[0].data(0, Qt.UserRole)) if items else None

    window.btn_delete.clicked.connect(lambda: delete_job(get_selected_ws()))
    window.btn_open_folder.clicked.connect(lambda: SystemUtils.open_file(get_selected_ws()))
    
    def delete_job(ws):
        if not ws: return
        if QMessageBox.question(window, "Confirm", "Delete this job?") == QMessageBox.Yes:
            try: shutil.rmtree(ws)
            except: pass
            window.refresh_job_list(None)

    def open_receipt():
        ws = Path(get_selected_ws())
        rpt = list((ws/"04_Reports").glob("*.html"))
        if rpt: SystemUtils.open_file(rpt[0])
    window.btn_receipt.clicked.connect(open_receipt)

    window.btn_logs.clicked.connect(lambda: InternalViewerDialog("Log", Path(LOG_PATH).read_text(encoding='utf-8', errors='ignore'), window).exec())
    
    def open_settings():
        dlg = SettingsDialog(window)
        dlg.btn_cl.clicked.disconnect()
        dlg.btn_ug.clicked.disconnect()
        def view_doc(f): 
            p = SystemUtils.find_doc_file(f)
            if p: InternalViewerDialog(f, p.read_text(encoding='utf-8', errors='ignore'), window).exec()
        dlg.btn_cl.clicked.connect(lambda: view_doc("CHANGELOG.md"))
        dlg.btn_ug.clicked.connect(lambda: view_doc("README.md"))
        dlg.btn_export_debug.clicked.disconnect()
        dlg.btn_export_debug.clicked.connect(lambda: threading.Thread(target=worker.run_debug_export, args=(get_selected_ws(),), daemon=True).start())
        dlg.exec()
    
    window.btn_settings.clicked.connect(open_settings)

    def resolve_file_path(file_id):
        ws = get_selected_ws()
        if not ws: return None, "No job"
        target = None
        for k, v in window.current_manifest.items():
            if v.get('id') == file_id: target = v; break
        if not target: return None, "ID not found"
        return Path(ws) / "01_Master_Files" / target['uid'], "OK"

    def on_inspector_open(file_id):
        p, err = resolve_file_path(file_id)
        if p and p.exists(): SystemUtils.open_file(p)
        else: QMessageBox.warning(window, "Error", f"File missing: {err}")

    def on_inspector_reveal(file_id):
        p, err = resolve_file_path(file_id)
        if p:
            clean = str(p.resolve())
            if SystemUtils.IS_WIN:
                # FIX: Use ShellExecute to bypass SW_HIDE patch
                # 1 = SW_SHOWNORMAL
                ctypes.windll.shell32.ShellExecuteW(None, "open", "explorer.exe", f'/select,"{clean}"', None, 1)
            else:
                SystemUtils.reveal_file(clean)

    window.req_open_file.connect(on_inspector_open)
    window.req_reveal_file.connect(on_inspector_reveal)
    window.insp_tree.itemDoubleClicked.connect(lambda item, _: on_inspector_open(item.text(0)))

    # --- PROCESS START ---
    def start_process(target, args, multi_threaded=False):
        window.set_processing_state(True, multi_threaded=multi_threaded)
        threading.Thread(target=target, args=args, daemon=True).start()

    def launch_new_job():
        d = NewJobDialog(window)
        if d.exec():
            start_process(worker.run_inventory, (d.selected_path, d.selected_mode), multi_threaded=False)
    window.btn_new_job.clicked.connect(launch_new_job)

    def launch_refine():
        ws = get_selected_ws()
        if not ws: return
        opts = {
            "resize": window.chk_resize.isChecked(),
            "img2pdf": window.chk_img2pdf.isChecked(),
            "sanitize": window.chk_sanitize.isChecked(),
            "pdf_mode": ['none','flatten','ocr'][window.cb_pdf_mode.currentIndex()],
            "dpi": [150, 300, 600][window.cb_dpi.currentIndex()]
        }
        start_process(worker.run_batch, (ws, opts), multi_threaded=True)
    window.btn_run_refine.clicked.connect(launch_refine)
    
    window.btn_preview.clicked.connect(lambda: start_process(worker.run_preview, (get_selected_ws(), [150,300,600][window.cb_dpi.currentIndex()])))
    window.btn_org.clicked.connect(lambda: start_process(worker.run_organize, (get_selected_ws(), window.cb_prio.currentText())))
    window.btn_dist.clicked.connect(lambda: start_process(worker.run_distribute, (get_selected_ws(), None, window.cb_prio.currentText())))
    window.btn_csv.clicked.connect(lambda: start_process(worker.run_full_export, (get_selected_ws(),)))

    window.refresh_job_list()
    window.show()
    sys.exit(app.exec())