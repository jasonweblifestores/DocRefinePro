# SAVE AS: docrefine/gui/app_qt.py
import sys
import threading
import shutil
import ctypes
from pathlib import Path
from PySide6.QtWidgets import QApplication, QMessageBox, QFileDialog
from PySide6.QtCore import Qt
from .main_window import MainWindow
from .dialogs import NewJobDialog, SettingsDialog, InternalViewerDialog, RebrandDialog, PipelineDialog
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

        # Rebrand history
        self.window.btn_run_open.clicked.connect(self.open_run_output)
        self.window.btn_run_sheet.clicked.connect(self.open_run_sheet)
        self.window.btn_run_forget.clicked.connect(self.forget_run)

        self.window.req_open_file.connect(self.on_inspector_open)
        self.window.req_reveal_file.connect(self.on_inspector_reveal)
        self.window.insp_tree.itemDoubleClicked.connect(lambda item, _: self.on_inspector_open(item.text(0)))

        # PROCESS LAUNCHERS
        self.window.btn_new_job.clicked.connect(self.launch_new_job)
        self.window.btn_rebrand.clicked.connect(self.launch_rebrand)
        self.window.btn_pipeline.clicked.connect(self.launch_pipeline)
        self.window.btn_run_refine.clicked.connect(self.launch_refine)
        self.window.btn_preview.clicked.connect(self.launch_preview)
        self.window.btn_org.clicked.connect(self.launch_organize)
        self.window.btn_dist.clicked.connect(self.launch_distribute)
        self.window.btn_csv.clicked.connect(lambda: self.start_process(self.worker.run_full_export, (self.get_selected_ws(),)))
        self.window.btn_rebrand_job.clicked.connect(self.launch_rebrand_job)

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
        if 'uid' not in entry:
            QMessageBox.information(self.window, "Info", "This is a quarantined file (no master to compare).")
            return
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

    def open_run_output(self):
        rec = self.window.selected_run()
        if rec and rec.get("output"):
            SystemUtils.open_file(rec["output"])

    def open_run_sheet(self):
        rec = self.window.selected_run()
        if rec and rec.get("sheet"):
            SystemUtils.open_file(rec["sheet"])

    def forget_run(self):
        """Drop a run from the history. Deliberately touches no files."""
        rec = self.window.selected_run()
        if not rec:
            return
        if QMessageBox.question(self.window, "Forget this run?",
                                "Remove this run from the history?\n\n"
                                "The rebranded files and the review sheet are not deleted."
                                ) != QMessageBox.Yes:
            return
        from docrefine import runs
        runs.remove(rec.get("ts"))
        self.window.refresh_run_list()

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
        if 'uid' not in target: return None, "Quarantined file (no master copy)"
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

    def launch_rebrand(self):
        self._open_rebrand_dialog()

    def launch_rebrand_job(self):
        ws = self.get_selected_ws()
        if not ws:
            return
        masters = Path(ws) / Constants.DIR_MASTER
        if not masters.exists():
            QMessageBox.information(self.window, "Info", "This job has no ingested master files yet.")
            return
        self._open_rebrand_dialog(default_source=str(masters))

    def _open_rebrand_dialog(self, default_source=""):
        from docrefine.config import CFG
        # Analyze closes the dialog, so the user reopens it to Apply — don't make
        # them re-browse to the same folder every time.
        source = default_source or CFG.get("last_rebrand_source")
        if source and not Path(source).is_dir():
            source = ""
        d = RebrandDialog(self.window, default_kit=CFG.get("last_brand_kit"), default_source=source)
        if d.exec():
            CFG.set("last_rebrand_source", d.source_path or "")
            if d.mode == "analyze":
                CFG.set("rebrand_vision_pass", d.vision_pass)
                self._start_analyze(d.source_path, d.vision_pass)
            else:
                CFG.set("last_brand_kit", d.kit_path)
                CFG.set("rebrand_complete_set", d.complete_set)
                CFG.set("rebrand_show_attribution", d.show_attribution)
                CFG.set("rebrand_keep_original_names", d.keep_original_names)
                self._remember_stamps(d.stamp_opts)
                self.start_process(self.worker.run_rebrand_apply,
                                   (d.source_path, d.kit_path, d.plan_path, None,
                                    d.complete_set, d.show_attribution,
                                    d.keep_original_names, d.stamp_opts),
                                   multi_threaded=True)

    @staticmethod
    def _remember_stamps(opts):
        """Persist each page-stamp toggle, like every other rebrand setting."""
        from docrefine.gui.dialogs import StampOptions
        from docrefine.config import CFG
        for key in StampOptions.KEYS:
            CFG.set(f"rebrand_{key}", bool((opts or {}).get(key)))

    def _start_analyze(self, source, vision_pass=False):
        """Ensure the local AI is usable before analyzing; guide the user if not."""
        from docrefine import classify
        status = classify.ollama_status()
        st = status["state"]
        if st == "not_running" and classify.start_server():
            st = "ready" if classify.has_model() else "no_model"
        if st == "ready":
            if vision_pass and not classify.has_vision_model():
                m = classify.DEFAULT_VISION_MODEL
                box = QMessageBox(self.window)
                box.setWindowTitle("Visual pass")
                box.setIcon(QMessageBox.Question)
                box.setText(f"Looking at pages needs the vision model '{m}', which isn't "
                            f"downloaded yet (about 6 GB, one time).\n\n"
                            f"Download it now, or analyze without looking at pages?")
                get = box.addButton("Download model", QMessageBox.AcceptRole)
                skip = box.addButton("Analyze without it", QMessageBox.DestructiveRole)
                box.addButton(QMessageBox.Cancel)
                box.exec()
                if box.clickedButton() is get:
                    self.start_process(self.worker.run_pull_model, (m,), multi_threaded=False)
                    return
                if box.clickedButton() is not skip:
                    return
                vision_pass = False
            self.start_process(self.worker.run_rebrand_analyze, (source, None, vision_pass),
                               multi_threaded=False)
            return

        box = QMessageBox(self.window)
        box.setWindowTitle("Local AI (Ollama)")
        box.setIcon(QMessageBox.Question)
        get_btn = None
        if st == "not_installed":
            box.setText("The local AI that reads your documents (Ollama) isn't installed.\n\n"
                        "It's free and runs entirely on your machine. Install it, download the model, "
                        "and every PDF is auto-classified. Or continue now using filenames only.")
            get_btn = box.addButton("Download Ollama…", QMessageBox.AcceptRole)
        elif st == "no_model":
            box.setText(f"Ollama is running, but the model '{classify.DEFAULT_MODEL}' (~2 GB) isn't "
                        "downloaded yet.\n\nDownload it now (one-time), or continue using filenames only.")
            get_btn = box.addButton("Download model (~2 GB)", QMessageBox.AcceptRole)
        else:
            box.setText("Ollama is installed but couldn't be started automatically.\n\n"
                        "Please start Ollama, then try again — or continue using filenames only.")
        fb_btn = box.addButton("Use filenames", QMessageBox.DestructiveRole)
        box.addButton(QMessageBox.Cancel)
        box.exec()

        clicked = box.clickedButton()
        if clicked is fb_btn:
            # This branch is reached only when Ollama is unusable, so there is no
            # vision model to consult either — say so rather than leaving it to config.
            self.start_process(self.worker.run_rebrand_analyze, (source, None, False),
                               multi_threaded=False)
        elif get_btn is not None and clicked is get_btn:
            if st == "not_installed":
                import webbrowser
                webbrowser.open(classify.OLLAMA_DOWNLOAD_URL)
                QMessageBox.information(self.window, "Ollama",
                    "After installing Ollama, come back and click Analyze again.")
            else:  # no_model → download it, then the user re-runs Analyze
                self.start_process(self.worker.run_pull_model, (classify.DEFAULT_MODEL,), multi_threaded=False)

    def launch_pipeline(self):
        from docrefine.config import CFG
        d = PipelineDialog(self.window, default_kit=CFG.get("last_brand_kit"))
        if d.exec():
            if d.do_rebrand and d.kit_path:
                CFG.set("last_brand_kit", d.kit_path)
            CFG.set("rebrand_complete_set", d.complete_set)
            CFG.set("rebrand_show_attribution", d.show_attribution)
            CFG.set("rebrand_keep_original_names", d.keep_original_names)
            self._remember_stamps(d.stamp_opts)
            self.start_process(self.worker.run_pipeline,
                               (d.source_path, d.do_flatten, d.do_rebrand, d.do_ocr, d.kit_path,
                                300, None, d.complete_set, d.show_attribution,
                                d.keep_original_names, d.stamp_opts),
                               multi_threaded=True)

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