from PySide6.QtCore import QObject, Signal, Slot
from docrefine.core.events import AppEvent, EventType

class DocRefineAdapter(QObject):
    """
    The Diplomat.
    Translates 'Worker Events' (Data) -> 'Qt Signals' (UI Actions).
    """
    
    # Define strongly-typed Qt Signals matching our Event Types
    sig_log = Signal(str, str)              # msg, level
    sig_progress_main = Signal(float, str)  # percent, text
    sig_slot_update = Signal(dict)          # {tid, text, percent}
    sig_worker_config = Signal(int)         # num_workers
    sig_status = Signal(str, str, str)      # stage, msg, color_hint
    sig_job_data = Signal(str)              # workspace_path
    sig_notification = Signal(dict)         # {title, msg, open_path}
    sig_error = Signal(str)
    sig_done = Signal()

    def __init__(self):
        super().__init__()

    def ingest_event(self, event: AppEvent):
        """
        The Callback. This runs on the WORKER THREAD.
        Qt Signals are thread-safe, so emitting them here safely 
        transfers execution to the Main Thread.
        """
        p = event.payload
        
        if event.type == EventType.LOG:
            self.sig_log.emit(p['msg'], p['level'])
            
        elif event.type == EventType.PROGRESS_MAIN:
            self.sig_progress_main.emit(p['percent'], p['text'])
        
        elif event.type == EventType.SLOT_UPDATE:
            self.sig_slot_update.emit(p)
            
        elif event.type == EventType.WORKER_CONFIG:
            self.sig_worker_config.emit(p)
            
        elif event.type == EventType.STATUS_CHANGE:
            self.sig_status.emit(p['stage'], p['msg'], p['color'])
            
        elif event.type == EventType.JOB_DATA:
            self.sig_job_data.emit(p)
            
        elif event.type == EventType.NOTIFICATION:
            self.sig_notification.emit(p)
            
        elif event.type == EventType.ERROR:
            self.sig_error.emit(str(p))
            
        elif event.type == EventType.DONE:
            self.sig_done.emit()