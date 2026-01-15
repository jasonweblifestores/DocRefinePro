# SAVE AS: docrefine/core/events.py
from dataclasses import dataclass
from enum import Enum, auto
from typing import Any, Optional

class EventType(Enum):
    LOG = auto()
    PROGRESS_MAIN = auto()
    SLOT_UPDATE = auto()    # Individual thread update
    WORKER_CONFIG = auto()  # Setup thread slots
    STATUS_CHANGE = auto()
    JOB_DATA = auto()       # Refresh job list
    NOTIFICATION = auto()   # Popups/Dialogs
    ERROR = auto()
    DONE = auto()

@dataclass
class AppEvent:
    type: EventType
    payload: Any = None

    @staticmethod
    def log(msg: str, level: str = "INFO"):
        return AppEvent(EventType.LOG, {"msg": msg, "level": level})

    @staticmethod
    def progress(percent: float, text: str = ""):
        return AppEvent(EventType.PROGRESS_MAIN, {"percent": percent, "text": text})

    @staticmethod
    def status(stage: str, message: str, color_hint: str = "blue"):
        return AppEvent(EventType.STATUS_CHANGE, {"stage": stage, "msg": message, "color": color_hint})