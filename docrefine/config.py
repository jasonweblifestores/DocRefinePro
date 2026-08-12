# SAVE AS: docrefine/config.py
import sys
import os
import subprocess
import json
import logging
import platform
import shutil
import time
from pathlib import Path
from logging.handlers import RotatingFileHandler
from datetime import datetime
from pydantic import BaseModel

def get_hidden_startupinfo():
    if os.name == 'nt':
        si = subprocess.STARTUPINFO()
        si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        si.wShowWindow = subprocess.SW_HIDE
        return si
    return None

class Constants:
    DIR_MASTER = "01_Master_Files"
    DIR_READY = "02_Ready_For_Redistribution"
    DIR_ORGANIZED = "03_Organized_Output"
    DIR_REPORTS = "04_Reports"
    DIR_QUARANTINE = "00_Quarantine"

class SystemUtils:
    IS_WIN = platform.system() == 'Windows'
    IS_MAC = platform.system() == 'Darwin'
    # ---------------------------------------------------------
    # VERSION SYNC: Keep in step with CHANGELOG.md
    # ---------------------------------------------------------
    CURRENT_VERSION = "v156"
    UPDATE_MANIFEST_URL = "https://gist.githubusercontent.com/jasonweblifestores/53752cda3c39550673fc5dafb96c4bed/raw/docrefine_version.json"

    @staticmethod
    def get_resource_dir():
        if getattr(sys, 'frozen', False): return Path(sys._MEIPASS)
        return Path(__file__).parent.parent

    @staticmethod
    def get_user_data_dir():
        if SystemUtils.IS_MAC or SystemUtils.IS_WIN:
            p = Path.home() / "Documents" / "DocRefinePro_Data"
            p.mkdir(parents=True, exist_ok=True)
            return p
        if getattr(sys, 'frozen', False): return Path(sys.executable).parent
        return Path(__file__).parent.parent

    @staticmethod
    def find_doc_file(filename):
        res = SystemUtils.get_resource_dir() / filename
        if res.exists(): return res
        if getattr(sys, 'frozen', False):
            exe_path = Path(sys.executable).parent / filename
            if exe_path.exists(): return exe_path
        cwd = Path(__file__).parent.parent / filename
        if cwd.exists(): return cwd
        return None

    @staticmethod
    def open_file(path):
        p = str(path)
        try:
            if not Path(p).exists(): return
            if SystemUtils.IS_WIN: os.startfile(p)
            elif SystemUtils.IS_MAC: subprocess.call(['open', p])
            else: subprocess.call(['xdg-open', p])
        except Exception as e: print(f"Error opening file: {e}")

    @staticmethod
    def reveal_file(path):
        p = str(Path(path).resolve())
        try:
            if not Path(p).exists(): return
            if SystemUtils.IS_WIN:
                subprocess.Popen(f'explorer /select,"{p}"')
            elif SystemUtils.IS_MAC:
                subprocess.Popen(["open", "-R", p])
            else:
                subprocess.call(['xdg-open', str(Path(p).parent)])
        except Exception as e:
            print(f"Error revealing file: {e}")
            SystemUtils.open_file(Path(p).parent)

    @staticmethod
    def find_binary(bin_name):
        res_dir = SystemUtils.get_resource_dir()
        if (res_dir / bin_name).exists(): return str(res_dir / bin_name)
        if (res_dir / "bin" / bin_name).exists(): return str(res_dir / "bin" / bin_name)

        # Bundled engines committed to the repo (Windows builds ship these folders).
        for bundled in (res_dir / "Tesseract-OCR", res_dir / "poppler" / "Library" / "bin"):
            if (bundled / bin_name).exists(): return str(bundled / bin_name)

        portable_target = res_dir / "DocRefine_Portable"
        if portable_target.exists():
             if (portable_target / bin_name).exists(): return str(portable_target)
             if (portable_target / "bin" / bin_name).exists(): return str(portable_target / "bin")

        sys_path = shutil.which(bin_name)
        if sys_path: return str(Path(sys_path).resolve())

        if SystemUtils.IS_MAC:
            for loc in ["/opt/homebrew/bin", "/usr/local/bin"]:
                brew_path = Path(loc) / bin_name
                if brew_path.exists(): return str(brew_path)
        return None

class ConfigData(BaseModel):
    ram_warning_mb: int = 1024
    resize_width: int = 1920
    log_level: str = "INFO"
    max_pixels: int = 500000000
    max_threads: int = 0
    default_export_prio: str = "Auto (Best Available)"
    default_ingest_mode: str = "Standard"
    ocr_lang: str = "eng"
    last_workspace: str = ""
    last_brand_kit: str = ""
    last_rebrand_source: str = ""
    rebrand_complete_set: bool = True
    # Batch 1 and 2 covers carry the title alone. The task brief also asks for a
    # "Manufactured by [X] | Sold by ..." line; this decides which wins.
    rebrand_show_attribution: bool = False
    # Batch 1/2 kept every original filename; the Batch 4 brief asks for
    # <product>-<asset-type>-budget-mailboxes.pdf instead. This picks which.
    rebrand_keep_original_names: bool = False
    # Per-page stamps the rebranding SOP asks for. All off by default, so the
    # output matches the signed-off Batch 1/2 sets until someone decides
    # otherwise; the wording itself lives in the brand kit's brand.json.
    rebrand_footer_attribution: bool = False   # attribution in the page footer, per SOP
    rebrand_stamp_tagline: bool = False
    rebrand_stamp_version: bool = False        # "Version 1.0 · Last Updated [Month Year]"
    rebrand_stamp_disclaimer: bool = False
    # Look at the page with a local vision model wherever the text pass is blind
    # (no extractable text) or hedging. Slower — seconds per file instead of
    # milliseconds — so it is opt-in, and it needs the model downloaded.
    rebrand_vision_pass: bool = False
    rebrand_vision_model: str = "qwen2.5vl:7b"
    # Sleep the machine once a long run finishes on its own. Per-run in the
    # dialog, remembered here. Never applied to a run that was stopped or failed.
    sleep_when_done: bool = False
    sleep_when_done_action: str = "sleep"      # "sleep" or "hibernate"
    last_geometry: str = "1024x700"
    last_tab: int = 0

class Config:
    GITHUB_REPO = "jasonweblifestores/DocRefinePro" 
    
    def __init__(self):
        self.path = SystemUtils.get_user_data_dir() / "config.json"
        
        # Load or create defaults using Pydantic
        if self.path.exists():
            try:
                with open(self.path, 'r') as f:
                    raw_data = json.load(f)
                self._data = ConfigData(**raw_data)
            except Exception as e:
                print(f"Config validation error: {e}. Falling back to defaults.")
                self._data = ConfigData()
        else:
            self._data = ConfigData()

    def get(self, key):
        return getattr(self._data, key)

    def set(self, key, val):
        setattr(self._data, key, val)
        self.save()

    def reset(self):
        self._data = ConfigData()
        self.save()

    def save(self):
        try:
            with open(self.path, 'w') as f:
                f.write(self._data.model_dump_json(indent=4))
        except Exception as e: print(f"Config Save Error: {e}")

CFG = Config()

USER_DIR = SystemUtils.get_user_data_dir()
LOG_PATH = USER_DIR / "app_debug.log"
JSON_LOG_PATH = USER_DIR / "app_events.jsonl"
WORKSPACES_ROOT = USER_DIR / "Workspaces"
WORKSPACES_ROOT.mkdir(parents=True, exist_ok=True)
# One fixed home for rebrand review sheets, so they never get lost inside a
# source tree (and never get swept into a "complete set" upload). See reviews.py.
REVIEWS_ROOT = USER_DIR / "Rebrand Reviews"
REVIEWS_ROOT.mkdir(parents=True, exist_ok=True)

logger = logging.getLogger("DocRefine")
logger.setLevel(getattr(logging, CFG.get("log_level").upper(), logging.INFO))
c_handler = logging.StreamHandler()
c_handler.setFormatter(logging.Formatter('[%(levelname)s] %(message)s'))
logger.addHandler(c_handler)

try:
    f_handler = RotatingFileHandler(LOG_PATH, maxBytes=1024*1024, backupCount=5, encoding='utf-8', mode='a')
    f_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
    logger.addHandler(f_handler)
except: pass

def log_app(msg, level="INFO", structured_data=None):
    if level == "ERROR": logger.error(msg)
    elif level == "WARN": logger.warning(msg)
    else: logger.info(msg)
    try:
        entry = {
            "timestamp": datetime.now().isoformat(),
            "level": level,
            "message": msg,
            "os": platform.system(),
            "version": SystemUtils.CURRENT_VERSION
        }
        if structured_data: entry.update(structured_data)
        with open(JSON_LOG_PATH, "a", encoding="utf-8") as f:
            f.write(json.dumps(entry) + "\n")
    except: pass
