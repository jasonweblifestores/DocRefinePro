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

# ==============================================================================
#   WINDOWS GHOST WINDOW FIX (Global Patch)
# ==============================================================================
if os.name == 'nt':
    try:
        _original_popen = subprocess.Popen
        def safe_popen(*args, **kwargs):
            if 'startupinfo' not in kwargs:
                si = subprocess.STARTUPINFO()
                si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                si.wShowWindow = subprocess.SW_HIDE
                kwargs['startupinfo'] = si
                if 'creationflags' not in kwargs:
                    kwargs['creationflags'] = 0x08000000 
            return _original_popen(*args, **kwargs)
        subprocess.Popen = safe_popen
    except Exception as e: print(f"Warning: Could not patch subprocess: {e}")

# ==============================================================================
#   SYSTEM UTILITIES
# ==============================================================================
class SystemUtils:
    IS_WIN = platform.system() == 'Windows'
    IS_MAC = platform.system() == 'Darwin'
    CURRENT_VERSION = "v128.2"
    UPDATE_MANIFEST_URL = "https://gist.githubusercontent.com/jasonweblifestores/53752cda3c39550673fc5dafb96c4bed/raw/docrefine_version.json"

    @staticmethod
    def get_resource_dir():
        if getattr(sys, 'frozen', False): return Path(sys._MEIPASS)
        return Path(__file__).parent.parent # Adjusted for package structure

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
        # 1. Look in bundled resources (PyInstaller)
        res = SystemUtils.get_resource_dir() / filename
        if res.exists(): return res
        
        # 2. Look next to executable (Portable/Distribution)
        if getattr(sys, 'frozen', False):
            exe_path = Path(sys.executable).parent / filename
            if exe_path.exists(): return exe_path
        
        # 3. Look in Project Root (Dev Mode - One level up from this file)
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

# ==============================================================================
#   CONFIGURATION
# ==============================================================================
class Config:
    GITHUB_REPO = "jasonweblifestores/DocRefinePro" 
    DEFAULTS = { 
        "ram_warning_mb": 1024, 
        "resize_width": 1920, 
        "log_level": "INFO",
        "max_pixels": 500000000,
        "max_threads": 0, 
        "default_export_prio": "Auto (Best Available)",
        "default_ingest_mode": "Standard", 
        "ocr_lang": "eng",
        "last_workspace": "",
        "last_geometry": "1024x700",
        "last_tab": 0
    }
    
    def __init__(self):
        self.data = self.DEFAULTS.copy()
        self.path = SystemUtils.get_user_data_dir() / "config.json"
        if self.path.exists():
            try:
                with open(self.path, 'r') as f: self.data.update(json.load(f))
            except: pass

    def get(self, key): return self.data.get(key, self.DEFAULTS.get(key))
    def set(self, key, val): self.data[key] = val; self.save()
    def reset(self): self.data = self.DEFAULTS.copy(); self.save()
    def save(self):
        try:
            with open(self.path, 'w') as f: json.dump(self.data, f, indent=4)
        except Exception as e: print(f"Config Save Error: {e}")

# Global Config Instance
CFG = Config()

# ==============================================================================
#   LOGGING
# ==============================================================================
USER_DIR = SystemUtils.get_user_data_dir()
LOG_PATH = USER_DIR / "app_debug.log"
JSON_LOG_PATH = USER_DIR / "app_events.jsonl"
WORKSPACES_ROOT = USER_DIR / "Workspaces"
WORKSPACES_ROOT.mkdir(parents=True, exist_ok=True)

logger = logging.getLogger("DocRefine")
logger.setLevel(getattr(logging, CFG.get("log_level").upper(), logging.INFO))
c_handler = logging.StreamHandler()
c_handler.setFormatter(logging.Formatter('[%(levelname)s] %(message)s'))
logger.addHandler(c_handler)

try:
    f_handler = RotatingFileHandler(LOG_PATH, maxBytes=1024*1024, backupCount=5, encoding='utf-8')
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