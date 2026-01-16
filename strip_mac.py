import os
import shutil
from pathlib import Path

# The Target: Your built macOS App Bundle
APP_PATH = Path("dist/DocRefinePro.app")
FRAMEWORKS_DIR = APP_PATH / "Contents" / "Frameworks"
PLUGINS_DIR = APP_PATH / "Contents" / "Resources" / "PySide6" / "plugins"

# The Hit List: Modules verified in logs as BLOAT
BLOAT_PATTERNS = [
    "QtWebEngine", "QtQuick", "QtQml", "Qt3D", 
    "QtVirtualKeyboard", "QtSerialBus", "QtSerialPort",
    "QtSensors", "QtCharts", "QtDataVisualization",
    "QtTest", "QtTextToSpeech", "QtDesigner", "QtHelp",
    "QtMultimedia", "QtLocation", "QtPositioning",
    "QtNetworkAuth", "QtScxml", "QtRemoteObjects",
    "QtStateMachine", "QtXml", "QtSql"
]

def get_size(path):
    total = 0
    if not path.exists(): return 0
    if path.is_file(): return path.stat().st_size
    for entry in os.scandir(path):
        if entry.is_file(): total += entry.stat().st_size
        elif entry.is_dir(): total += get_size(Path(entry.path))
    return total

def nuke_path(path):
    """Smart delete that handles symlinks and their targets."""
    if not path.exists(): return 0
    
    deleted_bytes = 0
    
    # If it's a symlink, resolve it first to find the payload
    if path.is_symlink():
        try:
            target = path.resolve()
            # Security check: Ensure target is inside our app bundle
            if str(APP_PATH) in str(target) and target.exists():
                print(f"   ↳ Following symlink to payload: {target.name}")
                if target.is_dir():
                    deleted_bytes += get_size(target)
                    shutil.rmtree(target)
                else:
                    deleted_bytes += target.stat().st_size
                    target.unlink()
            path.unlink() # Delete the link itself
        except Exception as e:
            print(f"   ⚠️ Error resolving symlink {path.name}: {e}")
            path.unlink()
    
    # Standard file/folder
    elif path.is_dir():
        deleted_bytes += get_size(path)
        shutil.rmtree(path)
    else:
        deleted_bytes += path.stat().st_size
        path.unlink()
        
    return deleted_bytes

def nuke_bloat():
    print(f"🚀 STARTING SURGICAL REMOVAL ON: {APP_PATH}")
    if not APP_PATH.exists():
        print(f"❌ CRITICAL: App bundle not found at {APP_PATH}"); return

    deleted_size = 0

    # 1. SCAN FRAMEWORKS
    if FRAMEWORKS_DIR.exists():
        print(f"🔍 Scanning Frameworks at {FRAMEWORKS_DIR}...")
        for item in FRAMEWORKS_DIR.iterdir():
            if any(pattern in item.name for pattern in BLOAT_PATTERNS):
                print(f"   💣 NUKE: {item.name}")
                deleted_size += nuke_path(item)

    # 2. SCAN PLUGINS
    if PLUGINS_DIR.exists():
        print(f"🔍 Scanning Plugins at {PLUGINS_DIR}...")
        for root, dirs, files in os.walk(PLUGINS_DIR):
            for d in dirs[:]: 
                if any(pattern in d for pattern in BLOAT_PATTERNS):
                    full_path = Path(root) / d
                    print(f"   💣 NUKE PLUGIN: {d}")
                    deleted_size += nuke_path(full_path)
                    dirs.remove(d)

    print("-" * 60)
    print(f"✅ CLEANUP COMPLETE. FREED: {deleted_size / 1024 / 1024:.2f} MB")
    print("-" * 60)

if __name__ == "__main__":
    nuke_bloat()