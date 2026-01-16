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
    if not path.exists() and not path.is_symlink(): return 0
    deleted = 0
    try:
        # If it's a directory, get size and remove tree
        if path.is_dir() and not path.is_symlink():
            deleted += get_size(path)
            shutil.rmtree(path)
        # If it's a file or symlink, just unlink
        else:
            if not path.is_symlink():
                deleted += path.stat().st_size
            path.unlink()
    except Exception as e:
        print(f"Error nuking {path}: {e}")
    return deleted

def cleanup_broken_symlinks():
    print("🧹 SCANNING FOR BROKEN SYMLINKS...")
    broken_count = 0
    for root, dirs, files in os.walk(APP_PATH):
        for filename in files:
            p = Path(root) / filename
            if p.is_symlink():
                if not p.exists(): # Checks if target exists
                    print(f"   ✂️ Removing broken link: {p.name}")
                    p.unlink()
                    broken_count += 1
    print(f"   Fixed {broken_count} broken links.")

def nuke_bloat():
    print(f"🚀 STARTING SURGICAL REMOVAL ON: {APP_PATH}")
    if not APP_PATH.exists():
        print(f"❌ CRITICAL: App bundle not found at {APP_PATH}"); return

    deleted_size = 0

    # 1. SCAN FRAMEWORKS
    if FRAMEWORKS_DIR.exists():
        for item in FRAMEWORKS_DIR.iterdir():
            if any(pattern in item.name for pattern in BLOAT_PATTERNS):
                print(f"   💣 NUKE FRAMEWORK: {item.name}")
                deleted_size += nuke_path(item)

    # 2. SCAN PLUGINS
    if PLUGINS_DIR.exists():
        for root, dirs, files in os.walk(PLUGINS_DIR):
            for d in dirs[:]: 
                if any(pattern in d for pattern in BLOAT_PATTERNS):
                    full_path = Path(root) / d
                    print(f"   💣 NUKE PLUGIN: {d}")
                    deleted_size += nuke_path(full_path)
                    dirs.remove(d)
                    
    # 3. GARBAGE COLLECTION (Crucial Step)
    cleanup_broken_symlinks()

    print("-" * 60)
    print(f"✅ CLEANUP COMPLETE. REMOVED: {deleted_size / 1024 / 1024:.2f} MB")
    print("-" * 60)

if __name__ == "__main__":
    nuke_bloat()