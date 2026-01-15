import shutil
import os
from pathlib import Path

# Targeted bloat frameworks to delete
BLOAT = [
    "QtDesigner", "QtNetwork", "QtDBus", "QtQml", "QtQuick", 
    "QtVirtualKeyboard", "QtWebEngine", "Qt3D", "QtCharts", 
    "QtSensors", "QtMultimedia", "QtTest", "QtTextToSpeech",
    "QtSql", "QtStateMachine"
]

app_path = Path("dist/DocRefinePro.app/Contents/Frameworks")

print(f"--- STARTING INDUSTRIAL STRIPPING AT {app_path} ---")

if app_path.exists():
    for folder in app_path.iterdir():
        # Check if this folder is in our hit list
        if any(b in folder.name for b in BLOAT):
            print(f"CRITICAL SLIMMING: Targeting {folder.name}")
            try:
                # FIX: Check if it's a symlink FIRST
                if folder.is_symlink() or os.path.islink(folder):
                    print(f"   -> Unlinking symlink: {folder.name}")
                    folder.unlink()
                else:
                    print(f"   -> Deleting directory: {folder.name}")
                    shutil.rmtree(folder)
            except Exception as e:
                print(f"   !! FAILED to remove {folder.name}: {e}")
else:
    print(f"ERROR: Frameworks path not found at {app_path}")
    
print("--- STRIPPING COMPLETE ---")