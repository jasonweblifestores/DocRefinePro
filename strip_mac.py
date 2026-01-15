import shutil
from pathlib import Path

# Targeted bloat frameworks to delete
BLOAT = [
    "QtDesigner", "QtNetwork", "QtDBus", "QtQml", "QtQuick", 
    "QtVirtualKeyboard", "QtWebEngineCore", "QtWebEngineWidgets",
    "Qt3DCore", "Qt3DRender", "QtCharts", "QtSensors", "QtMultimedia"
]

app_path = Path("dist/DocRefinePro.app/Contents/Frameworks")

if app_path.exists():
    for folder in app_path.iterdir():
        # Check if this folder is in our hit list
        if any(b in folder.name for b in BLOAT):
            print(f"CRITICAL SLIMMING: Removing {folder.name}")
            shutil.rmtree(folder)
else:
    print(f"ERROR: Frameworks path not found at {app_path}")