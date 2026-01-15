import sys
from docrefine.config import log_app

if __name__ == "__main__":
    try:
        log_app("Booting DocRefine Pro (Qt/PySide6 Edition)...")
        # Import the new Qt App Runner
        from docrefine.gui.app_qt import run
        run()
    except Exception as e:
        print(f"Fatal Boot Error: {e}")