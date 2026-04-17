import sys
from docrefine.config import log_app

if __name__ == "__main__":
    if "--dry-run" in sys.argv:
        print("Dry run initialized.")
        log_app("Booting DocRefine Pro (Dry Run Mode)...")
        from docrefine.gui.app_qt import dry_run
        dry_run()
        print("Dry run OK.")
        sys.exit(0)

    try:
        log_app("Booting DocRefine Pro (Qt/PySide6 Edition)...")
        # Import the new Qt App Runner
        from docrefine.gui.app_qt import run
        run()
    except Exception as e:
        print(f"Fatal Boot Error: {e}")