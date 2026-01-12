# DocRefine Pro - Changelog

## v117
* [Fixed] **Ingest Crash:** Resolved a critical regression where starting a new job caused a `TypeError` due to mismatched arguments.
* [Fixed] **Debug Export:** Fixed a threading violation that caused the "Export Debug Bundle" button to freeze the application without producing a file.
* [Refactor] **Thread Safety:** Hardened the separation between UI logic and background workers to prevent "Ghost" locks.

## v116 (Current)
* [Added] **In-App Documentation:** Added a "Documentation" section to Settings. Users can now view the **Changelog** and **User Guide** directly within the application.
* [Added] **Intelligent Resource Loader:** The app now intelligently locates documentation files whether running from the Python source, a compiled bundle, or an external folder.
* [Improved] **Build Pipeline:** Updated build configuration to bundle `.md` documentation files directly inside the executable/app bundle for offline access.

## v115
* [Fixed] Mac UI freeze issue by implementing thread-safe queue throttling.
* [Fixed] "Ghost Windows" on Windows 11 by patching subprocess.STARTUPINFO.
* [Added] New "Permanent Channel" deployment protocol.