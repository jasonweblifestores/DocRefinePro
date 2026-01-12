# DocRefine Pro - Changelog

## v116 (Current)
* [Added] **In-App Documentation:** Added a "Documentation" section to Settings. Users can now view the **Changelog** and **User Guide** directly within the application.
* [Added] **Intelligent Resource Loader:** The app now intelligently locates documentation files whether running from the Python source, a compiled bundle, or an external folder.
* [Improved] **Build Pipeline:** Updated build configuration to bundle `.md` documentation files directly inside the executable/app bundle for offline access.

## v115
* [Fixed] Mac UI freeze issue by implementing thread-safe queue throttling.
* [Fixed] "Ghost Windows" on Windows 11 by patching subprocess.STARTUPINFO.
* [Added] New "Permanent Channel" deployment protocol.