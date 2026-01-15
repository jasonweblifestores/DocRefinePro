# DocRefine Pro - Changelog

## [v125] - 2026-01-15
### Infrastructure
- **Mac Optimization:** Implemented manual framework stripping in CI/CD to reduce DMG size.
- **Cleanup:** Removed unused Qt translations and debug symbols from the macOS binary.

## [v124] - 2026-01-15
### Fixed
- **Build Fix**: Resolved `TypeError` in Spec file caused by deprecated `include_pycache` argument in PyInstaller 6.18.

## [v123] - 2026-01-15
### Infrastructure
- **Size Optimization:** Reduced DMG/EXE footprint by ~50% via targeted PySide6 stripping.
- **Compression:** Integrated UPX compression and enabled binary symbol stripping.
- **Architecture:** Moved from `collect_all` to manual dependency mapping to prevent "Universal" binary bloat.

## [v122] - 2026-01-15
### Infrastructure
- **Build Fix:** Added `BUNDLE` block to Spec file for correct macOS `.app` generation.
- **Asset Safety:** Added fallback logic for missing application icons during build.

## [v121] - 2026-01-15
### Maintenance
- **Release:** Fixed Git tag synchronization for CI/CD pipeline.

## [v120] - 2026-01-15
### Infrastructure
- **Spec-Based Build:** Switched from CLI commands to `DocRefinePro.spec` for release builds.
- **Cross-Platform Fix:** Added `collect_all('PySide6')` to the build spec to ensure Mac/Win DLLs are bundled correctly.

## v119 (The Great Refactor)
* [Architecture] **UI Migration:** Complete rewrite of the UI layer from Tkinter to **PySide6 (Qt)**.
    * Modern "Fusion" theme with Dark Mode support.
    * Non-blocking, thread-safe architecture using Signals & Slots.
* [New] **Forensic Comparator 2.0:**
    * Synchronized Zoom & Pan.
    * Dark background for high-contrast inspection.
    * Smooth page scrolling.
* [New] **Active Worker Visualizer:** Real-time grid showing multi-threaded status.
* [Improved] **Timer Logic:** Job timer now respects "Pause" state.
* [Fixed] **Windows Explorer:** Fixed "Reveal in Folder" failing due to subprocess restrictions.

## v118-patch1
* [Fixed] **Context Menu Bug:** Fixed "Reveal in Folder" failure by using unique ID lookup instead of fragile filename matching.

## v118 (Modular Foundation)
* [Architecture] **Modular Restructure:** Application logic split into `gui`, `worker`, `processing`, and `config` modules for improved stability.
* [Improvement] **Update Signal:** Hardcoded verification of v118.
* [Docs] Updated bundled documentation paths.

## v117 (Hotfixes)
* [Fixed] **Ingest Crash:** Resolved regression in `run_inventory` arguments.
* [Fixed] **Debug Export:** Fixed threading violation in export tool.
* [Refactor] **Thread Safety:** Hardened UI/Worker separation.

## v116
* [Added] In-App Documentation Viewer.
* [Improved] Intelligent Resource Loader.

## v115
* [Fixed] Mac UI freeze (queue throttling).
* [Fixed] Windows "Ghost Windows" patch.