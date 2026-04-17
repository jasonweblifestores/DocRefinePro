# DocRefine Pro - Changelog

## [v133.1] - 2026-04-17
### Hotfix
* **CI/CD Fix:** Resolved `ModuleNotFoundError` during build phase by syncing the GitHub pipeline installer process with `requirements.txt`.

## [v133] - 2026-04-17
### Infrastructure & CI/CD
* **Zero-Touch Deployments:** Fully automated GitHub Releases pipeline, permanently replacing the manual Google Drive deployment process.
* **Mac App Optimization:** Successfully removed the dangerous post-build Python stripping script. PyInstaller now natively filters out `QtWebEngine`, `QtQuick`, and `Qt3D` at the binary analysis level, retaining the massive file size reduction without corrupting Apple code signatures.
* **Mac Smoke Testing:** Programmed a `--dry-run` flag into the core engine to enable fully automated, headless sanity verification. GitHub Actions now tests the compiled MacOS binary to ensure it boots without crashing before releasing it.
* **Architectural Fixes:** Re-wrote Python internal configuration and validation using `pydantic`. Integrated `Jinja2` templating for all HTML reporting engines. Eliminated closure-loop memory leaks in PySide6 rendering.

## [v132] - 2026-03-12
### Added
* **Chained Workflow Bypass:** Added a new UI toggle in the Refine tab (`Source from Flattened Cache`) allowing users to execute sequential refinement actions (e.g. OCR) directly on files residing in the output cache rather than forcing a pull from the strict master hub. This fully supports intermediate manual editing steps like third-party visual rebranding.

## [v131] - 2026-03-12
### Fixed
* **Reconstruction Safety:** Re-wrote `get_best_source` in the Worker module to support intelligent ID prefix matching for processed files. The engine now ignores minor filename variations or unexpected double extensions (e.g. `.pdf.pdf`) applied by external rebranding tools.
* **Crash Protection:** Added strict validation checks to prevent `[Errno 2]` fatal failures if a fallback master file has been manually moved or deleted from the workspace.

## [v130] - 2026-03-12
### Fixed
* **Reconstruction Pipeline:** Wired up the "Override Source" checkbox to properly prompt for an external directory. This allows users to seamlessly inject externally modified files (like rebranded PDFs) into the redistribution engine, ignoring exact file extension mismatches as long as the base Unique ID matches.

## [v129] - 2026-01-19
### Maintenance
* **Legacy Cleanup:** Permanently removed the deprecated Tkinter UI module (`docrefine/gui/app.py`).
* **Architecture:** The application is now exclusively Qt/PySide6.

## [v128.5] - 2026-01-16
### UX & Stability
* **Inspector First:** Changed default tab order to prioritize Forensic Inspection.
* **Progress Fix:** Resolved issue where progress bar would hang at 99%.
* **Pause Safety:** Fixed critical bug where pausing the refinement process caused files to be copied without processing.
* **Smart Skip:** The engine now detects existing output files and skips them to save time.
* **Log Hygiene:** App logs now reset on startup to prevent bloat.
* **Timing:** Added detailed breakdown of Ingest vs. Refine times in the dashboard stats.

## [v128.4] - 2026-01-16
### Architecture & Optimization
* **Spec-First Build System:** Migrated from CLI overrides to a pure Python Spec file architecture for consistent cross-platform building.
* **Mac Diet (Surgical):** Reduced macOS app bundle size from **1.3GB to ~200MB** via:
    * **Pre-Build:** Filtering `hiddenimports` to prevent PyInstaller hooks from loading unwanted Qt frameworks (WebEngine, Quick, 3D).
    * **Post-Build:** "Nuclear" stripping script that physically removes any surviving bloat frameworks and cleans up broken symlinks to prevent installer crashes.
* **Stability:** Fixed Windows build crash caused by tuple unpacking errors in the Spec file.
* **Maintenance:** Added `tools/inventory.py` for project auditing.

## [v128] - 2026-01-16
### Architecture
* **Build System Overhaul:** Switched to a "Spec-First" build architecture.
* **Mac Diet:** Implemented aggressive binary filtering at the PyInstaller Spec level to block `QtWebEngine`, `QtQuick`, and `Qt3D` *before* bundling. This targets the 1.3GB bloat issue directly.
* **Inventory Control:** Added `tools/inventory.py` for project auditing.
* **Cleanup:** Removed deprecated CLI overrides from build scripts.

## [v127] - 2026-01-15
### Fixed
- **Mac Build:** Resolved `OSError` in stripping script by handling symlinks correctly.
- **Optimization:** Refined `strip_mac.py` to differentiate between directories (`rmtree`) and symbolic links (`unlink`) during framework cleanup.

## [v126] - 2026-01-15
### Infrastructure
- **Mac Optimization:** Implemented `strip_mac.py` to programmatically remove unused Qt Frameworks (`QtQuick`, `QtQml`, `QtWebEngine`) post-build.
- **CI/CD:** Replaced fragile bash commands with Python scripting for reliable path resolution during the build process.
- **Size Reduction:** Forced removal of PyInstaller-protected frameworks to reduce DMG size from ~1.2GB to target (~400MB).

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