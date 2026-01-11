DocRefine Pro: Project Manifesto & Architecture Guidelines
Current Version: v112 Status: Internal Beta / Production Pilot

1. Core Philosophy
Stability is King: We prefer a "boring," stable solution over a "clever," risky one. If a feature causes a 1% chance of a crash, we cut the feature.

The "Defensive" Mindset: We assume the OS is hostile. Files will be locked, permissions will be denied, and disks will be full. Every file operation must be wrapped in try/except.

No Internet Runtime: The core processing (Ingest, Refine, Export) must function 100% offline. Internet is only allowed for the "Check Updates" button.

Monolithic Simplicity: For now, we maintain a single file (main.py) to simplify PyInstaller compilation. We do not refactor into modules unless the file exceeds 3,000 lines.

2. Platform "Blood-Written" Rules
These rules exist because we broke them before. Do not violate them.

Windows Specifics
Ghost Windows: Every subprocess.Popen call (e.g., for Poppler/Tesseract) MUST use the STARTUPINFO patch to hide the black command prompt window.

File Locking: Never attempt to Zip or Move a log file that the application is currently writing to. Use the "Copy-First" strategy (binary read/write) to bypass Windows file locks.

Path Lengths: Always truncate filenames in the UI (e.g., Sync View labels) to prevent the window from stretching infinitely on deep folder structures.

macOS Specifics
The "Aqua" Constraint: Do NOT use bg= (background colors) on Buttons or Progress Bars. macOS renders them as invisible or illegible. Use standard system buttons.

Dock Safety: When calculating window geometry, always subtract 120px from the screen height to avoid hiding the status bar behind the Dock.

Magic Mouse Sensitivity: The MouseWheel event on Mac fires hundreds of times per second. Any scrolling logic (like changing pages in Sync View) MUST have a "Debounce/Cooldown" timer (min 0.4s).

Permissions: We cannot assume write access to ~/Documents. Always have a fallback to ~/Downloads or /tmp/ for exports.

3. UI & Threading Architecture
The Golden Rule: NEVER touch the UI from a background thread.

Wrong: worker_thread calls label.config(text="Done").

Right: worker_thread puts ("status", "Done") into self.q. The Main Thread's poll() loop reads the queue and updates the label.

Feedback Loops:

Any button that triggers a process > 1 second must immediately disable itself and change text (e.g., "Exporting...").

The UI must never freeze. Use threading.Thread for File IO, Zipping, and Processing.

4. Workflows & State Management
Ingest: Must support "Standard" (Smart Hash), "Lightning" (Binary Hash), and "Deep" modes.

Refine: This is a destructive process. We must always allow the user to Stop.

Stop Rule: If stop_sig is True, we must abort cleanly and NOT trigger the "Success" popup. We must ensure the stop_sig is reset to False before starting the next job.

Export:

Debug Bundle: Our primary support tool. It captures app_debug.log, app_events.jsonl, config.json, and the active workspace's session_log.txt.

5. Deployment Strategy (Hybrid)
Code Storage: Private GitHub Repository (Source of Truth).

Update Trigger: Public GitHub Gist (docrefine_version.json).

Distribution: Google Drive / Shared Link (Hosted Zip containing Executable + PDF Readme).

Version History: We never delete old releases from the repo tags. They are our safety net.

6. Technical Debt & Known Constraints
Code Size: The run_batch function in main.py is becoming dense. If we add new processors (e.g., Audio/Video), we must refactor this into a dedicated class structure.

Input Peripherals: Sync View scrolling is debounced (0.4s) for Mac Magic Mice. This might feel slightly sluggish on high-performance Windows gaming mice, but stability is preferred over speed.

Poppler Noise: We currently suppress standard error output from Poppler to keep logs clean; deeper PDF corruption errors might require temporarily enabling verbose logging for diagnosis.

PyInstaller Flags: Compilation requires specific hidden imports for pdf2image to locate the Poppler binaries correctly on Windows. Do not remove these flags from the build script.