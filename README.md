# DocRefine Pro v119

**Enterprise-Grade Document Processing & Organization Tool**

**DocRefine Pro** is a standalone desktop application for batch processing document workflows (Ingestion, Deduplication, Flattening, OCR). It runs 100% locally on your machine—no cloud uploads.

**v119 Update (The Great Refactor):**
* **New Engine:** Migrated to PySide6 (Qt) for improved stability and Dark Mode support.
* **Forensic Viewer 2.0:** Side-by-side comparison with synchronized zoom and panning.
* **Multi-Threading:** Real-time visualization of active worker threads.
* **Controls:** Pause/Resume support for long-running batch jobs.

---

## 📥 Installation Instructions

### 🪟 Windows
1.  Download `DocRefinePro_Win_v119.zip`.
2.  Right-click the zip file -> **Extract All**.
3.  Open the extracted folder.
4.  Double-click **DocRefine Pro.exe**.
    * *Note: If Windows SmartScreen appears, click "More Info" -> "Run Anyway".*

### 🍎 macOS
1.  Download `DocRefinePro_Mac_v119.dmg`.
2.  Double-click the `.dmg` file to mount it.
3.  **Drag the DocRefine Pro app** into your **Applications** folder.
    * *Note: The application size is larger (~230MB) due to the inclusion of the complete Qt6 Framework for native performance.*

#### ⚠️ Critical: "App is Damaged" Fix
Because this is an internal tool not signed by the Apple Store, macOS will likely block it with a message saying *"The app is damaged"* or *"Cannot be opened."*

**To fix this (One-time setup):**
1.  Open your Mac's **Terminal** app (Command+Space, type "Terminal").
2.  Paste the following command and hit Enter:
    ```bash
    xattr -cr /Applications/DocRefinePro.app
    ```
3.  You can now open the app normally from your Applications folder.

---

## 🚀 Quick Start Guide

### 1. Ingest
* Click **+ New Ingest Job**.
* Select your source folder containing raw documents.
* **Standard Mode:** Best for most PDFs.
* **Lightning Mode:** Fastest (Exact duplicate detection only).

### 2. Process
* **Refine Tab:** Flatten, OCR, or Sanitize files.
* **Pause/Resume:** You can now pause processing to free up system resources without cancelling the job.
* **Forensic Viewer:** Go to the Inspector tab, right-click a duplicate, and select "Compare Duplicates" to visually verify files.

### 3. Output
* **Option A (Unique Masters):** Export a clean folder containing one copy of every unique file.
* **Option B (Reconstruction):** Re-create the original folder structure using the optimized master files.

### 4. Support
For bugs or feature requests, contact the development team directly (Jason Diaz - Task Specialist : jason@weblifestores.com).
