# DocRefine Pro

**Enterprise-Grade Document Processing & Organization Tool**

**DocRefine Pro** is a standalone desktop application for batch processing document workflows (Ingestion, Deduplication, Flattening, OCR). It runs 100% locally on your machine—no cloud uploads.

---

## 📥 Installation Instructions

### 🪟 Windows
1.  You should have received a file named `DocRefinePro_Win_v113.zip`.
2.  Right-click the zip file -> **Extract All**.
3.  Open the extracted folder.
4.  Double-click **DocRefine Pro.exe**.
    * *Note: If Windows SmartScreen appears, click "More Info" -> "Run Anyway".*

### 🍎 macOS
1.  Download `DocRefinePro_Mac_vXX.dmg`.
2.  Double-click the `.dmg` file to mount it.
3.  **Drag the DocRefine Pro app** into your **Applications** folder.

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
* **Option A (Modify):** Use the "Refine" tab to Flatten, OCR, or Sanitize files.
* **Option B (Organize):** Use "Export > Option A" to extract unique master files without modifying them.

### 3. Support
For bugs or feature requests, contact the development team directly.
