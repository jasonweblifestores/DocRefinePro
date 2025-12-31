# DocRefine Pro

**Enterprise-Grade Document Processing & Organization Tool**

**DocRefine Pro** is a standalone desktop application built for batch processing heavy document workflows. It automates the ingestion, deduplication, flattening, and OCR (Optical Character Recognition) of PDF and image files without requiring expensive cloud subscriptions.

Works natively on **Windows 10/11** and **macOS (Intel & Apple Silicon)**.

---

## 🚀 Key Features

* **Smart Ingestion:** Automatically scans folders, detects duplicates using hashing (Binary or Smart Text), and quarantines corrupt files.
* **PDF Flattening:** Converts searchable PDFs into flattened images to remove hidden layers or metadata.
* **Batch OCR:** Bulk-converts scanned images or flat PDFs into searchable text documents using Tesseract 5.
* **Sanitization:** Removes metadata (Authors, Edit Time) from Microsoft Office documents (.docx, .xlsx).
* **Privacy First:** All processing happens **locally** on your machine. No data is ever uploaded to the cloud.

---

## 📥 Download & Installation

Go to the [Releases Page](https://github.com/jasonweblifestores/DocRefinePro/releases) to download the latest version.

### 🪟 Windows
1.  Download `DocRefinePro_Win_vXX.zip`.
2.  Right-click the zip file -> **Extract All**.
3.  Open the folder and run **DocRefine Pro.exe**.
    * *Note: If Windows SmartScreen appears, click "More Info" -> "Run Anyway".*

### 🍎 macOS
1.  Download `DocRefinePro_Mac_vXX.dmg`.
2.  Double-click the `.dmg` file to mount it.
3.  **Drag the DocRefine Pro app** into your **Applications** folder.

#### ⚠️ Important: First-Run Security Warning
Since this is an open-source tool signed with a community certificate, macOS will block it by default.

1.  Go to your **Applications** folder.
2.  **Right-Click (or Control+Click)** on `DocRefinePro`.
3.  Select **Open** from the menu.
4.  A popup will appear saying "macOS cannot verify the developer". Click **Open**.
    * *You only need to do this once.*

---

## 🛠️ Usage Workflow

### 1. Ingest
* Click **+ New Ingest Job**.
* Select your source folder.
* Choose a mode:
    * **Standard:** Best for general use (Smart text hashing for PDFs).
    * **Lightning:** Fastest (Strict binary hashing only).
    * **Deep Scan:** Slowest but most accurate (Full text scan).

### 2. Process
* Select your job from the dashboard.
* Go to the **Process** tab.
* Check the actions you want (e.g., **Flatten PDFs**, **Resize Images**).
* Click **Run Actions**.

### 3. Distribute
* Go to the **Distribute** tab.
* Choose if you want to output searchable (OCR) copies or flat copies.
* Click **Run Distribution** to generate the final "Clean" folder structure.

---

## 👨‍💻 Building from Source

If you prefer to run the raw Python code or build it yourself:

**Prerequisites:**
* Python 3.10+
* **Windows:** Install Poppler and Tesseract-OCR and add them to your PATH.
* **macOS:** Install via Homebrew: `brew install poppler tesseract tesseract-lang`.

**Setup:**

```bash
# Clone the repo
git clone [https://github.com/jasonweblifestores/DocRefinePro.git](https://github.com/jasonweblifestores/DocRefinePro.git)
cd DocRefinePro

# Install dependencies
pip install -r requirements.txt

# Run the app
python main.py
