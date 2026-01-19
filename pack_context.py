import os
from pathlib import Path

# --- CONFIGURATION ---
OUTPUT_FILE = "FULL_PROJECT_CONTEXT.txt"

# Folders to completely ignore (Recursion stops here)
IGNORE_DIRS = {
    ".git", ".vscode", ".idea", "__pycache__", "venv", "env", 
    "build", "dist", "Tesseract-OCR", "poppler", "DocRefinePro.app",
    "53752cda3c39550673fc5dafb96c4bed" # The gist folder
}

# Files to ignore (Binaries, heavy assets, or the output file itself)
IGNORE_EXTENSIONS = {
    ".exe", ".dll", ".so", ".dylib", ".bin", ".zip", ".7z", ".rar", 
    ".pdf", ".jpg", ".png", ".ico", ".icns", ".pyc", ".pyd", 
    ".git", ".gitignore", ".DS_Store"
}

IGNORE_FILES = {
    OUTPUT_FILE, 
    "pack_context.py", 
    "poetry.lock", 
    "yarn.lock"
}

def is_text_file(file_path):
    """Peek at the file to check if it's text or binary."""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            f.read(1024)
        return True
    except (UnicodeDecodeError, PermissionError):
        return False

def pack_project():
    root_dir = Path.cwd()
    print(f"📦 Packing project from: {root_dir}")
    
    with open(OUTPUT_FILE, "w", encoding="utf-8") as out:
        out.write(f"================================================================================\n")
        out.write(f"PROJECT CONTEXT DUMP\n")
        out.write(f"Source: DocRefine Pro (v128.5)\n")
        out.write(f"================================================================================\n\n")

        file_count = 0
        skipped_count = 0

        for root, dirs, files in os.walk(root_dir):
            # 1. Prune ignored directories in-place
            dirs[:] = [d for d in dirs if d not in IGNORE_DIRS]
            
            for file in files:
                if file in IGNORE_FILES: continue
                
                file_path = Path(root) / file
                
                # 2. Skip binary extensions
                if file_path.suffix.lower() in IGNORE_EXTENSIONS:
                    continue

                # 3. Skip heavy binary files detection
                if not is_text_file(file_path):
                    skipped_count += 1
                    continue

                # 4. Write Content
                rel_path = file_path.relative_to(root_dir)
                try:
                    content = file_path.read_text(encoding='utf-8', errors='replace')
                    
                    out.write(f"\n==================== START FILE: {rel_path} ====================\n")
                    out.write(content)
                    if not content.endswith("\n"): out.write("\n")
                    out.write(f"==================== END FILE: {rel_path} ====================\n\n")
                    
                    print(f" + Added: {rel_path}")
                    file_count += 1
                except Exception as e:
                    print(f" ! Error reading {rel_path}: {e}")

    print("-" * 50)
    print(f"✅ DONE. Packed {file_count} files into {OUTPUT_FILE}")
    print(f"🚫 Skipped {skipped_count} binary files.")

if __name__ == "__main__":
    pack_project()