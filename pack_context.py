import os
from pathlib import Path

# --- CONFIGURATION ---
OUTPUT_FILE = "FULL_PROJECT_CONTEXT.txt"

# Folders to completely ignore
IGNORE_DIRS = {
    ".git", "__pycache__", "dist", "build", "env", "venv", 
    ".idea", ".vscode", "DocRefinePro.app", "dmg_content"
}

# File extensions to include (Source Code)
INCLUDE_EXT = {
    ".py", ".spec", ".yml", ".yaml", ".json", ".md", ".txt", ".bat", ".ps1"
}

# Specific files to exclude
IGNORE_FILES = {
    "pack_context.py",  # Don't pack the packer itself
    ".DS_Store",
    "desktop.ini"
}

def pack_project():
    root = Path.cwd()
    output_path = root / OUTPUT_FILE
    
    print(f"📦 Packing project from: {root}")
    
    with open(output_path, "w", encoding="utf-8") as out:
        # Write Header
        out.write("="*80 + "\n")
        out.write(f"PROJECT CONTEXT DUMP\n")
        out.write(f"Source: {root.name}\n")
        out.write("="*80 + "\n\n")

        file_count = 0
        
        for dirpath, dirnames, filenames in os.walk(root):
            # 1. Modify dirnames in-place to skip ignored folders
            dirnames[:] = [d for d in dirnames if d not in IGNORE_DIRS]
            
            for f in filenames:
                if f in IGNORE_FILES: continue
                
                path = Path(dirpath) / f
                if path.suffix.lower() not in INCLUDE_EXT: continue
                
                # Calculate relative path for clarity
                rel_path = path.relative_to(root)
                
                try:
                    content = path.read_text(encoding="utf-8", errors="ignore")
                    
                    # Write File Banner
                    out.write(f"\n{'='*20} START FILE: {rel_path} {'='*20}\n")
                    out.write(content)
                    out.write(f"\n{'='*20} END FILE: {rel_path} {'='*20}\n\n")
                    
                    print(f"  + Added: {rel_path}")
                    file_count += 1
                except Exception as e:
                    print(f"  ! Skipped {rel_path}: {e}")

    print(f"\n✅ Done! Packed {file_count} files into '{OUTPUT_FILE}'.")
    print(f"👉 Upload '{OUTPUT_FILE}' to the Gem Knowledge section.")

if __name__ == "__main__":
    pack_project()