import os
import datetime
from pathlib import Path

# ==============================================================================
#   DOCREFINE PRO - PROJECT INVENTORY TOOL
#   Run this to generate 'project_inventory.txt'
# ==============================================================================

SKIP_DIRS = {'.git', '__pycache__', 'venv', 'env', '.idea', '.vscode'}
OUTPUT_FILE = "project_inventory.txt"

def get_size_str(size_bytes):
    if size_bytes < 1024: return f"{size_bytes} B"
    elif size_bytes < 1024**2: return f"{size_bytes/1024:.2f} KB"
    elif size_bytes < 1024**3: return f"{size_bytes/(1024**2):.2f} MB"
    else: return f"{size_bytes/(1024**3):.2f} GB"

def run_inventory():
    root_dir = Path.cwd()
    print(f"Scanning: {root_dir}")
    print("This may take a moment...")

    with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
        f.write(f"PROJECT INVENTORY SCAN\n")
        f.write(f"Root: {root_dir}\n")
        f.write(f"Date: {datetime.datetime.now()}\n")
        f.write("="*80 + "\n\n")

        total_size = 0
        file_count = 0

        for root, dirs, files in os.walk(root_dir):
            # Filter excluded directories in-place
            dirs[:] = [d for d in dirs if d not in SKIP_DIRS]
            
            level = root.replace(str(root_dir), '').count(os.sep)
            indent = ' ' * 4 * (level)
            rel_path = os.path.relpath(root, root_dir)
            
            if rel_path != ".":
                f.write(f"{indent}[{os.path.basename(root)}/]\n")
            
            subindent = ' ' * 4 * (level + 1)
            for fname in files:
                fpath = Path(root) / fname
                try:
                    size = fpath.stat().st_size
                    total_size += size
                    file_count += 1
                    f.write(f"{subindent}{fname}  ({get_size_str(size)})\n")
                except Exception as e:
                    f.write(f"{subindent}{fname}  [ERROR reading file]\n")

        f.write("\n" + "="*80 + "\n")
        f.write(f"TOTAL FILES: {file_count}\n")
        f.write(f"TOTAL SIZE:  {get_size_str(total_size)}\n")

    print(f"✅ Inventory saved to {OUTPUT_FILE}")

if __name__ == "__main__":
    run_inventory()