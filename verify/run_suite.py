"""Run every verify_*.py in the suite and tally the checks.

Usage: python run_suite.py <repo> <samples> <outdir> [--only name,name]
Each verify script prints "RESULT: n/m passed" and exits 0 on success.
"""
import re
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
# The scripts sit beside this runner, so the suite is self-contained and
# survives a temp-directory cleanup.
SUITE = Path(__file__).resolve().parent

REPO, SAMPLES, OUTDIR = sys.argv[1], sys.argv[2], sys.argv[3]
PY = str(Path(REPO) / ".venv" / "Scripts" / "python.exe")

# verify_phase0 needs a third argument (an output dir); the rest take repo + samples.
SCRIPTS = [
    "verify_sleep.py", "verify_analyze_route.py", "verify_checkpoint.py", "verify_repair.py",
    "verify_engine.py", "verify_boot.py", "verify_regression.py",
    "verify_phase0.py", "verify_phase1.py", "verify_phase2.py", "verify_phase3.py",
    "verify_collision.py", "verify_naming.py", "verify_keepnames.py",
    "verify_dedup.py", "verify_attrib.py", "verify_pagebox.py", "verify_nav.py",
    "verify_v139.py", "verify_v140.py", "verify_v141.py", "verify_v147.py",
    "verify_vision.py",
]

only = None
for a in sys.argv[4:]:
    if a.startswith("--only"):
        only = set(a.split("=", 1)[1].split(","))
if only:
    SCRIPTS = [s for s in SCRIPTS if s in only or s[:-3] in only]

rows, total, passed, failed_scripts = [], 0, 0, []
for name in SCRIPTS:
    path = SUITE / name
    if not path.exists():
        rows.append((name, "MISSING", 0, 0)); failed_scripts.append(name); continue
    args = [PY, str(path), REPO, SAMPLES]
    if name == "verify_phase0.py":
        args.append(OUTDIR)
    print(f"\n{'=' * 70}\n>>> {name}\n{'=' * 70}", flush=True)
    p = subprocess.run(args, capture_output=True, text=True, timeout=3600)
    out = p.stdout + p.stderr
    # echo only failures and the result line, to keep this readable
    for line in out.splitlines():
        if "[FAIL]" in line or "RESULT:" in line or "Traceback" in line:
            print(line, flush=True)
    m = re.search(r"RESULT:\s*(\d+)/(\d+)", out)
    if m:
        n, d = int(m.group(1)), int(m.group(2))
        total += d; passed += n
        ok = (n == d and p.returncode == 0)
        rows.append((name, "PASS" if ok else "FAIL", n, d))
        if not ok:
            failed_scripts.append(name)
            print(out[-3000:], flush=True)
    else:
        rows.append((name, f"NO RESULT LINE (rc={p.returncode})", 0, 0))
        failed_scripts.append(name)
        print(out[-3000:], flush=True)

print("\n" + "=" * 70)
print("SUITE SUMMARY")
print("=" * 70)
for name, status, n, d in rows:
    print(f"  {status:<28} {n:>4}/{d:<4}  {name}")
print("-" * 70)
print(f"  TOTAL {passed}/{total} checks across {len(rows)} scripts")
if failed_scripts:
    print(f"  FAILED SCRIPTS: {', '.join(failed_scripts)}")
print("=" * 70)
sys.exit(1 if failed_scripts else 0)
