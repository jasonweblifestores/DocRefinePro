import sys
from docrefine.config import log_app

if __name__ == "__main__":
    # Let a PACKAGED build prove it can actually brand a PDF, stamps and all.
    # A windowed build has no console, so the report goes to a file next to the
    # work directory and the exit code carries the verdict.
    #   DocRefinePro --self-test-rebrand <workdir> [source-dir] [brand-kit]
    if "--self-test-rebrand" in sys.argv:
        i = sys.argv.index("--self-test-rebrand")
        rest = [a for a in sys.argv[i + 1:] if not a.startswith("-")]
        if not rest:
            print("usage: --self-test-rebrand <workdir> [source-dir] [brand-kit]")
            sys.exit(2)
        log_app("Running the packaged rebrand self-test...")
        from docrefine.selftest import run as _selftest
        ok, report = _selftest(rest[0],
                               src_dir=rest[1] if len(rest) > 1 else None,
                               kit_dir=rest[2] if len(rest) > 2 else None)
        print(report)
        log_app(f"Self-test {'passed' if ok else 'FAILED'}.")
        sys.exit(0 if ok else 1)

    if "--dry-run" in sys.argv:
        print("Dry run initialized.")
        log_app("Booting DocRefine Pro (Dry Run Mode)...")
        from docrefine.gui.app_qt import dry_run
        dry_run()
        print("Dry run OK.")
        sys.exit(0)

    try:
        log_app("Booting DocRefine Pro (Qt/PySide6 Edition)...")
        # Import the new Qt App Runner
        from docrefine.gui.app_qt import run
        run()
    except Exception as e:
        print(f"Fatal Boot Error: {e}")