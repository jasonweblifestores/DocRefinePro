"""Bring frank_lloyd_wright_collection.pdf under the 50 MB limit, per Kunchana.

Brief: recompress the images as needed, keep the pages intact and legible; if it
cannot get under without becoming unreadable, flag it rather than ship it.

So this touches images only — no rasterising, no page changes, and the text
layer is left alone. Every image is resampled to a floor of TARGET_DPI.

The DPI figure is deliberately conservative: it is computed as though the image
spanned the full page width, which is the *lowest* resolution it could possibly
be displayed at. An image drawn at half the page width has twice the true DPI,
so resampling to a 200 DPI estimate never puts the real figure below 200.

Which batch it acts on comes from batch_paths (DRP_BATCH, default mbw). Kunchana
re-approved this treatment for MBW on 86eyaantb comment 90180247980961: "FLW
collection (58.2 MB) -> 200 DPI floor, same as Batch 4 - go ahead."

With --install the result replaces the staged master, keeping the original beside
it as .original, so Apply produces a compliant file directly rather than needing
the post-hoc swap Batch 4 used.

Usage: [set DRP_BATCH=mbw] python downsample_flw.py [dpi] [quality] [--install]
"""
import io
import logging
import shutil
import sys
from pathlib import Path

import batch_paths                      # also puts the repo on sys.path
logging.getLogger("pypdf").setLevel(logging.CRITICAL)

from PIL import Image
from pypdf import PdfReader, PdfWriter

NAME = "frank_lloyd_wright_collection.pdf"
SRC = batch_paths.SRC / NAME
SP = Path(__file__).resolve().parent
OUT = SP / "flw_downsampled.pdf"

INSTALL = "--install" in sys.argv
_args = [a for a in sys.argv[1:] if not a.startswith("--")]
TARGET_DPI = int(_args[0]) if _args else 200
QUALITY = int(_args[1]) if len(_args) > 1 else 80


def main():
    rd = PdfReader(str(SRC))
    before = SRC.stat().st_size
    print(f"source: {before / 1e6:.1f} MB, {len(rd.pages)} pages")
    print(f"target: images resampled to a floor of {TARGET_DPI} dpi, JPEG q{QUALITY}\n")

    # Images can only be replaced through a writer that owns them, so clone
    # first and edit the clone's pages.
    writer = PdfWriter(clone_from=str(SRC))

    changed = skipped = failed = 0
    for i, page in enumerate(writer.pages):
        pw = float(abs(page.mediabox.width)) / 72.0      # page width in inches
        if pw <= 0:
            continue
        try:
            imgs = list(page.images)
        except Exception as e:
            print(f"  p{i + 1}: cannot enumerate images ({e})")
            continue
        for img in imgs:
            try:
                pil = img.image
                if pil is None:
                    skipped += 1
                    continue
                w, h = pil.size
                dpi = w / pw
                if dpi <= TARGET_DPI:
                    skipped += 1
                    continue
                scale = TARGET_DPI / dpi
                nw, nh = max(1, int(w * scale)), max(1, int(h * scale))
                small = pil.convert("RGB").resize((nw, nh), Image.LANCZOS)
                buf = io.BytesIO()
                small.save(buf, format="JPEG", quality=QUALITY, optimize=True)
                buf.seek(0)
                img.replace(Image.open(buf))
                changed += 1
            except Exception as e:
                failed += 1
                if failed <= 5:
                    print(f"  p{i + 1} {img.name}: left alone ({type(e).__name__}: {e})")

    writer.compress_identical_objects()
    with open(OUT, "wb") as fh:
        writer.write(fh)

    after = OUT.stat().st_size
    print(f"\nresampled {changed} images, left {skipped} already at or under target, "
          f"{failed} could not be touched")
    print(f"result: {before / 1e6:.1f} MB -> {after / 1e6:.1f} MB "
          f"({100 * after / before:.0f}% of original)")

    # --- did we keep what we promised to keep? ---------------------------
    nd = PdfReader(str(OUT))
    src_txt = "".join((p.extract_text() or "") for p in rd.pages)
    new_txt = "".join((p.extract_text() or "") for p in nd.pages)
    print(f"pages: {len(rd.pages)} -> {len(nd.pages)}")
    print(f"text chars: {len(src_txt)} -> {len(new_txt)}")

    # lowest resolution now present, on the same conservative basis
    worst = None
    for i, page in enumerate(nd.pages):
        pwi = float(abs(page.mediabox.width)) / 72.0
        if pwi <= 0:
            continue
        try:
            for img in page.images:
                if img.image is None:
                    continue
                d = img.image.size[0] / pwi
                if worst is None or d < worst[0]:
                    worst = (d, i + 1, img.name, img.image.size)
        except Exception:
            pass
    if worst:
        print(f"lowest image resolution now: ~{worst[0]:.0f} dpi "
              f"(p{worst[1]} {worst[2]} {worst[3][0]}x{worst[3][1]})")

    ok = (after < 48e6 and len(nd.pages) == len(rd.pages)
          and len(new_txt) >= len(src_txt) * 0.99)
    print(f"\nunder 48 MB (leaves room for branding): {after < 48e6}")
    print(f"pages intact: {len(nd.pages) == len(rd.pages)}")
    print(f"text intact:  {len(new_txt) >= len(src_txt) * 0.99}")
    print("VERDICT:", "usable" if ok else "NOT usable — do not ship, flag it")

    if not ok:
        print("\nnot installing an unusable result")
        return 1
    if not INSTALL:
        print(f"\nnot installed. Re-run with --install to replace\n  {SRC}")
        return 0

    keep = SRC.with_suffix(SRC.suffix + ".original")
    if not keep.exists():
        shutil.copy2(SRC, keep)
        print(f"\noriginal kept as {keep.name} ({keep.stat().st_size / 1e6:.1f} MB)")
    shutil.copy2(OUT, SRC)
    # The staged file is what Apply reads, so confirm the swap on disk rather
    # than trusting the copy.
    now = PdfReader(str(SRC))
    print(f"installed: {SRC.name} is now {SRC.stat().st_size / 1e6:.1f} MB, "
          f"{len(now.pages)} pages")
    assert SRC.stat().st_size == OUT.stat().st_size, "install did not take"
    assert len(now.pages) == len(rd.pages), "page count changed on install"
    # .original must not ship: it is not a .pdf, so complete_set=False excludes it,
    # but say so out loud because a stray file in the source tree has bitten us.
    print(f"note: {keep.name} sits beside it and is NOT a .pdf, so Apply ignores it")
    return 0


if __name__ == "__main__":
    sys.exit(main())
